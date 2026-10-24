"""
tally_sync_app.py  -  Tally Sync window for accounts (the same program with a friendlier face).

What it does, in plain words:
  "Tally mein bhejo"            runs tally_itemwise.py exactly like Tally_Sync.bat, then shows the day's count
  table of waiting bills        bills the program stopped on, with the reason; accounts enter them by hand
  "Tick wale ... chadh gaye"    the ticked bills were entered BY HAND -> tally_itemwise.py --mark-manual
  "Hisaab dobara dekho"         reads the count again (sends nothing)

The window itself never writes to Tally or to the sent list: every change still goes through
tally_itemwise.py, which keeps all its checks. Tally_Sync.bat stays as the backup way.

Start: double-click Tally_Sync_App.bat
    pythonw tally_sync_app.py                         # settings.ini
    pythonw tally_sync_app.py --test-date 01-10-2026  # TEST COMPANY demo (Educational Tally)
    pythonw tally_sync_app.py --settings settings_gst.ini
"""
import argparse
import csv
import os
import queue
import re
import sqlite3
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

import tally_itemwise as ti

HERE = Path(__file__).resolve().parent
NO_WINDOW = 0x08000000 if os.name == "nt" else 0          # no black window flashing up on Windows

# dark theme
BG, SURFACE, SURFACE2, LINE = "#18181B", "#232327", "#2E2E33", "#3A3A40"
TEXT, MUTED = "#EDEDEF", "#9A9AA3"
GREEN_BG, GREEN_FG = "#16301F", "#7EDC9A"
AMBER_BG, AMBER_FG = "#33270F", "#F5C451"
GREY_BG, GREY_FG = SURFACE2, TEXT
RED_FG, BLUE, BLUE_HOVER = "#FF8A80", "#378ADD", "#4C9BE8"
FONT = "Segoe UI"
TICK, NO_TICK = "\u2611", "\u2610"


def python_exe():
    """python.exe next to pythonw.exe, so the program's output can be read."""
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe" and exe.with_name("python.exe").exists():
        return str(exe.with_name("python.exe"))
    return str(exe)


def dark_title_bar(root):
    """Dark title bar on Windows 10/11 (matches the window). Does nothing anywhere else or if not possible."""
    if os.name != "nt":
        return
    try:
        import ctypes
        root.update()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        on = ctypes.c_int(1)
        for attr in (20, 19):            # 20 on newer Windows, 19 on older Windows 10
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(on), ctypes.sizeof(on)) == 0:
                break
    except Exception:
        pass


class App:
    def __init__(self, root, args):
        self.root, self.args = root, args
        self.q = queue.Queue()
        self.running = False
        self.ticked = set()
        self.base_cmd = [python_exe(), str(HERE / "tally_itemwise.py")]
        if args.settings:
            self.base_cmd += ["--settings", args.settings]

        root.title("Tally Sync")
        root.geometry("900x680")
        root.minsize(760, 560)
        root.configure(bg=BG)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE, foreground=TEXT,
                        rowheight=30, font=(FONT, 10), borderwidth=0)
        style.configure("Treeview.Heading", background=SURFACE2, foreground=MUTED, relief="flat",
                        font=(FONT, 10), borderwidth=0, padding=(6, 6))
        style.map("Treeview.Heading", background=[("active", SURFACE2)])
        style.map("Treeview", background=[("selected", SURFACE)], foreground=[("selected", TEXT)])
        style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])     # no light border
        style.configure("Dark.Vertical.TScrollbar", background=SURFACE2, troughcolor=SURFACE, bordercolor=SURFACE,
                        lightcolor=SURFACE2, darkcolor=SURFACE2, arrowcolor=MUTED, gripcount=0, relief="flat")
        style.map("Dark.Vertical.TScrollbar", background=[("active", LINE)])
        style.configure("Blue.Horizontal.TProgressbar", background=BLUE, troughcolor=SURFACE2,
                        bordercolor=SURFACE2, lightcolor=BLUE, darkcolor=BLUE, thickness=8)

        outer = tk.Frame(root, bg=BG, padx=22, pady=18)
        outer.pack(fill="both", expand=True)

        # --- heading
        head = tk.Frame(outer, bg=BG)
        head.pack(fill="x")
        tk.Label(head, text="Tally Sync", font=(FONT, 20, "bold"), bg=BG, fg=TEXT).pack(side="left")
        if args.test_date:
            tk.Label(head, text=f"  TEST  {args.test_date}  ", font=(FONT, 9, "bold"), bg=AMBER_BG,
                     fg=AMBER_FG).pack(side="left", padx=12, pady=(6, 0))
        tk.Label(outer, text=f"Company: {ti.COMPANY}", font=(FONT, 10), bg=BG, fg=MUTED).pack(anchor="w", pady=(0, 14))

        # --- buttons
        row = tk.Frame(outer, bg=BG)
        row.pack(fill="x")
        self.btn_send = self._button(row, "Tally mein bhejo", self.send, primary=True)
        self.btn_send.pack(side="left")
        self.btn_refresh = self._button(row, "Hisaab dobara dekho", self.refresh)
        self.btn_refresh.pack(side="left", padx=10)

        # --- status line + progress
        self.status = tk.Label(outer, text="", font=(FONT, 10), anchor="w", justify="left", wraplength=840,
                               bg=BG, fg=MUTED)
        self.status.pack(fill="x", pady=(14, 6))
        self.progress = ttk.Progressbar(outer, mode="determinate", style="Blue.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 14))

        # --- three count cards
        cards = tk.Frame(outer, bg=BG)
        cards.pack(fill="x")
        self.n_done = self._card(cards, "Tally mein chadh gaye", GREEN_BG, GREEN_FG)
        self.n_wait = self._card(cards, "Haath se chadhane hain", AMBER_BG, AMBER_FG)
        self.n_total = self._card(cards, "Aaj ki ERP file mein bills", GREY_BG, GREY_FG)
        self.split = tk.Label(outer, text="", font=(FONT, 9), fg=MUTED, bg=BG, anchor="w")
        self.split.pack(fill="x", pady=(6, 8))

        # --- changed-bills warning (shown only when needed)
        self.changed_lbl = tk.Label(outer, text="", font=(FONT, 10, "bold"), fg=RED_FG, bg=BG, anchor="w",
                                    justify="left", wraplength=840)
        self.changed_lbl.pack(fill="x")

        # --- table of waiting bills
        tk.Label(outer, text="Haath se chadhane wale bills", font=(FONT, 12, "bold"), bg=BG, fg=TEXT,
                 anchor="w").pack(fill="x", pady=(8, 0))
        tk.Label(outer, text="Tally mein wahi ERP bill number aur date se chadhao, phir bill pe click karke tick lagao.",
                 font=(FONT, 9), bg=BG, fg=MUTED, anchor="w").pack(fill="x", pady=(0, 6))
        # bottom controls are packed first (side="bottom") so a small window never hides them
        self.detail_btn = tk.Label(outer, text="Detail dekho", fg=BLUE, bg=BG, cursor="hand2", font=(FONT, 9, "underline"))
        self.detail_btn.pack(side="bottom", anchor="w", pady=(8, 0))
        self.detail_btn.bind("<Button-1>", self.toggle_detail)
        self.detail = tk.Text(outer, height=8, font=("Consolas", 9), wrap="none", bg=SURFACE, fg=MUTED,
                              insertbackground=TEXT, relief="flat", highlightthickness=0)
        self.detail_shown = False
        bottom = tk.Frame(outer, bg=BG)
        bottom.pack(side="bottom", fill="x", pady=(12, 0))
        self._button(bottom, "Sab tick", lambda: self.tick_all(True), small=True).pack(side="left")
        self._button(bottom, "Tick hatao", lambda: self.tick_all(False), small=True).pack(side="left", padx=8)
        self.btn_mark = self._button(bottom, "Tick wale bills haath se chadh gaye", self.mark, primary=True)
        self.btn_mark.pack(side="right")


        tframe = tk.Frame(outer, bg=SURFACE, highlightbackground=LINE, highlightthickness=1)
        tframe.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tframe, columns=("tick", "bill", "party", "why"), show="headings",
                                 selectmode="none")
        for col, text, width, stretch in (("tick", "", 40, False), ("bill", "Bill no.", 90, False),
                                          ("party", "Party", 260, True), ("why", "Kyun ruka", 360, True)):
            self.tree.heading(col, text=text, anchor="center" if col == "tick" else "w")
            self.tree.column(col, width=width, stretch=stretch, anchor="center" if col == "tick" else "w")
        sb = ttk.Scrollbar(tframe, orient="vertical", command=self.tree.yview, style="Dark.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("ticked", background=GREEN_BG, foreground=GREEN_FG)
        self.tree.bind("<Button-1>", self.on_click)

        self.refresh()

    # ---------- small parts
    def _card(self, parent, label, bg, fg):
        f = tk.Frame(parent, bg=bg, padx=16, pady=12)
        f.pack(side="left", fill="x", expand=True, padx=(0, 12))
        tk.Label(f, text=label, bg=bg, fg=fg, font=(FONT, 10)).pack(anchor="w")
        num = tk.Label(f, text="-", bg=bg, fg=fg, font=(FONT, 28, "bold"))
        num.pack(anchor="w")
        return num

    def _button(self, parent, text, command, primary=False, small=False):
        bg, fg, hover = (BLUE, "#FFFFFF", BLUE_HOVER) if primary else (SURFACE2, TEXT, LINE)
        b = tk.Button(parent, text=text, command=command, bg=bg, fg=fg, activebackground=hover,
                      activeforeground=fg, disabledforeground=MUTED, relief="flat", bd=0, cursor="hand2",
                      font=(FONT, 10 if small else 11, "bold" if primary else "normal"),
                      padx=12 if small else 18, pady=5 if small else 9, highlightthickness=0)
        b.bind("<Enter>", lambda e: b.config(bg=hover) if str(b["state"]) != "disabled" else None)
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def say(self, text, colour=GREY_FG):
        self.status.config(text=text, fg=colour)

    def log(self, text):
        self.detail.insert("end", text)
        self.detail.see("end")

    def toggle_detail(self, _e=None):
        if self.detail_shown:
            self.detail.pack_forget()
            self.detail_btn.config(text="Detail dekho")
        else:
            self.detail.pack(side="bottom", fill="x", before=self.detail_btn, pady=(4, 0))
            self.detail_btn.config(text="Detail chhupao")
        self.detail_shown = not self.detail_shown

    def busy(self, on):
        self.running = on
        for b in (self.btn_send, self.btn_refresh, self.btn_mark):
            b.config(state="disabled" if on else "normal")
        self.root.config(cursor="watch" if on else "")

    # ---------- reading the day's count (sends nothing)
    def refresh(self):
        header_file = Path(ti.DEFAULT_HEADER)
        if not header_file.exists():
            self.say(f"ERP file nahi mili: {header_file}\nPehle ERP se export karke isi naam se save karo.", RED_FG)
            return
        try:
            headers = {r["Inv No."].strip(): r for r in ti.read_csv(header_file)}
        except Exception as e:
            self.say(f"ERP file padh nahi paye ({header_file.name}): {e}", RED_FG)
            return
        if Path(ti.DB_FILE).exists():
            db = sqlite3.connect(ti.DB_FILE)
        else:
            db = sqlite3.connect(":memory:")
        try:
            db.execute("SELECT invoice_no, sent_at FROM sent LIMIT 1")
        except sqlite3.Error:
            db = sqlite3.connect(":memory:")
            db.execute("CREATE TABLE sent (invoice_no TEXT PRIMARY KEY, sent_at TEXT)")
        auto, manual, waiting, before = ti.day_status(headers, db)
        db.close()

        reasons = {}
        if ti.FAILED_FILE.exists():
            with open(ti.FAILED_FILE, newline="", encoding="utf-8") as f:
                reasons = {r["inv_no"]: ti._why_hindi(r["reason"]) for r in csv.DictReader(f)
                           if not str(r["reason"]).startswith(ti.NO_TALLY)}
        # stopped by the program (has a reason) -> enter by hand; no reason yet -> simply not sent yet
        by_hand = [inv for inv in waiting if inv in reasons]
        not_sent = [inv for inv in waiting if inv not in reasons]

        self.n_done.config(text=str(len(auto) + len(manual)))
        self.n_wait.config(text=str(len(by_hand)))
        self.n_total.config(text=str(len(auto) + len(manual) + len(waiting)))
        extra = f"   |   {before} bills start-from se pehle ke, gine nahi" if before else ""
        self.split.config(text=f"Program se: {len(auto)}     Haath se: {len(manual)}{extra}")

        self.tree.delete(*self.tree.get_children())
        self.ticked = set()
        for inv in by_hand:
            party = str(headers[inv].get("Customer Name", "")).strip()
            self.tree.insert("", "end", iid=inv, values=(NO_TICK, inv, party, reasons[inv]))

        changed = []
        if ti.CHANGED_FILE.exists():
            with open(ti.CHANGED_FILE, newline="", encoding="utf-8") as f:
                changed = [r["inv_no"] for r in csv.DictReader(f)]
        self.changed_lbl.config(text=("DHYAN: yeh bills Tally mein jaane ke BAAD ERP mein badle gaye: "
                                      + ", ".join(changed) + "\nTally mein haath se theek karo, phir "
                                      "Accept_Change.bat chalao.") if changed else "")
        if not self.running:
            if not_sent:
                self.say(f"{len(not_sent)} bills abhi Tally mein nahi gaye. 'Tally mein bhejo' dabao.", BLUE)
            elif by_hand:
                self.say(f"{len(by_hand)} bills haath se chadhane hain. Chadhane ke baad table mein tick lagao "
                         "aur neeche wala button dabao.", AMBER_FG)
            else:
                self.say("Sab bills Tally mein hain. Kuch baaki nahi.", GREEN_FG)

    # ---------- ticking
    def on_click(self, event):
        inv = self.tree.identify_row(event.y)
        if not inv:
            return
        self.set_tick(inv, inv not in self.ticked)

    def set_tick(self, inv, on):
        vals = list(self.tree.item(inv, "values"))
        vals[0] = TICK if on else NO_TICK
        self.tree.item(inv, values=vals, tags=("ticked",) if on else ())
        (self.ticked.add if on else self.ticked.discard)(inv)

    def tick_all(self, on):
        for inv in self.tree.get_children():
            self.set_tick(inv, on)

    # ---------- running tally_itemwise.py
    def run(self, cmd, on_line, on_done):
        def worker():
            env = dict(os.environ, PYTHONIOENCODING="utf-8")
            try:
                p = subprocess.Popen(cmd, cwd=HERE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                     text=True, encoding="utf-8", errors="replace", env=env,
                                     creationflags=NO_WINDOW)
                for line in p.stdout:
                    self.q.put(("line", line))
                p.wait()
                self.q.put(("done", p.returncode))
            except Exception as e:
                self.q.put(("line", f"Program chal nahi paya: {e}\n"))
                self.q.put(("done", -1))
        self.on_line, self.on_done = on_line, on_done
        self.busy(True)
        threading.Thread(target=worker, daemon=True).start()
        self.root.after(100, self.poll)

    def poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "line":
                    self.log(val)
                    self.on_line(val)
                else:
                    self.busy(False)
                    self.on_done(val)
                    return
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def send(self):
        if not messagebox.askyesno("Tally Sync", "Tally khula hai aur sahi company load hai?\n\n"
                                                 f"Company: {ti.COMPANY}"):
            return
        cmd = self.base_cmd + (["--test-date", self.args.test_date] if self.args.test_date else [])
        self.result = {"seen": 0, "created": 0, "stopped": False, "stop": ""}
        try:
            total = len(ti.read_csv(Path(ti.DEFAULT_HEADER)))
        except Exception:
            total = 0
        self.progress.config(maximum=max(total, 1), value=0)
        self.log(f"\n--- {datetime.now():%d-%m-%Y %H:%M} Tally mein bhej rahe hain ---\n")
        self.say("Bills dekh rahe hain aur Tally mein bhej rahe hain... window band mat karna.", BLUE)

        def on_line(line):
            if re.match(r"(OK|SKIP|FAILED|CHANGED|ACCEPTED)\b", line):
                self.result["seen"] += 1
                self.progress.config(value=self.result["seen"])
            if line.startswith("STOPPED at"):
                self.result["stopped"] = True
            if line.startswith("STOP:"):
                self.result["stop"] = line[5:].strip()
            m = re.match(r"Done: (\d+) created", line)
            if m:
                self.result["created"] = int(m.group(1))

        def on_done(code):
            r = self.result
            if r["stopped"]:
                self.say("RUK GAYA: Tally se baat nahi ho payi (Tally band hai ya company load nahi). "
                         "Kuch kharab nahi hua. Tally kholo, company load karo, aur dobara 'Tally mein bhejo' dabao.",
                         RED_FG)
                return                      # the table stays as it was: unsent bills are NOT "enter by hand"
            if r["stop"] or code not in (0, None):
                self.say("Program ruk gaya: " + (r["stop"] or "'Detail dekho' mein dekho."), RED_FG)
                messagebox.showerror("Tally Sync", r["stop"] or "Program ruk gaya. 'Detail dekho' mein dekho.")
                return
            self.progress.config(value=self.progress["maximum"])
            self.refresh()
            self.say(f"Ho gaya: is baar {r['created']} naye bills Tally mein gaye. " + self.status.cget("text"),
                     GREEN_FG if not self.tree.get_children() else AMBER_FG)

        self.run(cmd, on_line, on_done)

    def mark(self):
        bills = sorted(self.ticked)
        if not bills:
            messagebox.showinfo("Tally Sync", "Pehle table mein un bills pe tick lagao jo aapne Tally mein "
                                              "haath se chadha diye hain.")
            return
        if not messagebox.askyesno("Tally Sync", f"Yeh {len(bills)} bills Tally mein HAATH SE chadh chuke hain?\n\n"
                                                 + ", ".join(bills) + "\n\nProgram inhe kabhi nahi bhejega."):
            return
        self.log(f"\n--- {datetime.now():%d-%m-%Y %H:%M} haath se chadhe bills: {' '.join(bills)} ---\n")
        missing = []

        def on_line(line):
            if "NAHI MILA" in line:
                missing.append(line.strip())

        def on_done(code):
            self.refresh()
            if missing or code not in (0, None):
                messagebox.showwarning("Tally Sync", "Kuch bills yaad nahi hue. 'Detail dekho' mein dekho.\n\n"
                                       + "\n".join(missing))

        self.run(self.base_cmd + ["--mark-manual"] + bills, on_line, on_done)

    def on_close(self):
        if self.running and not messagebox.askyesno("Tally Sync", "Program abhi chal raha hai. Phir bhi band karein?"):
            return
        self.root.destroy()


def main():
    ap = argparse.ArgumentParser(description="Tally Sync window for accounts")
    ap.add_argument("--settings", help="another settings file (default: settings.ini)")
    ap.add_argument("--test-date", help="DD-MM-YYYY, TEST COMPANY only (Educational Tally)")
    a = ap.parse_args()
    root = tk.Tk()
    if not ti.acquire_lock("tally_sync_app.lock"):
        messagebox.showinfo("Tally Sync", "Tally Sync ki window pehle se khuli hai. Neeche taskbar mein dekho.")
        root.destroy()
        return
    try:
        if a.settings:
            path = Path(a.settings) if Path(a.settings).exists() else HERE / a.settings
            if not path.exists():
                raise SystemExit(f"Settings file {a.settings} nahi mili.")
            ti.apply_settings(path)
        App(root, a)
        dark_title_bar(root)
    except BaseException as e:            # pythonw has no screen: show the problem instead of vanishing
        messagebox.showerror("Tally Sync", f"Window khul nahi payi:\n\n{e}")
        root.destroy()
        return
    root.mainloop()


if __name__ == "__main__":
    main()
