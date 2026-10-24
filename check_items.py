"""
check_items.py  -  does every Tally item in the glass mapping really exist in Tally?

Why: a wrong stock item name means wrong stock and wrong HSN in the GST return, and it passes
every total check. This script catches it BEFORE any bill is sent.

What it does, in plain words (it only READS from Tally, it never creates or changes anything):
1. Reads the glass mapping file named in the settings (Glass Name (ERP) -> Tally stock item).
2. Asks Tally for the list of all stock items, with each item's unit.
3. For every mapping line it decides:
     OK         the item exists in Tally with exactly this name, and its unit is the settings unit
     SPELLING   Tally has it with different capitals or spaces -> use the Tally spelling
     SIMILAR    no exact match, but a Tally item looks close -> a person must check
     MISSING    nothing like it in Tally -> fix the mapping, or accounts create the item
     UNIT       the item exists but its unit is not the settings unit (for example SQMT)
4. Also reads the ERP item-wise file and lists glass names that are NOT in the mapping (UNMAPPED),
   so new glass names in a new ERP day show up here first.
5. Writes item_check.csv next to the mapping file.

Run (Tally open with the company loaded):
    python check_items.py                               # uses settings.ini
    python check_items.py --settings settings_gst.ini   # GST TEST COMPANY
    python check_items.py --items "C:\\erp_data\\07_Oct_26 (item wise).xls"   # another ERP day
"""
import argparse
import csv
import difflib
import html
import re
from datetime import datetime
from pathlib import Path

import tally_itemwise as ti     # whole module, so the values read from --settings are seen here


def stock_list_xml():
    return f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>AllStockItems</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES>
      <SVCURRENTCOMPANY>{ti.escape(ti.COMPANY)}</SVCURRENTCOMPANY>
      <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    </STATICVARIABLES>
    <TDL><TDLMESSAGE>
      <COLLECTION NAME="AllStockItems" ISMODIFY="No">
        <TYPE>StockItem</TYPE>
        <NATIVEMETHOD>Name</NATIVEMETHOD>
        <NATIVEMETHOD>BaseUnits</NATIVEMETHOD>
      </COLLECTION>
    </TDLMESSAGE></TDL>
  </DESC></BODY>
</ENVELOPE>"""


def tally_stock_items():
    """Return {exact Tally item name: its unit}."""
    ok, reply = ti.send(stock_list_xml())
    if not ok:
        raise SystemExit(f"STOP: {reply}")
    items = {}
    for name, body in re.findall(r'<STOCKITEM\s+NAME="([^"]*)"[^>]*>(.*?)</STOCKITEM>', reply, re.S):
        unit = re.search(r"<BASEUNITS[^>]*>([^<]*)</BASEUNITS>", body)
        name = html.unescape(name).strip()
        if name:
            items[name] = html.unescape(unit.group(1)).strip() if unit else ""
    return items


def load_mapping_rows():
    """Every line of the mapping file, in file order: (ERP glass name, Tally item name)."""
    rows = []
    if not ti.MAPPING_FILE.exists():
        raise SystemExit(f"STOP: mapping file {ti.MAPPING_FILE.name} not found.")
    with open(ti.MAPPING_FILE, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            erp = (r.get("Glass Name (ERP)") or "").strip()
            tally = (r.get("Tally stock item name (fill in)") or "").strip()
            if erp:
                rows.append((erp, tally))
    return rows


def erp_glass_names(items_file):
    """Glass names used in the ERP item-wise file (empty set if the file is not there)."""
    p = Path(items_file)
    if not p.exists():
        print(f"(Note: ERP item-wise file {p.name} not found, so UNMAPPED glass names are not checked.)")
        return set()
    return {ti.norm(r["Glass Name"]) for r in ti.read_csv(p) if str(r.get("Glass Name", "")).strip()}


def main():
    ap = argparse.ArgumentParser(description="Check that every mapped Tally stock item exists (read only)")
    ap.add_argument("--settings", help="another settings file (default: settings.ini)")
    ap.add_argument("--items", help="ERP item-wise file to look for unmapped glass names (default: from settings)")
    a = ap.parse_args()
    if a.settings:
        path = Path(a.settings) if Path(a.settings).exists() else ti.HERE / a.settings
        if not path.exists():
            raise SystemExit(f"STOP: settings file {a.settings} not found.")
        ti.apply_settings(path)

    mapping_rows = load_mapping_rows()
    tally_items = tally_stock_items()
    if not tally_items:
        print(f"Tally returned no stock items in '{ti.COMPANY}'. "
              "Check that Tally is open with this company loaded.")
        return
    used = erp_glass_names(a.items or ti.DEFAULT_ITEMS)

    by_norm = {ti.norm(t): t for t in tally_items}
    norm_list = list(by_norm)
    unit_wanted = ti.norm(ti.UNIT)

    rows = []
    counts = {k: 0 for k in ("OK", "SPELLING", "SIMILAR", "MISSING", "UNIT", "UNMAPPED")}
    for erp, tally in mapping_rows:
        unit = ""
        if not tally:
            status, suggestion = "MISSING", ""           # Tally item column left empty
        elif tally in tally_items:
            unit = tally_items[tally]
            status, suggestion = ("OK" if ti.norm(unit) == unit_wanted else "UNIT"), tally
        elif ti.norm(tally) in by_norm:
            suggestion = by_norm[ti.norm(tally)]
            unit = tally_items[suggestion]
            status = "SPELLING"
        else:
            close = difflib.get_close_matches(ti.norm(tally), norm_list, n=1, cutoff=0.8)
            if close:
                suggestion = by_norm[close[0]]
                unit = tally_items[suggestion]
                status = "SIMILAR"
            else:
                status, suggestion = "MISSING", ""
        counts[status] += 1
        rows.append({"Glass Name (ERP)": erp, "Tally item in mapping": tally, "Status": status,
                     "Tally name to use": suggestion, "Unit in Tally": unit,
                     "Used in ERP file": "yes" if ti.norm(erp) in used else ""})

    mapped = {ti.norm(erp) for erp, tally in mapping_rows if tally}
    for g in sorted(used - mapped):
        counts["UNMAPPED"] += 1
        rows.append({"Glass Name (ERP)": g, "Tally item in mapping": "", "Status": "UNMAPPED",
                     "Tally name to use": "", "Unit in Tally": "", "Used in ERP file": "yes"})

    out_file = ti.MAPPING_FILE.with_name("item_check.csv")
    with open(out_file, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)

    print(f"Company: {ti.COMPANY}  |  {len(tally_items)} stock items in Tally  |  "
          f"{len(mapping_rows)} lines in {ti.MAPPING_FILE.name}  |  unit expected: {ti.UNIT}")
    for st in counts:
        print(f"  {st:9} {counts[st]}")
    for r in rows:
        if r["Status"] != "OK":
            extra = f"  (unit in Tally: {r['Unit in Tally'] or 'none'})" if r["Status"] == "UNIT" else ""
            print(f"  {r['Status']:9} {r['Glass Name (ERP)']}  ->  "
                  f"{r['Tally name to use'] or r['Tally item in mapping'] or '(no item)'}{extra}")
    print(f"\nSaved {out_file.name} ({datetime.now():%d-%m-%Y %H:%M}).")
    if counts["SPELLING"] or counts["SIMILAR"]:
        print("SPELLING/SIMILAR: check the name, then put the Tally spelling in the mapping file.")
    if counts["MISSING"]:
        print("MISSING: fix the name in the mapping file, or ask accounts to create the item. Never guess.")
    if counts["UNIT"]:
        print(f"UNIT: the item's unit in Tally is not {ti.UNIT}. Ask accounts before using it.")
    if counts["UNMAPPED"]:
        print("UNMAPPED: new glass names in the ERP file. Add them to the mapping (confirm with accounts).")


if __name__ == "__main__":
    main()
