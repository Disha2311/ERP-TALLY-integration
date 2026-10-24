"""
Item-wise version: join the ERP header export and the ERP item-wise export, then post each
bill to TallyPrime as an item invoice (stock items in SQM, taxes, Hole/Cutout, Freight, Round Off).

USE ONLY WITH THE ANONYMISED CSV FILES AND THE DUMMY "TEST COMPANY".

What it does for every bill
  1. Joins the two files on Inv No.
  2. Maps each ERP Glass Name to a Tally stock item (glass_name_mapping.csv)
  3. SAFETY CHECKS (a bill that fails any check is NOT posted, it goes to failed_itemwise.csv):
       - sum of item amounts + freight  must equal  Ass. Value
       - Ass. Value + tax + round off   must equal  Net Amt (round off must be under Rs 1)
       - IGST and CGST/SGST must not both be present
       - every glass name must have a Tally stock item
  4. Builds the voucher: voucher number = ERP invoice number, party, one line per stock item (SQM),
     Hole / Cutout / Freight / Output tax / Round Off as separate ledger lines
  5. Skips bills already sent (tally_sent_itemwise.db), logs failures, keeps going

Run (PowerShell, from the folder that holds this file and the CSVs):
    python tally_itemwise.py --dry-run                 # checks everything, sends NOTHING to Tally
    python tally_itemwise.py --only 6635 --create-parties --fallback-item "5 MM CLEAR TOUGHENED"
    python tally_itemwise.py --create-parties --fallback-item "5 MM CLEAR TOUGHENED"

Options
    --dry-run          build and check all bills, do not contact Tally
    --only 6635        process just this invoice number
    --show 6635        print the XML of this invoice (works with --dry-run)
    --create-parties   create missing party ledgers (Sundry Debtors). DUMMY COMPANY ONLY.
    --fallback-item X  glass names with no mapping use stock item X. TESTING ONLY.
    --test-date DD-MM-YYYY  post every bill with this date. TallyPrime Educational Mode only accepts
                       the 1st, 2nd or last day of a month, and the sample bills are dated 03-10-2026.
                       Use --test-date 01-10-2026 in TEST COMPANY. Never use it on a real company.
    --mark-manual 6651 6652  bills that accounts entered in Tally BY HAND (for example a bill the
                       program stopped on): put them in the sent list, so the program never sends
                       them later. Sends NOTHING to Tally. Use the same --settings as for sending.
"""
import argparse
import configparser
import csv
import hashlib
import os
import re
import sqlite3
from collections import OrderedDict, defaultdict
from datetime import datetime
from decimal import Decimal, ROUND_CEILING, ROUND_HALF_UP
from pathlib import Path
from xml.etree import ElementTree
from xml.sax.saxutils import escape

import requests

# --simple (used by Tally_Sync.bat): a short, coloured screen for accounts. Every detail still goes to
# run_log.txt, failed_itemwise.csv and changed_after_sent.csv, and a normal run prints everything as before.
SIMPLE = False
COLOURS_ON = False


def _enable_colours():
    """Colours on the Windows console (Windows 10 and later). Returns False if not possible: plain text then."""
    if os.name != "nt":
        return False
    try:
        import ctypes
        k = ctypes.windll.kernel32
        h = k.GetStdHandle(-11)
        mode = ctypes.c_uint32()
        return bool(k.GetConsoleMode(h, ctypes.byref(mode)) and k.SetConsoleMode(h, mode.value | 4))
    except Exception:
        return False


def c(text, colour):
    codes = {"green": "92", "yellow": "93", "red": "91", "cyan": "96", "bold": "1", "grey": "90"}
    return f"\033[{codes[colour]}m{text}\033[0m" if COLOURS_ON else str(text)


_LOCKS = []   # open lock files stay referenced here until the program ends


def acquire_lock(name="tally_sync.lock"):
    """Make sure only ONE copy runs at a time (handover rule: never two copies together).
    It is an operating-system lock on a small file next to the program: Windows frees it by itself
    when the program ends, even after a crash or a closed window, so it can never get stuck.
    Returns True if this copy got the lock, False if another copy is already running."""
    f = open(HERE / name, "a+")
    try:
        if os.name == "nt":
            import msvcrt
            f.seek(0)
            msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        f.close()
        return False
    _LOCKS.append(f)
    return True


def detail(*args, **kw):
    """A line for the full screen only; hidden on the --simple screen."""
    if not SIMPLE:
        print(*args, **kw)

# ---------------- settings (change here, nowhere else) ----------------
TALLY_URL = "http://localhost:9000"
COMPANY = "TEST COMPANY"            # must match the company name in Tally exactly
VOUCHER_TYPE = "Sales"
SALES_LEDGER = "Sales"
# Accounts' Tally has one sales ledger for local bills (CGST+SGST) and one for other states (IGST).
# Empty = use sales_ledger for both. When set, only 18% GST bills are sent (others stop for hand entry).
SALES_LEDGER_LOCAL = ""
SALES_LEDGER_IGST = ""
GST_RATE = Decimal("18")
# Voucher number in Tally. {no} = ERP bill number, {fy} = financial year of the bill date (26-27).
# Accounts' real Tally uses SAR/7301/26-27  ->  voucher_number_format = SAR/{no}/{fy}
VOUCHER_NUMBER_FORMAT = "{no}"
UNIT = "SQM"
PARTY_GROUP = "Sundry Debtors"
LEDGER_NAMES = {
    "cgst": "Output CGST",
    "sgst": "Output SGST",
    "igst": "Output IGST",
    # Accounts' Tally has 4 separate ledgers, one per ERP column (confirmed 08 Oct 2026).
    # Exact spelling still to confirm (for example "Cutout" or "CutOut").
    "hole": "Hole",
    "cutout": "Cutout",
    "big_hole": "Big Hole",
    "big_cutout": "Big Cutout",
    "freight": "Freight Charges",
    "document_charge": "Document Charge",     # ERP "Document Charge" line (not in the export; see below)
    "roundoff": "Round Off",
}
# The ERP item amounts already include Hole/Cutout charges, but Tally shows them as separate lines.
# True  = take Hole/Cutout out of the glass lines and post them as separate lines (total unchanged).
# False = leave them inside the glass lines and post no separate Hole/Cutout lines.
SPLIT_HOLE_CUTOUT = True
TOLERANCE = Decimal("0.05")
# Tag used for the party / tax / charges lines.
# Tally's own exports of item invoices ("Invoice Voucher View") use LEDGERENTRIES.LIST.
# ALLLEDGERENTRIES.LIST belongs to "Accounting Voucher View". Mixing them is a likely cause
# of "EXCEPTIONS 1" with no reason text. Change back to "ALLLEDGERENTRIES.LIST" only to compare.
LEDGER_TAG = "LEDGERENTRIES.LIST"

HERE = Path(__file__).parent
DEFAULT_HEADER = HERE / "header_ANON.csv"
DEFAULT_ITEMS = HERE / "items_ANON.csv"
DEFAULT_PI = HERE / "PI_21Aug_03Oct_ANON.csv"   # optional: ERP PI report (PI No, Total Area)
AREA_TOLERANCE = Decimal("0.001")                # SQM
MAPPING_FILE = HERE / "glass_name_mapping.csv"
PARTY_MAPPING_FILE = HERE / "party_mapping.csv"   # optional: ERP customer name -> Tally ledger name
DB_FILE = HERE / "tally_sent_itemwise.db"
FAILED_FILE = HERE / "failed_itemwise.csv"
CHANGED_FILE = HERE / "changed_after_sent.csv"   # bills changed in the ERP after they reached Tally
LOG_FILE = HERE / "run_log.txt"           # one line per run: when, what, how many, total
FAILED_XML_DIR = HERE / "failed_xml"       # XML + Tally's full reply of every bill Tally rejects
START_FROM_INVOICE = 0     # bills with a lower ERP number are never sent (0 = off). Set in settings.ini
INFER_DOCUMENT_CHARGE = True   # ERP exports leave out "Document Charge": take it as Ass. Value - items - freight
DOCUMENT_CHARGE_MAX = Decimal("100")   # ...only when that gap is whole rupees, above 0 and not above this
GST_PARTY_DETAILS = False   # True = add the party's state / registration type / place of supply to each bill
LIVE_MODE = False          # True = dummy-company shortcuts (--create-parties, --fallback-item, --test-date) refused
# ----------------------------------------------------------------------
# settings.ini (same folder) overrides the values above, so live use needs no code change.
SETTINGS_FILE = HERE / "settings.ini"
SETTINGS_USED = False


def apply_settings(path=SETTINGS_FILE):
    """Read settings.ini if it exists and replace the defaults above. Missing keys keep their default."""
    global TALLY_URL, COMPANY, VOUCHER_TYPE, SALES_LEDGER, UNIT, PARTY_GROUP, VOUCHER_NUMBER_FORMAT
    global SALES_LEDGER_LOCAL, SALES_LEDGER_IGST
    global DEFAULT_HEADER, DEFAULT_ITEMS, DEFAULT_PI, MAPPING_FILE, PARTY_MAPPING_FILE, DB_FILE, SETTINGS_FILE
    global START_FROM_INVOICE, LIVE_MODE, SETTINGS_USED, INFER_DOCUMENT_CHARGE, DOCUMENT_CHARGE_MAX, GST_PARTY_DETAILS
    path = Path(path)
    if not path.exists():
        return
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.optionxform = str                     # keep key case
    cfg.read(path, encoding="utf-8-sig")
    t = cfg["tally"] if cfg.has_section("tally") else {}
    TALLY_URL = t.get("url", TALLY_URL).strip()
    COMPANY = t.get("company", COMPANY).strip()
    VOUCHER_TYPE = t.get("voucher_type", VOUCHER_TYPE).strip()
    SALES_LEDGER = t.get("sales_ledger", SALES_LEDGER).strip()
    SALES_LEDGER_LOCAL = t.get("sales_ledger_local", SALES_LEDGER_LOCAL).strip()
    SALES_LEDGER_IGST = t.get("sales_ledger_igst", SALES_LEDGER_IGST).strip()
    if bool(SALES_LEDGER_LOCAL) != bool(SALES_LEDGER_IGST):
        raise SystemExit(f"STOP: {path.name} needs BOTH sales_ledger_local and sales_ledger_igst (or neither).")
    VOUCHER_NUMBER_FORMAT = t.get("voucher_number_format", VOUCHER_NUMBER_FORMAT).strip() or "{no}"
    if "{no}" not in VOUCHER_NUMBER_FORMAT:
        raise SystemExit(f"STOP: voucher_number_format in {path.name} must contain {{no}} (the ERP bill number).")
    UNIT = t.get("unit", UNIT).strip()
    PARTY_GROUP = t.get("party_group", PARTY_GROUP).strip()
    if cfg.has_section("ledgers"):
        for key in LEDGER_NAMES:
            if cfg["ledgers"].get(key, "").strip():
                LEDGER_NAMES[key] = cfg["ledgers"][key].strip()
    f = cfg["files"] if cfg.has_section("files") else {}
    resolve = lambda v, d: (HERE / v.strip()) if v and v.strip() else d
    DEFAULT_HEADER = resolve(f.get("header"), DEFAULT_HEADER)
    DEFAULT_ITEMS = resolve(f.get("items"), DEFAULT_ITEMS)
    DEFAULT_PI = resolve(f.get("pi"), DEFAULT_PI)
    MAPPING_FILE = resolve(f.get("mapping"), MAPPING_FILE)
    PARTY_MAPPING_FILE = resolve(f.get("party_mapping"), PARTY_MAPPING_FILE)
    DB_FILE = resolve(f.get("sent_list"), DB_FILE)   # one sent-list per Tally company
    sf = cfg["safety"] if cfg.has_section("safety") else {}
    START_FROM_INVOICE = int(str(sf.get("start_from_invoice", START_FROM_INVOICE)).strip() or 0)
    LIVE_MODE = str(sf.get("live_mode", "no")).strip().lower() in ("yes", "true", "1")
    INFER_DOCUMENT_CHARGE = str(sf.get("infer_document_charge", "yes")).strip().lower() in ("yes", "true", "1")
    DOCUMENT_CHARGE_MAX = Decimal(str(sf.get("document_charge_max", DOCUMENT_CHARGE_MAX)).strip() or "100")
    g = cfg["gst"] if cfg.has_section("gst") else {}
    GST_PARTY_DETAILS = str(g.get("party_details", "no")).strip().lower() in ("yes", "true", "1")
    SETTINGS_USED = True
    SETTINGS_FILE = path


apply_settings()


class BillError(Exception):
    """A bill that must not be posted. The message is written to the failed list."""


def D(x):
    s = str(x).strip() if x is not None else ""
    return Decimal(s) if s and s.lower() != "nan" else Decimal("0")


def money(x):
    return Decimal(x).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def norm(name):
    return re.sub(r"\s+", " ", str(name)).strip().upper()


def billed_area_sqm(row):
    """DUMMY RULE, for testing only.
    The ERP item-wise export has NO billed area. This rounds each side UP to the next 20 mm
    (Job Work uses exact size). It reproduced the ERP area for only about half of the bills.
    Replace this function when the real area source is known."""
    l1, l2, qty = D(row["L1-MM"]), D(row["L2-MM"]), D(row["Qty"])
    if str(row["Product Name"]).strip().upper() != "JOB WORK":
        step = Decimal("20")
        l1 = (l1 / step).to_integral_value(rounding=ROUND_CEILING) * step
        l2 = (l2 / step).to_integral_value(rounding=ROUND_CEILING) * step
    return l1 / 1000 * l2 / 1000 * qty


def load_pi_areas(path):
    """PI No -> Total Area (SQM) from the ERP PI report. Optional: returns {} if the file is missing."""
    areas = {}
    p = Path(path)
    if not p.exists():
        return areas
    for r in read_csv(p):
        no, area = str(r.get("PI No", "")).strip(), str(r.get("Total Area", "")).strip()
        if no and area:
            try:
                areas[no] = D(area)
            except Exception:
                pass
    return areas


def allocate_area(lines, bill_area, pi_areas):
    """Decide the SQM of each stock-item line. Returns (method, {item: area}).

    The header file's Area column is the ERP's billed area for the WHOLE bill (the real number).
    1. One stock-item line          -> that line gets the bill Area.                    method "exact"
    2. Several lines, and every PI belongs to only one line, and the PI Total Areas add up
       to the bill Area             -> each line gets the Total Area of its own PIs.     method "exact-PI"
    3. Otherwise                    -> the bill Area is shared out in proportion to the
       size-based dummy area of each line (last line takes the remainder, so the total is exact).
                                                                                         method "estimated"
    If the header has no Area at all, the dummy area is used as it is.                   method "dummy"
    """
    q = lambda x: x.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
    items = list(lines)
    if bill_area <= 0:
        return "dummy", {i: q(lines[i]["area"]) for i in items}
    if len(items) == 1:
        return "exact", {items[0]: q(bill_area)}

    owner = {}
    shared = False
    for i in items:
        for pi in lines[i]["pis"]:
            if owner.setdefault(pi, i) != i:
                shared = True
    if pi_areas and not shared and all(pi in pi_areas for pi in owner):
        by_pi = {i: sum((pi_areas[pi] for pi in lines[i]["pis"]), Decimal("0")) for i in items}
        if abs(sum(by_pi.values()) - bill_area) <= AREA_TOLERANCE and all(v > 0 for v in by_pi.values()):
            # PI areas have 3 decimals, the bill Area 4: the last line absorbs the tiny
            # difference, so the Tally total is exactly the ERP bill Area
            out = {i: q(by_pi[i]) for i in items[:-1]}
            out[items[-1]] = q(bill_area - sum(out.values(), Decimal("0")))
            return "exact-PI", out

    dummy_total = sum(lines[i]["area"] for i in items)
    out, given = {}, Decimal("0")
    for n, i in enumerate(items):
        if n == len(items) - 1:
            out[i] = q(bill_area - given)
        else:
            out[i] = q(bill_area * lines[i]["area"] / dummy_total)
            given += out[i]
    return "estimated", out


def load_mapping():
    mapping = {}
    if MAPPING_FILE.exists():
        with open(MAPPING_FILE, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                erp = norm(r.get("Glass Name (ERP)", ""))
                tally = (r.get("Tally stock item name (fill in)") or "").strip()
                if erp and tally:
                    mapping[erp] = tally
    return mapping


PARTY_MAP = {}   # filled in main() from party_mapping.csv


def load_party_mapping():
    """ERP customer name (spaces collapsed, capitals ignored) -> exact Tally ledger name.
    Only customers whose Tally name is spelt differently need a line; everyone else is used as is."""
    pm = {}
    if PARTY_MAPPING_FILE.exists():
        with open(PARTY_MAPPING_FILE, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                erp = norm(r.get("Customer Name (ERP)", ""))
                tally = (r.get("Tally ledger name") or "").strip()
                if erp and tally:
                    pm[erp] = tally
    return pm


def parse_date(text):
    """ERP dates come as 03-10-2026 in one export and 05/10/2026 in another (day first). -> 'YYYYMMDD' or None."""
    text = str(text or "").strip().split(" ")[0]
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y%m%d")
        except ValueError:
            pass
    return None


def voucher_number(inv, yyyymmdd):
    """Tally voucher number from the settings format: SAR/{no}/{fy} -> SAR/7301/26-27.
    The financial year runs April to March, so a 15-03-2027 bill is 26-27 and 01-04-2027 is 27-28."""
    y, m = int(yyyymmdd[:4]), int(yyyymmdd[4:6])
    start = y if m >= 4 else y - 1
    fy = f"{start % 100:02d}-{(start + 1) % 100:02d}"
    return VOUCHER_NUMBER_FORMAT.replace("{no}", inv).replace("{fy}", fy)


def build_bill(inv, header, rows, mapping, fallback, test_date=None, pi_areas=None):
    """Return a dict describing the voucher, or raise BillError."""
    party = header["Customer Name"].strip()
    party = PARTY_MAP.get(norm(party), party)      # Tally spelling, if party_mapping.csv has one
    if not party:
        raise BillError("party name is empty")
    date = parse_date(test_date or header["Inv Date"])
    if not date:
        raise BillError(f"bad invoice date: {header['Inv Date']!r}")

    # --- group the piece rows into Tally stock-item lines
    # One line per Product + Glass Name, like the ERP print and accounts' Tally bills: bill 6652 has the
    # same glass as Toughened and as Annealed, and accounts keep them as two lines (even with the same item).
    lines = OrderedDict()
    missing = []
    pieces = Decimal("0")
    for r in rows:
        erp_name = norm(r["Glass Name"])
        item = mapping.get(erp_name) or fallback
        if not item:
            if erp_name not in missing:
                missing.append(erp_name)
            continue
        key = (norm(r.get("Product Name", "")), erp_name)
        ln = lines.setdefault(key, {"item": item, "amount": Decimal("0"), "area": Decimal("0"), "pis": set()})
        ln["amount"] += D(r["Amount"])
        ln["area"] += billed_area_sqm(r)          # only used to share out the bill Area (method "estimated")
        if str(r.get("PI No", "")).strip():
            ln["pis"].add(str(r["PI No"]).strip())
        pieces += D(r["Qty"])
    if missing:
        raise BillError("no Tally stock item for: " + " | ".join(missing))

    raw_total = sum(l["amount"] for l in lines.values())
    glass_total = money(raw_total)
    freight = money(D(header["Freight Amt Billed"]))
    ass_value = money(D(header["Ass. Value"]))
    gap = money(ass_value - glass_total - freight)
    doc_charge = Decimal("0")
    if abs(gap) > TOLERANCE:
        # ERP bills 6647, 6687, 6711, 6712 have a "Document Charge" (Rs 20-50) that the exports leave out:
        # Basic Amount + Document Charge + Freight = Ass. Value (seen on the ERP invoice screen of 6687).
        if INFER_DOCUMENT_CHARGE and Decimal("0") < gap <= DOCUMENT_CHARGE_MAX and gap == gap.to_integral():
            doc_charge = gap
        else:
            raise BillError(f"item total {glass_total} + freight {freight} does not match Ass. Value {ass_value} "
                            f"(difference {gap})")

    igst, cgst, sgst = (money(D(header[k])) for k in ("I-GST Amt", "C-GST Amt", "S-GST Amt"))
    if igst > 0 and (cgst > 0 or sgst > 0):
        raise BillError("both IGST and CGST/SGST present")
    if SALES_LEDGER_LOCAL:
        # two sales ledgers in Tally: pick by tax type, and only for 18% bills
        tax = igst + cgst + sgst
        if tax <= 0:
            raise BillError("bill par GST nahi hai - haath se chadhao")
        rate = (tax * 100 / ass_value) if ass_value else Decimal("0")
        if abs(rate - GST_RATE) > Decimal("0.5"):
            raise BillError(f"GST rate {rate:.1f}% hai, {GST_RATE}% nahi - haath se chadhao")
        sales_ledger = SALES_LEDGER_IGST if igst > 0 else SALES_LEDGER_LOCAL
    else:
        sales_ledger = SALES_LEDGER
    net = money(D(header["Net Amt"]))
    roundoff = money(net - (ass_value + igst + cgst + sgst))
    if abs(roundoff) >= 1:
        raise BillError(f"round off {roundoff} is Rs 1 or more (Net Amt {net})")

    # --- Hole / Cutout carved out of the glass lines
    # One ERP column -> one Tally ledger (Hole, Cutout, Big Hole, Big Cutout)
    charges = [
        ("hole", money(D(header["Hole Amt"]))),
        ("cutout", money(D(header["CutOut Amt"]))),
        ("big_hole", money(D(header["Big Hole Amt"]))),
        ("big_cutout", money(D(header["Big CutOut Amt"]))),
    ]
    extras = sum((amt for _, amt in charges), Decimal("0")) if SPLIT_HOLE_CUTOUT else Decimal("0")
    if extras > glass_total:
        raise BillError(f"hole+cutout {extras} is more than glass total {glass_total}")
    items = list(lines.items())
    try:
        bill_area = D(header.get("Area") or 0)
    except Exception:
        bill_area = Decimal("0")
    area_method, areas = allocate_area(lines, bill_area, pi_areas or {})
    # The ERP's Ass. Value is the authority: glass lines + Hole/Cutout must add up to
    # Ass. Value minus freight to the paisa, so a difference of a few paise is absorbed here.
    remaining = (ass_value - freight - doc_charge) - extras
    out_items, allocated = [], Decimal("0")
    for i, (key, ln) in enumerate(items):
        item = ln["item"]
        if i == len(items) - 1:
            amt = money(remaining - allocated)
        else:
            amt = money(ln["amount"] * remaining / raw_total) if raw_total else Decimal("0")
        allocated += amt
        area = areas[key]
        if area <= 0 or amt < 0:
            raise BillError(f"stock item {item}: area {area} / amount {amt} is not valid")
        out_items.append({"item": item, "area": area, "amount": amt, "rate": money(amt / area)})

    # --- ledger lines other than the party and the stock items (credit side unless stated)
    others = []
    if SPLIT_HOLE_CUTOUT:
        for key, amt in charges:
            if amt:
                others.append((LEDGER_NAMES[key], amt, False))
    if doc_charge:
        others.append((LEDGER_NAMES["document_charge"], doc_charge, False))
    if freight:
        others.append((LEDGER_NAMES["freight"], freight, False))
    if igst:
        others.append((LEDGER_NAMES["igst"], igst, False))
    if cgst:
        others.append((LEDGER_NAMES["cgst"], cgst, False))
    if sgst:
        others.append((LEDGER_NAMES["sgst"], sgst, False))
    if roundoff:
        others.append((LEDGER_NAMES["roundoff"], abs(roundoff), roundoff < 0))

    credits = sum(i["amount"] for i in out_items) + sum(a for _, a, dr in others if not dr)
    debits = sum(a for _, a, dr in others if dr)
    if abs(credits - debits - net) > Decimal("0.005"):
        raise BillError(f"voucher does not balance: credits {credits} debits {debits} party {net}")

    for it in out_items:
        it["sales_ledger"] = sales_ledger
    return {"inv": inv, "vno": voucher_number(inv, date), "sales_ledger": sales_ledger, "party": party, "date": date, "net": net, "items": out_items,
            "others": others, "pieces": pieces, "area_method": area_method, "doc_charge": doc_charge}


def ledger_xml(name, amount, is_debit, party=False):
    # The minus sign on AMOUNT decides Debit (minus) or Credit (plus).
    # ISDEEMEDPOSITIVE tells Tally's invoice screen which SIDE the line belongs to:
    # in a sales invoice only the party is on the debit side ("Yes"); every other line
    # (sales, taxes, charges, round off) is on the credit side ("No"), even when its amount
    # is minus. Learned from Tally's own export of a hand-made bill (T1, 08 Oct 2026):
    # Round Off -0.54 is written as ISDEEMEDPOSITIVE No + AMOUNT -0.54.
    # With "Yes" the books were right but the invoice screen showed 0.48 instead of (-)0.48.
    sign = "-" if is_debit else ""
    side = "Yes" if party else "No"
    return f"""
            <{LEDGER_TAG}>
              <LEDGERNAME>{escape(name)}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>{side}</ISDEEMEDPOSITIVE>
              <ISLASTDEEMEDPOSITIVE>{side}</ISLASTDEEMEDPOSITIVE>{"<ISPARTYLEDGER>Yes</ISPARTYLEDGER>" if party else ""}
              <AMOUNT>{sign}{amount}</AMOUNT>
            </{LEDGER_TAG}>"""


def item_xml(it):
    qty = f" {it['area']} {UNIT}"
    return f"""
            <ALLINVENTORYENTRIES.LIST>
              <STOCKITEMNAME>{escape(it['item'])}</STOCKITEMNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <RATE>{it['rate']}/{UNIT}</RATE>
              <AMOUNT>{it['amount']}</AMOUNT>
              <ACTUALQTY>{qty}</ACTUALQTY>
              <BILLEDQTY>{qty}</BILLEDQTY>
              <ACCOUNTINGALLOCATIONS.LIST>
                <LEDGERNAME>{escape(it.get('sales_ledger', SALES_LEDGER))}</LEDGERNAME>
                <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
                <AMOUNT>{it['amount']}</AMOUNT>
              </ACCOUNTINGALLOCATIONS.LIST>
            </ALLINVENTORYENTRIES.LIST>"""


def build_xml(b):
    body = ledger_xml(b["party"], b["net"], True, party=True)
    body += "".join(item_xml(i) for i in b["items"])
    body += "".join(ledger_xml(*o) for o in b["others"])
    return f"""<ENVELOPE>
  <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES><SVCURRENTCOMPANY>{escape(COMPANY)}</SVCURRENTCOMPANY></STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <VOUCHER VCHTYPE="{VOUCHER_TYPE}" ACTION="Create" OBJVIEW="Invoice Voucher View">
            <DATE>{b['date']}</DATE>
            <VOUCHERTYPENAME>{VOUCHER_TYPE}</VOUCHERTYPENAME>
            <VOUCHERNUMBER>{escape(b['vno'])}</VOUCHERNUMBER>
            <PARTYLEDGERNAME>{escape(b['party'])}</PARTYLEDGERNAME>
            <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
            <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>{party_gst_xml(b["party"], b["gst"]) if b.get("gst") else ""}
            <ISINVOICE>Yes</ISINVOICE>
            <NARRATION>Auto-posted from ERP invoice {escape(b['inv'])} ({b['pieces']} pieces){"; glass-wise area estimated" if b['area_method'] in ("estimated", "dummy") else ""}{"; document charge taken as Ass. Value minus items" if b.get('doc_charge') else ""}</NARRATION>{body}
          </VOUCHER>
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>"""


PARTY_GST = {}   # norm(party name) -> {"state", "regtype", "gstin", "pos", "country"}; filled in main()
PARTY_GST_REPLY = HERE / "party_gst_reply.xml"


def fetch_party_gst():
    """Ask Tally (read only) for every party's state, GST registration type, GSTIN and place of supply.
    A GST company needs these inside each sales bill: a bill made by hand gets them from the Party
    Details screen; without them GSTR-1 lists the bill as 'Uncertain' (found 21 Oct 2026, bill T2 vs 6651)."""
    xml = f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>PartyGST</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES><SVCURRENTCOMPANY>{escape(COMPANY)}</SVCURRENTCOMPANY><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>
    <TDL><TDLMESSAGE>
      <COLLECTION NAME="PartyGST" ISMODIFY="No">
        <TYPE>Ledger</TYPE><CHILDOF>{escape(PARTY_GROUP)}</CHILDOF><BELONGSTO>Yes</BELONGSTO>
        <NATIVEMETHOD>Name</NATIVEMETHOD><NATIVEMETHOD>LedStateName</NATIVEMETHOD><NATIVEMETHOD>CountryName</NATIVEMETHOD>
        <NATIVEMETHOD>GSTRegistrationType</NATIVEMETHOD><NATIVEMETHOD>PartyGSTIN</NATIVEMETHOD>
        <NATIVEMETHOD>LedGSTRegDetails</NATIVEMETHOD><NATIVEMETHOD>LedMailingDetails</NATIVEMETHOD>
      </COLLECTION>
    </TDLMESSAGE></TDL>
  </DESC></BODY>
</ENVELOPE>"""
    ok, reply = send(xml)
    if not ok:
        return None, reply
    try:
        PARTY_GST_REPLY.write_text(reply, encoding="utf-8")   # kept for checking if something is missing
    except OSError:
        pass
    import html as _html
    out = {}
    for name, block in re.findall(r'<LEDGER\s+NAME="([^"]*)"[^>]*>(.*?)</LEDGER>', reply, re.S):
        def first(*tags):
            for t in tags:
                for v in re.findall(rf"<{t}(?:\s[^>]*)?>([^<]*)</{t}>", block):
                    if v.strip():
                        return _html.unescape(v.strip())
            return ""
        out[norm(_html.unescape(name))] = {
            "state": first("LEDSTATENAME", "STATENAME", "STATE"),
            "regtype": first("GSTREGISTRATIONTYPE"),
            "gstin": first("PARTYGSTIN", "GSTIN"),
            "pos": first("PLACEOFSUPPLY"),
            "country": first("COUNTRYNAME", "COUNTRYOFRESIDENCE") or "India",
        }
    return out, ""


def party_gst_xml(party, g):
    """Voucher-level party details, written the way Tally writes them for a bill made by hand (bill T2)."""
    state, pos = escape(g["state"]), escape(g.get("pos") or g["state"])
    country, p = escape(g.get("country") or "India"), escape(party)
    lines = [f"<GSTREGISTRATIONTYPE>{escape(g.get('regtype') or 'Unregistered/Consumer')}</GSTREGISTRATIONTYPE>",
             f"<STATENAME>{state}</STATENAME>", f"<COUNTRYOFRESIDENCE>{country}</COUNTRYOFRESIDENCE>",
             f"<PLACEOFSUPPLY>{pos}</PLACEOFSUPPLY>", f"<PARTYNAME>{p}</PARTYNAME>",
             f"<BASICBUYERNAME>{p}</BASICBUYERNAME>", f"<PARTYMAILINGNAME>{p}</PARTYMAILINGNAME>",
             f"<BASICBASEPARTYNAME>{p}</BASICBASEPARTYNAME>",
             f"<CONSIGNEEMAILINGNAME>{p}</CONSIGNEEMAILINGNAME>", f"<CONSIGNEESTATENAME>{state}</CONSIGNEESTATENAME>",
             f"<CONSIGNEECOUNTRYNAME>{country}</CONSIGNEECOUNTRYNAME>"]
    if g.get("gstin"):
        lines.append(f"<PARTYGSTIN>{escape(g['gstin'])}</PARTYGSTIN>")
    return "".join("\n            " + x for x in lines)


def party_xml(name):
    q = escape(name, {'"': "&quot;"})
    return f"""<ENVELOPE>
  <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
  <BODY><IMPORTDATA>
    <REQUESTDESC><REPORTNAME>All Masters</REPORTNAME>
      <STATICVARIABLES><SVCURRENTCOMPANY>{escape(COMPANY)}</SVCURRENTCOMPANY></STATICVARIABLES></REQUESTDESC>
    <REQUESTDATA><TALLYMESSAGE xmlns:UDF="TallyUDF">
      <LEDGER NAME="{q}" ACTION="Create"><NAME.LIST><NAME>{escape(name)}</NAME></NAME.LIST>
        <PARENT>{escape(PARTY_GROUP)}</PARENT></LEDGER>
    </TALLYMESSAGE></REQUESTDATA>
  </IMPORTDATA></BODY>
</ENVELOPE>"""


def count(tag, text):
    m = re.search(rf"<{tag}>(\d+)</{tag}>", text)
    return int(m.group(1)) if m else 0


EXISTING_REPLY = HERE / "existing_vouchers_reply.xml"
ALREADY_IN_TALLY = "already in Tally:"


def fetch_existing_vouchers(from_date, to_date):
    """Ask Tally (read only) for the voucher numbers of every voucher of our type in the given period.
    Accounts' Tally does not prevent duplicate numbers, so the program checks itself: a bill whose
    number is already there (for example typed by hand and not marked) is never sent again.
    Returns (set of normalised numbers, None) or (None, error text)."""
    xml = f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>ExistingVchNos</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES>
      <SVCURRENTCOMPANY>{escape(COMPANY)}</SVCURRENTCOMPANY><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      <SVFROMDATE>{from_date}</SVFROMDATE><SVTODATE>{to_date}</SVTODATE>
    </STATICVARIABLES>
    <TDL><TDLMESSAGE>
      <COLLECTION NAME="ExistingVchNos" ISMODIFY="No">
        <TYPE>Voucher</TYPE>
        <NATIVEMETHOD>VoucherNumber</NATIVEMETHOD><NATIVEMETHOD>VoucherTypeName</NATIVEMETHOD>
      </COLLECTION>
    </TDLMESSAGE></TDL>
  </DESC></BODY>
</ENVELOPE>"""
    ok, reply = send(xml)
    if not ok:
        return None, reply
    try:
        EXISTING_REPLY.write_text(reply, encoding="utf-8")   # kept for checking
    except OSError:
        pass
    import html as _html
    found = set()
    for block in re.findall(r"<VOUCHER\b[^>]*>(.*?)</VOUCHER>", reply, re.S):
        vtype = re.search(r"<VOUCHERTYPENAME[^>]*>([^<]*)</VOUCHERTYPENAME>", block)
        if vtype and _html.unescape(vtype.group(1)).strip().lower() != VOUCHER_TYPE.lower():
            continue                                   # receipts, payments etc. have their own numbers
        for num in re.findall(r"<VOUCHERNUMBER[^>]*>([^<]*)</VOUCHERNUMBER>", block):
            if num.strip():
                found.add(norm(_html.unescape(num)))
    return found, None


TALLY_CLOSED_HINDI = ("Tally se baat nahi ho payi (Tally band hai ya company load nahi). Kuch kharab nahi hua. "
                      "Tally kholo, company load karo, aur dobara chalao.")
NO_TALLY = "Could not reach Tally: is Tally open with the company loaded, and Connectivity (port 9000) on?"


def send(xml):
    """Return (ok, text)."""
    try:
        r = requests.post(TALLY_URL, data=xml.encode("utf-8"), timeout=60)
    except (requests.exceptions.MissingSchema, requests.exceptions.InvalidSchema,
            requests.exceptions.InvalidURL) as e:
        return False, (f"{NO_TALLY} - the url in the settings file looks wrong: {TALLY_URL!r} "
                       f"(it must be like http://localhost:9000)")
    except requests.exceptions.RequestException as e:
        return False, f"{NO_TALLY} ({type(e).__name__})"
    return True, r.text


def save_failed(inv, xml, reply):
    """Keep what we sent and what Tally answered, so the failure can be studied later."""
    FAILED_XML_DIR.mkdir(exist_ok=True)
    (FAILED_XML_DIR / f"{inv}.xml").write_text(xml, encoding="utf-8")
    (FAILED_XML_DIR / f"{inv}_reply.txt").write_text(reply, encoding="utf-8")


def post_voucher(xml):
    ok, text = send(xml)
    if not ok:
        return False, text, text
    if count("CREATED", text) == 1 and count("ERRORS", text) == 0 and count("EXCEPTIONS", text) == 0:
        return True, "created", text
    reasons = re.findall(r"<LINEERROR>(.*?)</LINEERROR>", text, re.S)
    summary = (f"CREATED {count('CREATED', text)}, ERRORS {count('ERRORS', text)}, "
               f"EXCEPTIONS {count('EXCEPTIONS', text)}")
    return False, summary + ("; " + "; ".join(x.strip() for x in reasons) if reasons else ""), text


def _cell_text(v):
    """Excel cell -> the text a CSV export would have: 6635.0 -> '6635', empty -> ''."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:                      # NaN = empty cell
            return ""
        if v.is_integer():
            return str(int(v))
        return repr(v)
    if hasattr(v, "strftime"):          # a real Excel date
        return v.strftime("%d-%m-%Y")
    return str(v).strip()


FP_HEADER_FIELDS = ("Inv Date", "Customer Name", "Qty", "Area", "Ass. Value", "S-GST Amt", "C-GST Amt",
                    "I-GST Amt", "Net Amt", "Hole Amt", "CutOut Amt", "Big Hole Amt", "Big CutOut Amt",
                    "Freight Amt Billed")
FP_ROW_FIELDS = ("Product Name", "Glass Name", "L1-MM", "L2-MM", "Qty", "Amount", "PI No")


def _fp_value(v):
    v = re.sub(r"\s+", " ", str(v or "")).strip().upper()
    try:
        d = Decimal(v)
        return format(d.normalize(), "f") if d == d.to_integral() else str(d.normalize())
    except Exception:
        return v


def fingerprint(hdr, rows):
    """A short code that changes whenever the ERP bill changes (amounts, tax, area, pieces, party, date).
    Built from the ERP data itself, so a change in this program never looks like a change in the bill.
    Returns (code, summary) - the summary is what people read in changed_after_sent.csv."""
    parts = [_fp_value(hdr.get(k, "")) for k in FP_HEADER_FIELDS]
    parts += sorted("|".join(_fp_value(r.get(k, "")) for k in FP_ROW_FIELDS) for r in rows)
    code = hashlib.sha256("\n".join(parts).encode("utf-8")).hexdigest()[:16]
    summary = (f"Net Rs {_fp_value(hdr.get('Net Amt'))}, Area {_fp_value(hdr.get('Area'))}, "
               f"{len(rows)} piece rows, Ass. Value {_fp_value(hdr.get('Ass. Value'))}")
    return code, summary


def read_csv(path):
    """Read an export as a list of rows (dicts). Works for .csv and for the ERP's own .xls / .xlsx,
    so live use needs no CSV conversion (and no anonymising: live data stays on the accounts PC)."""
    path = Path(path)
    if path.suffix.lower() in (".xls", ".xlsx"):
        try:
            import pandas as pd
        except ImportError:
            raise SystemExit("STOP: reading .xls needs pandas and xlrd:  python -m pip install pandas xlrd openpyxl")
        engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"
        df = pd.read_excel(path, engine=engine, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        return [{c: _cell_text(v) for c, v in row.items()} for row in df.to_dict("records")]
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def main():
    # --settings FILE (optional): use another settings file, e.g. settings_gst.ini for the GST test company
    pre = argparse.ArgumentParser(add_help=False)
    pre.add_argument("--settings")
    known, _ = pre.parse_known_args()
    if known.settings:
        if not (HERE / known.settings).exists() and not Path(known.settings).exists():
            print(f"STOP: settings file {known.settings} not found.")
            return
        apply_settings(known.settings if Path(known.settings).exists() else HERE / known.settings)
    ap = argparse.ArgumentParser(description="Post item-wise ERP bills to Tally (dummy data only)")
    ap.add_argument("header_csv", nargs="?", default=str(DEFAULT_HEADER))
    ap.add_argument("items_csv", nargs="?", default=str(DEFAULT_ITEMS))
    ap.add_argument("--settings", help="another settings file (default: settings.ini)")
    ap.add_argument("--pi", default=str(DEFAULT_PI), help="anonymised ERP PI report (optional)")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--only")
    ap.add_argument("--show")
    ap.add_argument("--create-parties", action="store_true")
    ap.add_argument("--fallback-item")
    ap.add_argument("--test-date", help="DD-MM-YYYY, overrides the invoice date (TEST COMPANY only)")
    ap.add_argument("--accept-changes", nargs="+", metavar="BILL",
                    help="after fixing changed bills in Tally by hand: bill numbers (or ALL) to stop reporting")
    ap.add_argument("--simple", action="store_true",
                    help="short coloured screen for accounts (Tally_Sync.bat); details stay in the log files")
    ap.add_argument("--mark-manual", nargs="+", metavar="BILL",
                    help="bills entered in Tally by hand: mark them as done so they are never sent (sends nothing)")
    a = ap.parse_args()
    global SIMPLE, COLOURS_ON
    SIMPLE = bool(a.simple and not a.dry_run and not a.only and not a.show)
    COLOURS_ON = (SIMPLE or bool(a.mark_manual)) and _enable_colours()
    if not a.dry_run and not acquire_lock():
        print("STOP: Tally Sync pehle se chal raha hai (doosri window ya button mein). "
              "Use khatam hone do, phir dobara chalao.")
        return
    if a.mark_manual:
        mark_manual(a)
        return

    if LIVE_MODE and (a.create_parties or a.fallback_item or a.test_date):
        print("STOP: live_mode is on in settings.ini, so --create-parties, --fallback-item and --test-date "
              "are not allowed (they are for the dummy company only).")
        return
    if LIVE_MODE and not START_FROM_INVOICE and not a.dry_run:
        # a real (or restored copy) company already holds the bills accounts typed by hand;
        # without a start-from bill, a range export would send those again
        print("STOP: live_mode is on but start_from_invoice is 0. Set start_from_invoice in the settings "
              "to the first ERP bill that is NOT yet in this Tally company, then run again.")
        return
    detail(f"Company: {COMPANY}  |  settings: {Path(SETTINGS_FILE).name if SETTINGS_USED else 'defaults in the script'}"
           f"  |  start from bill: {START_FROM_INVOICE or 'off'}  |  live mode: {'yes' if LIVE_MODE else 'no'}")
    if SIMPLE:
        print("\n" + c("=" * 72, "cyan"))
        print(c(f"  TALLY SYNC", "bold") + f"      Company: {COMPANY}      {datetime.now():%d-%m-%Y  %H:%M}")
        print(c("=" * 72, "cyan"))

    headers = {r["Inv No."].strip(): r for r in read_csv(a.header_csv)}
    items = defaultdict(list)
    for r in read_csv(a.items_csv):
        items[r["Inv No."].strip()].append(r)
    mapping = load_mapping()
    PARTY_MAP.update(load_party_mapping())
    pi_areas = load_pi_areas(a.pi)
    if GST_PARTY_DETAILS:
        got, err = fetch_party_gst()
        if got is None:
            if not a.dry_run:
                print("STOP: " + (TALLY_CLOSED_HINDI if str(err).startswith(NO_TALLY)
                                  else f"could not read party GST details from Tally: {err}"))
                return
            print("(dry run: Tally not reachable, party GST details not checked)")
        else:
            PARTY_GST.update(got)
            detail(f"Party GST details read from Tally for {len(got)} parties.")
    # --- duplicate check: voucher numbers already in Tally for the bills' financial year(s)
    existing = None
    bill_dates = [d for d in (parse_date(a.test_date or h.get("Inv Date")) for h in headers.values()) if d]
    if bill_dates:
        lo, hi = min(bill_dates), max(bill_dates)
        fy_lo = int(lo[:4]) if int(lo[4:6]) >= 4 else int(lo[:4]) - 1
        fy_hi = int(hi[:4]) if int(hi[4:6]) >= 4 else int(hi[:4]) - 1
        existing, err = fetch_existing_vouchers(f"{fy_lo}0401", f"{fy_hi + 1}0331")
        if existing is None:
            if not a.dry_run:
                print("STOP: " + (TALLY_CLOSED_HINDI if str(err).startswith(NO_TALLY)
                                  else f"could not read existing voucher numbers from Tally: {err}"))
                return
            detail("(dry run: Tally not reachable, duplicate voucher numbers not checked)")
        else:
            detail(f"Duplicate check: {len(existing)} '{VOUCHER_TYPE}' voucher numbers already in Tally.")
            if LIVE_MODE and not existing and not a.dry_run:
                print(f"STOP: Tally returned 0 '{VOUCHER_TYPE}' vouchers for this period, so the duplicate check "
                      f"cannot work (see {EXISTING_REPLY.name}). Nothing sent.")
                return
    detail(f"{len(headers)} bills in header file, {sum(len(v) for v in items.values())} item rows, "
          f"{len(mapping)} glass names mapped, {len(PARTY_MAP)} party names mapped, {len(pi_areas)} PI areas loaded.")

    db = None
    if not a.dry_run:
        db = sqlite3.connect(DB_FILE)
        db.execute("CREATE TABLE IF NOT EXISTS sent (invoice_no TEXT PRIMARY KEY, sent_at TEXT)")
        cols = {row[1] for row in db.execute("PRAGMA table_info(sent)")}
        for col in ("fingerprint", "summary"):            # older databases get the two new columns
            if col not in cols:
                db.execute(f"ALTER TABLE sent ADD COLUMN {col} TEXT")
        db.commit()
    accept = {x.strip().strip(",").upper() for x in (a.accept_changes or [])}
    changed = []

    done, skipped, failed, total = 0, 0, [], Decimal("0")
    before_start = 0
    created_parties = set()
    tally_stopped = False
    for n_i, inv in enumerate(sorted(headers), 1):
        if SIMPLE:
            print(f"\r  Bills dekh rahe hain aur Tally mein bhej rahe hain:  {n_i} / {len(headers)}   ", end="", flush=True)
        if a.only and inv != a.only.strip():
            continue
        if START_FROM_INVOICE and inv.isdigit() and int(inv) < START_FROM_INVOICE:
            before_start += 1                      # entered by hand before go-live: never send
            continue
        hdr = headers[inv]
        row_db = db.execute("SELECT fingerprint, summary FROM sent WHERE invoice_no=?", (inv,)).fetchone() if db else None
        if row_db:
            fp_now, sum_now = fingerprint(hdr, items.get(inv, []))
            fp_then, sum_then = row_db
            if fp_then is None or inv.upper() in accept or "ALL" in accept:
                # first look at an older record, or accounts fixed Tally by hand: remember the ERP bill as it is now
                db.execute("UPDATE sent SET fingerprint=?, summary=? WHERE invoice_no=?", (fp_now, sum_now, inv))
                db.commit()
                if fp_then is not None and fp_then != fp_now:
                    print(f"ACCEPTED {inv}  (change marked as fixed in Tally)")
            elif fp_then != fp_now:
                detail(f"CHANGED {inv}  sent earlier, but the ERP bill changed since:  then [{sum_then}]  now [{sum_now}]")
                changed.append({"inv_no": inv, "party": hdr["Customer Name"].strip(),
                                "when_sent": sum_then, "erp_now": sum_now})
                skipped += 1
                continue
            detail(f"SKIP   {inv}  (already sent earlier)")
            skipped += 1
            continue
        try:
            if inv not in items:
                raise BillError("no item rows found in the item-wise file")
            bill = build_bill(inv, hdr, items[inv], mapping, a.fallback_item, a.test_date, pi_areas)
            if GST_PARTY_DETAILS and PARTY_GST:
                g = PARTY_GST.get(norm(bill["party"]))
                if not g or not g.get("state"):
                    raise BillError(f"GST details (state) of party {bill['party']} not found in Tally "
                                    f"(check its ledger; see {PARTY_GST_REPLY.name})")
                bill["gst"] = g
            if existing is not None and norm(bill["vno"]) in existing:
                raise BillError(f"{ALREADY_IN_TALLY} {bill['vno']} (probably typed by hand) - not sent again")
            xml = build_xml(bill)
            ElementTree.fromstring(xml)                    # the XML must be well formed
        except BillError as e:
            detail(f"FAILED {inv}  {hdr['Customer Name'].strip()}  -> {e}")
            failed.append({"inv_no": inv, "party": hdr["Customer Name"].strip(), "reason": str(e)})
            continue
        except ElementTree.ParseError as e:
            detail(f"FAILED {inv}  (XML error: {e})")
            failed.append({"inv_no": inv, "party": hdr["Customer Name"].strip(), "reason": f"XML error: {e}"})
            continue

        if a.show and inv == a.show.strip():
            print(xml)
        if a.dry_run:
            print(f"OK(dry) {inv}  ({bill['vno']})  {bill['party']}  Rs {bill['net']}  ({len(bill['items'])} stock lines, area {bill['area_method']})")
            done += 1
            total += bill["net"]
            continue

        if a.create_parties and bill["party"] not in created_parties:
            ok_p, reply_p = send(party_xml(bill["party"]))
            if ok_p and count("CREATED", reply_p) == 1:
                detail(f"       party ledger created: {bill['party']}")
            # anything else usually means "already exists", which is fine
            created_parties.add(bill["party"])
        ok, msg, reply = post_voucher(xml)
        if not ok and msg.startswith(NO_TALLY):
            # Tally is closed: stop now instead of trying every bill; the rest go next run
            if SIMPLE:
                print("\n\n" + c("  RUK GAYA: Tally se baat nahi ho payi (Tally band hai ya company load nahi).", "red"))
                print("  Kuch kharab nahi hua. Tally kholo, company load karo, aur yeh button dobara chalao.")
            else:
                print(f"STOPPED at {inv}: {msg}")
                print("       Nothing was lost. Open Tally and run again: bills not sent yet will go then.")
            failed.append({"inv_no": inv, "party": bill["party"], "reason": msg})
            tally_stopped = True
            break
        if ok:
            if existing is not None:
                existing.add(norm(bill["vno"]))
            fp, summ = fingerprint(hdr, items[inv])
            db.execute("INSERT INTO sent (invoice_no, sent_at, fingerprint, summary) VALUES (?, ?, ?, ?)",
                       (inv, datetime.now().isoformat(timespec="seconds"), fp, summ))
            db.commit()
            detail(f"OK     {inv}  ({bill['vno']})  {bill['party']}  Rs {bill['net']}  (area {bill['area_method']})")
            done += 1
            total += bill["net"]
        else:
            save_failed(inv, xml, reply)
            detail(f"FAILED {inv}  {bill['party']}  -> {msg}")
            detail(f"       XML and Tally reply saved in {FAILED_XML_DIR.name}\\{inv}.xml / {inv}_reply.txt")
            failed.append({"inv_no": inv, "party": bill["party"], "reason": msg})

    if failed:
        with open(FAILED_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["inv_no", "party", "reason"])
            w.writeheader()
            w.writerows(failed)
    elif FAILED_FILE.exists() and not a.only:
        FAILED_FILE.unlink()                           # no failures this run: an old list would mislead
    if db:
        if changed:
            with open(CHANGED_FILE, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=["inv_no", "party", "when_sent", "erp_now"])
                w.writeheader()
                w.writerows(changed)
        elif CHANGED_FILE.exists():
            CHANGED_FILE.unlink()                          # nothing open any more
    label = "checked OK (nothing sent)" if a.dry_run else "created"
    if SIMPLE:
        print()                                    # end the running count line
    if before_start:
        detail(f"\n{before_start} bills are numbered below {START_FROM_INVOICE} (start from bill) and were left out.")
    detail(f"\nDone: {done} {label}, {skipped} skipped, {len(failed)} failed.  Total of OK bills: Rs {total}")
    if failed:
        detail(f"Failed bills saved in {FAILED_FILE.name}.")
    if changed and not SIMPLE:
        print(f"{len(changed)} bill(s) CHANGED in the ERP after reaching Tally: see {CHANGED_FILE.name}. "
              "Correct them in Tally by hand, then double-click Accept_Change.bat.")
    if tally_stopped:
        pass        # Tally was closed: a day's count now would wrongly call every unsent bill "enter by hand"
    elif db and not a.only:
        if SIMPLE:
            print(f"  Is baar program ne {c(done, 'green')} naye bills Tally mein chadhaye  (Rs {total}).")
        day_summary(headers, db, {f["inv_no"]: f["reason"] for f in failed})
    if SIMPLE and changed:
        print(c("  DHYAN: yeh bills Tally mein jaane ke BAAD ERP mein badle gaye:", "red"))
        for ch in changed:
            print(f"    {ch['inv_no']:<10} {ch['party']}")
        print("  Tally mein haath se theek karo, phir Accept_Change.bat chalao.")
        print(c("=" * 72, "cyan"))
    write_log(a, done, skipped, failed, total, changed)


def _why_hindi(reason):
    """Short, plain reason for the day's count (the full English reason stays in failed_itemwise.csv)."""
    r = (reason or "").strip()
    if r.startswith("no Tally stock item for:"):
        return "glass mapping mein nahi: " + r.split(":", 1)[1].strip()
    if r.startswith(ALREADY_IN_TALLY):
        num = r[len(ALREADY_IN_TALLY):].split("(")[0].strip()
        return f"Tally mein pehle se hai ({num}) - check karo, sahi ho to tick lagao"
    if r.startswith("no item rows"):
        return "item-wise file mein is bill ki lines nahi"
    if "GST details" in r:
        return "party ki GST state Tally mein nahi"
    if r.startswith(NO_TALLY):
        return "Tally band tha - program dobara chalao"
    return r or "program ne nahi bheja (failed_itemwise.csv dekho)"


def day_status(headers, db):
    """For the bills in this ERP file: (sent by the program, entered by hand, still waiting, count below
    the start-from bill). Used by the day's count on screen and by the Tally Sync app window."""
    auto, manual, waiting, before = [], [], [], 0
    for inv in sorted(headers):
        if START_FROM_INVOICE and inv.isdigit() and int(inv) < START_FROM_INVOICE:
            before += 1
            continue
        row = db.execute("SELECT sent_at FROM sent WHERE invoice_no=?", (inv,)).fetchone()
        if not row:
            waiting.append(inv)
        elif str(row[0]).endswith("(manual)"):
            manual.append(inv)
        else:
            auto.append(inv)
    return auto, manual, waiting, before


def day_summary(headers, db, reasons=None):
    """The day's count for the bills in this ERP file: sent by the program, entered by hand
    (--mark-manual / Tally_Sync.bat), and still waiting. The waiting numbers are printed on one
    line, ready to copy. Returns the waiting bill numbers."""
    reasons = reasons or {}
    auto, manual, waiting, before = day_status(headers, db)
    total = len(auto) + len(manual) + len(waiting)
    line, thin = c("=" * 72, "cyan"), c("-" * 72, "grey")
    print("\n" + line)
    print(c("  AAJ KA HISAAB", "bold") + f"      ERP file mein {total} bills"
          + (f"   ({before} bills start-from se pehle ke, gine nahi)" if before else ""))
    print(line)
    print(f"  Tally mein chadh gaye :  {c(f'{len(auto) + len(manual):>4}', 'green')}"
          f"    (program se: {len(auto)},  haath se: {len(manual)})")
    print(f"  Abhi baaki            :  {c(f'{len(waiting):>4}', 'yellow' if waiting else 'green')}")
    if not waiting:
        print(thin)
        print(c("  Sab bills Tally mein hain. Kuch baaki nahi.", "green"))
        print(line)
        return waiting
    print(thin)
    print(c("  Yeh bills HAATH SE Tally mein chadhane hain", "yellow") + " (wahi ERP bill number aur date):\n")
    print(c(f"  {'Bill no.':<10} {'Party':<28} Kyun ruka", "grey"))
    for inv in waiting:
        party = str(headers[inv].get("Customer Name", "")).strip()
        party = party if len(party) <= 28 else party[:25] + "..."
        why = _why_hindi(reasons[inv]) if inv in reasons else "abhi tak bheja nahi gaya"
        print(f"  {inv:<10} {party:<28} {why}")
    print(thin)
    print("  Haath se chadhane ke baad: yeh line copy karke neeche 'Bill numbers:' mein paste karo, Enter:\n")
    print("      " + c(" ".join(waiting), "cyan") + "\n")
    print("  (Copy: mouse se line select karo, phir Enter ya right-click.  Paste: right-click.)")
    print(line)
    return waiting


def mark_manual(a):
    """--mark-manual: bills accounts entered by hand go into the sent list, so they are never sent.
    Nothing is sent to Tally. The ERP bill as it is now is remembered, so a later change is still reported."""
    bills = []
    for x in a.mark_manual:                       # accepts "6651 6652" and "6651,6652"
        bills += [b.strip().upper() for b in x.split(",") if b.strip()]
    headers = {r["Inv No."].strip(): r for r in read_csv(a.header_csv)}
    items = defaultdict(list)
    for r in read_csv(a.items_csv):
        items[r["Inv No."].strip()].append(r)
    by_upper = {k.upper(): k for k in headers}    # stored exactly as the ERP writes it, like a normal send
    db = sqlite3.connect(DB_FILE)
    db.execute("CREATE TABLE IF NOT EXISTS sent (invoice_no TEXT PRIMARY KEY, sent_at TEXT)")
    cols = {row[1] for row in db.execute("PRAGMA table_info(sent)")}
    for col in ("fingerprint", "summary"):
        if col not in cols:
            db.execute(f"ALTER TABLE sent ADD COLUMN {col} TEXT")
    print(f"\n  Haath se chadhe bills yaad kar rahe hain   (Company: {COMPANY}; Tally ko kuch nahi bheja jata)\n")
    marked = []
    for inv in bills:
        inv = by_upper.get(inv, inv)
        if db.execute("SELECT 1 FROM sent WHERE invoice_no=?", (inv,)).fetchone():
            print(c("PEHLE SE ", "yellow") + f" {inv}  yeh bill pehle se Tally wali list mein hai, kuch nahi badla")
            continue
        if inv not in headers:
            # a typed-wrong number would block a real bill forever, so only bills in the ERP file are accepted
            print(c("NAHI MILA", "red") + f" {inv}  yeh number aaj ki ERP file ({Path(a.header_csv).name}) mein nahi hai. "
                  "Kuch nahi badla - number check karo.")
            continue
        fp, summ = fingerprint(headers[inv], items.get(inv, []))
        db.execute("INSERT INTO sent (invoice_no, sent_at, fingerprint, summary) VALUES (?, ?, ?, ?)",
                   (inv, datetime.now().isoformat(timespec="seconds") + " (manual)", fp, summ))
        marked.append(inv)
        print(c("HO GAYA  ", "green") + f" {inv}  haath se chadha hua yaad kar liya - program ise kabhi nahi bhejega")
    db.commit()
    print(f"\n{len(marked)} bill(s) haath se chadhe hue yaad kar liye.")
    day_summary(headers, db)
    db.close()
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%d-%m-%Y %H:%M:%S} | MANUAL  | {COMPANY} | marked as entered by hand: "
                    f"{','.join(marked) or '-'}\n")
    except OSError as e:
        print(f"(could not write {LOG_FILE.name}: {e})")


def write_log(a, done, skipped, failed, total, changed=()):
    """Add one line to run_log.txt, so there is a record of every run (who checks it: accounts / me)."""
    mode = "DRY-RUN" if a.dry_run else "LIVE"
    extra = []
    if a.only:
        extra.append(f"only {a.only}")
    if START_FROM_INVOICE:
        extra.append(f"start-from {START_FROM_INVOICE}")
    if a.test_date:
        extra.append(f"test-date {a.test_date}")
    if a.fallback_item:
        extra.append("fallback-item used")
    failed_nos = ",".join(f["inv_no"] for f in failed) or "-"
    line = (f"{datetime.now():%d-%m-%Y %H:%M:%S} | {mode:7} | {COMPANY} | ok {done} | skipped {skipped} | "
            f"failed {len(failed)} ({failed_nos}) | total Rs {total}"
            + (f" | changed {len(changed)} ({','.join(c['inv_no'] for c in changed)})" if changed else "")
            + (f" | {'; '.join(extra)}" if extra else "") + "\n")
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except OSError as e:
        print(f"(could not write {LOG_FILE.name}: {e})")


if __name__ == "__main__":
    main()
