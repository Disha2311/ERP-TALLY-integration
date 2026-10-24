"""
check_parties.py  -  does every ERP customer have a ledger in Tally?

What it does, in plain words (it only READS from Tally, it never creates or changes anything):
1. Reads the customer names from the header export (the same file tally_itemwise.py uses).
2. Asks Tally for the list of ledgers under the party group (Sundry Debtors).
3. For every ERP customer it decides:
     OK         the same name exists in Tally (or party_mapping.csv points to one that exists)
     SPELLING   Tally has it with different capitals or spaces -> use the Tally spelling
     SIMILAR    no exact match, but a Tally name looks close -> a person must check
     MISSING    nothing like it in Tally -> accounts must create the ledger
4. Writes party_check.csv. For SPELLING and SIMILAR rows, copy the right Tally name into
   party_mapping.csv (columns: Customer Name (ERP), Tally ledger name, Notes).

Run (Tally open with the company loaded):
    python check_parties.py                               # uses settings.ini
    python check_parties.py --settings settings_gst.ini   # another company / settings file
"""
import argparse
import csv
import difflib
import html
import re
from datetime import datetime
from pathlib import Path

import tally_itemwise as ti     # whole module, so the values read from --settings are seen here
from tally_itemwise import escape, load_party_mapping, norm, read_csv, send


def ledger_list_xml():
    return f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>PartyLedgers</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES>
      <SVCURRENTCOMPANY>{escape(ti.COMPANY)}</SVCURRENTCOMPANY>
      <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    </STATICVARIABLES>
    <TDL><TDLMESSAGE>
      <COLLECTION NAME="PartyLedgers" ISMODIFY="No">
        <TYPE>Ledger</TYPE>
        <CHILDOF>{escape(ti.PARTY_GROUP)}</CHILDOF>
        <BELONGSTO>Yes</BELONGSTO>
        <NATIVEMETHOD>Name</NATIVEMETHOD>
      </COLLECTION>
    </TDLMESSAGE></TDL>
  </DESC></BODY>
</ENVELOPE>"""


def tally_party_names():
    ok, reply = send(ledger_list_xml())
    if not ok:
        raise SystemExit(f"STOP: {reply}")
    names = re.findall(r'<LEDGER\s+NAME="([^"]*)"', reply)
    if not names:   # some Tally versions put the name in a NAME tag instead
        names = re.findall(r"<LEDGER[^>]*>.*?<NAME[^>]*>([^<]*)</NAME>", reply, re.S)
    return sorted({html.unescape(n).strip() for n in names if n.strip()})


def main():
    ap = argparse.ArgumentParser(description="Check that every ERP customer has a ledger in Tally (read only)")
    ap.add_argument("--settings", help="another settings file (default: settings.ini)")
    a = ap.parse_args()
    if a.settings:
        path = Path(a.settings) if Path(a.settings).exists() else ti.HERE / a.settings
        if not path.exists():
            raise SystemExit(f"STOP: settings file {a.settings} not found.")
        ti.apply_settings(path)
    COMPANY, PARTY_GROUP = ti.COMPANY, ti.PARTY_GROUP
    OUT_FILE = ti.PARTY_MAPPING_FILE.with_name("party_check.csv")

    erp_names = sorted({r["Customer Name"].strip() for r in read_csv(ti.DEFAULT_HEADER) if r.get("Customer Name", "").strip()})
    tally_names = tally_party_names()
    if not tally_names:
        print(f"Tally returned no ledgers under '{PARTY_GROUP}' in '{COMPANY}'. "
              "Check that Tally is open with this company, and the group name in settings.ini.")
        return
    party_map = load_party_mapping()
    exact = set(tally_names)
    by_norm = {norm(t): t for t in tally_names}
    norm_list = list(by_norm)

    rows, counts = [], {"OK": 0, "SPELLING": 0, "SIMILAR": 0, "MISSING": 0}
    for erp in erp_names:
        target = party_map.get(norm(erp), erp)
        if target in exact:
            status, suggestion = "OK", target
        elif norm(target) in by_norm:
            status, suggestion = "SPELLING", by_norm[norm(target)]
        else:
            close = difflib.get_close_matches(norm(target), norm_list, n=1, cutoff=0.8)
            if close:
                status, suggestion = "SIMILAR", by_norm[close[0]]
            else:
                status, suggestion = "MISSING", ""
        counts[status] += 1
        rows.append({"Customer Name (ERP)": erp, "Status": status, "Tally name to use": suggestion,
                     "Mapped in party_mapping.csv": "yes" if norm(erp) in party_map else ""})

    with open(OUT_FILE, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)

    print(f"Company: {COMPANY}  |  {len(tally_names)} ledgers under '{PARTY_GROUP}' in Tally  |  "
          f"{len(erp_names)} customers in the ERP file")
    for st in ("OK", "SPELLING", "SIMILAR", "MISSING"):
        print(f"  {st:9} {counts[st]}")
    for r in rows:
        if r["Status"] != "OK":
            print(f"  {r['Status']:9} {r['Customer Name (ERP)']}  ->  {r['Tally name to use'] or '(create in Tally)'}")
    print(f"\nSaved {OUT_FILE.name} ({datetime.now():%d-%m-%Y %H:%M}).")
    if counts["SPELLING"] or counts["SIMILAR"]:
        print("For SPELLING/SIMILAR rows: check the name, then add a line to party_mapping.csv.")
    if counts["MISSING"]:
        print("For MISSING rows: accounts create the ledger in Tally (exact ERP name), then run this again.")


if __name__ == "__main__":
    main()
