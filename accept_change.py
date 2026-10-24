"""
accept_change.py  -  "I have corrected this changed bill in Tally by hand, stop reporting it."

When the program says  CHANGED 6651  (the ERP bill changed after it reached Tally):
  1. Accounts correct bill 6651 in Tally by hand.
  2. Run this (or double-click Accept_Change.bat) with the bill number.
It only updates the program's own record. It never sends anything to Tally.

    python accept_change.py 6651          one bill
    python accept_change.py 6651 6652     several
    python accept_change.py ALL           every bill now listed in changed_after_sent.csv
"""
import csv
import sqlite3
import sys
from collections import defaultdict
from datetime import datetime

from tally_itemwise import (CHANGED_FILE, COMPANY, DB_FILE, DEFAULT_HEADER, DEFAULT_ITEMS, LOG_FILE,
                            fingerprint, read_csv)


def main(args):
    bills = [a.strip().strip(",") for a in args if a.strip().strip(",")]
    if not bills:
        print(__doc__)
        return
    if not DB_FILE.exists():
        print("Nothing has been sent yet, so there is nothing to accept.")
        return
    if len(bills) == 1 and bills[0].upper() == "ALL":
        if not CHANGED_FILE.exists():
            print(f"{CHANGED_FILE.name} does not exist: no changed bills are open.")
            return
        with open(CHANGED_FILE, newline="", encoding="utf-8") as f:
            bills = [r["inv_no"] for r in csv.DictReader(f)]
    headers = {r["Inv No."].strip(): r for r in read_csv(DEFAULT_HEADER)}
    items = defaultdict(list)
    for r in read_csv(DEFAULT_ITEMS):
        items[r["Inv No."].strip()].append(r)

    print(f"Company: {COMPANY}")
    answer = input(f"Have you already corrected bill(s) {', '.join(bills)} in Tally by hand? Type Y to continue: ")
    if answer.strip().upper() != "Y":
        print("Stopped. Nothing was changed.")
        return
    db = sqlite3.connect(DB_FILE)
    done = []
    try:
        for b in bills:
            if b not in headers:
                print(f"  {b}: not in the current ERP export, skipped")
                continue
            fp, summ = fingerprint(headers[b], items.get(b, []))
            n = db.execute("UPDATE sent SET fingerprint=?, summary=? WHERE invoice_no=?", (fp, summ, b)).rowcount
            print(f"  {b}: {'accepted' if n else 'was never sent, nothing to accept'}")
            if n:
                done.append(b)
        db.commit()
    finally:
        db.close()
    if done and CHANGED_FILE.exists():                      # remove accepted bills from the open list
        with open(CHANGED_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
        left = [r for r in rows if r["inv_no"] not in done]
        if left:
            with open(CHANGED_FILE, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=list(rows[0]))
                w.writeheader()
                w.writerows(left)
        else:
            CHANGED_FILE.unlink()
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%d-%m-%Y %H:%M:%S} | ACCEPT  | {COMPANY} | bill(s) {', '.join(done) or '-'}\n")


if __name__ == "__main__":
    main(sys.argv[1:])
