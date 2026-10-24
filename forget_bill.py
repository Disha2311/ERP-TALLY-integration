"""
forget_bill.py  -  make the program "forget" bills, so they are sent to Tally again.

Use it ONLY after the same bills were deleted in Tally (Day Book, select, Alt+D).
If a bill is still in Tally and you forget it here, the next run creates it TWICE.

    python forget_bill.py 6651            one bill
    python forget_bill.py 6651 6652       several bills
    python forget_bill.py ALL             every bill (after deleting all of them in Tally)
"""
import sqlite3
import sys
from datetime import datetime

from tally_itemwise import COMPANY, DB_FILE, LOG_FILE


def main(args):
    bills = [a.strip().strip(",") for a in args if a.strip().strip(",")]
    if not bills:
        print(__doc__)
        return
    if not DB_FILE.exists():
        print(f"{DB_FILE.name} does not exist yet: nothing has been sent, nothing to forget.")
        return
    forget_all = len(bills) == 1 and bills[0].upper() == "ALL"
    what = "ALL bills" if forget_all else "bill(s) " + ", ".join(bills)
    print(f"Company: {COMPANY}")
    answer = input(f"Have you already deleted {what} in Tally? Type Y to continue: ").strip().upper()
    if answer != "Y":
        print("Stopped. Nothing was changed.")
        return
    db = sqlite3.connect(DB_FILE)
    try:
        db.execute("CREATE TABLE IF NOT EXISTS sent (invoice_no TEXT PRIMARY KEY, sent_at TEXT)")
        if forget_all:
            n = db.execute("DELETE FROM sent").rowcount
        else:
            n = sum(db.execute("DELETE FROM sent WHERE invoice_no=?", (b,)).rowcount for b in bills)
        db.commit()
    finally:
        db.close()
    print(f"Forgotten: {n} bill(s). The next run will send them to Tally again.")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%d-%m-%Y %H:%M:%S} | FORGET  | {COMPANY} | {what} | removed {n}\n")


if __name__ == "__main__":
    main(sys.argv[1:])
