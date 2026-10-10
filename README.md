# ERP to Tally: automatic sales bill posting

**Status (10 Oct 2026):** working on DUMMY data in a dummy Tally company on my laptop. Not connected to the live ERP or the live Tally. Going live needs my manager's approval, testing on a backup copy, and a supervised trial (see "Going live").

**Owner:** Disha (AI&TECH developer). Last updated 10 Oct 2026.

---

## 1. What this does, in one paragraph

Today accounts type every sales bill twice: once in the ERP (esyGlass / RecEmsPro) and again in TallyPrime, about 2 hours a day. This program reads the bills from the ERP's export files and creates the same bills in Tally automatically. The ERP is not changed in any way (it is only read). Before sending a bill, the program checks that all totals match the ERP. If anything does not match, the bill is not sent; it goes to a "failed" list with the reason.

## 2. How it works

```
ERP export files (.xls)  ->  program checks every bill  ->  Tally (localhost:9000)
   header, item-wise, PI        totals, names, duplicates       creates the Sales voucher
```

For each bill the program:

1. Joins the header file (one row per bill) with the item-wise file (one row per glass piece) on the invoice number.
2. Turns each ERP glass name into the Tally stock item name (using `glass_name_mapping.csv`).
3. Makes one stock line per glass type, with quantity in SQM (square metres).
4. Adds separate lines for Hole, Cutout, Big Hole, Big Cutout, Freight, GST (IGST, or CGST + SGST) and Round Off.
5. Uses the ERP invoice number as the Tally voucher number.
6. Runs the safety checks (section 6). Only if all pass, sends the bill to Tally.
7. Remembers every bill it sent, so it never sends the same bill twice.

## 3. Files in the folder

| File | What it is |
|---|---|
| `tally_itemwise.py` | **The main program.** |
| `Tally_Sync.bat` | Double-click button that runs the program and shows the last 5 runs. |
| `glass_name_mapping.csv` | ERP glass name to Tally stock item name. One row per glass name. |
| `header_ANON.csv`, `items_ANON.csv` | Anonymised ERP exports for 03 Oct 2026 (dummy customer names). |
| `PI_21Aug_03Oct_ANON.csv` | Anonymised ERP PI report (PI number and total area). Optional. |
| `anonymise_exports.py` | Makes the anonymised copies from the real ERP exports. |
| `create_stock_items.py` | Creates dummy stock items in TEST COMPANY. **Dummy company only.** |
| `tally_sent_itemwise.db` | Created by the program: list of bills already sent (duplicate protection). |
| `run_log.txt` | Created by the program: one line per run (date, time, sent, skipped, failed, total). |
| `failed_itemwise.csv` | Created by the program: bills that were stopped, with the reason. |
| `failed_xml\` | Created by the program: what was sent and Tally's reply, for bills Tally rejected. |
| `tally_test.py`, `tally_invoice_test.py`, `tally_batch.py`, `bills.csv`, `check_drivers.py` | Early tests. Not needed any more. |

Real (not anonymised) ERP files are kept in `C:\erp_data`, **outside OneDrive**. Never put them in this folder.

## 4. One-time setup (laptop)

1. **Python** installed (3.x). Then in PowerShell:
   ```
   python -m pip install requests pandas xlrd
   ```
2. **Tally as server:** in TallyPrime, F1 Help > Settings > Connectivity > Client/Server configuration: TallyPrime acts as **Server**, port **9000**.
3. **Dummy company "TEST COMPANY"** with: unit `SQM` (4 decimals), the stock items, party ledgers (Sundry Debtors), and the ledgers in section 5. Sales voucher type numbering = **Manual** (needed so the ERP number becomes the voucher number).

Note: my laptop has TallyPrime in **Educational mode**, which only allows vouchers on the 1st, 2nd or last day of a month. That is why tests use `--test-date 01-10-2026`.

## 5. Settings (top of `tally_itemwise.py`)

Change settings only in this block, nowhere else.

| Setting | Now (dummy) | Meaning |
|---|---|---|
| `TALLY_URL` | `http://localhost:9000` | Where Tally listens. |
| `COMPANY` | `TEST COMPANY` | Tally company name, exactly. |
| `VOUCHER_TYPE` | `Sales` | Voucher type name. |
| `SALES_LEDGER` | `Sales` | Sales ledger used for the glass lines. |
| `UNIT` | `SQM` | Unit of the stock items. |
| `LEDGER_NAMES` | Output CGST / SGST / IGST, Hole, Cutout, Big Hole, Big Cutout, Freight Charges, Round Off | Ledger names in Tally. Must match exactly. |
| `SPLIT_HOLE_CUTOUT` | `True` | ERP item amounts include Hole/Cutout; Tally shows them as separate lines. |
| `TOLERANCE` | `0.05` | Allowed paise difference in the checks. |

## 6. Safety checks (a bill is sent only if all pass)

1. Glass total + freight = ERP **Ass. Value**.
2. Ass. Value + GST + round off = ERP **Net Amt**, and round off is less than Rs 1.
3. Not both IGST and CGST/SGST on one bill.
4. Every glass name has a Tally stock item.
5. Hole/Cutout total is not more than the glass total.
6. The voucher balances (debits = credits).
7. The bill was not sent before (`tally_sent_itemwise.db`).
8. Tally's reply must say CREATED 1, ERRORS 0, EXCEPTIONS 0. Otherwise it counts as failed.

A failed bill never stops the others. Example on 03 Oct data: bill 6647 is stopped because item total and Ass. Value differ by Rs 50 (asked accounts why).

## 7. How the SQM area is decided

The ERP header file has the billed **Area** of the whole bill (the real ERP number). The program always makes the Tally total SQM equal to it.

| Case | Area of each glass line | Label in output |
|---|---|---|
| Bill has one glass type | The bill Area | `area exact` |
| Several glass types, each PI has only one glass type | Total Area of that glass's own PIs (from the PI report) | `area exact-PI` |
| Several glass types, PIs mixed | Bill Area shared out by piece sizes. **An estimate.** The bill narration says "glass-wise area estimated". | `area estimated` |

On 03 Oct data: 20 exact, 5 exact-PI, 15 estimated. The real glass-wise area for the estimated bills is not in any ERP report; it is probably in the ERP database (needs permission, see section 12).

## 8. How to run (dummy company)

Tally must be open with TEST COMPANY loaded. Open PowerShell in the project folder.

| What | Command |
|---|---|
| Check everything, send nothing | `python tally_itemwise.py --dry-run --test-date 01-10-2026` |
| Send all bills | `python tally_itemwise.py --test-date 01-10-2026 --create-parties` |
| One bill only | `python tally_itemwise.py --only 6651 --test-date 01-10-2026` |
| Show the XML of one bill | add `--show 6651` |
| Or just | double-click `Tally_Sync.bat` |

Options: `--dry-run` (send nothing), `--only N` (one bill), `--show N` (print XML), `--test-date DD-MM-YYYY` (dummy company only), `--create-parties` (create missing party ledgers; **dummy company only**), `--fallback-item NAME` (use one item for unmapped glass; **testing only**), `--pi FILE` (PI report file).

Expected on 03 Oct data: `40 created, 0 skipped, 1 failed`, total Rs 14,69,335. Trial Balance: Current Assets 14,69,335.00.

## 9. Common tasks

**Send one bill again** (dummy company only):
1. Tally: Day Book, select the bill, Alt+D, Yes.
2. Remove it from the program's memory:
   ```
   python -c "import sqlite3; c=sqlite3.connect('tally_sent_itemwise.db'); n=c.execute('DELETE FROM sent WHERE invoice_no=?', ('6651',)).rowcount; c.commit(); print('removed', n)"
   ```
3. Run with `--only 6651`.

**Start again from clean** (dummy company only):
1. Tally: Day Book, F2 date 1-10-2026, Ctrl+Space (select all), Alt+D, Yes.
2. Rename `tally_sent_itemwise.db` (for example add `_OLD`).
3. Send all bills again.

**Backup / restore Tally:** Alt+Y (Data) > Backup, path `C:\TallyBackup`. Restore: Alt+Y > Restore.

**New anonymised PI file:**
```
python anonymise_exports.py "C:\erp_data\PI_21Aug_03Oct.xls"
Copy-Item .\anonymised\PI_21Aug_03Oct_ANON.csv .
```
When anonymising several files together, always give them in the same order (header, item-wise, PI) so the TEST PARTY numbers stay the same.

## 10. Problems I met, and the fix

| Problem | Cause | Fix |
|---|---|---|
| Tally replies EXCEPTIONS 1, no reason | Ledger lines used `ALLLEDGERENTRIES.LIST` in an item invoice | Use `LEDGERENTRIES.LIST` (setting `LEDGER_TAG`) |
| Minus round off shows "0.48" and wrong total on the invoice screen (books were right) | `ISDEEMEDPOSITIVE` was Yes on that line | Only the party line is Yes; all other lines No; the minus sign goes on AMOUNT. Learned from Tally's own export of a hand-made bill. |
| "Ledger does not exist" | Party or ledger missing in Tally | Create it in Tally, run again (only that bill is sent) |
| "no Tally stock item for ..." | Glass name not in `glass_name_mapping.csv` | Add the row (and the item in Tally) |
| Tally ignores the voucher number | Numbering is Automatic | Sales voucher type: Method of numbering = Manual |
| "Prevent creating duplicate Voucher Nos." cannot be selected | Locked on my laptop (probably Educational mode) | Program's own duplicate check is used; ask accounts about their Tally |
| `No module named 'xlrd'` | Add-on missing | `python -m pip install xlrd` |
| Python one-liner fails with "unterminated string" | PowerShell quotes | Use only single quotes inside the double quotes |
| New bill not where expected in Day Book | Day Book lists bills in the order they were made | Scroll to the bottom |
| "Negative stock" warning | No purchases in the dummy company | Only a warning in the dummy company |
| `can't open file ...py` | Downloaded file still in Downloads | Move it into the project folder |

## 11. Going live (only with approval)

**What changes from the demo**

| Demo | Live |
|---|---|
| TEST COMPANY | Accounts' real company name |
| Anonymised CSV files | Real ERP data (daily `.xls` export, or a read-only copy of the ERP database). Not anonymised: party names must match Tally. |
| `--test-date`, `--fallback-item`, `--create-parties` | Never used |
| Dummy ledger and item names | Exact names from accounts' Tally |
| Old `tally_sent_itemwise.db` | New, empty one |

**Still to build before live**

1. **Start point:** "send only bills after invoice number X", so bills already typed by hand are not sent again.
2. **Party name check:** a report of ERP customer names that do not match a Tally ledger, plus a party mapping file.
3. **Settings in a separate file** (`settings.ini`), so the code is never edited on the main PC.
4. **Read `.xls` directly** (no CSV step).
5. A rule for bills **changed or cancelled in the ERP after being sent** (the program only creates bills, it never changes them).

**Steps**

1. Get exact names from accounts (ledgers, stock items, voucher type, company).
2. Make the changes above; dry run.
3. Restore a **backup copy** of the real Tally on a test PC; send a few real days; compare with the bills accounts typed.
4. Supervised run on the accounts PC, into a **separate copy company**.
5. One week side by side: accounts keep typing in the real company, the program posts into the copy; compare totals daily.
6. Go-live day (with manager approval): Tally backup, set the start invoice number, stop typing in Tally, turn on the scheduler.
7. First weeks: check `run_log.txt` and the failed list daily. If something goes wrong, restore that day's backup.

**What to copy to the main PC** (new folder, for example `C:\TallyAuto`, not in OneDrive):
`tally_itemwise.py`, `glass_name_mapping.csv` (real names), `Tally_Sync.bat` (live line), `settings.ini`, party mapping file.
**Do not copy:** anonymised CSVs, `anonymised\`, any `tally_sent_itemwise*.db`, `failed_*`, `run_log.txt`, `create_stock_items.py`, `anonymise_exports.py`, old test scripts.

**Main PC needs:** Python + `requests` + `xlrd` (or one `.exe` built with PyInstaller if Python cannot be installed), and Tally server on port 9000.

**Running by itself (Windows Task Scheduler):** Start > Task Scheduler > Create Basic Task > name "Tally Sync" > Daily > time (for example 6:00 PM) > Start a program > choose `C:\TallyAuto\Tally_Sync.bat` > Finish. Then open the task's Properties and set "Start in" to `C:\TallyAuto`. The PC must be on and Tally open at that time. Start with once a day; later every 30 minutes is possible, picking only bills older than about 1 hour (so a bill still being edited is not sent).

**In `Tally_Sync.bat` for live:** remove the line with `--test-date` and use `python tally_itemwise.py`. For scheduled runs, also remove `pause` at the end so the window closes by itself.

**Rules for accounts after go-live**
- Make bills in the ERP as usual.
- Do **not** type the same bills in Tally (they would be doubled).
- Tally must be open at the scheduled time.
- Check the run summary for about 5 minutes a day; create a missing party ledger or stock item when a bill is stopped, and it will go in the next run.

## 12. Open items

1. Exact names in accounts' Tally: Sales ledger (local and other state?), Output taxes, Freight, Round Off, voucher type, stock items; "Cutout" or "CutOut".
2. Where the piece count is written on a Tally bill line.
3. Bill 6647: why the Rs 50 difference.
4. Accounts' Tally version; is "Prevent creating duplicate Voucher Nos." on; is stock maintained (negative stock).
5. Who creates new customers and new glass types in Tally.
6. What to do when an ERP bill is changed or cancelled after posting.
7. How often the program should run (once a day or every 30 minutes).
8. **Permissions from my manager:** a read-only backup copy of the ERP database (for the real glass-wise area), running the program on the accounts PC, contact with the vendor (Recursive Infotech) if needed, and who approves go-live.

## 13. Rules I follow

- Never write to the live ERP. The ERP is only read.
- Never touch the live Tally without a backup copy, supervised testing and my manager's approval.
- Real customer names and GST numbers stay out of chats and out of OneDrive. Only anonymised files are shared.
- In the ERP never use Year Ending, Admin Edit or Option screens.
