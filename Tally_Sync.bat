@echo off
REM ============================================================
REM  Tally_Sync.bat  -  double-click to send ERP bills to Tally
REM  Needs: Tally open with the company loaded, Python installed.
REM  1. Sends the bills whose glass is in the mapping, shows AAJ KA HISAAB.
REM  2. Then asks for the bills entered BY HAND (Enter = skip for now).
REM  Every run is recorded in run_log.txt; details in failed_itemwise.csv.
REM  Safe to run again: bills already in Tally are never sent twice.
REM ============================================================
title Tally Sync
cd /d "%~dp0"
cls

REM --- DEMO / TEST COMPANY (Educational mode needs a 1st/2nd/last-day date) ---
python tally_itemwise.py --simple --test-date 01-10-2026

REM --- LIVE (later, only after approval): remove the line above and use ---
REM python tally_itemwise.py --simple

:manual
echo.
echo   Haath se chadhaye bills ke number yahan paste karo (right-click), phir Enter.
echo   Abhi nahi chadhaye? Sirf Enter dabao - baad mein yeh button dobara chala lena.
echo.
set "BILLS="
set /p "BILLS=  Bill numbers: "
if not defined BILLS goto done
python tally_itemwise.py --mark-manual %BILLS%
goto manual

:done
echo.
echo   Pichhle 5 runs ka record (run_log.txt):
powershell -NoProfile -Command "Get-Content run_log.txt -Tail 5"
echo.
echo   Window band karne ke liye koi bhi key dabao.
pause >nul
