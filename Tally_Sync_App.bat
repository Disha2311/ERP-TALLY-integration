@echo off
REM Tally_Sync_App.bat - opens the Tally Sync window (no black window).
REM Tally_Sync.bat stays as the backup way of doing the same thing.
cd /d "%~dp0"

REM --- DEMO / TEST COMPANY (Educational mode needs a 1st/2nd/last-day date) ---
start "" pythonw tally_sync_app.py --test-date 01-10-2026

REM --- LIVE (later, only after approval): remove the line above and use ---
REM start "" pythonw tally_sync_app.py
