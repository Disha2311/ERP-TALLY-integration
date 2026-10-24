@echo off
REM ============================================================
REM  Accept_Change.bat  -  after correcting a CHANGED bill in Tally by hand
REM  (see changed_after_sent.csv). Never sends anything to Tally.
REM ============================================================
cd /d "%~dp0"
set /p BILLS=Bill number(s) you corrected in Tally (for example 6651, or ALL): 
python accept_change.py %BILLS%
echo.
pause
