@echo off
REM ============================================================
REM  Forget_Bill.bat  -  send a bill to Tally again
REM  FIRST delete the bill in Tally (Day Book, select, Alt+D).
REM ============================================================
cd /d "%~dp0"
set /p BILLS=Bill number(s) to send again (for example 6651, or ALL): 
python forget_bill.py %BILLS%
echo.
pause
