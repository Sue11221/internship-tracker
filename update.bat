@echo off
REM Daily refresh. Double-click to run by hand, or the scheduled task runs it as:  update.bat quiet
REM   1. fetch every company board          (run.py update)
REM   2. rebuild the four-group shortlists   (shortlist.py)
REM   3. email whatever is new               (send_email.py, needs email_config.json)
REM   4. optional Telegram/Discord alerts    (run.py notify, no-op unless configured)
REM   5. open the dashboard                  (skipped in quiet mode)
setlocal
cd /d "%~dp0"
set PY=.venv\Scripts\python.exe
if not exist logs mkdir logs
set LOG=logs\last_run.log

if /i "%~1"=="quiet" (
    call :run > "%LOG%" 2>&1
    exit /b %ERRORLEVEL%
)
call :run
echo.
echo Opening the dashboard...
start "" "docs\index.html"
exit /b %ERRORLEVEL%

:run
echo ===== %DATE% %TIME% =====
echo [1/4] Fetching internship postings from all company boards (takes ~20 minutes)...
"%PY%" run.py update
if errorlevel 1 (
    echo Update failed; shortlist and email skipped.
    exit /b 1
)
echo [2/4] Reading summer dates of new postings, then picking today's 20...
"%PY%" term_dates.py
"%PY%" shortlist.py
echo [3/4] Emailing new postings...
"%PY%" send_email.py
echo [4/4] Other alerts...
"%PY%" run.py notify
echo Done %DATE% %TIME%
exit /b 0
