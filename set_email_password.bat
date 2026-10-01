@echo off
REM One-time: store the Gmail app password in Windows Credential Manager, then send a test email.
cd /d "%~dp0"
.venv\Scripts\python.exe send_email.py --set-password
echo.
pause
