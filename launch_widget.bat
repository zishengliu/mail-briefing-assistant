@echo off
setlocal
set "APP=%LOCALAPPDATA%\MailBriefingAssistant"
if exist "%APP%\.venv\Scripts\pythonw.exe" (
  start "" "%APP%\.venv\Scripts\pythonw.exe" "%APP%\mail_briefing.py" --widget
) else (
  start "" pyw -3 "%~dp0mail_briefing.py" --widget
)
endlocal
