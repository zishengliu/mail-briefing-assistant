@echo off
setlocal
set "APP=%LOCALAPPDATA%\MailBriefingAssistant"
if exist "%APP%\.venv\Scripts\python.exe" (
  "%APP%\.venv\Scripts\python.exe" "%APP%\mail_briefing.py"
) else (
  py -3 "%~dp0mail_briefing.py"
)
endlocal
