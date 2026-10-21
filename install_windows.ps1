$ErrorActionPreference = 'Stop'

$InstallDir = Join-Path $env:LOCALAPPDATA 'MailBriefingAssistant'
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

Get-ChildItem -LiteralPath $PSScriptRoot -Force |
    Where-Object { $_.Name -notin @('.venv', '__pycache__', '.git') } |
    ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $InstallDir -Recurse -Force
    }

$Py = Get-Command py -ErrorAction SilentlyContinue
if (-not $Py) {
    throw 'Python Launcher not found. Install Python 3.11+ and enable Add Python to PATH, then run this installer again.'
}

$VenvDir = Join-Path $InstallDir '.venv'
$Python = Join-Path $VenvDir 'Scripts\python.exe'
$PythonW = Join-Path $VenvDir 'Scripts\pythonw.exe'
$Script = Join-Path $InstallDir 'mail_briefing.py'
$Requirements = Join-Path $InstallDir 'requirements.txt'

if (-not (Test-Path -LiteralPath $Python)) {
    & py -3 -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) {
        throw 'Could not create the Python virtual environment.'
    }
}
if (-not (Test-Path -LiteralPath $Python)) {
    throw 'The virtual environment was not created. Check your Python installation.'
}

& $Python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    throw 'Could not upgrade pip.'
}
& $Python -m pip install -r $Requirements
if ($LASTEXITCODE -ne 0) {
    throw 'Could not install the requirements.'
}

$ActionArgs = '"' + $Script + '" --background'
$Action = New-ScheduledTaskAction -Execute $PythonW -Argument $ActionArgs -WorkingDirectory $InstallDir
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$Trigger = New-ScheduledTaskTrigger -AtLogOn -User $CurrentUser
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero)

Register-ScheduledTask -TaskName 'MailBriefingAssistant' -Action $Action -Trigger $Trigger -Settings $Settings -Description 'Mail briefing desktop widget; starts at user logon.' -Force | Out-Null

$Programs = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'
New-Item -ItemType Directory -Force -Path $Programs | Out-Null
$ShortcutPath = Join-Path $Programs '邮件晨报助手.lnk'
$LaunchPath = Join-Path $InstallDir 'launch.bat'
$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $LaunchPath
$Shortcut.WorkingDirectory = $InstallDir
$Shortcut.Save()

Write-Host "Installation complete. Program folder: $InstallDir"
Write-Host 'Open Mail Briefing Assistant from the Start menu and configure your mail accounts.'
Start-Process -FilePath $LaunchPath
