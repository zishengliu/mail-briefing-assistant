$ErrorActionPreference = 'Stop'

# Use Unicode code points for localized shortcut names so Windows PowerShell 5.1
# can parse this ASCII-only script even when the file is downloaded without a BOM.
function ConvertFrom-CodePoints([int[]]$CodePoints) {
    return -join ($CodePoints | ForEach-Object { [char]$_ })
}

$AppName = ConvertFrom-CodePoints @(0x90AE, 0x4EF6, 0x52A9, 0x624B)
$SettingsName = ConvertFrom-CodePoints @(0x8BBE, 0x7F6E)
$OldAppName = ConvertFrom-CodePoints @(0x90AE, 0x4EF6, 0x6668, 0x62A5, 0x52A9, 0x624B)
$DescriptionWidget = 'Open the Mail Assistant desktop widget'
$DescriptionSettings = 'Open Mail Assistant account and schedule settings'

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
$Shell = New-Object -ComObject WScript.Shell
$WidgetLaunchPath = Join-Path $InstallDir 'launch_widget.bat'
$SettingsLaunchPath = Join-Path $InstallDir 'launch.bat'

$WidgetShortcutPath = Join-Path $Programs ($AppName + '.lnk')
$WidgetShortcut = $Shell.CreateShortcut($WidgetShortcutPath)
$WidgetShortcut.TargetPath = $WidgetLaunchPath
$WidgetShortcut.WorkingDirectory = $InstallDir
$WidgetShortcut.Description = $DescriptionWidget
$WidgetShortcut.Save()

$SettingsShortcutPath = Join-Path $Programs ($AppName + ' - ' + $SettingsName + '.lnk')
$SettingsShortcut = $Shell.CreateShortcut($SettingsShortcutPath)
$SettingsShortcut.TargetPath = $SettingsLaunchPath
$SettingsShortcut.WorkingDirectory = $InstallDir
$SettingsShortcut.Description = $DescriptionSettings
$SettingsShortcut.Save()

# Remove the previous shortcut name after the new shortcuts are created.
$OldShortcutPath = Join-Path $Programs ($OldAppName + '.lnk')
if (Test-Path -LiteralPath $OldShortcutPath) {
    Remove-Item -LiteralPath $OldShortcutPath -Force
}

Write-Host "Installation complete. Program folder: $InstallDir"
Write-Host 'Open Mail Assistant from the Start menu to launch the desktop widget.'
Write-Host 'Open Mail Assistant - Settings to configure email accounts and schedule.'
Start-Process -FilePath $WidgetLaunchPath
