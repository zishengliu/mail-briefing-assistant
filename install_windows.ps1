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
$Shell = New-Object -ComObject WScript.Shell
$WidgetLaunchPath = Join-Path $InstallDir 'launch_widget.bat'
$SettingsLaunchPath = Join-Path $InstallDir 'launch.bat'

$WidgetShortcut = $Shell.CreateShortcut((Join-Path $Programs '邮件助手.lnk'))
$WidgetShortcut.TargetPath = $WidgetLaunchPath
$WidgetShortcut.WorkingDirectory = $InstallDir
$WidgetShortcut.Description = '打开邮件助手桌面悬浮小插件'
$WidgetShortcut.Save()

$SettingsShortcut = $Shell.CreateShortcut((Join-Path $Programs '邮件助手 - 设置.lnk'))
$SettingsShortcut.TargetPath = $SettingsLaunchPath
$SettingsShortcut.WorkingDirectory = $InstallDir
$SettingsShortcut.Description = '打开邮件助手的邮箱与定时设置'
$SettingsShortcut.Save()

# Remove the prior shortcut name after successfully creating the new shortcuts.
$OldShortcutPath = Join-Path $Programs '邮件晨报助手.lnk'
if (Test-Path -LiteralPath $OldShortcutPath) {
    Remove-Item -LiteralPath $OldShortcutPath -Force
}

Write-Host "Installation complete. Program folder: $InstallDir"
Write-Host '从开始菜单打开“邮件助手”启动桌面悬浮小插件。'
Write-Host '需要配置邮箱时，打开“邮件助手 - 设置”。'
Start-Process -FilePath $WidgetLaunchPath
