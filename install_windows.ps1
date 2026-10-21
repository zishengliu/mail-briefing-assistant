$ErrorActionPreference = 'Stop'
$InstallDir = Join-Path $env:LOCALAPPDATA 'MailBriefingAssistant'
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
Get-ChildItem -LiteralPath $PSScriptRoot -Force | Where-Object { $_.Name -notin @('.venv', '__pycache__') } | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $InstallDir -Recurse -Force
}
$Py = Get-Command py -ErrorAction SilentlyContinue
if (-not $Py) {
    throw '未找到 Python Launcher。请先安装 Python 3.11+，并勾选 Add Python to PATH，然后重新运行本脚本。'
}
if (-not (Test-Path (Join-Path $InstallDir '.venv\Scripts\python.exe'))) {
    & py -3 -m venv (Join-Path $InstallDir '.venv')
}
$Python = Join-Path $InstallDir '.venv\Scripts\python.exe'
& $Python -m pip install --upgrade pip
& $Python -m pip install -r (Join-Path $InstallDir 'requirements.txt')
$PythonW = Join-Path $InstallDir '.venv\Scripts\pythonw.exe'
$Script = Join-Path $InstallDir 'mail_briefing.py'
$Action = New-ScheduledTaskAction -Execute $PythonW -Argument ('"{0}" --background' -f $Script) -WorkingDirectory $InstallDir
$Trigger = New-ScheduledTaskTrigger -AtLogOn -User ("{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero)
Register-ScheduledTask -TaskName 'MailBriefingAssistant' -Action $Action -Trigger $Trigger -Settings $Settings -Description '本地邮件晨报助手：登录时启动桌面悬浮小插件，每天按设置时间汇总邮件。' -Force | Out-Null
$Shell = New-Object -ComObject WScript.Shell
$ShortcutPath = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\邮件晨报助手.lnk'
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = (Join-Path $InstallDir 'launch.bat')
$Shortcut.WorkingDirectory = $InstallDir
$Shortcut.Save()
Write-Host "安装完成。程序目录：$InstallDir"
Write-Host '请从开始菜单打开“邮件晨报助手”，配置邮箱并测试连接。'
Start-Process -FilePath (Join-Path $InstallDir 'launch.bat')
