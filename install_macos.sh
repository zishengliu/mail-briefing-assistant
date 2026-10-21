#!/bin/bash
set -euo pipefail

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 was not found. Install Python 3.11+ from python.org, including Tcl/Tk support, then rerun this installer."
  exit 1
fi

if ! python3 -c 'import tkinter, keyring' >/dev/null 2>&1; then
  echo "Python is missing Tkinter or keyring. Install a Python 3.11+ build with Tcl/Tk support, then rerun this installer."
  echo "A python.org macOS installer is recommended."
  exit 1
fi

SOURCE_DIR="$(cd "$(dirname "$BASH_SOURCE")" && pwd)"
INSTALL_DIR="$HOME/Library/Application Support/MailAssistant"
APP_BUNDLE="$HOME/Applications/Mail Assistant.app"
AGENT_DIR="$HOME/Library/LaunchAgents"
AGENT_PATH="$AGENT_DIR/com.zishengliu.mailassistant.plist"

mkdir -p "$INSTALL_DIR" "$HOME/Applications" "$APP_BUNDLE/Contents/MacOS" "$AGENT_DIR"
cp -f "$SOURCE_DIR/mail_briefing.py" "$INSTALL_DIR/mail_briefing.py"
cp -f "$SOURCE_DIR/requirements.txt" "$INSTALL_DIR/requirements.txt"

if [ ! -x "$INSTALL_DIR/.venv/bin/python" ]; then
  python3 -m venv "$INSTALL_DIR/.venv"
fi
"$INSTALL_DIR/.venv/bin/python" -m pip install --upgrade pip
"$INSTALL_DIR/.venv/bin/python" -m pip install -r "$INSTALL_DIR/requirements.txt"

# Generate a clickable app bundle and LaunchAgent. Unicode display text is built from code points.
python3 - "$APP_BUNDLE" "$INSTALL_DIR" "$AGENT_PATH" <<'PY'
import os
import plistlib
import shlex
import sys

app_bundle, install_dir, agent_path = sys.argv[1:4]
display_name = "".join(chr(n) for n in [0x90AE, 0x4EF6, 0x52A9, 0x624B])
contents = os.path.join(app_bundle, "Contents")
macos_dir = os.path.join(contents, "MacOS")
os.makedirs(macos_dir, exist_ok=True)

python_path = os.path.join(install_dir, ".venv", "bin", "python")
script_path = os.path.join(install_dir, "mail_briefing.py")
launcher_path = os.path.join(macos_dir, "MailAssistant")
launcher = "#!/bin/bash\ncd " + shlex.quote(install_dir) + "\nexec " + shlex.quote(python_path) + " " + shlex.quote(script_path) + " --background\n"
with open(launcher_path, "w", encoding="utf-8") as f:
    f.write(launcher)
os.chmod(launcher_path, 0o755)

info = {
    "CFBundleName": display_name,
    "CFBundleDisplayName": display_name,
    "CFBundleExecutable": "MailAssistant",
    "CFBundleIdentifier": "com.zishengliu.mailassistant",
    "CFBundlePackageType": "APPL",
    "CFBundleVersion": "1.0",
    "CFBundleShortVersionString": "1.0",
    "NSHighResolutionCapable": True,
    "LSUIElement": True,
}
with open(os.path.join(contents, "Info.plist"), "wb") as f:
    plistlib.dump(info, f)

agent = {
    "Label": "com.zishengliu.mailassistant",
    "ProgramArguments": ["/usr/bin/open", "-a", app_bundle],
    "RunAtLoad": True,
    "KeepAlive": False,
}
os.makedirs(os.path.dirname(agent_path), exist_ok=True)
with open(agent_path, "wb") as f:
    plistlib.dump(agent, f)
PY

launchctl bootout "gui/$(id -u)" "$AGENT_PATH" >/dev/null 2>&1 || true
launchctl bootstrap "gui/$(id -u)" "$AGENT_PATH" >/dev/null 2>&1 || echo "Login auto-start could not be registered. You can still open Mail Assistant from Applications."
open "$APP_BUNDLE"

echo "Mail Assistant has been installed."
echo "Open the Mail Assistant app from ~/Applications; it is configured to start at login."
echo "App data and Keychain credentials remain in your local macOS user account."
