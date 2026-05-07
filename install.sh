#!/usr/bin/env bash
#
# Mouseless Mac Excel installer.
#
# Writes ~/.hammerspoon/init.lua so Hammerspoon loads the plugin from
# this clone. Idempotent and safe to re-run: any existing init.lua is
# backed up to init.lua.backup-YYYYMMDD-HHMMSS.
#
# Run from inside the clone:
#
#   ./install.sh
#
# Or from anywhere:
#
#   /absolute/path/to/clone/install.sh

set -euo pipefail

# Resolve this script's directory (the repo root) regardless of where
# it was invoked from or whether it was run via a symlink.
SCRIPT_PATH="${BASH_SOURCE[0]}"
while [ -L "$SCRIPT_PATH" ]; do
  SCRIPT_DIR="$(cd "$(dirname "$SCRIPT_PATH")" && pwd)"
  SCRIPT_PATH="$(readlink "$SCRIPT_PATH")"
  case "$SCRIPT_PATH" in
    /*) ;;
    *) SCRIPT_PATH="$SCRIPT_DIR/$SCRIPT_PATH" ;;
  esac
done
PROJECT_DIR="$(cd "$(dirname "$SCRIPT_PATH")" && pwd)"

TEMPLATE="$PROJECT_DIR/hammerspoon-bootstrap.lua"
HS_DIR="$HOME/.hammerspoon"
HS_INIT="$HS_DIR/init.lua"

if [ ! -f "$TEMPLATE" ]; then
  echo "Error: template not found at $TEMPLATE" >&2
  echo "Are you running install.sh from inside the cloned repo?" >&2
  exit 1
fi

# Sanity-check: is Hammerspoon itself installed?
if [ ! -d "/Applications/Hammerspoon.app" ]; then
  cat <<'EOF' >&2
Warning: /Applications/Hammerspoon.app is not present.

Install Hammerspoon first, e.g.:
  brew install --cask hammerspoon

Continuing anyway so this script is idempotent.
EOF
fi

mkdir -p "$HS_DIR"

if [ -f "$HS_INIT" ]; then
  BACKUP="$HS_INIT.backup-$(date +%Y%m%d-%H%M%S)"
  cp "$HS_INIT" "$BACKUP"
  echo "Backed up existing $HS_INIT -> $BACKUP"
fi

# Substitute the project path. Escape the few characters sed treats
# specially in a replacement string.
ESCAPED_PATH=$(printf '%s' "$PROJECT_DIR" | sed -e 's/[\\|&]/\\&/g')
sed "s|__MOUSELESS_EXCEL_PATH__|$ESCAPED_PATH|g" "$TEMPLATE" > "$HS_INIT"

echo "Wrote $HS_INIT pointing at $PROJECT_DIR"

cat <<EOF

Next steps:

  1. Launch Hammerspoon (or reload if it is already running):
       open -a Hammerspoon
     If it was already running, you can also click the menu-bar icon ->
     Reload Config.

  2. Grant Accessibility permission when macOS prompts:
       System Settings -> Privacy & Security -> Accessibility
       -> tick Hammerspoon
     If you grant the permission while Hammerspoon is already running,
     fully quit and relaunch Hammerspoon for it to take effect.

  3. Open Microsoft Excel. Tap and release the Option key alone.
     A small "Excel menu" indicator should appear.

  4. With something on the clipboard, type  e  s  v  to run
     Paste Special > Values.

Logs from the plugin go to /tmp/mouseless-mac-excel.log.
EOF
