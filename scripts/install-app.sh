#!/bin/zsh
# Builds Quire.app and installs it to /Applications. A running Quire is quit first: replacing an
# Electron app's files underneath a live instance corrupts that instance.
set -euo pipefail
cd "$(dirname "$0")/.."
npm run dist
if pgrep -xq Quire; then
  osascript -e 'quit app "Quire"' || true
  for _ in {1..40}; do pgrep -xq Quire || break; sleep 0.25; done
  if pgrep -xq Quire; then
    echo "Quire is still running (unsaved changes?). Quit it, then rerun." >&2
    exit 1
  fi
fi
rm -rf /Applications/Quire.app
cp -R dist-app/mac-arm64/Quire.app /Applications/
touch /Applications/Quire.app
/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f /Applications/Quire.app
# Pay macOS's first-launch scan now (headless, separate profile, never takes focus).
WARM=$(mktemp -d)
QUIRE_USER_DATA="$WARM" QUIRE_HARNESS="$PWD/scripts/harness/warmup.js" /Applications/Quire.app/Contents/MacOS/Quire >/dev/null 2>&1 || true
rm -rf "$WARM"
# Each build has a new CFBundleVersion, which makes macOS drop its cached icon; restart the
# Dock so it shows the current one.
killall Dock 2>/dev/null || true
echo "Installed /Applications/Quire.app"
