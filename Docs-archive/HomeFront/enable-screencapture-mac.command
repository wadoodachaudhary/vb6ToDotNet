#!/bin/bash
# ============================================================================
#  HomeFront / Precision Builder — one-click Chrome screen-capture setup (macOS)
#
#  Double-click this file to run it.
#
#  Note: macOS Chrome ignores `defaults write` for capture policies — it only
#  honors them from a configuration PROFILE. So this opens the bundled profile
#  for you to install, then relaunches Chrome. After this, the Feedback capture
#  (F8 / Shift+F8) runs with no "Allow ... to see this tab?" prompt.
#
#  If your app origin is NOT http://localhost:5065, edit that string inside
#  HomeFront-Chrome-ScreenCapture.mobileconfig (same folder) before running.
# ============================================================================
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
PROFILE="$DIR/HomeFront-Chrome-ScreenCapture.mobileconfig"

if [ ! -f "$PROFILE" ]; then
    echo "ERROR: profile not found next to this script:"
    echo "  $PROFILE"
    echo "Keep enable-screencapture-mac.command and HomeFront-Chrome-ScreenCapture.mobileconfig together."
    exit 1
fi

echo "Opening the Chrome screen-capture profile for installation..."
open "$PROFILE"

cat <<'EOF'

  ── Finish installing the profile ───────────────────────────────────────────
   System Settings → General → Device Management   (older macOS: Profiles)
     → "HomeFront — Chrome Screen Capture" → Install   (enter your password;
     an "unsigned profile" warning is expected for a local profile).
  ────────────────────────────────────────────────────────────────────────────

EOF

read -r -p "Press RETURN here AFTER the profile is installed to relaunch Chrome... " _

echo "Relaunching Google Chrome..."
osascript -e 'quit app "Google Chrome"' >/dev/null 2>&1 || true
sleep 2
open -a "Google Chrome" || true

cat <<'EOF'

  Done. Verify it took effect:
    1. In Chrome open  chrome://policy
    2. Click "Reload policies"
    3. Confirm TabCaptureAllowedByOrigins (and SameOriginTabCaptureAllowedByOrigins)
       list your app origin.

  Then open the app and press F8 — it should capture with no prompt.
EOF
