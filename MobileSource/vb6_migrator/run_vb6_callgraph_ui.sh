#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
QT_SCRIPT="$SCRIPT_DIR/vb6_callgraph_ui_qt.py"
TK_SCRIPT="$SCRIPT_DIR/vb6_callgraph_ui.py"
LOG_DIR="$SCRIPT_DIR/.logs"
mkdir -p "$LOG_DIR"
ALLOW_TK_FALLBACK="${VB6_UI_ALLOW_TK_FALLBACK:-0}"

if [ "$#" -gt 0 ]; then
  echo "[launcher] CLI parameters are ignored. Using migrator_config.json defaults." >&2
fi

if [ "${VB6_UI_FORCE_TK:-0}" = "1" ]; then
  exec python3 "$TK_SCRIPT" --project-dir "$SCRIPT_DIR/projects/PrescionBuilder" --prefer-hfest true --ui true
fi

QT_LOG="$LOG_DIR/qt-launch.log"
{
  echo "=== Qt launcher start $(date '+%Y-%m-%d %H:%M:%S') ==="
  echo "Script: $0"
  echo "Python: $(command -v python3 || echo 'not-found')"
  python3 --version 2>&1 || true
  echo "Qt script: $QT_SCRIPT"
  echo "VB6_UI_FORCE_TK=${VB6_UI_FORCE_TK:-0}"
  echo "VB6_UI_ALLOW_TK_FALLBACK=$ALLOW_TK_FALLBACK"
  echo "Args: (none; using migrator_config.json)"
} >"$QT_LOG"
set +e
python3 "$QT_SCRIPT" >>"$QT_LOG" 2>&1
QT_EXIT=$?
set -e
echo "Qt exit code: $QT_EXIT" >>"$QT_LOG"
if [ "$QT_EXIT" -eq 0 ]; then
  exit 0
fi

if grep -Eq "No module named 'PySide6'|No module named 'PyQt6'" "$QT_LOG"; then
  if [ "$ALLOW_TK_FALLBACK" != "1" ]; then
    echo "Qt bindings are not available and Tk fallback is disabled." >&2
    echo "Install PySide6/PyQt6 or run with VB6_UI_ALLOW_TK_FALLBACK=1." >&2
    echo "Qt stderr log: $QT_LOG" >&2
    tail -n 80 "$QT_LOG" >&2 || true
    exit "$QT_EXIT"
  fi
  echo "Qt bindings are not available. Falling back to Tkinter UI..." >&2
  exec python3 "$TK_SCRIPT" --project-dir "$SCRIPT_DIR/projects/PrescionBuilder" --prefer-hfest true --ui true
fi

echo "Qt shell failed (exit $QT_EXIT). Tk fallback skipped so you can see/fix the Qt error." >&2
echo "Qt stderr log: $QT_LOG" >&2
tail -n 80 "$QT_LOG" >&2 || true
exit "$QT_EXIT"
