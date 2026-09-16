#!/bin/bash
# install.sh — make /update-repositories available to EVERY Claude account on this Mac.
#
# Claude Code discovers personal skills only under ~/.claude/skills/<name>/, and ~/.claude is
# keyed to the macOS user — not the Anthropic account — so one install serves both the
# Wadood (wadood@gmail.com) and Innovatix (wchaudhary@innovatixinc.com) accounts.
# The directory name IS the slash command and must be kebab-case: update-repositories.
#
# This folder (HomeFront/Docs-Skill/UpdateRepositories) is the canonical, git-tracked source.
# The install is a real copy, not a symlink: symlinked skill directories are not documented as
# supported. Re-run this after editing anything here; preflight.sh warns when they drift.
set -euo pipefail
SRC="$(cd "$(dirname "$0")" && pwd)"
DST="$HOME/.claude/skills/update-repositories"
mkdir -p "$DST"
rsync -a --delete --exclude='.DS_Store' "$SRC/" "$DST/"
chmod +x "$DST"/install.sh "$DST"/scripts/*.sh
diff -rq "$SRC" "$DST" >/dev/null && echo "Installed: $DST  (invoke with /update-repositories)"
