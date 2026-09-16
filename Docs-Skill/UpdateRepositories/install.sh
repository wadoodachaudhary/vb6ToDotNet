#!/bin/bash
# install.sh — make /update-repositories available to EVERY Claude account on this Mac.
#
# Claude Code discovers personal skills only under ~/.claude/skills/<name>/, and ~/.claude is keyed to the
# macOS user — not the Anthropic account — so one install serves both the Wadood (wadood@gmail.com) and
# Innovatix (wchaudhary@innovatixinc.com) accounts. The directory name IS the slash command and must be
# kebab-case: update-repositories.
#
# Always installs FROM the canonical, git-tracked copy below, whatever copy you run it from. The installed
# copy is made READ-ONLY: an agent that loads the skill from ~/.claude/skills and edits a reference file
# there gets a permission error instead of a change the next install silently erases. Edit the canonical
# copy, commit it in the HomeFront repo, then run this. It is a real copy, not a symlink: symlinked skill
# directories are not documented as supported.
set -euo pipefail
SRC=/Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories
DST="$HOME/.claude/skills/update-repositories"
HERE="$(cd "$(dirname "$0")" && pwd -P)"

if [ "$HERE" = "$(cd "$DST" 2>/dev/null && pwd -P)" ]; then
  echo "✗ this is the INSTALLED copy. Edit and install from the canonical source instead:" >&2
  echo "    bash $SRC/install.sh" >&2
  exit 1
fi
[ -f "$SRC/SKILL.md" ] || { echo "✗ canonical source missing: $SRC/SKILL.md" >&2; exit 1; }

mkdir -p "$DST"
chmod -R u+w "$DST"                       # previous install left it read-only
rsync -a --delete --exclude='.DS_Store' "$SRC/" "$DST/"
find "$DST" -type f -exec chmod a-w {} +
find "$DST" -type f -name '*.sh' -exec chmod a+x {} +
if diff -rq -x .DS_Store "$SRC" "$DST" >/dev/null; then
  echo "✓ Installed $DST (read-only) — invoke with /update-repositories"
else
  echo "✗ installed copy differs from $SRC after rsync:" >&2
  diff -rq -x .DS_Store "$SRC" "$DST" >&2 || true
  exit 1
fi
