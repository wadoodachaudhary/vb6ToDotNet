#!/usr/bin/env bash
# SessionStart hook: in a cloud session, run session-start.sh in the background and tell
# Claude where to look. Silent no-op everywhere else (the Mac).
[ "${HF_CLOUD:-}" = 1 ] || [ "${CLAUDE_CODE_REMOTE:-}" = true ] || exit 0
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [ ! -f /tmp/hf-cloud-ready ] && ! pgrep -f "tools/cloud/session-start.sh" >/dev/null 2>&1; then
    nohup bash "$HERE/session-start.sh" >/tmp/hf-cloud-start.log 2>&1 &
fi
echo "HomeFront cloud session: repo layout, SQL Server and HOMEFRONTSQL are being prepared in the background by tools/cloud/session-start.sh. Before building or querying, wait until /tmp/hf-cloud-ready exists (progress and errors: /tmp/hf-cloud-start.log; rerun it in the foreground if it failed). Query the DB with: bash tools/cloud/sql.sh \"SELECT ...\". See tools/cloud/README.md."
exit 0
