#!/usr/bin/env bash
# Per-session start for a HomeFront cloud session. Lays the repos out exactly like the
# Mac (this checkout = VBToCSharp/HomeFront, the app at MobileSource/HomeFront, FlexKit
# beside this checkout, so HomeFront.csproj's ..\..\..\FlexKit resolves), starts SQL
# Server with HOMEFRONTSQL, and gives the app its SA password through user-secrets.
# Writes /tmp/hf-cloud-ready when done; log in /tmp/hf-cloud-start.log (via the hook).
# Refuses to run outside a cloud sandbox so it can never touch the Mac's trees.
set -uo pipefail
if [ "${HF_CLOUD:-}" != 1 ] && [ "${CLAUDE_CODE_REMOTE:-}" != true ]; then
    echo "session-start.sh: not a cloud session (set HF_CLOUD=1 in the cloud environment); nothing done."
    exit 0
fi
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUTER="$(cd "$HERE/../.." && pwd)"
APP="$OUTER/MobileSource/HomeFront"
FLEXKIT="$(cd "$OUTER/.." && pwd)/FlexKit"
CACHE_ROOT=/opt/hf-cache/VBToCSharp
READY=/tmp/hf-cloud-ready
rm -f "$READY"
status=0

place() { # <dir> <cache dir> <url> <branch>
    local dir=$1 cached=$2 url=$3 branch=$4
    if [ ! -d "$dir/.git" ]; then
        if [ -d "$cached/.git" ]; then git clone -q "$cached" "$dir"; else git clone -q --branch "$branch" "$url" "$dir"; fi \
            || { echo "ERROR: could not place $url at $dir"; status=1; return; }
        git -C "$dir" remote set-url origin "$url"
    fi
    if git -C "$dir" fetch -q origin "$branch"; then
        git -C "$dir" checkout -q "$branch" 2>/dev/null || git -C "$dir" checkout -q -b "$branch" "origin/$branch"
        git -C "$dir" branch -q --set-upstream-to "origin/$branch" "$branch"
        git -C "$dir" merge -q --ff-only "origin/$branch" || echo "WARN: $dir has local commits; not fast-forwarded"
    else
        echo "WARN: could not fetch $url (GitHub access?); using the cached copy"
    fi
}
place "$APP" "$CACHE_ROOT/HomeFront/MobileSource/HomeFront" https://github.com/wadoodachaudhary/HomeFront.git main
place "$FLEXKIT" "$CACHE_ROOT/FlexKit" https://github.com/wadoodachaudhary/FlexKit.git telerik-parity-20260904

bash "$HERE/db.sh" || status=1

# The app reads Database:Password (DbWrapperSqlServer); give it this sandbox's SA password.
SECRETS_ID="$(sed -n 's:.*<UserSecretsId>\(.*\)</UserSecretsId>.*:\1:p' "$APP/HomeFront.csproj" 2>/dev/null | head -1)"
if [ -n "$SECRETS_ID" ] && [ -s "$HOME/.hf-sa" ]; then
    mkdir -p "$HOME/.microsoft/usersecrets/$SECRETS_ID"
    (umask 077; python3 -c 'import json,sys; print(json.dumps({"Database:Password": open(sys.argv[1]).read().strip()}))' \
        "$HOME/.hf-sa" > "$HOME/.microsoft/usersecrets/$SECRETS_ID/secrets.json")
else
    echo "WARN: app user-secrets not written (no UserSecretsId or SA password)"; status=1
fi

if [ $status -eq 0 ]; then
    date > "$READY"
    echo "HomeFront cloud session ready: app $APP, FlexKit $FLEXKIT, HOMEFRONTSQL on localhost,1433."
else
    echo "HomeFront cloud session start finished WITH ERRORS (see above)."
fi
exit $status
