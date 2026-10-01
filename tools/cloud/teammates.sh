#!/usr/bin/env bash
# Brings teammates' merged GitLab work into the cloud trees, so cloud fixes start from
# current code. Cloud only; it can never push to GitLab (the token is read-only, and it
# only fetches).
#
#   GitLab hyphen-pb main  ->  app (MobileSource/HomeFront) main
#   GitLab flexkit main    ->  FlexKit telerik-parity-20260904
#
# The merge is deploy_to_repos.sh --pull's: only files the teammates' commits touched,
# 3-way merged per file. The base is the last "Deploy <repo> — <time>" commit on the GitLab
# main (the last tree the Mac shipped), or the last sync if newer; teammates' work is
# everything after it. Re-running is harmless: a change our tree already holds merges to itself.
#
# Branch-owned files (csproj/sln, appsettings, pipelines, ...) and upstream deletions are
# reported, never applied. A conflicting file is left untouched and listed. The result is
# committed with a "GitLab-Synced: <repo> <sha>" trailer, which the next run uses as its base
# when it is newer than the last deploy. It is pushed to the GitHub mirror only when nothing
# conflicted and the app builds; otherwise it stays a local commit for the session to finish.
#
# Needs HF_GITLAB_TOKEN (a GitLab token with read_repository) in the cloud environment.
# Usage: bash tools/cloud/teammates.sh [--no-push]
set -uo pipefail
if [ "${HF_CLOUD:-}" != 1 ] && [ "${CLAUDE_CODE_REMOTE:-}" != true ]; then
    echo "teammates.sh: not a cloud session; nothing done (the Mac uses deploy_to_repos.sh --pull)."
    exit 0
fi
PUSH=true; [ "${1:-}" = --no-push ] && PUSH=false
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUTER="$(cd "$HERE/../.." && pwd)"
APP="$OUTER/MobileSource/HomeFront"
FLEXKIT="$(cd "$OUTER/.." && pwd)/FlexKit"
MIRRORS="${HF_GITLAB_MIRRORS:-$HOME/.hf-gitlab}"
GITLAB="${HF_GITLAB_URL:-https://gitlab.innovatixinc.com/application-modernization}"
STATE="$MIRRORS/last-sync.txt"

if [ -z "${HF_GITLAB_TOKEN:-}" ]; then
    echo "WARN: HF_GITLAB_TOKEN is not set; teammates' GitLab work was NOT brought in."
    exit 3
fi
mkdir -p "$MIRRORS"
: > "$STATE"
attention=0
pushable=()

# The token goes in an Authorization header passed through the environment, so it is
# never written to a git config, a remote URL or a command line.
gitlab_fetch() { # <bare dir> <repo name>
    [ -d "$1" ] || git init -q --bare "$1"
    GIT_TERMINAL_PROMPT=0 GIT_CONFIG_COUNT=1 GIT_CONFIG_KEY_0=http.extraHeader \
        GIT_CONFIG_VALUE_0="Authorization: Basic $(printf 'oauth2:%s' "$HF_GITLAB_TOKEN" | base64 | tr -d '\n')" \
        git -C "$1" fetch -q --depth 400 "$GITLAB/$2.git" +main:refs/heads/main
}

sync_repo() { # <gitlab name> <target tree> <target branch> <pathspec...>
    local name=$1 tree=$2 branch=$3; shift 3
    local bare="$MIRRORS/$name.git" base head deploy synced f rc
    echo "== $name main -> $(basename "$tree") $branch"
    if ! gitlab_fetch "$bare" "$name"; then
        echo "ERROR: could not fetch $name from GitLab (token, scope or network)"; attention=1; return
    fi
    head=$(git -C "$bare" rev-parse main)
    deploy=$(git -C "$bare" log --first-parent --format='%H %s' main | awk -v n="$name" '$2 == "Deploy" && $3 == n { print $1; exit }')
    if [ -z "$deploy" ]; then
        echo "ERROR: no 'Deploy $name' commit in the last 400 on $name main; cannot tell teammates' work apart"
        attention=1; return
    fi
    # The newer of the last deploy and the last sync this branch records, so a conflict
    # resolved in the cloud is not merged (and conflicted) again.
    synced=$(git -C "$tree" log -n 500 --format='%(trailers:key=GitLab-Synced,valueonly)' "$branch" \
        | awk -v n="$name" '$1 == n { print $2; exit }')
    base=$deploy
    if [ -n "$synced" ] && git -C "$bare" cat-file -e "$synced^{commit}" 2>/dev/null \
        && git -C "$bare" merge-base --is-ancestor "$deploy" "$synced"; then base=$synced; fi
    echo "$name deploy $deploy synced ${synced:-none} base $base head $head" >> "$STATE"
    if [ "$base" = "$head" ]; then echo "  nothing new since the last deploy or sync"; return; fi
    if [ "$(git -C "$tree" branch --show-current)" != "$branch" ] || [ -n "$(git -C "$tree" status --porcelain)" ]; then
        echo "ATTENTION: $tree is not a clean $branch; teammates' work not merged"; attention=1; return
    fi

    echo "  teammates' commits:"
    git -C "$bare" log --no-merges --format='    %h %an | %s' "$base..$head"
    local changed deleted owned conflicts=()
    changed=$(git -C "$bare" diff --name-only --diff-filter=d "$base" "$head" -- "$@")
    deleted=$(git -C "$bare" diff --name-only --diff-filter=D "$base" "$head" -- "$@")
    owned=$(git -C "$bare" diff --name-only "$base" "$head" -- \
        HomeFront.csproj HomeFront.sln 'appsettings*.json' 'azure-pipelines*.yml' NuGet.Config)

    local tmp; tmp=$(mktemp -d)
    while IFS= read -r f; do
        [ -n "$f" ] || continue
        git -C "$bare" show "$head:$f" > "$tmp/theirs"
        if [ ! -f "$tree/$f" ]; then
            mkdir -p "$tree/$(dirname "$f")"; cp "$tmp/theirs" "$tree/$f"; continue
        fi
        git -C "$bare" show "$base:$f" > "$tmp/base" 2>/dev/null || : > "$tmp/base"
        cp "$tree/$f" "$tmp/ours"
        git merge-file -q -L ours -L base -L "$name main" "$tmp/ours" "$tmp/base" "$tmp/theirs"
        rc=$?
        if [ $rc -eq 0 ]; then cp "$tmp/ours" "$tree/$f"; else conflicts+=("$f"); fi
    done <<< "$changed"
    rm -rf "$tmp"

    # Reported, not applied, and not blocking: the Mac's deploy owns these, as with --pull.
    if [ -n "$owned" ]; then
        echo "NOTE: branch-owned files changed on $name main (not synced; merge by hand if the change belongs in source):"
        echo "$owned" | sed 's/^/    /'
    fi
    if [ -n "$deleted" ]; then
        echo "NOTE: deleted on $name main (not applied; remove by hand if intended):"
        echo "$deleted" | sed 's/^/    /'
    fi
    if [ ${#conflicts[@]} -gt 0 ]; then
        echo "ATTENTION: ${#conflicts[@]} file(s) CONFLICT, left unchanged. Merge by hand, then commit with the"
        echo "    trailer 'GitLab-Synced: $name $head'. Compare: git -C $bare diff $base $head -- <file>"
        printf '    %s\n' "${conflicts[@]}"
        attention=1
    fi

    # The trailer marks $head as synced only when every file merged.
    local trailer="" subject
    [ ${#conflicts[@]} -eq 0 ] && trailer="GitLab-Synced: $name $head"
    subject="Merge teammates' work from GitLab $name main ($(git -C "$bare" rev-parse --short "$base")..$(git -C "$bare" rev-parse --short "$head"))"
    if [ -z "$(git -C "$tree" status --porcelain)" ]; then
        [ -n "$trailer" ] || { echo "  nothing merged cleanly"; return; }
        git -C "$tree" commit -q --allow-empty -m "Record GitLab $name main as synced (our tree already held it)" -m "$trailer"
    else
        git -C "$tree" add -A
        git -C "$tree" commit -q -m "$subject" \
            -m "$(git -C "$bare" log --no-merges --format='%h %an | %s' "$base..$head")" ${trailer:+-m "$trailer"}
    fi
    echo "  committed $(git -C "$tree" rev-parse --short HEAD) on $branch"
    pushable+=("$tree|$branch")
}

# Paths that never come back to source, as in deploy_to_repos.sh --pull.
COMMON=(':!appsettings*.json' ':!App_Data/jira-settings.json' ':!azure-pipelines*.yml' ':!NuGet.Config'
        ':!local-packages/' ':!certs/' ':!wwwroot/tickets/' ':!wwwroot/feedback/' ':!Logs/')
sync_repo hyphen-pb "$APP" main ':!HomeFront.csproj' ':!HomeFront.sln' "${COMMON[@]}"
sync_repo flexkit "$FLEXKIT" telerik-parity-20260904 "${COMMON[@]}"

if [ ${#pushable[@]} -gt 0 ]; then
    if [ $attention -ne 0 ]; then
        echo "ATTENTION: merge commits left UNPUSHED until the items above are resolved."
    elif ! dotnet build "$APP/HomeFront.sln" -v q -nologo > "$MIRRORS/build.log" 2>&1; then
        echo "ATTENTION: the app does not build after the merge (see $MIRRORS/build.log); commits left UNPUSHED."
        attention=1
    elif $PUSH; then
        for p in "${pushable[@]}"; do
            IFS='|' read -r tree branch <<< "$p"
            if git -C "$tree" push -q origin "$branch"; then echo "  pushed $(basename "$tree") $branch to GitHub"
            else echo "ATTENTION: push of $tree $branch failed (GitHub moved? merge origin/$branch, then push)"; attention=1; fi
        done
    fi
fi
[ $attention -eq 0 ] && echo "Teammates' GitLab work is in the cloud trees." || exit 2
