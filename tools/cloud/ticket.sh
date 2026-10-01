#!/usr/bin/env bash
# Mac side of the cloud Jira workflow. A cloud session pushes each ticket's fix to a branch
# jira/<KEY> on the GitHub mirrors (app and/or FlexKit, whichever it changed). This script
# brings those branches to the Mac for testing, and merges the ones you accept.
#
#   bash tools/cloud/ticket.sh list            ticket branches on GitHub, and whether merged
#   bash tools/cloud/ticket.sh test HHM-1234   check out jira/HHM-1234 in each repo that has it
#                                              (the other repo goes to its current main branch)
#   bash tools/cloud/ticket.sh back            return both repos to main / telerik-parity-20260904
#   bash tools/cloud/ticket.sh accept HHM-1234 merge jira/HHM-1234 into main / telerik-parity-20260904
#                                              and push those to GitHub (never GitLab)
#
# Deploying stays with /update-repositories. Every FlexKit change still needs its FlexCore
# port on the Mac. Refuses to touch a repo with uncommitted changes.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUTER="$(cd "$HERE/../.." && pwd)"
REPOS=("$OUTER/MobileSource/HomeFront|main" "$(cd "$OUTER/.." && pwd)/FlexKit|telerik-parity-20260904")

remote() { git -C "$1" remote | grep -qx github && echo github || echo origin; }
clean() { [ -z "$(git -C "$1" status --porcelain --untracked-files=no)" ] || { echo "ERROR: $1 has uncommitted changes; commit or stash them first."; exit 1; }; }
fetch() { git -C "$1" fetch -q --prune "$(remote "$1")" "+refs/heads/$2:refs/remotes/$(remote "$1")/$2" "+refs/heads/jira/*:refs/remotes/$(remote "$1")/jira/*"; }
key() { [[ "${1:-}" =~ ^[A-Z][A-Z0-9]+-[0-9]+$ ]] || { echo "usage: $0 $cmd <JIRA-KEY>, e.g. HHM-1234"; exit 1; }; }
to_main() { # <dir> <branch>
    local r; r=$(remote "$1")
    git -C "$1" checkout -q "$2"
    git -C "$1" merge -q --ff-only "$r/$2" || echo "WARN: $1 $2 has local commits not on $r; left as is"
}

cmd="${1:-}"
case "$cmd" in
list)
    for e in "${REPOS[@]}"; do
        IFS='|' read -r dir main <<< "$e"; r=$(remote "$dir"); fetch "$dir" "$main"
        echo "== $(basename "$dir") (base $main)"
        git -C "$dir" for-each-ref --format='%(refname:short)' "refs/remotes/$r/jira/" | while read -r b; do
            if git -C "$dir" merge-base --is-ancestor "$b" "$r/$main"; then s=merged; else s="$(git -C "$dir" rev-list --count "$r/$main..$b") commit(s) to test"; fi
            echo "  ${b#$r/}  $(git -C "$dir" log -1 --format='%cs %s' "$b")  [$s]"
        done
    done ;;
test)
    key "${2:-}"; found=0
    for e in "${REPOS[@]}"; do IFS='|' read -r dir main <<< "$e"; clean "$dir"; done
    for e in "${REPOS[@]}"; do
        IFS='|' read -r dir main <<< "$e"; r=$(remote "$dir"); fetch "$dir" "$main"
        if git -C "$dir" rev-parse -q --verify "$r/jira/$2" >/dev/null; then
            git -C "$dir" checkout -q -B "jira/$2" "$r/jira/$2"; found=1
            echo "$(basename "$dir"): jira/$2 at $(git -C "$dir" log -1 --format='%h %s')"
        else
            to_main "$dir" "$main"; echo "$(basename "$dir"): no jira/$2; on $main"
        fi
    done
    [ $found -eq 1 ] || { echo "ERROR: no branch jira/$2 on GitHub"; exit 1; }
    echo "Build and test now. Then: $0 accept $2   (or $0 back)" ;;
back)
    for e in "${REPOS[@]}"; do IFS='|' read -r dir main <<< "$e"; clean "$dir"; fetch "$dir" "$main"; to_main "$dir" "$main"; echo "$(basename "$dir"): $main"; done ;;
accept)
    key "${2:-}"
    for e in "${REPOS[@]}"; do IFS='|' read -r dir main <<< "$e"; clean "$dir"; done
    for e in "${REPOS[@]}"; do
        IFS='|' read -r dir main <<< "$e"; r=$(remote "$dir"); fetch "$dir" "$main"
        git -C "$dir" rev-parse -q --verify "$r/jira/$2" >/dev/null || { to_main "$dir" "$main"; continue; }
        to_main "$dir" "$main"
        git -C "$dir" merge -q --no-ff -m "Merge jira/$2" "$r/jira/$2" \
            || { echo "ERROR: merge conflict in $dir; resolve, commit, then git -C $dir push $r $main"; exit 1; }
        git -C "$dir" push -q "$r" "$main"
        echo "$(basename "$dir"): jira/$2 merged into $main and pushed to $r"
    done
    echo "Accepted. Deploy with /update-repositories when ready; port any FlexKit change to FlexCore." ;;
*)
    sed -n '2,15p' "$0" | sed 's/^# \{0,1\}//'; exit 1 ;;
esac
