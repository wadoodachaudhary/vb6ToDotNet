#!/bin/bash
# _common.sh — shared paths and helpers for the update-repositories scripts. Sourced, never run.
# Everything here is read-only. Never name a variable path/fpath/status/argv: zsh ties those to
# PATH and friends, so sourcing this into zsh would break every command inside the function.

# UR_* overrides exist only so fixture tests can point the helpers at throwaway repos.
V=${UR_V:-/Users/wadood/projects/VBToCSharp}
H=${UR_H:-$V/HomeFront}
HF=${UR_HF:-$H/MobileSource/HomeFront}
PB=$H/HomeFrontPB
D=${UR_D:-$H/Deploy/repos}
CANON=$H/Docs-Skill/UpdateRepositories
INSTALLED=$HOME/.claude/skills/update-repositories
GATE_STATE=$H/Deploy/.update-repositories-gate.state     # written by pre_deploy_gate.sh on PASS (Deploy/ is git-ignored)
PULL_CONFLICTS=$H/Deploy/pull-conflicts.txt               # the agent saves --pull's CONFLICT list here
PREFLIGHT_ACK=$H/Deploy/preflight-ack.txt                # "<sha> <path> <reason>": teammate hunks we deliberately did not take

# Staging clone -> the source tree it is fed from.
src_for() { case "$1" in hyphen-pb|homefront) echo "$HF" ;; flexkit) echo "$V/FlexKit" ;; flexcore) echo "$V/FlexCore" ;; esac; }

# Files the deploy never syncs from our tree (the branch owns them).
is_repo_owned() {
  case "$1" in azure-pipelines*.yml|NuGet.Config|local-packages/*|certs/*|appsettings*.json|App_Data/jira-settings.json) return 0 ;; esac
  return 1
}

# Files the deploy TRANSFORMS on the way out, so our copy never equals the branch blob.
is_transformed() { case "$1" in Program.cs|HomeFront.csproj|HomeFront.sln) return 0 ;; esac; return 1; }

# Resolve <rev>:<path> to a blob id, or print NOTHING. Plain `git rev-parse` echoes an argument it cannot
# resolve on stdout, which makes every "does this file exist there?" test pass.
blob_at() { git -C "$1" rev-parse --verify --quiet "$2:$3" 2>/dev/null; }

# hunks_present <clone> <sha> <path> <our file>
#   0 = the change <sha> made to <path> is in <our file>: every DISTINCTIVE line it added is present, and
#       (for a commit that only removed lines) every distinctive line it removed is absent
#   1 = at least one such line is missing / still present, or <our file> does not exist
#   2 = <sha> changed no distinctive lines (only braces, blank lines, short tags) — cannot judge by lines
# Line presence, not blob equality: our file legitimately differs when we ALSO edited it, or when the
# deploy transforms it (Program.cs, HomeFront.csproj, HomeFront.sln). "Distinctive" = 8+ non-space
# characters, so a generic added line such as "}" or "</div>" cannot make an unmerged change look merged.
hunks_present() {
  local clone=$1 sha=$2 rel=$3 ours=$4 diff added removed line
  [ -f "$ours" ] || return 1
  diff=$(git -C "$clone" diff --no-renames "${sha}^" "$sha" -- "$rel" 2>/dev/null)
  added=$(printf '%s\n' "$diff" | sed -e '/^+++ /d' | sed -n 's/^+//p' | awk '{t=$0; gsub(/[[:space:]]/,"",t); if (length(t)>=8) print}')
  removed=$(printf '%s\n' "$diff" | sed -e '/^--- /d' | sed -n 's/^-//p' | awk '{t=$0; gsub(/[[:space:]]/,"",t); if (length(t)>=8) print}')
  if [ -n "$added" ]; then
    while IFS= read -r line; do grep -Fxq -- "$line" "$ours" || return 1; done <<< "$added"
    return 0
  fi
  if [ -n "$removed" ]; then
    # A removed line may legitimately survive elsewhere in the file; only lines the commit's own result
    # no longer contains count.
    local post; post=$(git -C "$clone" show "$sha:$rel" 2>/dev/null)
    local judged=0
    while IFS= read -r line; do
      grep -Fxq -- "$line" <<< "$post" && continue     # here-string, not a pipe: grep -q + pipefail = SIGPIPE
      judged=1
      grep -Fxq -- "$line" "$ours" && return 1
    done <<< "$removed"
    [ $judged = 1 ] && return 0
  fi
  return 2
}

# ever_in_our_history <clone> <sha> <path> <source repo>
#   Looks only at the distinctive lines <sha> added that are MISSING from our file now.
#   0 = one of them was added to or removed from OUR source repo's history after <sha> was committed —
#       we had it and later changed it ourselves (superseded, not lost)
#   1 = none of them ever appeared in our history — our tree never had that part of the change
ever_in_our_history() {
  local clone=$1 sha=$2 rel=$3 src=$4 when added line n=0
  when=$(git -C "$clone" log -1 --format=%cI "$sha" 2>/dev/null)
  added=$(git -C "$clone" diff --no-renames "${sha}^" "$sha" -- "$rel" 2>/dev/null \
          | sed -e '/^+++ /d' | sed -n 's/^+//p' | awk '{t=$0; gsub(/[[:space:]]/,"",t); if (length(t)>=12) print}')
  while IFS= read -r line; do
    [ -z "$line" ] && continue
    grep -Fxq -- "$line" "$src/$rel" 2>/dev/null && continue      # present now: says nothing about loss
    n=$((n+1)); [ $n -gt 6 ] && break
    if [ -n "$(git -C "$src" log --since="$when" -F -S"$line" --format=%h -- "$rel" 2>/dev/null | head -1)" ]; then return 0; fi
  done <<< "$added"
  return 1
}

# is_acked <sha> <path>  -> prints the recorded reason, returns 0 when the owner/agent recorded that this
# teammate change was deliberately NOT taken (e.g. MR !33's PO-vendor plain text, 2026-09-12).
is_acked() {
  [ -f "$PREFLIGHT_ACK" ] || return 1
  local hit; hit=$(awk -v s="$1" -v f="$2" '$1==s && $2==f {$1="";$2=""; sub(/^  */,""); print; exit}' "$PREFLIGHT_ACK")
  [ -n "$hit" ] || grep -qE "^$1[[:space:]]+$2([[:space:]]|$)" "$PREFLIGHT_ACK" || return 1
  echo "${hit:-acknowledged}"; return 0
}

# Latest "Deploy <repo> — …" commit on origin/main: everything after it has not shipped from this machine.
last_deploy() { git -C "$D/$1" log -1 --format=%H --grep="^Deploy $1" origin/main 2>/dev/null; }

# Teammate (non-deploy, non-merge) commits on origin/main that came AFTER our last deploy commit.
# Independent of the clone's local main, which any push-mode run (even --dry-run) resets.
unshipped_commits() {
  local c=$D/$1 ld
  ld=$(last_deploy "$1")
  if [ -n "$ld" ]; then git -C "$c" log --no-merges --format='%h|%cs|%an|%s' "$ld..origin/main" 2>/dev/null | grep -v '|Deploy '
  else git -C "$c" log --no-merges --since="14 days ago" --format='%h|%cs|%an|%s' origin/main 2>/dev/null | grep -v '|Deploy '; fi
}

# classify_file <repo> <sha> <path>   -> prints "<STATE>|<human text>"
#   STATE: ok | NOTE | REVERT | LOST | ABSENT | RESURRECT | HANDMERGE   (NOTE is informational)
# REVERT/ABSENT/RESURRECT/HANDMERGE = the next push would undo this teammate change.
# LOST = our earlier deploy ALREADY undid it on origin/main (the MR !33 case).
classify_file() {
  local r=$1 sha=$2 f=$3 c=$D/$1 src ours post cur lastsubj bydeploy=false unshipped=false ld
  src=$(src_for "$r")
  if is_repo_owned "$f"; then echo "ok|repo-owned — never synced from our tree"; return; fi
  local ack; if ack=$(is_acked "$sha" "$f"); then echo "ok|acknowledged: $ack"; return; fi
  post=$(blob_at "$c" "$sha" "$f"); cur=$(blob_at "$c" origin/main "$f")
  if [ -e "$src/$f" ]; then ours=$(git -C "$src" hash-object "$src/$f"); else ours=""; fi
  lastsubj=$(git -C "$c" log -1 --format=%s origin/main -- "$f" 2>/dev/null)
  case "$lastsubj" in "Deploy $r"*) bydeploy=true ;; esac
  ld=$(last_deploy "$r")
  if [ -z "$ld" ] || ! git -C "$c" merge-base --is-ancestor "$sha" "$ld" 2>/dev/null; then unshipped=true; fi

  if [ -z "$post" ]; then                                   # the teammate DELETED this path
    if [ -z "$ours" ]; then echo "ok|deleted upstream — gone here too"
    elif [ -n "$cur" ]; then echo "ok|deleted upstream, re-added later — read it"
    elif $unshipped; then echo "RESURRECT|deleted upstream but still in our tree — the push would re-add it (git rm it in the source after checking)"
    else echo "LOST|deleted upstream, then re-added on origin by our deploy — remove it from the source"; fi
    return
  fi
  if [ -z "$ours" ]; then
    if [ -z "$cur" ] && ! $bydeploy; then echo "ok|moved/removed upstream later"
    elif [ -z "$cur" ]; then echo "LOST|added upstream, then removed from origin by our deploy — restore it"
    else echo "ABSENT|on origin/main but missing in our tree — the push would delete it"; fi
    return
  fi
  if [ "$ours" = "$post" ]; then echo "ok|present"; return; fi
  # Matching origin/main proves nothing when OUR deploy was the last to write the file there: that is
  # exactly how a teammate's change disappears (the MR !33 case). Only trust it otherwise.
  if [ -n "$cur" ] && [ "$ours" = "$cur" ] && ! $bydeploy; then echo "ok|matches origin/main"; return; fi
  hunks_present "$c" "$sha" "$f" "$src/$f"; local hp=$?
  if [ $hp -eq 0 ]; then echo "ok|teammate's lines are present (merged)"; return; fi
  if is_transformed "$f" && $unshipped; then
    echo "HANDMERGE|the deploy rewrites this file on the way out — merge the teammate's hunk by hand (keep our ProjectReference)"; return
  fi
  if $unshipped; then echo "REVERT|teammate's change is not in our tree — the push would revert it"; return; fi
  if $bydeploy; then
    if ever_in_our_history "$c" "$sha" "$f" "$src"; then
      echo "NOTE|lines no longer present, but our own history had them and changed them later — superseded"
    else
      echo "LOST|our tree never had this change and our deploy overwrote it on origin/main — restore it, or record it in $PREFLIGHT_ACK if deliberate"
    fi
    return
  fi
  echo "ok|changed again upstream later (not ours to restore)"
}
