#!/bin/bash
# pre_deploy_gate.sh — last check IMMEDIATELY before deploy_to_repos.sh. Exits non-zero to block.
#   bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh && bash tools/deploy_to_repos.sh
#
# Read-only apart from `git fetch` in the staging clones and, on PASS, the state file that
# verify_deploy.sh and publish_flexcore.sh read.
#
#   bash pre_deploy_gate.sh               # normal push
#   bash pre_deploy_gate.sh --with-r1uat  # the run will also deploy R1-UAT
#   bash pre_deploy_gate.sh --with-r2uat  # the run will also merge main into R2-UAT
#   QUIET_MINUTES=20 bash pre_deploy_gate.sh
set -uo pipefail

WITH_R1UAT=false
WITH_R2UAT=false
for a in "$@"; do
  case "$a" in
    --with-r1uat) WITH_R1UAT=true ;;
    --with-r2uat) WITH_R2UAT=true ;;
    *) echo "✗ unknown argument '$a' (deploy_to_repos.sh would silently ignore a typo like this and push for real)"; exit 2 ;;
  esac
done
QUIET="${QUIET_MINUTES:-10}"
source "$(cd "$(dirname "$0")" && pwd)/_common.sh"
BLOCK=0
no()  { printf '✗ %s\n' "$1"; BLOCK=1; }
yes() { printf '✓ %s\n' "$1"; }

# FlexCore ships from main only. While its checkout is on another branch the deploy skips the repo
# (nothing is pushed to GitHub FlexCore), so its upstream state cannot be reverted by this run and is
# not judged here.
fcbr=$(git -C "$V/FlexCore" branch --show-current 2>/dev/null)
gate_repos="hyphen-pb homefront flexkit flexcore"
if [ "$fcbr" != "main" ]; then
  gate_repos="hyphen-pb homefront flexkit"
  printf '! FlexCore is on %s, not main — the deploy will SKIP flexcore; GitHub FlexCore is left as it is\n' "'$fcbr'"
fi

# 1. Upstream vs the clone's local main, and unpushed leftovers from a failed push.
refs=""; for r in $gate_repos; do refs="$refs $r:main"; done
$WITH_R1UAT && refs="$refs hyphen-pb:R1-UAT"
for spec in $refs; do
  r=${spec%%:*}; b=${spec#*:}
  if ! git -C "$D/$r" fetch -q origin 2>/dev/null; then no "$r: fetch failed — cannot prove upstream is unchanged"; continue; fi
  a=$(git -C "$D/$r" rev-list --count "origin/${b}..${b}" 2>/dev/null || echo "?")
  [ "$a" != "0" ] && no "$r $b has $a UNPUSHED commit(s) from a failed push — first: git -C $D/$r reset --hard \$(git -C $D/$r merge-base $b origin/$b)"
  n=$(git -C "$D/$r" rev-list --count "${b}..origin/${b}" 2>/dev/null || echo "?")
  if [ "$n" = "0" ]; then yes "$r $b: no upstream commits past the clone's last push"
  elif [ "$r" = hyphen-pb ] && [ "$b" = main ]; then no "$r $b: $n upstream commit(s) not merged — run: bash tools/deploy_to_repos.sh --pull (then commit the app)"
  else no "$r $b: $n upstream commit(s) not merged — --pull does not cover $r $b: merge them by hand (SKILL.md step 3)"; fi
done

# 2. Teammate content that has not shipped from here must be IN our tree. This does not trust the clone's
#    local main — any push-mode run, including the --dry-run just before this gate, resets it to
#    origin/main and hides upstream commits from check 1. It judges every commit after our last
#    "Deploy <repo>" commit by its content.
content_bad=0
for r in $gate_repos; do
  while IFS='|' read -r sha date author subj; do
    [ -z "$sha" ] && continue
    while IFS= read -r f; do
      [ -z "$f" ] && continue
      res=$(classify_file "$r" "$sha" "$f"); state=${res%%|*}
      case "$state" in ok|NOTE) ;; *) no "$state: $r $f ($sha, $author) — ${res#*|}"; content_bad=1 ;; esac
    done < <(git -C "$D/$r" show --name-only --no-renames --format='' "$sha" 2>/dev/null | grep -v '^$')
  done < <(unshipped_commits "$r")
done
[ $content_bad = 0 ] && yes "every unshipped teammate change is in our trees"

# 3. --pull conflicts that are still unresolved.
if [ -s "$PULL_CONFLICTS" ]; then no "unresolved --pull conflicts listed in $PULL_CONFLICTS — hand-merge each, then delete the file"
else yes "no unresolved --pull conflicts recorded"; fi

# 4. Conflict markers anywhere in the trees that ship (a failed hand-merge writes them into source, and
#    no build checks CSS or JS).
markers=""
for t in "$HF" "$V/FlexKit" "$V/FlexCore"; do
  hits=$(git -C "$t" grep --untracked -nE '^(<<<<<<<|>>>>>>>)( |$)' 2>/dev/null | head -3)
  [ -n "$hits" ] && markers="$markers$(printf '%s\n' "$hits" | sed "s|^|     $(basename "$t"): |")\n"
done
if [ -z "$markers" ]; then yes "no merge-conflict markers in the app, FlexKit or FlexCore"
else no "merge-conflict markers found:"; printf "$markers"; fi

# 5. Nobody saved source in the last $QUIET minutes (the deploy ships working trees). Runtime data the
#    running app writes itself (User/ prefs, App_Data, tickets, feedback, Logs) is not source.
recent=""
[ "$QUIET" -gt 0 ] && recent=$(find "$HF" "$V/FlexKit" "$V/FlexCore" \
  \( -name bin -o -name obj -o -name .git -o -name node_modules \
     -o -path '*/HomeFront/User' -o -path '*/HomeFront/App_Data' -o -path '*/wwwroot/tickets' \
     -o -path '*/wwwroot/feedback' -o -path '*/HomeFront/Logs' \) -prune -o \
  -type f \( -name '*.cs' -o -name '*.razor' -o -name '*.css' -o -name '*.js' -o -name '*.csproj' -o -name '*.mjs' -o -name '*.json' \) \
  -mmin "-$QUIET" -print 2>/dev/null | head -5)
if [ "$QUIET" -eq 0 ]; then yes "recent-edit check disabled (QUIET_MINUTES=0 — fixtures only)"
elif [ -z "$recent" ]; then yes "no source file changed in the last $QUIET minutes"
else no "source changed in the last $QUIET minutes — another session may be mid-edit:"; echo "$recent" | sed 's/^/     /'; fi

# 6. FlexKit ships from its working branch.
br=$(git -C "$V/FlexKit" branch --show-current)
[ "$br" = "telerik-parity-20260904" ] && yes "FlexKit on telerik-parity-20260904" || no "FlexKit is on '$br', not telerik-parity-20260904"
[ "$fcbr" = "main" ] && yes "FlexCore on main"

# 7. FlexKit packaging. deploy_to_repos.sh repacks when the version changed or a .cs/.razor/.css/.js/.csproj
#    file is newer than the feed nupkg — so (a) a repack under a version already extracted in
#    ~/.nuget/packages ships bytes the staged build never verified, and (b) a changed asset (image, json,
#    README) under an unchanged version is NOT repacked at all and the stale package ships.
ver=$(grep -o '<Version>[^<]*' "$V/FlexKit/FlexKit.csproj" | head -1 | cut -c10-)
feed="$H/Deploy/local-packages/FlexKit.$ver.nupkg"
prune=( \( -name bin -o -name obj -o -name .git -o -name node_modules -o -name .agents -o -name docs -o -name tests \) -prune -o )
if [ ! -f "$feed" ]; then will_pack=true
elif [ -n "$(find "$V/FlexKit" "${prune[@]}" -type f \( -name '*.cs' -o -name '*.razor' -o -name '*.css' -o -name '*.js' -o -name '*.csproj' \) -newer "$feed" -print 2>/dev/null | head -1)" ]; then will_pack=true
else will_pack=false; fi
if $will_pack && [ -d "$HOME/.nuget/packages/flexkit/$ver" ]; then
  no "FlexKit $ver would be REPACKED but $ver is already extracted in ~/.nuget/packages — bump FlexKit.csproj"
elif $will_pack; then yes "FlexKit $ver will be packed fresh"
else
  assets=$(find "$V/FlexKit" "${prune[@]}" -type f ! \( -name '*.cs' -o -name '*.razor' -o -name '*.css' -o -name '*.js' -o -name '*.csproj' \
           -o -name CLAUDE.md -o -name AGENTS.md -o -name PUBLISHING.md -o -name '.DS_Store' \) -newer "$feed" -print 2>/dev/null | head -3)
  if [ -n "$assets" ]; then
    no "FlexKit non-code files changed since the $ver package, but the script only repacks on code changes — bump FlexKit.csproj to force a repack:"
    echo "$assets" | sed 's/^/     /'
  else yes "FlexKit $ver feed package is current (no repack needed)"; fi
fi

echo
if [ $BLOCK = 0 ]; then
  {
    echo "time=$(date +%s)"
    echo "with_r1uat=$WITH_R1UAT"
    echo "with_r2uat=$WITH_R2UAT"
    echo "hyphen-pb.r2uat=$(git -C "$D/hyphen-pb" rev-parse origin/R2-UAT 2>/dev/null)"
    for r in hyphen-pb homefront flexkit flexcore; do echo "$r.origin=$(git -C "$D/$r" rev-parse origin/main)"; done
    echo "hyphen-pb.r1uat=$(git -C "$D/hyphen-pb" rev-parse origin/R1-UAT 2>/dev/null)"
    echo "flexkit.version=$ver"
    echo "flexcore.version=$(grep -o '<Version>[^<]*' "$V/FlexCore/FlexCore.csproj" | head -1 | cut -c10-)"
    echo "flexcore.head=$(git -C "$V/FlexCore" rev-parse HEAD)"
    echo "flexcore.dirty=$(git -C "$V/FlexCore" status --porcelain | wc -l | tr -d ' ')"
  } > "$GATE_STATE"
  echo "GATE PASSED — clear to run deploy_to_repos.sh (state recorded for verify_deploy.sh)"
else
  echo "GATE BLOCKED — do not deploy"
fi
exit $BLOCK
