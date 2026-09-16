#!/bin/bash
# pre_deploy_gate.sh — last check IMMEDIATELY before deploy_to_repos.sh. Exits non-zero to block.
# Chain it:   bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh && bash tools/deploy_to_repos.sh
#
# Closes the gap between preflight and the push: a teammate can merge, or a session can save a
# file, in the minutes in between. Read-only apart from `git fetch` in the staging clones.
#
#   bash pre_deploy_gate.sh               # normal push
#   bash pre_deploy_gate.sh --with-r1uat  # the run will also deploy R1-UAT
#   QUIET_MINUTES=20 bash pre_deploy_gate.sh
set -uo pipefail

WITH_R1UAT=false
for a in "$@"; do
  case "$a" in
    --with-r1uat) WITH_R1UAT=true ;;
    *) echo "✗ unknown argument '$a' (deploy_to_repos.sh would silently ignore a typo like this and push for real)"; exit 2 ;;
  esac
done
QUIET="${QUIET_MINUTES:-10}"
V=/Users/wadood/projects/VBToCSharp
H=$V/HomeFront
D=$H/Deploy/repos
BLOCK=0
no()  { printf '✗ %s\n' "$1"; BLOCK=1; }
yes() { printf '✓ %s\n' "$1"; }

# 1. Upstream must not have moved past the sha we last pushed. The deploy's own fetch swallows
#    errors, and ANY push-mode run (even --dry-run) resets the clone's main to origin/main — after
#    which --pull reports "No new changes" and the push reverts teammates' work. So check first.
refs="hyphen-pb:main homefront:main flexkit:main flexcore:main"
$WITH_R1UAT && refs="$refs hyphen-pb:R1-UAT"
for spec in $refs; do
  r=${spec%%:*}; b=${spec#*:}
  if ! git -C "$D/$r" fetch -q origin 2>/dev/null; then no "$r: fetch failed — cannot prove upstream is unchanged"; continue; fi
  n=$(git -C "$D/$r" rev-list --count "${b}..origin/${b}" 2>/dev/null || echo "?")
  if [ "$n" = "0" ]; then yes "$r $b: no upstream commits past our last push"
  else no "$r $b: $n upstream commit(s) not merged — run: bash tools/deploy_to_repos.sh --pull (then commit the app)"; fi
done

# 2. Nobody saved source in the last $QUIET minutes (the deploy ships working trees). Runtime data
#    the running app writes itself (User/ prefs, App_Data, tickets, feedback, Logs) is not source.
recent=$(find "$H/MobileSource/HomeFront" "$V/FlexKit" "$V/FlexCore" \
  \( -name bin -o -name obj -o -name .git -o -name node_modules \
     -o -path '*/HomeFront/User' -o -path '*/HomeFront/App_Data' -o -path '*/wwwroot/tickets' \
     -o -path '*/wwwroot/feedback' -o -path '*/HomeFront/Logs' \) -prune -o \
  -type f \( -name '*.cs' -o -name '*.razor' -o -name '*.css' -o -name '*.js' -o -name '*.csproj' -o -name '*.mjs' -o -name '*.json' \) \
  -mmin "-$QUIET" -print 2>/dev/null | head -5)
if [ -z "$recent" ]; then yes "no source file changed in the last $QUIET minutes"
else no "source changed in the last $QUIET minutes — another session may be mid-edit:"; echo "$recent" | sed 's/^/     /'; fi

# 3. FlexKit ships from its working branch.
br=$(git -C "$V/FlexKit" branch --show-current)
[ "$br" = "telerik-parity-20260904" ] && yes "FlexKit on telerik-parity-20260904" || no "FlexKit is on '$br', not telerik-parity-20260904"

# 4. The script will repack FlexKit when the version changed OR sources are newer than the feed
#    nupkg. Repacking a version already extracted in ~/.nuget/packages ships bytes the staged
#    build never verified — so if it WILL repack, the version must be fresh.
ver=$(grep -o '<Version>[^<]*' "$V/FlexKit/FlexKit.csproj" | head -1 | cut -c10-)
feed="$H/Deploy/local-packages/FlexKit.$ver.nupkg"
will_pack=false
if [ ! -f "$feed" ]; then will_pack=true
elif [ -n "$(find "$V/FlexKit" \( -name bin -o -name obj -o -name .git \) -prune -o -type f \( -name '*.cs' -o -name '*.razor' -o -name '*.css' -o -name '*.js' -o -name '*.csproj' \) -newer "$feed" -print 2>/dev/null | head -1)" ]; then will_pack=true; fi
if $will_pack && [ -d "$HOME/.nuget/packages/flexkit/$ver" ]; then
  no "FlexKit $ver will be REPACKED but $ver is already extracted in ~/.nuget/packages — bump FlexKit.csproj"
elif $will_pack; then yes "FlexKit $ver will be packed fresh"
else yes "FlexKit $ver feed package is current (no repack)"; fi

echo
if [ $BLOCK = 0 ]; then echo "GATE PASSED — clear to run deploy_to_repos.sh"; else echo "GATE BLOCKED — do not deploy"; fi
exit $BLOCK
