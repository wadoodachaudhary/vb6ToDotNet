#!/bin/bash
# verify_deploy.sh — prove what a deploy actually pushed. Run AFTER deploy_to_repos.sh.
# Read-only apart from `git fetch` in the staging clones. Exits 1 if any hard check fails.
# Needs the state pre_deploy_gate.sh records on PASS: without it, "the branches look right" cannot be
# told apart from "nothing was pushed at all".
#   bash verify_deploy.sh
set -uo pipefail
source "$(cd "$(dirname "$0")" && pwd)/_common.sh"
FAIL=0
bad()  { printf '   ✗ %s\n' "$1"; FAIL=1; }
good() { printf '   ✓ %s\n' "$1"; }
note() { printf '   · %s\n' "$1"; }
hdr()  { printf '\n== %s\n' "$1"; }
state() { grep -m1 "^$1=" "$GATE_STATE" 2>/dev/null | cut -d= -f2-; }

for r in hyphen-pb homefront flexkit flexcore; do git -C "$D/$r" fetch -q origin 2>/dev/null || bad "$r: fetch failed"; done

hdr "What shipped since the gate passed"
if [ ! -s "$GATE_STATE" ]; then
  bad "no gate state ($GATE_STATE) — run pre_deploy_gate.sh before deploying; nothing can be proven"
  gate_time=0
else
  gate_time=$(state time); age=$(( ($(date +%s) - gate_time) / 60 ))
  note "gate passed $age min ago"
  [ "$age" -gt 180 ] && bad "the gate state is $age min old — it does not describe this deploy; gate again before the next run"
fi
WITH_R1UAT=$(state with_r1uat); WITH_R1UAT=${WITH_R1UAT:-false}
shipped=0
for r in hyphen-pb homefront flexkit flexcore; do
  if [ "$gate_time" -eq 0 ]; then note "$r: origin/main $(git -C "$D/$r" rev-parse --short origin/main) (no gate baseline — cannot tell whether this run shipped it)"; continue; fi
  c=$D/$r; o=$(git -C "$c" rev-parse origin/main); before=$(state "$r.origin")
  ahead=$(git -C "$c" rev-list --count origin/main..main 2>/dev/null || echo 0)
  subj=$(git -C "$c" log -1 --format=%s origin/main); when=$(git -C "$c" log -1 --format=%ct origin/main)
  if [ "${ahead:-0}" -gt 0 ]; then
    bad "$r: $ahead UNPUSHED commit(s) in the clone — the push was rejected or aborted (step 3 recovery; never force-push)"
  elif [ -n "$before" ] && [ "$o" = "$before" ]; then
    note "$r: unchanged since the gate (expected only if the run printed '✓ No changes to push' for it)"
  else
    case "$subj" in
      "Deploy $r"*) if [ "$when" -ge "${gate_time:-0}" ]; then good "$r: shipped $(git -C "$c" rev-parse --short origin/main) — $subj"; shipped=$((shipped+1))
                    else bad "$r: origin/main moved to an OLD deploy commit — not this run"; fi ;;
      *) bad "$r: origin/main moved to a non-deploy commit ('$subj') — a teammate pushed after the gate; re-check before publishing" ;;
    esac
  fi
done
[ "$gate_time" -gt 0 ] && [ "$shipped" -eq 0 ] && bad "no repository received a deploy commit after the gate — nothing was deployed"

APP_REFS="hyphen-pb:origin/main homefront:origin/main"
[ "$WITH_R1UAT" = true ] && APP_REFS="$APP_REFS hyphen-pb:origin/R1-UAT"
NON_RUNTIME='^(CLAUDE\.md|AGENTS\.md|GEMINI\.md|PUBLISHING\.md|Data/DataControl\.API\.md|\.agents/|\.claude/|\.codex/|\.codex-backups/|verification/|tests/|docs/|Docs/)|\.vb6_gap\.txt$'

hdr "Login-skip must be gone from every pushed app branch (owner 2026-09-12)"
for spec in $APP_REFS; do
  repo=${spec%%:*}; ref=${spec#*:}
  hits=$(git -C "$D/$repo" grep -I -i -l -e devautologin "$ref" -- ':!*.md' ':!*.txt' 2>/dev/null | sed "s|^$ref:||" | tr '\n' ' ')
  [ -z "$hits" ] && good "$repo $ref: clean" || bad "$repo $ref still carries the login-skip: $hits"
done

hdr "Non-runtime files must be stripped from every pushed branch (owner 2026-09-13)"
for spec in $APP_REFS flexkit:origin/main flexcore:origin/main; do
  repo=${spec%%:*}; ref=${spec#*:}
  hits=$(git -C "$D/$repo" ls-tree -r --name-only "$ref" | grep -E "$NON_RUNTIME" | head -5 | tr '\n' ' ')
  [ -z "$hits" ] && good "$repo $ref: none" || bad "$repo $ref ships non-runtime files: $hits"
done

hdr "Password check still ships (positive assertion)"
for spec in $APP_REFS; do
  repo=${spec%%:*}; ref=${spec#*:}
  login=$(git -C "$D/$repo" ls-tree -r --name-only "$ref" | grep -E '(^|/)FLogin\.razor$' | head -1)
  body=""; [ -n "$login" ] && body=$(git -C "$D/$repo" cat-file blob "$ref:$login" 2>/dev/null)
  # Captured first: `cat-file | grep -q` under pipefail fails at random when grep exits early (SIGPIPE).
  if [[ "$body" == *"Sec.EnforcePasswordCheck"* ]]; then good "$repo $ref: $login consults Sec.EnforcePasswordCheck"
  else bad "$repo $ref: FLogin.razor missing or no longer checks the password"; fi
done

hdr "hyphen-pb main carries the FlexKit package its csproj references"
csproj=$(git -C "$D/hyphen-pb" cat-file blob origin/main:HomeFront.csproj 2>/dev/null)
pr=$(printf '%s\n' "$csproj" | grep -oE '<PackageReference[^>]*Include="FlexKit"[^>]*>' | grep -oE 'Version="[^"]+"' | cut -d'"' -f2 | head -1)
pk=$(git -C "$D/hyphen-pb" ls-tree --name-only origin/main local-packages/ 2>/dev/null | grep -oE 'FlexKit\.[0-9][0-9.]*[0-9]\.nupkg' | sed 's/^FlexKit\.//; s/\.nupkg$//' | head -1)
if [ -n "$pr" ] && [ "$pr" = "$pk" ]; then good "PackageReference $pr == local-packages/FlexKit.$pk.nupkg (a silent pack failure would break this)"
else bad "PackageReference '$pr' vs local-packages '$pk'"; fi

hdr "R1-UAT"
if [ "$WITH_R1UAT" = true ]; then
  flogin=$(git -C "$D/hyphen-pb" cat-file blob origin/R1-UAT:Components/Pages/FLogin.razor 2>/dev/null)
  n=$(printf '%s\n' "$flogin" | grep -o 'Disabled="true"' | wc -l | tr -d ' ')
  [ "$n" -eq 5 ] && good "lockdown intact (5/5 Disabled toggles)" || bad "lockdown broken ($n/5)"
  m1=$(git -C "$D/hyphen-pb" cat-file blob "origin/main:local-packages/FlexKit.$pk.nupkg" 2>/dev/null | md5)
  m2=$(git -C "$D/hyphen-pb" cat-file blob "origin/R1-UAT:local-packages/FlexKit.$pk.nupkg" 2>/dev/null | md5)
  [ -n "$pk" ] && [ "$m1" = "$m2" ] && good "FlexKit.$pk.nupkg byte-identical on main and R1-UAT" || bad "main and R1-UAT carry DIFFERENT FlexKit packages"
else
  now_r1=$(git -C "$D/hyphen-pb" rev-parse origin/R1-UAT); was_r1=$(state hyphen-pb.r1uat)
  if [ -n "$was_r1" ] && [ "$now_r1" != "$was_r1" ]; then bad "R1-UAT moved without --with-r1uat — it is frozen"
  else note "frozen — not deployed (head $(git -C "$D/hyphen-pb" rev-parse --short origin/R1-UAT))"; fi
fi

hdr "Staging clones handed back clean and on main"
for r in hyphen-pb homefront flexkit flexcore; do
  br=$(git -C "$D/$r" branch --show-current); dirty=$(git -C "$D/$r" status --porcelain | wc -l | tr -d ' ')
  [ "$br" = main ] && [ "$dirty" = 0 ] && good "$r: main, clean" || bad "$r: on '$br' with $dirty dirty entries"
done

hdr "LOCAL dev login-skip still works (owner: must stay functional on this machine)"
prog=$(cat "$HF/Program.cs" 2>/dev/null)
if [ -f "$HF/Services/DevAutoLogin.cs" ] && [[ "$prog" == *"app.UseDevAutoLogin();"* ]]; then good "Services/DevAutoLogin.cs + Program.cs registration present locally"
else bad "the LOCAL dev login-skip was removed — the strip must only touch staging clones"; fi

hdr "Result"
if [ $FAIL = 0 ]; then echo "   ALL CHECKS PASSED — $shipped repo(s) shipped since the gate"
else echo "   FAILURES ABOVE — do not publish FlexCore or record the deploy as done"; fi
exit $FAIL
