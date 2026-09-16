#!/bin/bash
# verify_deploy.sh — prove what a deploy actually pushed. Run AFTER deploy_to_repos.sh.
# Read-only apart from `git fetch` in the staging clones. Exits 1 if any hard check fails.
#
#   bash verify_deploy.sh               # normal run (R1-UAT frozen, not deployed)
#   bash verify_deploy.sh --with-r1uat  # the run also deployed R1-UAT
set -uo pipefail

WITH_R1UAT=false; [ "${1:-}" = "--with-r1uat" ] && WITH_R1UAT=true
V=/Users/wadood/projects/VBToCSharp
H=$V/HomeFront
HF=$H/MobileSource/HomeFront
D=$H/Deploy/repos
FAIL=0
bad()  { printf '   ✗ %s\n' "$1"; FAIL=1; }
good() { printf '   ✓ %s\n' "$1"; }
hdr()  { printf '\n== %s\n' "$1"; }

for r in hyphen-pb homefront flexkit flexcore; do git -C "$D/$r" fetch -q origin 2>/dev/null || bad "$r: fetch failed"; done

hdr "Remote heads"
printf '   hyphen-pb main=%s  R1-UAT=%s\n   homefront=%s  flexkit=%s  flexcore(GitHub)=%s\n' \
  "$(git -C $D/hyphen-pb rev-parse --short origin/main)" "$(git -C $D/hyphen-pb rev-parse --short origin/R1-UAT)" \
  "$(git -C $D/homefront rev-parse --short origin/main)" "$(git -C $D/flexkit rev-parse --short origin/main)" \
  "$(git -C $D/flexcore rev-parse --short origin/main)"

# Branches the app ships on, and the files that must NOT be on any pushed branch.
APP_REFS="hyphen-pb:origin/main homefront:origin/main"
$WITH_R1UAT && APP_REFS="$APP_REFS hyphen-pb:origin/R1-UAT"
NON_RUNTIME='^(CLAUDE\.md|AGENTS\.md|GEMINI\.md|\.agents/|\.claude/|\.codex/|\.codex-backups/|verification/|tests/|docs/|Docs/|PUBLISHING\.md)|\.vb6_gap\.txt$'

hdr "Login-skip must be gone from every pushed app branch (owner 2026-09-12)"
for spec in $APP_REFS; do
  repo=${spec%%:*}; ref=${spec#*:}
  hits=$(git -C "$D/$repo" grep -I -i -l -e devautologin "$ref" -- ':!*.md' 2>/dev/null | sed "s|^$ref:||" | tr '\n' ' ')
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
  if [ -n "$login" ] && git -C "$D/$repo" cat-file blob "$ref:$login" | grep -q 'Sec.EnforcePasswordCheck'; then good "$repo $ref: $login consults Sec.EnforcePasswordCheck"
  else bad "$repo $ref: FLogin.razor missing or no longer checks the password"; fi
done

hdr "FlexKit package on the app branch matches its PackageReference"
ref=origin/main
pr=$(git -C "$D/hyphen-pb" cat-file blob "$ref:HomeFront.csproj" 2>/dev/null | grep -o 'Include="FlexKit" Version="[^"]*"' | grep -o '[0-9][0-9.]*')
pk=$(git -C "$D/hyphen-pb" ls-tree --name-only "$ref" local-packages/ 2>/dev/null | grep -o 'FlexKit\.[0-9.]*[0-9]\.nupkg' | sed 's/FlexKit\.//; s/\.nupkg//' | head -1)
[ -n "$pr" ] && [ "$pr" = "$pk" ] && good "hyphen-pb main: PackageReference $pr == local-packages $pk" || bad "hyphen-pb main: PackageReference '$pr' vs local-packages '$pk'"

hdr "R1-UAT"
if $WITH_R1UAT; then
  n=$(git -C "$D/hyphen-pb" cat-file blob origin/R1-UAT:Components/Pages/FLogin.razor 2>/dev/null | grep -o 'Disabled="true"' | wc -l | tr -d ' ')
  [ "$n" -eq 5 ] && good "lockdown intact (5/5 Disabled toggles)" || bad "lockdown broken ($n/5)"
  m1=$(git -C "$D/hyphen-pb" cat-file blob "origin/main:local-packages/FlexKit.$pk.nupkg" 2>/dev/null | md5)
  m2=$(git -C "$D/hyphen-pb" cat-file blob "origin/R1-UAT:local-packages/FlexKit.$pk.nupkg" 2>/dev/null | md5)
  [ -n "$m1" ] && [ "$m1" = "$m2" ] && good "FlexKit.$pk.nupkg byte-identical on main and R1-UAT" || bad "main and R1-UAT carry DIFFERENT FlexKit packages"
else
  echo "   frozen — not deployed this run (head $(git -C "$D/hyphen-pb" rev-parse --short origin/R1-UAT))"
fi

hdr "Staging clones handed back clean and on main"
for r in hyphen-pb homefront flexkit flexcore; do
  br=$(git -C "$D/$r" branch --show-current); dirty=$(git -C "$D/$r" status --porcelain | wc -l | tr -d ' ')
  [ "$br" = main ] && [ "$dirty" = 0 ] && good "$r: main, clean" || bad "$r: on '$br' with $dirty dirty entries"
done

hdr "LOCAL dev login-skip still works (owner: must stay functional on this machine)"
[ -f "$HF/Services/DevAutoLogin.cs" ] && grep -q 'app.UseDevAutoLogin();' "$HF/Program.cs" \
  && good "Services/DevAutoLogin.cs + Program.cs registration present locally" \
  || bad "the LOCAL dev login-skip was removed — the strip must only touch staging clones"

hdr "Result"
[ $FAIL = 0 ] && echo "   ALL CHECKS PASSED" || echo "   FAILURES ABOVE — do not publish FlexCore or record the deploy as done"
exit $FAIL
