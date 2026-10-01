#!/bin/bash
# Fixture for deploy_r2uat (tools/deploy_to_repos.sh): throwaway origin + staging clone, the REAL
# function extracted from the script, verify_build stubbed. Touches nothing real.
#   bash tests/r2uat_fixture.sh
set -uo pipefail
SKILL_DIR="$(cd "$(dirname "$0")/.." && pwd)"          # resolve before any cd: $0 may be relative
SCRIPT="$(cd "$SKILL_DIR/../.." && pwd)/tools/deploy_to_repos.sh"
FAILS=0
t() { local name="$1" want="$2" got="$3"; if [ "$want" = "$got" ]; then printf '  %-62s PASS\n' "$name"; else printf '  %-62s FAIL (want %s, got %s)\n' "$name" "$want" "$got"; FAILS=$((FAILS+1)); fi; }

LOCK='<CheckBoxControl @bind-Checked="Sec.EnforceAuth" Disabled="true" />
<CheckBoxControl @bind-Checked="Sec.EncryptDbConnection" Disabled="true" />
<CheckBoxControl @bind-Checked="Sec.EnforcePasswordCheck" Disabled="true" />
<CheckBoxControl @bind-Checked="Sec.GetPasswordFromSystemSecrets" Disabled="true" />
Sec.EncryptDbConnection = false;'
OPEN='<CheckBoxControl @bind-Checked="Sec.EnforceAuth" />
<CheckBoxControl @bind-Checked="Sec.EncryptDbConnection" />
<CheckBoxControl @bind-Checked="Sec.EnforcePasswordCheck" />
<CheckBoxControl @bind-Checked="Sec.GetPasswordFromSystemSecrets" />'

setup() {   # $1 = lockdown text on R2-UAT, $2 = pipeline trigger on R2-UAT
  T=$(mktemp -d); O=$T/origin.git; W=$T/work; DEPLOY=$T/repos
  git init -q --bare "$O"; git init -q "$W"; cd "$W"; git checkout -q -b main
  git config user.email f@x; git config user.name fixture
  mkdir -p Components/Pages/Migrated; printf 'header\n%s\nfooter\n' "$OPEN" > Components/Pages/Migrated/FLogin.razor
  printf 'trigger:\n  branches:\n    include:\n      - R1-UAT\n' > azure-pipelines-uat.yml; echo one > page.razor
  git add -A; git commit -qm base; git push -q "$O" main
  git checkout -q -b R2-UAT
  printf 'header\n%s\nfooter\n' "$1" > Components/Pages/Migrated/FLogin.razor
  printf 'trigger:\n  branches:\n    include:\n      - %s\n' "$2" > azure-pipelines-uat.yml
  git commit -qam "R2-UAT lockdown"; git push -q "$O" R2-UAT; git checkout -q main
  mkdir -p "$DEPLOY"; git clone -q "$O" "$DEPLOY/hyphen-pb"; git -C "$DEPLOY/hyphen-pb" config user.email f@x; git -C "$DEPLOY/hyphen-pb" config user.name fixture
}
advance_main() { cd "$W"; git checkout -q main; "$@"; git add -A; git commit -qm "main moves"; git push -q "$O" main; }
run() {     # runs the real function; prints its exit code
  ( log(){ :; }; ok(){ :; }; err(){ :; }; verify_build(){ return ${STUB_BUILD:-0}; }
    DRY_RUN=${DRY:-false}
    eval "$(sed -n '/^deploy_r2uat() {/,/^}/p' "$SCRIPT")"
    deploy_r2uat >/dev/null 2>&1; echo $? )
}
r2() { git -C "$O" rev-parse R2-UAT; }
contains_main() { git -C "$O" merge-base --is-ancestor main R2-UAT && echo yes || echo no; }
parked() { echo "$(git -C "$DEPLOY/hyphen-pb" branch --show-current):$( [ -f "$DEPLOY/hyphen-pb/.git/MERGE_HEAD" ] && echo merging || echo clean )"; }

echo "deploy_r2uat fixture"
setup "$LOCK" R2-UAT; advance_main sh -c 'echo two > page.razor'; before=$(r2)
t "A. clean merge: exit 0" 0 "$(run)"; t "A. R2-UAT now contains main" yes "$(contains_main)"
t "A. lockdown kept on the pushed branch" 4 "$(git -C "$O" show R2-UAT:Components/Pages/Migrated/FLogin.razor | grep -c 'Disabled="true"')"
t "A. clone handed back on main, no merge in progress" "main:clean" "$(parked)"

setup "$LOCK" R2-UAT; before=$(r2)
t "B. nothing to merge: exit 0" 0 "$(run)"; t "B. branch untouched" "$before" "$(r2)"

setup "$LOCK" R2-UAT; advance_main sh -c "printf 'header\nMAIN REWROTE THE TOGGLES\nfooter\n' > Components/Pages/Migrated/FLogin.razor"; before=$(r2)
t "C. conflict in FLogin: exit 1" 1 "$(run)"; t "C. nothing pushed" "$before" "$(r2)"; t "C. clone back on main, merge aborted" "main:clean" "$(parked)"

setup "$OPEN" R2-UAT; advance_main sh -c 'echo two > page.razor'; before=$(r2)
t "D. lockdown missing on the branch: exit 1" 1 "$(run)"; t "D. nothing pushed" "$before" "$(r2)"; t "D. clone back on main" "main:clean" "$(parked)"

setup "$LOCK" R1-UAT; advance_main sh -c 'echo two > page.razor'; before=$(r2)
t "E. pipeline not targeting R2-UAT: exit 1" 1 "$(run)"; t "E. nothing pushed" "$before" "$(r2)"

setup "$LOCK" R2-UAT; advance_main sh -c 'echo two > page.razor'; before=$(r2)
t "F. dry run: exit 0" 0 "$(DRY=true run)"; t "F. dry run pushes nothing" "$before" "$(r2)"; t "F. clone back on main, merge aborted" "main:clean" "$(parked)"

setup "$LOCK" R2-UAT; advance_main sh -c 'echo two > page.razor'; before=$(r2)
t "G. build fails: exit 1" 1 "$(STUB_BUILD=1 run)"; t "G. nothing pushed" "$before" "$(r2)"; t "G. clone back on main" "main:clean" "$(parked)"

setup "$LOCK" R2-UAT; advance_main sh -c 'mkdir -p Services; echo "class DevAutoLogin{}" > Services/Leak.cs'; before=$(r2)
t "H. login-skip leaked into main: exit 1" 1 "$(run)"; t "H. nothing pushed" "$before" "$(r2)"

echo "  ---- $FAILS failure(s)"; exit $FAILS
