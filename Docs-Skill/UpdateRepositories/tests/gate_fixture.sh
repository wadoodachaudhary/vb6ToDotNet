#!/bin/bash
# Fixture test for scripts/pre_deploy_gate.sh: builds a throwaway copy of the whole layout (four bare
# remotes, staging clones, app/FlexKit/FlexCore source repos, a feed) and checks each gate branch.
# Touches nothing real.   bash tests/gate_fixture.sh
set -uo pipefail
GATE="$(cd "$(dirname "$0")/.." && pwd)/scripts/pre_deploy_gate.sh"
T=$(mktemp -d); export HOME=$T/home; mkdir -p $HOME
export UR_V=$T/V QUIET_MINUTES=0
V=$T/V; H=$V/HomeFront; HF=$H/MobileSource/HomeFront; D=$H/Deploy/repos
g() { git -c user.email=t@t -c user.name=t "$@"; }
mkdir -p $D $H/Deploy/local-packages
mkrepo() { mkdir -p "$1"; g -C "$1" init -q; }
mkrepo $HF; mkdir -p $HF/P; printf 'shared page body line\n' > $HF/P/Page.razor; g -C $HF add -A; g -C $HF commit -qm init
mkrepo $V/FlexKit; printf '<Project><PropertyGroup><Version>9.9.9</Version></PropertyGroup></Project>\n' > $V/FlexKit/FlexKit.csproj
mkdir -p $V/FlexKit/wwwroot; printf 'class A {}\n' > $V/FlexKit/A.cs; printf 'png\n' > $V/FlexKit/wwwroot/img.png
g -C $V/FlexKit add -A; g -C $V/FlexKit commit -qm init; g -C $V/FlexKit checkout -qb telerik-parity-20260904
mkrepo $V/FlexCore; printf '<Project><PropertyGroup><Version>1.0.0</Version></PropertyGroup></Project>\n' > $V/FlexCore/FlexCore.csproj; g -C $V/FlexCore add -A; g -C $V/FlexCore commit -qm init
for r in hyphen-pb homefront flexkit flexcore; do
  g init -q --bare $T/$r.git
  g clone -q $T/$r.git $D/$r 2>/dev/null
  mkdir -p $D/$r/P; printf 'shared page body line\n' > $D/$r/P/Page.razor
  g -C $D/$r add -A; g -C $D/$r commit -qm "Deploy $r — base"; g -C $D/$r branch -q -M main; g -C $D/$r push -q origin main 2>/dev/null
done
g -C $D/hyphen-pb push -q origin main:R1-UAT 2>/dev/null
# what verify_deploy.sh inspects: local dev login-skip, app branches with FLogin + csproj + package
mkdir -p $HF/Services; printf 'class DevAutoLogin {}\n' > $HF/Services/DevAutoLogin.cs
printf 'app.UseAuthorization();\napp.UseDevAutoLogin();\n' > $HF/Program.cs
for r in hyphen-pb homefront; do
  mkdir -p $D/$r/Components/Pages; printf '@if (Sec.EnforcePasswordCheck) {}\n' > $D/$r/Components/Pages/FLogin.razor
  g -C $D/$r add -A; g -C $D/$r commit -qm "Deploy $r — login page"; g -C $D/$r push -q origin main 2>/dev/null
done
printf '<Project><ItemGroup><PackageReference Include="FlexKit" Version="9.9.9" /></ItemGroup></Project>\n' > $D/hyphen-pb/HomeFront.csproj
mkdir -p $D/hyphen-pb/local-packages; printf 'pkg\n' > $D/hyphen-pb/local-packages/FlexKit.9.9.9.nupkg
g -C $D/hyphen-pb add -A; g -C $D/hyphen-pb commit -qm "Deploy hyphen-pb — package"; g -C $D/hyphen-pb push -q origin main 2>/dev/null
sleep 1; touch $H/Deploy/local-packages/FlexKit.9.9.9.nupkg
fails=0
check() {  # check <label> <expect: pass|block> [pattern that must appear]
  out=$(bash "$GATE" 2>&1); rc=$?
  want=$([ "$2" = pass ] && echo 0 || echo 1)
  ok=PASS; [ "$rc" = "$want" ] || ok=FAIL
  [ -n "${3:-}" ] && ! printf '%s' "$out" | grep -q -- "$3" && ok=FAIL
  [ $ok = FAIL ] && { fails=$((fails+1)); printf '%s\n' "$out" | sed 's/^/      | /'; }
  printf '  %-58s rc=%s expected %-5s %s\n' "$1" "$rc" "$2" "$ok"
}
check "A. clean layout" pass "GATE PASSED"
[ -s $H/Deploy/.update-repositories-gate.state ] && echo "     state file written: $(grep -c = $H/Deploy/.update-repositories-gate.state) keys" || { echo "     state file MISSING"; fails=$((fails+1)); }

# B. teammate commit, then a push-mode run (e.g. --dry-run) resets the clone: check 1 is blind, check 2 is not
g clone -q $T/hyphen-pb.git $T/mate 2>/dev/null; printf 'shared page body line\nteammate distinctive addition\n' > $T/mate/P/Page.razor
g -C $T/mate commit -qam "teammate fix"; g -C $T/mate push -q origin HEAD:main 2>/dev/null
g -C $D/hyphen-pb fetch -q; g -C $D/hyphen-pb reset -q --hard origin/main
check "B. upstream hidden by a push-mode reset -> content check" block "REVERT: hyphen-pb P/Page.razor"
printf 'shared page body line\nteammate distinctive addition\n' > $HF/P/Page.razor
check "B2. after merging the teammate line into our tree" pass "every unshipped teammate change is in our trees"

# C. conflict markers written into source by a failed hand-merge
printf '<<<<<<< ours\nx\n=======\ny\n>>>>>>> theirs\n' > $HF/P/Broken.razor
check "C. conflict markers in the app" block "merge-conflict markers found"
rm $HF/P/Broken.razor

# D. unresolved --pull conflicts recorded
printf 'P/Page.razor\n' > $H/Deploy/pull-conflicts.txt
check "D. pull-conflicts.txt present" block "unresolved --pull conflicts"
rm $H/Deploy/pull-conflicts.txt

# E. an asset changed under an unchanged FlexKit version (the script would not repack)
sleep 1; touch $V/FlexKit/wwwroot/img.png
check "E. FlexKit asset newer than the package" block "only repacks on code changes"
touch $H/Deploy/local-packages/FlexKit.9.9.9.nupkg

# F. code changed under a version already in the NuGet cache
mkdir -p $HOME/.nuget/packages/flexkit/9.9.9; sleep 1; touch $V/FlexKit/A.cs
check "F. repack under a burned FlexKit version" block "already extracted"
rm -rf $HOME/.nuget/packages/flexkit/9.9.9; touch $H/Deploy/local-packages/FlexKit.9.9.9.nupkg

# G. a rejected push left an unpushed deploy commit in a clone
printf 'x\n' > $D/flexcore/P/extra.txt; g -C $D/flexcore add -A; g -C $D/flexcore commit -qm "Deploy flexcore — rejected"
check "G. unpushed commit left by a failed push" block "UNPUSHED commit"

VERIFY="$(cd "$(dirname "$0")/.." && pwd)/scripts/verify_deploy.sh"
vcheck() {  # vcheck <label> <expect: pass|fail> <pattern>
  out=$(bash "$VERIFY" 2>&1); rc=$?
  want=$([ "$2" = pass ] && echo 0 || echo 1); ok=PASS; [ "$rc" = "$want" ] || ok=FAIL
  printf '%s' "$out" | grep -q -- "$3" || ok=FAIL
  [ $ok = FAIL ] && { fails=$((fails+1)); printf '%s\n' "$out" | sed 's/^/      | /'; }
  printf '  %-58s rc=%s expected %-5s %s\n' "$1" "$rc" "$2" "$ok"
}
g -C $D/flexcore reset -q --hard origin/main          # undo scenario G
rm -f $H/Deploy/.update-repositories-gate.state
vcheck "H. verify with no gate state" fail "no gate state"
bash "$GATE" >/dev/null 2>&1                           # gate passes -> baseline recorded
vcheck "I. verify right after the gate, nothing deployed" fail "nothing was deployed"
sleep 1; printf 'y\n' > $D/flexkit/P/new.txt; g -C $D/flexkit add -A; g -C $D/flexkit commit -qm "Deploy flexkit — run"; g -C $D/flexkit push -q origin main 2>/dev/null
vcheck "J. verify after a real deploy commit reached origin" pass "shipped"
printf 'z\n' > $D/homefront/P/z.txt; g -C $D/homefront add -A; g -C $D/homefront commit -qm "Deploy homefront — rejected"
vcheck "K. verify after a push was rejected (unpushed commit)" fail "UNPUSHED commit"
echo "  ---- $fails failure(s)"
rm -rf $T
exit $fails
