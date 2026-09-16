#!/bin/bash
# publish_flexcore.sh — publish FlexCore to nuget.org from the snapshot that was PUSHED to GitHub.
# A NuGet version is permanent (unlist-only), so every check here fails closed.
#
#   bash publish_flexcore.sh --check   # every step except the push (packs and inspects)
#   bash publish_flexcore.sh           # publish
#
# Never packs the live FlexCore working tree: reviews can hold the publish for hours while other sessions
# keep editing it. The package is built from a detached worktree of the flexcore staging clone at
# origin/main — byte-for-byte what GitHub has, which also makes SourceLink name a commit that exists there.
set -uo pipefail
CHECK=false
for a in "$@"; do case "$a" in --check) CHECK=true ;; *) echo "✗ unknown argument '$a'"; exit 2 ;; esac; done
source "$(cd "$(dirname "$0")" && pwd)/_common.sh"
C=$D/flexcore
die() { printf '✗ %s\n' "$1"; cleanup; exit 1; }
ok()  { printf '✓ %s\n' "$1"; }
# Physical path: macOS /var is a symlink to /private/var, and a project under /var/folders/... makes the
# Razor generator see two different roots, drop to bare file names, and fail on duplicate names
# (Editor/TextAreaControl.razor vs TextAreaControl.razor).
T=$(cd "$(mktemp -d)" && pwd -P)
cleanup() { git -C "$C" worktree remove --force "$T/src" >/dev/null 2>&1; git -C "$C" worktree prune >/dev/null 2>&1; rm -rf "$T"; }

# 1. The snapshot must be a real, completed push.
git -C "$C" fetch -q origin 2>/dev/null || die "fetch of the flexcore staging clone failed"
[ "$(git -C "$C" branch --show-current)" = main ] || die "flexcore staging clone is not on main"
[ "$(git -C "$C" rev-list --count origin/main..main)" = 0 ] || die "flexcore staging clone has UNPUSHED commits — the deploy did not complete"
snap=$(git -C "$C" rev-parse --short origin/main); subj=$(git -C "$C" log -1 --format=%s origin/main)
case "$subj" in "Deploy flexcore"*) ok "snapshot $snap — $subj" ;; *) die "GitHub main tip $snap is not a deploy commit ('$subj')" ;; esac
if [ -s "$GATE_STATE" ]; then
  gt=$(grep -m1 '^time=' "$GATE_STATE" | cut -d= -f2); age=$(( ($(date +%s) - gt) / 60 ))
  [ "$age" -le 720 ] || die "gate state is $age min old — run the full procedure (gate, deploy, verify) first"
  ok "gate passed $age min ago"
else die "no gate state — publish only after pre_deploy_gate.sh, the deploy, and verify_deploy.sh"; fi

# 2. Version: taken from the snapshot, must match the source, must be new and above the latest.
ver=$(git -C "$C" show origin/main:FlexCore.csproj 2>/dev/null | grep -o '<Version>[^<]*' | head -1 | cut -c10-)
[ -n "$ver" ] || die "cannot read <Version> from the snapshot's FlexCore.csproj"
src_ver=$(grep -o '<Version>[^<]*' "$V/FlexCore/FlexCore.csproj" | head -1 | cut -c10-)
[ "$ver" = "$src_ver" ] || die "pushed snapshot is $ver but FlexCore.csproj now says $src_ver — deploy again before publishing"
vcheck=$(curl -s --max-time 20 https://api.nuget.org/v3-flatcontainer/flexcore/index.json | python3 -c '
import sys, json
want = sys.argv[1]
try: vs = json.load(sys.stdin)["versions"]
except Exception: print("UNKNOWN nuget.org unreachable"); sys.exit()
key = lambda v: tuple(int(x) if x.isdigit() else 0 for x in v.split("-")[0].split("."))
latest = max(vs, key=key)
if want in vs: print(f"PUBLISHED {want} is already on nuget.org")
elif key(want) <= key(latest): print(f"LOWER {want} is not above the latest published {latest}")
else: print(f"OK {want} is above the latest published {latest}")' "$ver")
case "$vcheck" in
  OK*) ok "${vcheck#OK }" ;;
  *) if $CHECK; then printf '! %s (would block a real publish)\n' "${vcheck#* }"; else die "${vcheck#* } — bump FlexCore.csproj, deploy, then publish"; fi ;;
esac

# 3. Pack the snapshot.
git -C "$C" worktree add -q --detach "$T/src" origin/main 2>/dev/null || die "could not create a worktree of the snapshot"
if ! dotnet pack "$T/src/FlexCore.csproj" -c Release -o "$T/out" -v q -nologo > "$T/pack.log" 2>&1; then
  tail -15 "$T/pack.log"; die "dotnet pack failed"
fi
nupkg="$T/out/FlexCore.$ver.nupkg"
[ -f "$nupkg" ] || die "pack produced no FlexCore.$ver.nupkg"
ok "packed $(basename "$nupkg") ($(du -h "$nupkg" | cut -f1 | tr -d ' '))"

# 4. Inspect what would be published.
listing=$(unzip -Z1 "$nupkg" 2>/dev/null)
[[ "$listing" == *"lib/net10.0/FlexCore.dll"* ]] || die "package has no lib/net10.0/FlexCore.dll"
[[ "$listing" == *"README.md"* ]] || die "package has no README.md"
leak=$(printf '%s\n' "$listing" | grep -iE 'FlexCore\.Llm|(^|/)CLAUDE\.md$|(^|/)AGENTS\.md$|(^|/)docs/|PUBLISHING\.md|(^|/)tests/' | head -5)
[ -z "$leak" ] || die "package contains files that must not ship: $leak"
ok "contents: FlexCore.dll + README present; no FlexCore.Llm, agent notes, docs or tests"

if $CHECK; then ok "CHECK ONLY — not pushed"; cleanup; exit 0; fi

# 5. Push. No --skip-duplicate: a duplicate means the version was not bumped, which is a failure.
#    The key is read inside the command substitution and never echoed; long tokens are redacted.
out=$(dotnet nuget push "$nupkg" --api-key "$(security find-generic-password -s flexcore-nuget-key -w)" \
      --source https://api.nuget.org/v3/index.json 2>&1); rc=$?
printf '%s\n' "$out" | grep -viE 'api-?key' | sed -E 's/[A-Za-z0-9_-]{40,}/[redacted]/g' | sed 's/^/   /'
if [ $rc -ne 0 ] || [[ "$out" != *"Your package was pushed"* ]]; then die "NuGet push FAILED (rc=$rc) — FlexCore $ver is NOT published"; fi
ok "pushed FlexCore $ver"

# 6. Confirm nuget.org lists it (indexing can take a few minutes).
for i in $(seq 1 12); do
  if curl -s --max-time 20 https://api.nuget.org/v3-flatcontainer/flexcore/index.json | grep -q "\"$ver\""; then
    ok "nuget.org lists FlexCore $ver"; cleanup; exit 0
  fi
  sleep 30
done
printf '! pushed, but nuget.org did not list %s within 6 minutes — check https://www.nuget.org/packages/FlexCore before recording it\n' "$ver"
cleanup; exit 0
