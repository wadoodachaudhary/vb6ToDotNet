#!/bin/bash
# run_harnesses.sh — run every database-free console harness in the app.
# The harnesses are csproj-EXCLUDED and use reflection, so a clean HomeFront.sln build proves nothing about
# them: signature/option changes compile and then throw at runtime. Exits 1 if any fail.
#
#   bash run_harnesses.sh           # run them
#   bash run_harnesses.sh --list    # show what would run and what is skipped
#
# Skipped automatically: projects on Sdk="Microsoft.NET.Sdk.Web". Those are browser-driven fixture HOSTS
# (e.g. InputDialogBrowserChecks: `dotnet run --urls http://127.0.0.1:0`, then `node browser.mjs <url>`,
# per its README) — started bare they never exit, or fail to bind :5000 (AirPlay owns it on macOS).
set -uo pipefail
LIST=false; [ "${1:-}" = "--list" ] && LIST=true
HF=/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront
LOGS="${TMPDIR:-/tmp}/update-repositories-harness-$(date +%Y%m%d-%H%M%S)"
TIMEOUT=420
fail=0; total=0; skipped=0
cd "$HF" || exit 1
$LIST || mkdir -p "$LOGS"

# Timeout that kills the whole process group: `dotnet run` starts the app as a child, and an alarm that
# only kills `dotnet run` leaves the child running.
run_bounded() {
  perl -e '
    my $t = shift; my $pid = fork();
    die "fork failed" unless defined $pid;
    if ($pid == 0) { setpgrp(0, 0); exec @ARGV or exit 127; }
    local $SIG{ALRM} = sub { kill "TERM", -$pid; sleep 3; kill "KILL", -$pid; exit 124; };
    alarm $t; waitpid($pid, 0); alarm 0; exit($? >> 8);
  ' "$TIMEOUT" "$@"
}

for d in verification/*/; do
  proj=$(ls "$d"*.csproj 2>/dev/null | head -1); [ -z "$proj" ] && continue
  name=$(basename "$d")
  if grep -q 'Sdk="Microsoft.NET.Sdk.Web"' "$proj"; then
    skipped=$((skipped + 1)); printf '  %-26s SKIPPED  browser-driven fixture host — run it by hand per %sREADME.md\n' "$name" "$d"
    continue
  fi
  total=$((total + 1))
  if $LIST; then printf '  %-26s would run\n' "$name"; continue; fi
  ( cd "$d" && run_bounded dotnet run --project "$(basename "$proj")" -c Release > "$LOGS/$name.log" 2>&1 )
  rc=$?
  [ $rc -ne 0 ] && fail=$((fail + 1))
  printf '  %-26s rc=%-3s %s\n' "$name" "$rc" "$(grep -vE '^\s*$|warning |Determining|Restored| -> |^Build succeeded' "$LOGS/$name.log" | tail -1 | cut -c1-80)"
  [ $rc -eq 124 ] && printf '  %-26s          timed out after %ss (process group killed)\n' "" "$TIMEOUT"
done
if $LIST; then echo "  ---- $total would run, $skipped skipped"; exit 0; fi
echo "  ---- $total harnesses, $fail failed, $skipped skipped. Logs: $LOGS"
[ $fail -eq 0 ]
