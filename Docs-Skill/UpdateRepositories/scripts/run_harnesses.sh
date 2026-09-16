#!/bin/bash
# run_harnesses.sh — run every database-free verification harness in the app.
# The harnesses are csproj-EXCLUDED and use reflection, so a clean HomeFront.sln build proves
# nothing about them: signature/option changes compile and then throw at runtime.
# macOS has no GNU `timeout`; a perl alarm bounds each run instead. Exits 1 if any fail.
set -uo pipefail
HF=/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront
LOGS="${TMPDIR:-/tmp}/update-repositories-harness-$(date +%Y%m%d-%H%M%S)"
mkdir -p "$LOGS"
fail=0; total=0
cd "$HF" || exit 1
for d in verification/*/; do
  proj=$(ls "$d"*.csproj 2>/dev/null | head -1); [ -z "$proj" ] && continue
  name=$(basename "$d"); total=$((total + 1))
  ( cd "$d" && perl -e 'alarm shift; exec @ARGV' 420 dotnet run --project "$(basename "$proj")" -c Release > "$LOGS/$name.log" 2>&1 )
  rc=$?
  [ $rc -ne 0 ] && fail=$((fail + 1))
  printf '  %-24s rc=%-3s %s\n' "$name" "$rc" "$(grep -vE '^\s*$|warning |Determining|Restored| -> |^Build succeeded' "$LOGS/$name.log" | tail -1 | cut -c1-80)"
done
echo "  ---- $total harnesses, $fail failed. Logs: $LOGS"
[ $fail -eq 0 ]
