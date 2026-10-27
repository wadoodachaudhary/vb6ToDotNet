#!/bin/bash
# report_regression.sh — Report Writer regression over the nineteen HomeFront reports of
# MobileSource/HomeFront/wwwroot/resources/ReportTest.zip (owner 2026-10-26: "Use it when Report Writer is modified
# for regression testing. Add this as regression for update repositories session as well.").
#
#   bash report_regression.sh                  # FlexCore, then FlexKit, both against the one committed baseline set
#   bash report_regression.sh --kit | --core   # one library only
#   bash report_regression.sh --accept         # FlexKit rewrites the baselines, FlexCore must then agree —
#                                              #   only after the owner-intended change was reviewed page by page
#   bash report_regression.sh menu06 menu13    # case ids (verification/ReportRegressionChecks/cases.json) pass through
#
# Per report it proves: the .rpt and the VB reference PDF in the zip are the catalogued ones; the report converts,
# queries the database with the catalogued prompt values, paginates with real browser metrics and exports a PDF;
# no subreport error and no diagnostic other than the Calibri/Carlito font note; the same number of pages as
# VB/Crystal printed; and every page's band contents equal the committed baseline
# (verification/ReportRegressionChecks/baseline/<id>.json, order inside a page ignored — rows with equal sort keys
# follow the server's order). Both libraries run against the same baselines, so a FlexKit→FlexCore mirror that
# drifted shows here too. run_harnesses.sh leaves this project out of its sweep on purpose (csproj <HarnessSweep>).
#
# Needs: the SQL container (docker start sqlserver_new — port 1433), Chrome, Node with Playwright (NODE_PATH or
# PLAYWRIGHT_MODULE, else the newest ~/.npm/_npx install is used), pdfinfo (poppler). About 25 s per library.
# Output (XML, SQL, snapshots, HTML, PDF per case) and logs: ${TMPDIR:-/tmp}/update-repositories-report-regression-<stamp>/.
set -uo pipefail
HF=/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront
CORE=/Users/wadood/projects/VBToCSharp/FlexCore/FlexCore.csproj
PROJECT=verification/ReportRegressionChecks
OUT="${TMPDIR:-/tmp}"; OUT="${OUT%/}/update-repositories-report-regression-$(date +%Y%m%d-%H%M%S)"
kit=true; core=true; accept=false; pass=()
for a in "$@"; do
  case "$a" in
    --kit) core=false ;; --core) kit=false ;; --accept) accept=true ;;
    -h|--help) sed -n '2,22p' "$0"; exit 0 ;;
    *) pass+=("$a") ;;
  esac
done
cd "$HF" || exit 1
[ -f "$PROJECT/cases.json" ] || { echo "✗ $HF/$PROJECT/cases.json missing"; exit 2; }
# The harness reads the zip, or the extracted ReportTest/ PDFs when the zip is gone, with the same hash and page-count checks.
[ -f wwwroot/resources/ReportTest.zip ] || [ -d wwwroot/resources/ReportTest ] || { echo "✗ wwwroot/resources/ReportTest.zip (or its extracted ReportTest/ folder) missing (the owner's 19-report reference set)"; exit 2; }
nc -z -w 2 localhost 1433 >/dev/null 2>&1 || { echo "✗ SQL Server is not listening on localhost:1433 — docker start sqlserver_new"; exit 2; }
command -v node >/dev/null 2>&1 || { echo "✗ node not on PATH (Playwright drives the measurement browser)"; exit 2; }
command -v pdfinfo >/dev/null 2>&1 || echo "  (pdfinfo not found: reference page counts fall back to the PDF's page tree)"
mkdir -p "$OUT"
rc=0
run() {  # run <label> <ControlKitProject or ""> [harness args...]
  local label=$1 lib=$2; shift 2
  local props=(); [ -n "$lib" ] && props=(-p:ControlKitProject="$lib")
  echo "== $label"
  dotnet run -c Release --project "$PROJECT" ${props[@]+"${props[@]}"} -- "$OUT/$label" "$@" > "$OUT/$label.log" 2>&1
  local r=$?
  grep -E '^(Library |FAIL:|[0-9]+/[0-9]+ report regression checks|menu[^ ]*: baseline written)' "$OUT/$label.log" | sed 's/^/  /'
  if [ $r -ne 0 ]; then rc=1; echo "  rc=$r — log: $OUT/$label.log"; grep -E 'Unhandled exception|Exception:' "$OUT/$label.log" | head -3 | sed 's/^/  /'; fi
}
if $accept; then
  # The app's library writes the baselines; FlexCore is then held to them.
  $kit && run kit "" --accept ${pass[@]+"${pass[@]}"}
  $core && run core "$CORE" ${pass[@]+"${pass[@]}"}
  echo "  baselines rewritten in $HF/$PROJECT/baseline — review the diff and commit them with the change"
else
  $core && run core "$CORE" ${pass[@]+"${pass[@]}"}
  $kit && run kit "" ${pass[@]+"${pass[@]}"}      # last, so the harness binary is left built against FlexKit
fi
if [ $rc -eq 0 ]; then echo "✓ Report Writer regression: the nineteen reports print as baselined. Output: $OUT"
else echo "✗ Report Writer regression FAILED — a stale baseline (intended change: review, then --accept) or a regression. Output: $OUT"; fi
exit $rc
