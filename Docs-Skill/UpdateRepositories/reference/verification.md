# Step 7 — the evidence matrix

`run_harnesses.sh` is the floor: it runs one half of one of the four bodies of evidence this project
has. Everything below was checked against the project files on 2026-09-26; paths are absolute.

Neither solution carries any of it — `HomeFront.sln` holds `HomeFront.csproj` + `FlexKit.csproj`,
`FlexCore.Showcase.sln` holds `FlexCore.Showcase.csproj` + `FlexCore.csproj`, and the app csproj
excludes the harnesses outright (`HomeFront.csproj:7,17-20`). A green build proves nothing about them.

## What run_harnesses.sh actually does

```bash
R=/Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories/scripts/run_harnesses.sh
bash "$R"
bash "$R" --list     # what would run and what is skipped; always exits 0
```

- It `cd`s to the app (`:14,18`) and loops `verification/*/`, taking the first `*.csproj` in each.
  35 directories exist today: 27 run, 8 are skipped.
- **Skip rule:** `grep -q 'Sdk="Microsoft.NET.Sdk.Web"'` on the csproj (`:36`). Those eight are browser
  fixture *hosts* — started bare they never exit. They print as `SKIPPED` in both modes.
- **Working directory:** a harness whose top-level `.cs` mentions `GetCurrentDirectory` runs from the
  app directory, the rest from their own (`:43`). Two qualify —
  `verification/PricingWorksheetChecks/Program.cs` and `verification/GridLayoutSaveChecks/Program.cs`,
  both reading the app's `appsettings.json`; run either from elsewhere and it throws
  `FileNotFoundException`. The grep is top-level `.cs` only, so a harness that moved that call into a
  subdirectory would silently run from the wrong place.
- **Timeout:** `TIMEOUT=420` (`:16`), enforced by a `perl` fork that kills the whole process group
  because `dotnet run` starts the app as a child. A timed-out harness reports `rc=124`.
- **Logs and exit:** `${TMPDIR:-/tmp}/update-repositories-harness-<timestamp>/<Name>.log` — on macOS
  `TMPDIR` is a per-user `/var/folders/.../T/` path, not `/tmp`. Exits 0 only if every harness exited 0
  (`[ $fail -eq 0 ]`). `set -uo pipefail`, no `-e`.
- **It does not source `_common.sh`.** `HF` is hardcoded (`:14`), so no `UR_*` override reaches it and
  no fixture test covers it. The only variable it reads is `TMPDIR`; `QUIET_MINUTES` belongs to
  `pre_deploy_gate.sh:20`, not here.

Record the count each round: `FlexKit/CLAUDE.md:15` says "27 database-free harnesses", and fewer means a
harness lost its csproj or was added without one.

## rc=0 covers the C# stage only

Thirteen console-SDK harnesses also carry a Playwright script the runner never runs and never mentions.
Most are two-stage: the C# run writes static markup, the `.mjs` renders it in Chrome. Run the one whose
area changed, from the app directory.

| Harness | Second stage | Output directory |
|---|---|---|
| BillingChecks | `browser.mjs` (imports `palette.mjs`) | `/tmp/homefront-billing-checks` — **C# writes only when given the argument** (`Program.cs:253`) |
| GridDateCellChecks | `drag-selection.mjs` — no browser, server or Playwright; argv[1] is the FlexKit path | — |
| GridFilterPopupChecks | `browser.mjs` | `/tmp/homefront-grid-filter-popup-checks` (`Program.cs:21`) |
| GridRenameChecks | `browser.mjs` | `/tmp/homefront-grid-rename-checks` (`Program.cs:17`) |
| NavigationChecks | `browser.mjs`, `properties-tree.mjs` | writes JSON graphs only when given paths (`Program.cs:174-175`); scripts take `<graph.json> <FlexKit path>` |
| PasswordChecks | `browser.mjs` | **mismatch:** C# uses `$TMPDIR/homefront-password-checks` (`Program.cs:78`), the script `/tmp/homefront-password-checks` — pass the same directory to both |
| PoHeaderChecks | `browser.mjs` | `/tmp/homefront-po-header-checks` (`Program.cs:36`) |
| PriceComparisonChecks | `layout-browser.mjs` | `/private/tmp/price-comparison-checks` (`Program.cs:15`) |
| SplitterChecks | `browser.mjs` | `/tmp/homefront-splitter-checks` (`Program.cs:19`) |
| TakeoffSettingsChecks | `layout-browser.mjs` | `/private/tmp/takeoff-settings-checks` (`Program.cs:19`) |
| WizardChecks | `reassignment-browser.mjs`, `sending-browser.mjs` | `/private/tmp/hhm1062-wizard-checks`, `/tmp/homefront-wizard-checks`; **C# writes only when given the argument** (`Program.cs:38`) |
| WorksheetValidationChecks | `browser.mjs` | `/private/tmp/r2qa-worksheet-ui` — **C# writes only when given the argument** (`Program.cs:175`) |
| R2QaChecks | `mail-navigation.mjs` (self-contained); three others drive a live app — see below | — |

`run_harnesses.sh` passes no arguments, so for the four marked harnesses the browser stage has nothing
to read. Re-run the C# stage with the directory first:

```bash
cd /Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront
dotnet run --project verification/WorksheetValidationChecks -- /private/tmp/r2qa-worksheet-ui
node verification/WorksheetValidationChecks/browser.mjs /private/tmp/r2qa-worksheet-ui
```

**Debug build prerequisite.** Eleven of these read the Debug scoped-CSS bundles
`FlexKit/obj/Debug/net10.0/scopedcss/projectbundle/FlexKit.bundle.scp.css` and
`<app>/obj/Debug/net10.0/scopedcss/…` (e.g. `WorksheetValidationChecks/browser.mjs:12-15`) while the
runner builds `-c Release` (`:45`). A Debug build must exist or they throw ENOENT on a path that reads
like a missing file rather than a missing build.

**Playwright is machine-local.** Most scripts fall back to
`/Users/wadood/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright`
(`…/playwright/test` where they use `expect`); `GridFilterPopupChecks/latency.mjs:26` uses
`/Users/wadood/.npm/_npx/705bc6b22212b352/node_modules/playwright`. Both resolve today (node v22.16.0);
git carries neither. `PLAYWRIGHT_MODULE` overrides. Six scripts resolve a bare `'playwright'` and need
`NODE_PATH`: `NavigationChecks/browser.mjs`, `NavigationChecks/properties-tree.mjs`,
`R2QaChecks/mail-navigation.mjs`, and the `.cjs` fixtures `GridSortChecks/browser.cjs`,
`AttachmentColumnChecks/browser.cjs`, `DropDownOpeningChecks/browser.cjs` — those three read no env
override at all, so `NODE_PATH` is the only lever.

## The eight browser-driven fixture hosts

Each is its own isolated Blazor Server app with no database. Commands are each project's own
`README.md`; run from `/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront` unless noted.
Never `dotnet run` bare — :5000 is AirPlay's on macOS. **Stop every fixture afterwards.**

| Fixture | Start | Then |
|---|---|---|
| AttachmentColumnChecks | `dotnet run --project verification/AttachmentColumnChecks --no-launch-profile --urls http://127.0.0.1:5310` | `NODE_PATH=<pw>/node_modules node verification/AttachmentColumnChecks/browser.cjs http://127.0.0.1:5310` |
| DropDownOpeningChecks | from its own directory: `dotnet run --urls http://127.0.0.1:5308` | `NODE_PATH=<pw>/node_modules node browser.cjs http://127.0.0.1:5308` |
| EstimateTreeChecks | `dotnet run --project verification/EstimateTreeChecks/EstimateTreeChecks.csproj --no-build --no-launch-profile --urls http://127.0.0.1:5399` | `node verification/EstimateTreeChecks/browser.mjs` |
| GridSortChecks | `dotnet run --project verification/GridSortChecks --urls http://127.0.0.1:5309` (console mode: `-- --checks`) | `NODE_PATH=<pw>/node_modules node verification/GridSortChecks/browser.cjs` |
| InputDialogBrowserChecks | from its own directory: `dotnet run --urls http://127.0.0.1:0` (prints the port) | `node browser.mjs <printed url>` |
| ManualPoChecks | from its own directory: `dotnet build ManualPoChecks.csproj --ignore-failed-sources`, then `dotnet run --project ManualPoChecks.csproj --no-build -- --serve --urls http://127.0.0.1:5407` | `node browser.mjs` |
| OneTimeChecks | from its own directory: `dotnet bin/Debug/net10.0/OneTimeChecks.dll --urls http://127.0.0.1:5398` | `node verification/OneTimeChecks/browser.mjs` |
| TbdAssignmentChecks | from its own directory: `dotnet bin/Debug/net10.0/TbdAssignmentChecks.dll --urls http://127.0.0.1:5302` | `node browser.mjs`, then `node visual.mjs` |

`ManualPoChecks` has no project references and loads the existing
`<app>/bin/Debug/net10.0/HomeFront.dll` and `FlexKit.dll` — both present today; refresh them if child
component contracts changed. `TbdAssignmentChecks` must start from its own directory: it resolves the
app's wwwroot and scoped CSS from its content root, and elsewhere fails with
`DirectoryNotFoundException`. Overrides here: `CHECK_URL` / `CHECK_OUTPUT` (OneTimeChecks,
ManualPoChecks), `BENCH_URL` / `RESULTS_DIR` / `PNG_MODULE` (TbdAssignmentChecks), `PLAYWRIGHT_MODULE`
throughout.

## Suites that drive the OWNER'S live instance — explicit ask only

| Script | Default target |
|---|---|
| `verification/R2QaChecks/browser.mjs:8`, `po-numbering-browser.mjs:6` | `http://127.0.0.1:5071` |
| `verification/R2QaChecks/header-paint-browser.mjs:8` | `http://127.0.0.1:5065` |
| `verification/GridDateCellChecks/date-calendar-browser.mjs:6` | `http://127.0.0.1:5065` |
| `verification/GridFilterPopupChecks/latency.mjs:31` | `HF_BASE ?? http://localhost:5270` |

These hit a running HomeFront with its real database, not a fixture, and are out of the automated step.
`AGENTS.md` §7 is binding: the owner runs HomeFront himself, never auto-start `dotnet run`, live
execution requires an explicit ask, and a session started for one is killed as soon as the work is done.
Memory `dev_autologin_local_bypass` records :5270 as the dev instance and :5065 as the owner's. Each
script asserts a loopback host — a guard against pointing it at QA, not permission to start a server.
QA's port 81 serves the frozen HomeFrontPB (memory `qa_port81_is_homefrontpb`), never a HomeFront fix.

## The cross-library lever

Four harnesses build against either library via `ControlLibraryProject`, default
`../../../../../FlexKit/FlexKit.csproj`: `GridSelectionChecks`, `GridSortChecks`,
`DropDownOpeningChecks`, `InputDialogBrowserChecks` (each csproj, lines 6-9). Cheapest proof a
FlexKit→FlexCore port actually landed:

```bash
cd /Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront
dotnet run --project verification/GridSortChecks \
  -p:ControlLibraryProject=/Users/wadood/projects/VBToCSharp/FlexCore/FlexCore.csproj -- --checks
```

## ReportDesigner — run the runner, not the suites

`-p:ReportSourceDir=` is **dead**: no csproj consumes it. All 14 ReportDesigner projects declare only
`<ReportLibraryProject Condition="…">…/VBToCSharp/FlexKit/FlexKit.csproj</ReportLibraryProject>`. The
property survives in `tools/verify-report-audit.mjs:14` (passed, unused) and in
`ReportDesigner.RegressionTests/README.md:11,17`, which is wrong. The switch is
`-p:ReportLibraryProject=/Users/wadood/projects/VBToCSharp/FlexCore/FlexCore.csproj`.

```bash
cd /Users/wadood/projects/JavaToCSharp
node tools/verify-report-audit.mjs [output-dir]      # default /tmp/flex-report-audit-verification
```

That runs ten suites (RegressionTests, ComponentTests, RuntimeTests, AuthoringTests, ExportTests,
LayoutTests, AnalysisTests, RegionTests, ParityTests, AuditChecks) against **both** libraries, five
minutes each, writes `<out>/<library>/<suite>/run.log` and `<out>/results.json`, tolerates the one known
AuditChecks gap (`AUD-MAP-01`) and exits 1 on anything else (`verify-report-audit.mjs:7,18,23-26,35`).

Five suites sit outside it. Add `CrystalSemanticsTests` after any Crystal formula or native-reader
change — the 09-25 blocked-52 work added it (75 checks), and memory `crystal_blocked52_resolution`'s
"the 11 ReportDesigner suites pass on both libraries" means the runner's ten plus this one.
`NativeFormatTests` and `RdlAudit` are build-then-run-the-DLL suites with their own arguments (their
READMEs; both want `FLEXKIT_DEBUG_RPT=1` for a strict run). `ParameterAudit` writes JSON to an optional
`args[0]`.

Two are designed to exit non-zero, and doing so is not a regression.
`ReportDesigner.AuditChecks/README.md:16`: "Exit code 1 is expected while the audited defects remain."
`ReportDesigner.VisualTests` has no csproj at all — it is
`node tools/ReportDesigner.VisualTests/compare-crystal.cjs manifest.json output-dir` — and its
README:3-4 says it exits status 2 (blocked) and is never passed, because no approved Crystal baseline is
checked in. Only 5 of the 15 directories have a README: AuditChecks, NativeFormatTests, RdlAudit,
RegressionTests, VisualTests.

## The Crystal bench pack

The bench is `/Users/wadood/projects/VBToCSharp/HomeFront/FlexKitTester` (HTTP profile
`http://localhost:5043`). Its default target is FlexCore; `-p:UseFlexKit=true` targets FlexKit
(`CRYSTAL-REPORTS.md:5`). Run from `JavaToCSharp` — `CrystalBench.Tests/Program.cs:117` resolves
`../VBToCSharp/HomeFront/FlexKitTester/Data/CrystalSamples.db` relative to the working directory.

```bash
cd /Users/wadood/projects/JavaToCSharp
dotnet run --project tools/CrystalBench.Tests -p:UseFlexKit=true     # FlexKit
dotnet run --project tools/CrystalBench.Tests -p:UseFlexKit=false    # FlexCore (the bench default)
node tools/CrystalBench.Tests/verify-browser.cjs                     # reads /tmp/flex-crystal-bench-tests
```

**A Reports/ merge invalidates the pack** — schema fingerprints change, and a stale or unnumbered pack
is rejected rather than silently renumbered. Regenerate, then install at
`FlexKitTester/Data/CrystalSamples.db` (back the old pack up; existing output files are never
overwritten). `CRYSTAL_TEST_SAMPLE_DB` points `CrystalBench.Tests` at another pack.

```bash
cd /Users/wadood/projects/JavaToCSharp
dotnet run --project tools/CrystalSamples.Seed -p:UseFlexKit=true -- build  Reports /tmp/CrystalSamples-new.db /tmp/crystal-sample-audit
dotnet run --project tools/CrystalSamples.Seed -p:UseFlexKit=true -- verify Reports /tmp/CrystalSamples-new.db /tmp/crystal-sqlite-audit
```

Last regeneration 2026-09-25 23:45 against FlexCore: 530/530 paginate (220 Rendered, 310 Review,
0 Blocked), `CrystalBench.Tests` 107 pass; the same pack re-verified against FlexKit after the 0a473d8
port gave identical numbers (`FlexKitTester/CRYSTAL-REPORTS.md:78-85`, memory
`crystal_blocked52_resolution`). FlexKit 0.1.107 shipped on exactly that evidence.

## FlexCore's own test suites

`FlexCore/tests/` holds seven suites nothing in the skill runs. FlexKit has no `tests/` directory, so
these are FlexCore-only coverage.

```bash
cd /Users/wadood/projects/VBToCSharp/FlexCore
dotnet run --project tests/FlexCore.RegressionTests/FlexCore.RegressionTests.csproj   # console, non-zero on failure
dotnet run --project tests/FlexCore.Llm.Tests/FlexCore.Llm.Tests.csproj               # console
```

The five `*.BrowserTests` (Counterparts, Documents, Editor, RemainingControls, Trees) are Web-SDK hosts
binding `http://127.0.0.1:0` (`Program.cs:6` / `:8`), driven by **Python**, not node: start the host,
take the URL it prints, then `python3 tests/<Suite>/verify.py <url>`. They need `playwright`, `pypdf`
and `openpyxl` — importable on this machine today, none carried by git. `FLEXCORE_BROWSER` picks the
engine.

## None of this ships, so step 12 has to carry it

`verify_deploy.sh:50` strips `verification/` and `tests/` from every pushed branch and asserts they are
gone (`:59-64`), and FlexKitTester lives in the outer `HomeFront` repo, which the deploy never pushes.
Per round, record: the harness count and any failures, which browser suites ran and against what, the
ReportDesigner runner's `results.json` verdict, and the pack's paginate numbers if Reports/ moved.

## Consumer builds — serially

After a FlexCore change the owner wants its other three consumers built (owner 2026-09-14, memory
`apps_switched_to_flexcore_20260914`); each `ProjectReference`s `FlexCore.csproj`.

```bash
dotnet build /Users/wadood/projects/VBToCSharp/Mutarjim/Mutarjim.sln --no-incremental -v q -nologo
dotnet build /Users/wadood/GhostWriter/GhostWriter.sln               --no-incremental -v q -nologo
dotnet build /Users/wadood/DotNetCCM/DotNetCCM.sln                   --no-incremental -v q -nologo
```

**One at a time.** Two concurrent builds that both build the library race on its shared `obj/` and
produce phantom `CS0101` duplicate-definition errors in `GridEnums.cs` — rebuild serially before
believing them (memory `ghostwriter_nested_danayalfinance`). `GhostWriterLLM` was not named in that
directive and still references FlexKit. Mutarjim's checkout carries another session's uncommitted files
(memory `mutarjim_phase3_state`) — build it, never commit in it.

## Unverified

- The fixture ports (5302, 5308, 5309, 5310, 5398, 5399, 5407) come from READMEs and script defaults,
  not from a run in this session; if one is occupied, pick another and give the script the URL. The
  TbdAssignmentChecks command shape is inferred — its README names the DLL and the directory, not flags.
- Whether the eight README-less ReportDesigner suites need arguments run individually is documented
  nowhere. Use the runner, which knows LayoutTests, AnalysisTests and RegionTests take an output
  directory after `--` (`verify-report-audit.mjs:15`).
- No fixture test exists for `run_harnesses.sh`, `preflight.sh`, `publish_flexcore.sh` or `install.sh`.
  `tests/classify_fixture.sh` and `tests/gate_fixture.sh` cover `_common.sh`, `pre_deploy_gate.sh` and
  `verify_deploy.sh` only; both are safe (throwaway repos under `mktemp -d`).
