# HomeFront phase-1 menu bench

Run this standalone tester in the cloud with
`CrystalReportsBench__MenuScope=/workspace/HomeFront-Crystal-Cloud/corpus/homefront-menu.json`.
The report page at `/crystal-reports` selects exactly the 19 approved menu reports,
in menu order, with menu numbers 01–19 and the existing HF corpus IDs. Search supports
both identifiers. The shared FlexKit/FlexCore grid, conversion controls, designer and
viewer remain in use.

Menu mode validates the manifest against `corpus/reports.json`, the original RPT
hashes and source aliases. It hides and guards RPT upload and server browsing.
Conversion rechecks the isolated copy before passing original bytes to the native
converter. Missing, invalid or changed scope configuration fails closed, including
removal during a running circuit. Off-menu paths, symlinks and changed binaries are
rejected. No expected/preconverted XML participates in conversion.

Menu mode never reads the 530-entry synthetic catalog or runs its data. Data is
initially unselected. A layout-only preview has no records and does not validate
data semantics or pagination. Matching data fixtures and a separately configured
isolated database remain explicit options. Embedded report endpoints are never
connection sources. Conversion, XML comparison, data execution, pagination and
original-Crystal approval require independent evidence.

Focused regression command, from the repository root after activating the prepared
cloud tools:

```sh
source /workspace/.homefront-tools/activate.sh
dotnet run --project JavaToCSharp/tools/CrystalBench.Tests -p:UseFlexKit=true -- --menu-scope "$PWD"
```

Validation on 2026-10-20: 45 checks passed against each library on the Mac using
`/Users/wadood/projects/HomeFront-Crystal-Cloud` as the corpus root. The tester now
uses Microsoft.Data.Sqlite 10.0.8 with FlexKit and 10.0.12 with FlexCore, matching
their existing dependency graphs and resolving the prior NU1605 downgrade.
Existing package vulnerability warnings remain; these checks establish menu
scoping and conversion guards, not production readiness or visual Crystal parity.

The early `--menu-scope` entry point runs only these focused checks, including fresh
conversion of menu 01, without entering the existing broad sample/native modes.
The checks exercise allowlist metadata and hash changes, removed manifests,
traversal and symlinks, source and stream conversion guards, shared component
selection/search/navigation, blocked uploads, and disabled synthetic data. They do
not run report queries or establish Crystal parity.

Launch with the same scope environment variable and loopback-only binding:

```sh
CrystalReportsBench__MenuScope="$PWD/corpus/homefront-menu.json" \
CrystalReportsBench__EnableDatabase=false \
ASPNETCORE_ENVIRONMENT=Development \
ASPNETCORE_URLS=http://127.0.0.1:5043 \
dotnet run --project VBToCSharp/HomeFront/FlexKitTester -p:UseFlexKit=true --no-launch-profile
```

Use `-p:UseFlexKit=false` for FlexCore; keep the projects' different package graphs.
Do not build either HomeFront host application as part of this bench workflow.
