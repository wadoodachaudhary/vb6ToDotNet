# FlexKit ↔ FlexCore — the two identities and the port

FlexKit and FlexCore are the same library under two package identities. Namespace and CSS prefix never
change — `Fx.ControlKit.*` and `fx-*` in both (AGENTS.md:53; both libraries define the same 1,705
distinct `.fx-` classes today). Only project / assembly / package names differ.

Verified against the working trees on 2026-09-26: FlexKit `36885e8` on `telerik-parity-20260904`,
FlexCore `fa8011c` on `main`.

## The two identities

| | FlexKit | FlexCore |
|---|---|---|
| csproj | `/Users/wadood/projects/VBToCSharp/FlexKit/FlexKit.csproj` | `/Users/wadood/projects/VBToCSharp/FlexCore/FlexCore.csproj` |
| `RootNamespace` | `Fx.ControlKit` (:10) | `Fx.ControlKit` (:17) |
| `AssemblyName` / `PackageId` | FlexKit (:11 / :13) | FlexCore (:18 / :20) |
| `<Version>` today | 0.1.107 (:14) | 0.2.48 (:21) |
| On nuget.org | **no** — `api.nuget.org/v3-flatcontainer/flexkit/index.json` is 404; it ships only as a nupkg in `HomeFront/Deploy/local-packages` | **yes** — 48 versions, latest published 0.2.47; 0.2.48 is not published |
| `PackageProjectUrl` | github.com/wadoodachaudhary/FlexKit (:22) | flexcoreui.com (:29) |
| ClosedXML | in-tree dependency (:50) | none — it lives in `FlexCore.Documents` (:42) |
| Satellites | none | `FlexCore.Documents` 0.2.35, `FlexCore.Llm` 0.2.42 (neither published — both 404), plus `tests/` (7 projects) |
| Deploy target | GitLab `flexkit` main, rsynced from the working **branch** | GitHub `wadoodachaudhary/FlexCore` main + nuget.org |

FlexCore excludes its satellites and tests from its own globs — `tests/**` (:5),
`FlexCore.Documents/**` (:10), `FlexCore.Llm/**` (:13) — because `FlexCore.csproj` sits at the repo
root and the SDK globs would otherwise pull them back in with their dependencies. It grants
`InternalsVisibleTo FlexCore.Documents` (:57) rather than widening its public API for the split.

Consumers, all live `ProjectReference`, so a library edit lands on the next consumer build:

| Consumer | References |
|---|---|
| HomeFront (`HomeFront.csproj:61`) | FlexKit |
| HomeFrontPB (`HomeFrontPB.csproj:70`) | FlexKit — **frozen**, never built or deployed (AGENTS.md:38) |
| GhostWriterLLM (`:15`) | FlexKit |
| FlexCore.Showcase (`:11`, `:12`) | FlexCore **+ FlexCore.Documents** |
| Mutarjim (`:13`, `:14`) | FlexCore **+ FlexCore.Llm** |
| GhostWriter (`:59`), DotNetCCM (`:11`) | FlexCore |
| FlexKitTester (`:10`,`:11` default; `:15` under `-p:UseFlexKit=true`) | FlexCore + Documents by default, FlexKit on demand |

⚠ `~/.claude/CLAUDE.md:45`, `:61`, `:63` and AGENTS.md:40 still list **GhostWriter** under FlexKit.
The owner moved Mutarjim, GhostWriter and DotNetCCM to FlexCore on 2026-09-14 and the csprojs confirm
it; memory `apps_switched_to_flexcore_20260914` wins. GhostWriterLLM was not named, and is still on
FlexKit.

## Mirror direction (owner 2026-09-10)

Owner's words, in memory `operating_rules` §4a: **"Always move changes from FlexKit to FlexCore — but
not in the other direction unless told."** Also AGENTS.md:37.

- **FlexKit → FlexCore: always, as part of the same work.** Port the change, then build
  `FlexCore.Showcase.sln`.
- **FlexCore → FlexKit: only when the owner asks.**

This supersedes the 2026-07-25 and 2026-09-07 "the libraries may diverge / never sync" rules. Three
memory notes still state the old rule, and an account that reads memory before this skill will act on
them: `flexcore_never_merged_to_main` ("Never mirror a FlexKit edit into FlexCore or vice versa on your
own initiative" — obsolete in that half; its "never merge FlexCore, or FlexKit's working branch, into a
GitLab main" half still stands); `vb6todotnet_workstream_map`:28 ("the libraries are allowed to diverge
… and are NOT auto-mirrored" — obsolete); and `operating_rules` §2a, the "FlexKit is FROZEN / HANDS-OFF"
paragraphs of 2026-08-01/08, which no longer describe how the project runs — FlexKit has changed in
every deploy round since. Rule 4 of that same note also still says to build `HomeFrontPB.sln` after a
FlexKit edit, superseded by the PB freeze (AGENTS.md:38).

**What still gates FlexKit:** bench work (FlexKitTester → FlexCore) does not reach FlexKit until the
owner has tested it on the bench and said go — two approvals, never one (memory
`flexkit_port_needs_user_go_ahead`, owner 2026-08-19).

That gate is a deploy concern, not only a development one. `deploy_to_repos.sh:514-515` rsyncs each
library's **working tree**, so a premature port ships; and `verify_build` (`:492`) is called only for
hyphen-pb (`:725`), so the `flexkit` and `flexcore` clones are pushed with no staged build at all. The
only compile-time proof either library is sound is the step 6 build.

## Port, never copy

Most shared files sit at the same relative path in both repos. Nine source files and 206 vendored PDF
assets do not, because FlexCore split its document controls into the `FlexCore.Documents` satellite
while FlexKit keeps them in-tree:

| FlexKit path | FlexCore path |
|---|---|
| `Spreadsheet/**` (4 files) | `FlexCore.Documents/Spreadsheet/**` |
| `Pdf/**` (3 files) | `FlexCore.Documents/Pdf/**` |
| `Data/FilerControl.cs` | `FlexCore.Documents/Data/FilerControl.cs` |
| `wwwroot/pdf-viewer.js` | `FlexCore.Documents/wwwroot/pdf-viewer.js` |
| `wwwroot/vendor/**` (pdfjs + pdf-lib, 206 files) | `FlexCore.Documents/wwwroot/vendor/**` (206 files) |
| — (FlexKit's root `_Imports.razor` carries the usings) | `FlexCore.Documents/_Imports.razor` |

All nine source files are byte-identical to their satellite twins today. Copying one to FlexCore's
**root** path instead is silently inert: it is excluded from `FlexCore.csproj`'s globs (:10), so it
compiles nowhere while the real satellite copy stays stale.

### Never carry across

- **The csprojs** — package/assembly/version/description/URLs differ; FlexKit has ClosedXML, FlexCore
  has `InternalsVisibleTo` and the three `DefaultItemExcludes`.
- **`README.md`** (2.2 KB FlexKit vs 26 KB FlexCore; both packed as `PackageReadmeFile`) and
  **`.gitignore`**.
- **`CLAUDE.md`.** Both get a section per round (AGENTS.md:44), written from different positions:
  FlexKit's describes the fix, FlexCore's says "Mirrored the FlexKit correction…". Write each; never
  copy one over the other.
- **`_Imports.razor`** — FlexKit's root copy has two lines FlexCore's must not get:
  `@using Fx.ControlKit.Spreadsheet` and `@using Fx.ControlKit.Pdf`. In FlexCore those usings live in
  `FlexCore.Documents/_Imports.razor`.
- **FlexCore's `tests/` (57 files), `FlexCore.Documents/` (217), `FlexCore.Llm/` (38)** where they have
  no FlexKit twin, and any unrelated in-flight FlexCore work — perf work is developed there on the
  bench.

### The `_content/<AssemblyName>/` trap

A literal `_content/FlexKit/...` path copied into FlexCore compiles and then fails at **runtime** — the
JS module never loads (memory `flexkit_flexcore_convergence_20260901`). Library code is now immune by
construction: all 30 `_content/` path references in FlexKit's `.cs`/`.razor` source resolve the
assembly name at runtime, none is a literal, e.g. `FlexKit/Grid/GridControl.razor.cs:4829`

```csharp
FxJsAsset.Versioned($"./_content/{typeof(GridControl<TValue>).Assembly.GetName().Name}/grid-control.js");
```

Still grep `_content/Flex` after a port, for two reasons. **Host code hardcodes the name:** HomeFront
carries literal `_content/FlexKit/images/16/…` paths (`Components/Pages/Migrated/FPriceList.razor:407`)
and harnesses assert on them (`verification/BrowserInteropChecks/Program.cs:31`), so a host-side snippet
cannot move between a FlexKit host and a FlexCore host unedited. **One stale literal survives in both
libraries**, in a comment: `wwwroot/drawing-control.js:5` says `./_content/FlexKit/drawing-control.js`
in FlexKit *and* FlexCore. Harmless, and the two files must stay byte-identical — leave it.

### After any copy: two audits

Both from real incidents, recorded in `flexkit_flexcore_convergence_20260901`:

- **Lost-member audit.** A wholesale copy deletes the other side's work — FlexKit's `GridColumn.cs`
  once lacked FlexCore's `FilterTemplate` / `FilterMode` / `Footer` + `GroupHeaderTemplate`. Diff the
  member lists both ways before building.
- **Inert-API audit — "compiling is NOT working".** Dropping FlexCore's
  `GridControl.AdvancedFeatures.cs` + `GridControl.ItemsProvider.cs` into FlexKit compiled, yet **37 of
  39 AdvancedFeatures members and 20 of 21 ItemsProvider members were never invoked** by FlexKit's grid
  — `ItemsProvider`, `DetailTemplate`, `ShowExpandColumn`, `Density` would all have silently done
  nothing. Run an invoked-by check; never ship a `[Parameter]` with no consumer, because a host trusts
  it.

## Prove the port landed

Read-only, about 30 seconds.

```bash
cd /Users/wadood/projects/VBToCSharp
git -C FlexKit  ls-files --cached --others --exclude-standard | sort > /tmp/fk.txt
git -C FlexCore ls-files --cached --others --exclude-standard | sort > /tmp/fc.txt
comm -12 /tmp/fk.txt /tmp/fc.txt | while read -r f; do
  cmp -s "FlexKit/$f" "FlexCore/$f" || echo "DIFFER $f"
done
# the Documents split — same content, different home
for f in $(git -C FlexCore ls-files FlexCore.Documents | grep -v /vendor/ \
           | grep -vE '\.csproj$|_Imports' | sed 's|^FlexCore.Documents/||'); do
  cmp -s "FlexKit/$f" "FlexCore/FlexCore.Documents/$f" || echo "DIFFER(doc) $f"
done
grep -rl '_content/Flex' FlexKit FlexCore --exclude-dir=obj --exclude-dir=bin | grep -v /docs/
```

`--cached --others --exclude-standard`, not plain `ls-files`, on purpose: a brand-new file is untracked
until committed and plain `ls-files` cannot see it. Today's uncommitted round is exactly that case —
`Grid/IGridColumnCaptionLimit.cs` is untracked, identical on both sides, and invisible to the
tracked-only form of the check.

**Expected output: exactly four `DIFFER` lines and nothing else** — `.gitignore`, `CLAUDE.md`,
`README.md`, `_Imports.razor`. A fifth line is either an unported FlexKit change or unrelated FlexCore
work. Decide which before shipping; never "fix" it by copying.

Baseline measured 2026-09-26: FlexKit 896 files, FlexCore 993, **680 shared paths, 676 byte-identical,
the 4 expected divergences**; all nine Documents-split files identical; outside the shared set the only
non-satellite difference on either side is its own csproj. The libraries are in full parity right now,
including the uncommitted round — both trees carry the same five modified files (`CLAUDE.md`,
`Grid/GridControl.razor`, `.razor.cs`, `.razor.css`, `wwwroot/grid-control.js`) plus that one new file.

Cheapest runtime proof for a Grid / DropDown / Dialog port: four of the app's 35 `verification/*`
harnesses can target either library through `<ControlLibraryProject>` (default
`../../../../../FlexKit/FlexKit.csproj`, e.g. `GridSortChecks.csproj:7,10`) — `DropDownOpeningChecks`,
`GridSortChecks`, `GridSelectionChecks`, `InputDialogBrowserChecks`. From the active HomeFront
directory (`verification/GridSortChecks/README.md`):

```bash
dotnet run --project verification/GridSortChecks \
  -p:ControlLibraryProject=/Users/wadood/projects/VBToCSharp/FlexCore/FlexCore.csproj -- --checks
```

The other 31 build against FlexKit only, so a green harness run says nothing about FlexCore. FlexCore's
own `tests/` (7 projects) are in no solution and are not run by this procedure.

## Build matrix

| Edited | Build, serially, `--no-incremental` |
|---|---|
| FlexKit | `MobileSource/HomeFront/HomeFront.sln` — contains FlexKit + HomeFront, nothing else (AGENTS.md:41) |
| FlexCore | `../FlexCore.Showcase/FlexCore.Showcase.sln` — contains Showcase + FlexCore; `FlexCore.Documents` arrives transitively via `FlexCore.Showcase.csproj:12` (AGENTS.md:42) |
| FlexCore, additionally (owner 2026-09-14) | Mutarjim, GhostWriter, DotNetCCM — one at a time; concurrent builds race on FlexCore's `obj/` |
| anything | **never** `HomeFrontPB.sln` — frozen (AGENTS.md:38) |

Each build must print `0 Error(s)`.

`--no-incremental` because an incremental `HomeFront.sln` build after a FlexKit source edit can be a
cached no-op that masks whether FlexKit recompiled at all, and can hide pre-existing errors in other
FlexKit files (AGENTS.md:43; memory `flexkit_plusminus_inline_style_and_poc_projectref`).

Two gaps in that matrix:

- **FlexCore has no `.sln` of its own.** `FlexCore.Showcase.sln` is the only routine build of it.
- **`FlexCore.Llm` is in no solution anywhere** — not even `Mutarjim.sln`, which lists only
  `FlexCore.csproj` and `Mutarjim.csproj`. It compiles as a transitive `ProjectReference` of
  `Mutarjim.csproj:14`. A FlexCore.Llm edit is not covered by the standing build check.

## If the owner asks for a FlexCore → FlexKit port

Owner-requested only. The one that happened — FlexKit `4efb8e2` (2026-09-07), *"Sync source from
FlexCore 2026-09-07 — full parity, branding untouched"* — silently reverted FlexKit-only behaviour that
had shipped on owner directives. Casualties, from memory `flexcore_sync_reverted_gridcolumn_align`:

- `GridColumn` date right-align default (HHM-920). `TreeGridColumn` kept it, `GridColumn` lost it — the
  asymmetry between the two was the only tell. Restored 09-08, shipped in FlexKit 0.1.88.
- `DatePickerControl` AutoFocus, the owner's Enter-on-blank-opens-calendar rule, the mousedown
  `preventDefault`s; `DropDownListControl`'s 2026-09-01 no-trap fixes (focus hand-back after
  close/pick, Escape on a closed list forwarded, Tab on an open list commits). Restored 09-10, pinned
  by `verification/PropertiesTreeChecks`.
- Twelve further GridControl behaviours (date first-click mount, `_batchEditDirty` flush,
  `Task<bool> CommitBatchEdit` veto, dropdown-editor vertical arrows, `editorOwnsEnter`,
  `CancelActiveBatchEditAsync` + host Escape, invalid red flash, edit button pinned while editing,
  `gridHighlightsSelectedRows`, combo MaxLength, `fx-grid-core.css` rules). Restored in both libraries
  09-11, pinned by `verification/GridDateCellChecks`.

**The rule that came out of it:** after any FlexCore→FlexKit port, diff each ported file for
FlexKit-only behaviour and re-verify that the owner directives it carried still exist, before trusting
the build. The build stayed green through all of the above — the readers were reverted alongside the
writers.

Two findings from the 2026-09-01 convergence work, still binding: the two grids **cannot** be merged (a
true three-way merge of `GridControl.razor.cs` gives 46 conflict hunks, one of 15,591 lines — two
implementations, not two revisions; port wanted features individually, each runtime-tested); and
"textually clean" is not safe (`ServerOptimization.cs` and `GridModels.cs` merged with zero conflicts,
then failed with 21 duplicate-definition errors, because FlexKit already implements that cell/row
handler cache across different partial files).

Revert levers, all present today: FlexCore tag `pre-flexkit-sync-20260901`; FlexKit tags
`pre-flexcore-sync-20260907` (the lever for 4efb8e2 itself), `pre-flexcore-port-20260829`,
`pre-phase2-20260901`, `pre-phase3-20260901`, `pre-flexcore-grid-merge-20260901`.

Still lost from `4efb8e2`: `TextBoxControl.UseNativeAutoFocusOnly` is absent from both libraries today.
Restoring it is an owner decision, not a parity repair. The other two open items from that note are
closed — `Uncontrolled` is back in both `TextBoxControl.razor` and `DatePickerControl.razor`, and
password-reveal is in both copies of `TextBoxControl.razor`.

## Stale documents on disk

These contradict the standing rules. All are stripped from every pushed branch
(`deploy_to_repos.sh:93-108`: `CLAUDE.md`, `AGENTS.md`, `PUBLISHING.md`, `docs/`, `tests/`), so they
mislead only local sessions — this skill's own audience.

- **`FlexCore/PUBLISHING.md`** (`bea21b5`, 2026-08-13) and **`FlexKit/PUBLISHING.md`** (`b6594ee`,
  2026-06-21) are byte-identical, both titled "Publishing FlexCore" although FlexKit is not on
  nuget.org at all. Between them they instruct: push with `--skip-duplicate` (:78, the opposite of the
  current rule, under which a duplicate means the version was not bumped); pack the **live** tree into
  `./nupkg` (:60-61, exactly what `publish_flexcore.sh` exists to avoid); `git push` in FlexCore (:32,
  :53 — a direct push to `github/main` is rejected non-fast-forward); read the API key from
  `~/.flexcore-nuget-key` (:16); and "**HomeFrontPB** … now references FlexCore via NuGet — no longer a
  `<ProjectReference>`" (:17, whereas `HomeFrontPB.csproj:70` references **FlexKit** by
  ProjectReference). **`publish_flexcore.sh` and this skill are canonical; both `PUBLISHING.md` files
  are historical and wrong.** `~/.flexcore-nuget-key` does still exist (mode 600, dated 2026-08-12), so
  following that line would appear to work while bypassing the Keychain route the owner set up; what was
  deleted on 2026-08-12 was the other plaintext file, `~/.nuget-api-key` (memory
  `flexcore_nuget_publish`). Whether the on-disk key is still valid is unverified — it was not read.
- **`FlexKit/docs/`** is a byte-identical copy of `FlexCore/docs/`, so it tells readers to load
  `_content/FlexCore/fx-grid-core.css` (`docs/remaining-controls.md:48`,
  `docs/dedicated-controls.md:30`) — wrong for a FlexKit consumer. Keeping the two identical is what
  parity requires; do not "fix" one side.
- **`FlexCore.csproj:25`**, the shipped package `<Description>`, still advertises "native ClosedXML
  Excel export", untrue since the Documents split (ClosedXML is at `FlexCore.Documents.csproj:42`). It
  is live metadata on nuget.org, so a fix only reaches consumers with the next publish.
- **Retire one trap:** memory `flexcore_grid_control_js_latent_bug` (FlexCore calling an undefined
  `gridHighlightsSelectedRows`) no longer reproduces — the function is defined at
  `FlexCore/wwwroot/grid-control.js:4733` and the file is byte-identical between the libraries. Mark
  the note resolved rather than leave a warning that cannot fire.
