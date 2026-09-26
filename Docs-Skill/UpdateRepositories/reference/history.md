# Update Repositories — history

Append one row per run, newest first. Heads are the pushed tips; "local" = commits made in the working
repos for that run. Git cannot tell which Claude account ran a deploy — every deploy commit carries the
same staging-clone identity — so say so in the hand-off note if it matters.

| When (EDT) | hyphen-pb main | homefront | flexkit (FlexKit) | flexcore GitHub (FlexCore → nuget.org) | R1-UAT | Local commits | Notes |
|---|---|---|---|---|---|---|---|
| 09-26 02:17 | 4a6d61f | 5052586 | feb0d80 (0.1.107) | f4de318 (0.2.48 NOT published — still held) | ca518c4 frozen | app dff7e97 (`--pull` Deepika 601603f 438cad9 60dfe16, Irfan 98677c6 — fast-forward, no conflicts), 6a58c4b (R2-QA HHM-1165..1170/1172 worksheet + picker, Codex report-import reconversion, ship-review fixes); FlexKit 36885e8 (Crystal formula rewrite + blocked-52, FlexCore 0a473d8 ported, HHM-1132/1157/1170/901, GridColumn.DisplayFormatter); FlexCore fa8011c (mirror) | owner asked: verify R2-QA Dev Review + bring in the FlexCore reporting code. Review wf_0398afff-2e1 landed BEFORE the push: 8 confirmed (picker date filter/checklist, duplicate rows uncosted, dropdown focus 2 RTT + dispose leak, cell-mode keypress flash, price batch overwrite, Sales Pricing menu hidden) — all fixed. flexcore staging clone reset to origin/main 0a473d8 (the Crystal session had pushed the source repo to GitHub); the push re-stripped docs/ and tests/ from GitHub main. 12 stale self-history REVERT flags acked in preflight-ack.txt. |
| 09-23 17:19 | 316afba | 9f4fca6 | 0ff5af6 (0.1.106) | 4566b0e (0.2.48 NOT published — held) | ca518c4 frozen | app 8e10e6d (revert MR !39 — owner had closed it 09-21), 3f3df5d (wizard harness, 15 wizards), b021eb1 (MR !41 HHM-922/964/1159/1161), 6272974 (HHM-1159 DELETE gated on saved keys), 7822dee (Codex Reports → Manage + "Convert to Open XML" label); FlexKit b4e1978; FlexCore 99821fa | held three times for live Codex sessions (Crystal reader, report manager, record-161 fix; last quiet 16:51). Reviews before the push: wf_06f8c26a-ce4 (!41: HHM-1159 DELETE could wipe a live row's saved record — fixed, verified 3 rounds), wf_48805d75-d75 + wf_02a1512f-0d0 + wf_2ae6ecc6-545 (Crystal reader + Reports → Manage: security/data-loss and library defects — owner: "just ship it", report work is not QA's focus). NuGet publish of FlexCore 0.2.48 HELD until the confirmed library defects are fixed (a publish is permanent). No teammate commits; MR !41 taken. |
| 09-23 08:56 | 22b5713 | 2261abc | 243ca03 (0.1.105) | eccf524 (0.2.47 published, listed 09:04 — indexing lagged past the script's 6-min wait) | ca518c4 frozen | app 46c383a (`--pull` Irfan a42eb60, Deepika 6b9507b — fast-forward, no conflicts), 3d6f63b (Baaria MR !40 HHM-1100), 9174ec2 (Codex wizard/WizardControl batch + deleted dead pages + HHM-1156 TBD + review fixes); FlexKit 0eb4634; FlexCore 6b38787 | review wf_14d2d7a0-336 (4 dimensions + adversarial verify) landed BEFORE the push. Fixed: the filter checklist lost its full value set when Search values was cleared under a condition, so Apply stored a checked-value filter that outlived the condition (FlexKit+FlexCore); EstimateChecks asserted a constant for the disabled-Forecasting guard. Refuted after verification: deleting FRFPWizard (orphaned page, VB6 NewRFQ button never migrated), Field PO Requests removal (documented HHM-790 sunset + regression tests), TBD Ctrl+Delete (documented HHM-1156 spec). Open for the owner: Cost Forecasting disabled on all three surfaces with no owner attribution. New harness TbdAssignmentChecks (220 browser + 572 visual). |
| 09-22 01:05 | 34f41b9 | 392ed78 | 97e3bc6 (0.1.104) | d407e87 (0.2.46 published) | ca518c4 frozen | app 2eb1836 (`--pull` Deepika 7e4face, Irfan fa2c839 + 71497be; 2 conflicts safe-merged + acked), 9e36b26, 3c5cb0e, 5fb28a1 (HHM-1074, session 23165bb4), fa8be39 (HHM-1149 + job-band slider, session 5c72594a), afaf1b2 (Baaria MR !39 by intent); FlexKit 3b9ab51 a6b345a 2e8f6d2; FlexCore 725ce87 7b26080 6829c0b | held twice for live sessions (Codex Choose Columns 21:34-21:56, 23165bb4 HHM-1074 to 22:36, 5c72594a HHM-1149 to 00:01). Reviews landed BEFORE the push: pre-deploy review of 09-20/21 work (C1-C12: Field PO Requests, Manual PO VB6 parity, Takeoff Delete, Attachments header sort, keyless keys, report pane minimum), wf_6d3426a4-c65 (chooser slide-off discarded edits; Best Fit to Grid could not grow FitColumns grids), wf_40a65f51-808 (band combos clipped by the new slider; !39 view-switch race; slider double-click default; splitter fallback stuck drag), plus a verifier on the DialogControl fix. MRs !35-!38 already in; GitLab MR close NOT done (not signed in, Chrome extension disconnected). Staged build took 16 min. |
| 09-18 09:21 | d5ec26f | f41fe81 | d60ec9b (0.1.103) | 2bd8273 (0.2.45 published, listed 09:27) | ca518c4 frozen | app 910c3e0 (Irfan f16d2d3), 23bcb3a (Baaria MR !36 update 935b451), ac73ddf (Irfan 23ddbf1), 3b45a5e+ef2fdba (R2-QA 09-17), 8e3a6cc/59bd589/64fc269 (91ef9618 grid-filter harness), f690c82 (Codex HHM-1070/1071), a003489; FlexKit e238107 3aa3457 + 91ef9618 cf98ec4..7f4e66d + Codex 932a67b; FlexCore 27d582e 4cc08d5 + ports; outer 291e961.. | held overnight: 91ef9618's half-written GridControl edits broke the FlexKit build mid-chain; shipped after a 12-min quiet watch. Reviews wf_14845c13-5fc (Codex report work: blank viewer buttons, currency counts, uncleared pick lists, PDF reload — fixed) and wf_47238abd-a8a (FLogin autofilled Password never published, dropdown bottom-row open, Min/Max stepping, filter-row double apply — fixed); fixes verified by wf_69e310d1-677 / wf_4c5dd264-729 |
| 09-17 07:11 | c3b0ef3 | e79586c | d8941ad (0.1.102) | ea6de7c (0.2.44 pushed 07:12 listed 07:20) | ca518c4 frozen | app 1944b8e (`--pull` Irfan 40ef498 29f8287 432c842, Deepika 3581c61 32204cb) + aac9d99 (Baaria MR !35 + !36, still OPEN in GitLab) + f8d1a33; FlexKit 0b76a21, FlexCore 5eae357, outer 291e961 | review wf_af95484b-ac5 landed BEFORE push: blocker in Irfan 29f8287 (System Settings OK wiped every division's standard items) fixed + acked; 3 stale harnesses (HHM-977, HHM-1062, MR !35) |
| 09-15 00:30 | 1313b21 | 9b23aa8 | ef988bc (0.1.101) | 5be387c (0.2.43 published; FlexCore.Llm on GitHub only) | ca518c4 frozen | FlexKit 2b071f1, FlexCore 6d8d7bf, app 0327ccd | step 0 caught Irfan b23d656 + Deepika aede435 → `--pull` (0327ccd is a single-parent rsync commit, not a git merge) |
| 09-14 01:29 | 4a2794c | 768a1c8 | e05090f (0.1.100) | 7e27e76 (0.2.42 published) | ca518c4 frozen | FlexKit b2f5cba, FlexCore 274e5db, app 134938b + 6bc7058 | held until session 5c72594a was idle; review wf_f62e78d5-387 fixes shipped |
| 09-13 09:30 | 5d075fc | 6a92f8f | 7568d15 (0.1.99) | e21156b (0.2.41 published) | ca518c4 frozen | FlexKit ee0bc4d, FlexCore 2be495b, app 94202f2 | review landed AFTER the publish: 8 defects shipped, incl. a Multi-sort revert now permanent in 0.2.41 |
| 09-13 00:19 | 1f356cc | ee5aeb1 | 52f7e29 (0.1.97) | e4fc0aa (0.2.39) | ca518c4 frozen | app 80624d2, FlexKit 72a388f, FlexCore aeb1073 | second run of the night — first with the non-runtime strip (d7f6d95) |
| 09-13 00:02 | db81c50 | 6c4460d | 3abbebb (0.1.97) | d3f069c (0.2.39 published) | ca518c4 frozen | (same night) | first run with R1-UAT frozen by default (e1aae92) |
| 09-12 20:07 | ca3e8c9 | acbff0f | 676692e (0.1.96) | b684094 (0.2.38 published) | ca518c4 | app 144427c, FlexKit 4222c20, FlexCore 45a0f81, outer 78fb27d | first run of the login-skip guard; FlexKit packed once, byte-identical nupkg on both branches |
| 09-12 16:59 | 8f03fa3 | 9910e05 | 48eb73c (0.1.95) | 8cf89d7 (0.2.37 published) | 0519173 | app f50421f, FlexKit bfe8672, FlexCore f3ebcd2 | restored MR !33 that our 09-09 deploy had reverted; login-skip still shipped (guard not yet written) |
| 09-10 23:45 | 6de7a2b | 7c354ff | e7f6123 (0.1.94) | — | 5f23759 | app 63b16b4, FlexKit e7c8599 | FlexKit edited mid-run → R1-UAT got an unverified 0.1.94 |

nuget.org has no FlexCore 0.2.34 or 0.2.40 (never published; 0.2.34 exists only in the local NuGet cache). FlexKit is not on nuget.org — `Deploy/local-packages` is its only ship route.

## Lessons, in the order they were learned

1. **09-10** — "idle" judged from file mtimes missed a session whose background agent started editing
   FlexKit minutes into the deploy. Check transcripts and `subagents/` recursively.
2. **09-12** — a teammate's merged MR had been reverted by our own deploy three days earlier and stayed
   reverted through three more deploys. Blob-hash every recent upstream file (preflight §5).
3. **09-12** — the lockdown check, `--dry-run`, and a force-added secret were all silently wrong for
   weeks. Distrust green output; verify remote state after every run (`verify_deploy.sh`).
4. **09-13** — publishing to NuGet before a review finished made a defect permanent. Hold it.
5. **09-13** — a stale copy undid a fix byte-for-byte. Check edits against recent commits (§6).
6. **09-14** — deploying while another session was mid-way through a feature would have frozen it into
   a published package. Hold until idle.
7. **09-15** — step 0 + the blob check caught two teammates' commits the push would have reverted;
   `--pull` merged them (FOptions 3-way with our own edit, no conflict).

8. **09-17** — three harnesses failed at once and all were STALE, not regressions: a teammate's intro step
   (HHM-977), our own 09-16 HHM-1062 commit made without its harness, and MR !35's design-time columns.
   `git log` on the page settles it. The review of the merged teammate code found a data-loss blocker
   already live on hyphen-pb main (29f8287); fixing it makes preflight report REVERT for that commit, so
   the rejected hunk goes into `Deploy/preflight-ack.txt` after checking it is the ONLY missing part
   (diff our file against origin/main).

9. **09-18** — another session's implementation agent started editing FlexKit's GridControl in the
   middle of the verification chain; every build after that failed on members it had not written yet.
   Hold means hold: watch the session's transcript AND `subagents/` AND the source trees until all are
   quiet (12 min), then re-run everything. Other sessions' FINISHED work still needs the ship review:
   both reviews this round found real regressions in work its authors had verified (blank toolbar
   buttons in HomeFront, and an autofilled login password never reaching the server).
10. **09-23** — merge-request refs (`refs/merge-requests/N/head`) carry no MR STATE. "Not in main" was read
   as "open": !39, closed by the owner on 09-21 (its description asked for a real-division check first),
   was taken into main on 09-22 and had to be reverted (8e10e6d). Before taking an MR, confirm in the
   GitLab UI that it is OPEN; if the UI is unreachable, ask the owner rather than infer it from git.

## Open as of 2026-09-26 02:20 (after the 09-26 deploy)

- **FlexCore NuGet still held** (0.2.48 unburned): of the five 09-23 Crystal-reader defects, at least three are open —
  TslvStreamReader.LoadString decodes UTF-8 (Latin-1 needed), a dropped legacy join still renders as CROSS JOIN, and
  ConversionDiagnostics are still not surfaced at runtime. StartsWith-array and cross-tab percentage summaries not rechecked.
- **Jira:** HHM-1007/1028/1084/1156/1165/1167/1168/1169/1170 moved to Ready for Test (verified live + against VB6).
  Left in Development Review: HHM-1157 (rename cap 255 vs AppGridLayout.caption varchar(100) — owner asked for the DB max),
  HHM-1166 (mixed Add still INSERTs non-duplicate rows at add time; Save As DB error never reproduced), HHM-1138 (thumb/End
  fixed; wheel path unmeasured; owner question on the web-only wheel settings still open), HHM-1132 (one round trip now;
  VB6 is instant; owner to set a target). HHM-1172 (To Do) got its picker half (chronological date sort) in this deploy.
- **GitHub FlexCore:** the Crystal session pushed the FlexCore source repo (with docs/ and tests/) straight to GitHub on
  09-25; this deploy's strip removed docs/ and tests/ again. Decide whether GitHub main should carry them.
- **MRs:** branches HHM-1053, HHM-1057, HHM-1057-persist-on-first-use, HHM-922-964-1159-1161, festimateitems-phase-dropdown,
  wip/festimateitems-local-20260914 have recent commits; the 09-23 notes record every MR as settled — not re-checked in GitLab.

## Open as of 2026-09-23 17:30 (after the second 09-23 deploy)

- **FlexCore 0.2.48 is on GitHub but NOT on nuget.org** — held on purpose. Fix these library defects
  (FlexKit + FlexCore, byte-identical) first, then publish 0.2.48 (still unburned) or higher:
  1. pre-v9 strings: TslvArchiveReader.LoadString decodes UTF-8 while LegacyCrystalDatabaseParser.Text
     and the legacy field name use Latin-1 — non-ASCII names break the formulas that use them;
  2. a dropped legacy join runs as a CROSS JOIN; CrystalXmlReportLoader never surfaces
     ConversionDiagnostics (CRYSTAL_PARTIAL_EXTRACTION) at runtime (direct-.rpt viewer path is silent);
  3. CrystalFormula StartsWith with an array argument returns false instead of any-prefix;
  4. CompoundFileReader / CrystalRptBinaryReader: many directory entries sharing one sector chain make
     a small crafted .rpt allocate far beyond its size (reachable from Reports → Manage uploads);
  5. cross-tab percentage summaries (HasPercentSummary / SummaryType) convert as plain Sum/Count.
  Exact fixes are in the verdicts of wf_02a1512f-0d0 and wf_2ae6ecc6-545. FlexKit next version >= 0.1.107.
- **Reports → Manage shipped with known defects (owner: "just ship it", 09-23):** the private-storage
  check in UserReportLibrary.ResolveReportPath can be bypassed from the viewer (relative path resolves
  against the content root; Windows device/UNC paths) — one user can read another's import; User/Reports
  lives inside the deployed content root, so QA deploys can wipe imports; the manager filter no longer
  filters as you type; a file name ending in " ." corrupts the user's list; server paths shown in the
  browser. An uploaded report's SQL runs under the app's DB login — owner: acceptable for now.
- **Owner question — HHM-964 (MR !41):** CellEditablePredicate now makes multi-row edits (fill-down,
  column mass edit) skip rows VB6 writes to; VB6 gates only the active cell (frm:3195-3199).
- **Pre-existing, found this round:** FPricingWorkSheet.SaveDataAsync writes fewer columns than VB6
  SaveData on Design Center sheets (drops Pretax/Tax/Markup/Margin 2..10, Color/Style/Finish lists,
  SpecDocument, GraphicPath and more) — every save of a DC sheet loses them. HIGH for DC users.
  Also: SaveDataAsync sets WorksheetId before commit, so a failed new-sheet save leaves a dead id.
- **FAssemblyReplicator** completion page keeps Cancel enabled beside Finish (owner call).
- **MRs:** all of !35-!41 are settled (!41 taken; close it in GitLab — its content is on main 316afba).
  Before taking any MR, confirm it is OPEN in the GitLab UI (lesson 10).

## Open as of 2026-09-23 09:05 (after the 09-23 deploy)

- **MRs — settled 2026-09-23.** !35, !36, !38, !39 were already CLOSED (the "still open" reading came
  from git refs; see lesson above); !37 merged into !36's branch; !40 merged into main by the owner
  (9a3f78a, no file change; pulled). **!39 REVERTED** (8e10e6d, owner): it had been closed on 09-21, so
  the gItems layout is again written on view switch / close only, as VB6 does. **New MR !41**
  (HHM-922/964/1159/1161, branch HHM-922-964-1159-1161, ad18508, base 22b5713; FEstimateItems +
  FPricingWorkSheet) is OPEN — owner: take it in the NEXT update, reviewed; do not merge it in GitLab
  first (main is QA). The revert ships with that update.
- **Cost Forecasting — SETTLED 2026-09-23.** It stays disabled in the Tasks menu, the Vendor Pricing
  sidebar and FPriceList's toolbar (route and page still live). Owner: "VB6 has it enabled but for
  Web that is out of scope for this release. It will be enabled later." EstimateChecks pins the
  three disabled surfaces; re-enabling later means re-enabling those entry points only. Memory note:
  cost_forecasting_disabled_this_release.
- **Field PO Requests — SETTLED 2026-09-23: descoped.** It stays out of both FMain navigation lists
  (route, page and task handler still live). VB6 lists it in both (FMain.frm:1838, :2745 under
  IssuePOs Or GeneratePOs); the 09-21 restore was the mistake, not the 09-23 removal.
  EstimateChecks asserts both lists stay clear. Memory note: field_po_requests_descoped.
- **Owner FYI, remaining deliberate VB6 divergences in this batch:** the TBD wizard clears vendors
  with Ctrl+Delete instead of VB6's plain Delete and its Clear vendor button is gone (HHM-1156 spec,
  pinned by WizardChecks); Mass Change enables Next on the Choose-Action page where VB6 disables it
  for Add/Remove/Substitute (owner 2026-09-22 "no pickers on this page").
- **FRFPWizard** was deleted as dead. It is a VB6-shipped form (HFEst.vbp:50) whose entry point,
  FEstimateItems' NewRFQ toolbar button, was never migrated; restore from git history when that
  button is ported.
- **Still open from 09-22:** Jira HHM-1149 / HHM-1074 to
  Development Review; the low review findings listed under "Open as of 2026-09-22".
- **Next versions:** FlexKit >= 0.1.106 (0.1.105 burned), FlexCore >= 0.2.48.

## Open as of 2026-09-22 01:30 (after the 09-22 deploy)

- **Close the MRs in GitLab (owner asked; not done — GitLab not signed in, Chrome extension
  disconnected):** hyphen-pb !35 (HHM-1057), !36 (festimateitems-phase-dropdown), !37 (merged into
  !36's branch; close if still open), !38 (HHM-1053), !39 (HHM-1057-persist-on-first-use). All their
  content is on main 34f41b9 (!39 by intent, afaf1b2). `wip/festimateitems-local-20260914` is not an MR.
- **Owner decision:** !39 writes the VB6 design-time gItems layout on first use; VB6 writes the same
  rows only at Form_Unload, and the 09-13 HHM-1057 direction said "save on close". Taken because the
  owner asked for every MR; revert afaf1b2 if on-close only is wanted.
- **Jira:** session 5c72594a asks whether to move HHM-1149 to Development Review; HHM-1074 (session
  23165bb4) likewise.
- **Low review findings not fixed:** short windows clip the grid bottom under the 144 px band;
  the job splitter re-measures on every page render; SplitterControl leaks its DotNetObjectReference
  if disposed during its first registration (pre-existing); Best Fit to Grid on a FitColumns grid
  with the options rail counts the rail margin (pre-existing); TreeGrid best fit is on by default for
  trees that did not opt in; HHM-1074 deferral: a double Enter can reopen the picker, and closing the
  variance prompt with X/Escape leaves the "..." picker unopened; ArrowUp/Down in Choose Columns can
  scroll the page; the grid's client-buffered typing buffer can still take chooser keys (pre-existing).
- **From the 09-21 review, still for the owner:** Application.GetDb external-DB throw; tree
  double-click on a collapsed node; Manual PO: edit-time coding lists cached, blank-vendor POs not in
  the Open PO picker (VB6-faithful); FAttachments Remove/Find order after sorting;
  HeaderClickShowsColumns now unused by the app; Ctrl/Alt+Delete in Takeoff Settings;
  ParameterAudit/AuditChecks pre-existing defects.
- **Housekeeping:** the Codex Choose Columns bench (FlexKitTester choose-columns-popup + verification)
  and the Crystal bench work are uncommitted in the outer repo. Next versions: FlexKit >= 0.1.105
  (0.1.104 burned), FlexCore >= 0.2.47.

## Open as of 2026-09-18 09:30 (after the 09-18 deploy)

- **Baaria's MRs !35 (HHM-1057) and !36 (festimateitems-phase-dropdown, tip ee2db8a) are still OPEN in
  GitLab.** Their content shipped. From 935b451 (merged into !36) the HHM-1075 hover-text half was NOT
  taken — 3b45a5e already titles the tree via the VB6 PrettyName port (acked in Deploy/preflight-ack.txt).
- **Jira (session 91ef9618 asked the owner):** move HHM-988 and HHM-874 to Development Review with its
  one-line comment. The Codex session already moved HHM-1025/1070/1071.
- **Owner decisions from the reviews:** HomeFront's Reports menu now renders fresh native conversions of
  .rpt files, not the checked-in xml/ exports (Codex change, deliberate); FlexCore NuGet users need the
  unpublished FlexCore.Documents for the Crystal PDF preview; ReportVectorShape's constructor changed
  (binary break vs 0.2.44); the grid toolbar/side-panel searches stop page-level Blazor keydown shortcuts;
  MR !35 still refuses Project in Quote mode (owner 09-13 said literal VB6); Irfan's add-division skips
  the System_Setup copies; Deepika's HHM-1061 turned off FCustomQuote invalid-data colouring.
- **Follow-up tasks offered as chips (not started):** persist assembly-tree widths on fallback columns
  (FEstimateItems gAssemblies); make department delete select the next department (FOptions, Irfan's area).
- **Test gaps:** WizardChecks never calls FVendorChange.ChooseReplacementVendorAsync (guard, column filter).
- **Rule breaks seen:** the Codex report session built HomeFrontPB.sln on 09-17 22:12-22:16 (source untouched).
  FLogin still logs the typed password when debug_mode is on and NoPasswordLogging is off (pre-existing, opt 7).
- **Not an MR, not applied:** `wip/festimateitems-local-20260914` (d145722). Owner decision.
- **Still open from earlier:** Atlassian token rotation; QA login dev panel (`Security__HideDevPanel`);
  runtime-limitations panel; report-designs saves nothing reopens; session 5c72594a's PB edit + paste patch;
  FlexCore.Llm/Documents versions.
- **Next versions:** FlexKit ≥ 0.1.104 (0.1.103 burned), FlexCore ≥ 0.2.46.
- **Housekeeping:** outer repo still has the Crystal session's uncommitted FlexKitTester bench work and stray
  `MobileSource/m/` + `MobileSource/t_rec3.sh`; the deploy never pushes the outer repo.
