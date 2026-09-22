# Update Repositories — history

Append one row per run, newest first. Heads are the pushed tips; "local" = commits made in the working
repos for that run. Git cannot tell which Claude account ran a deploy — every deploy commit carries the
same staging-clone identity — so say so in the hand-off note if it matters.

| When (EDT) | hyphen-pb main | homefront | flexkit (FlexKit) | flexcore GitHub (FlexCore → nuget.org) | R1-UAT | Local commits | Notes |
|---|---|---|---|---|---|---|---|
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
