# Update Repositories — history

Append one row per run, newest first. Heads are the pushed tips; "local" = commits made in the working
repos for that run. Git cannot tell which Claude account ran a deploy — every deploy commit carries the
same staging-clone identity — so say so in the hand-off note if it matters.

| When (EDT) | hyphen-pb main | homefront | flexkit (FlexKit) | flexcore GitHub (FlexCore → nuget.org) | R1-UAT | Local commits | Notes |
|---|---|---|---|---|---|---|---|
| 09-15 00:30 | 1313b21 | 9b23aa8 | ef988bc (0.1.101) | 5be387c (0.2.43 published; FlexCore.Llm on GitHub only) | ca518c4 frozen | FlexKit 2b071f1, FlexCore 6d8d7bf, app 0327ccd | step 0 caught Irfan b23d656 + Deepika aede435 → `--pull` (0327ccd is a single-parent rsync commit, not a git merge) |
| 09-14 01:29 | 4a2794c | 768a1c8 | e05090f (0.1.100) | 7e27e76 (0.2.42 published) | ca518c4 frozen | FlexKit b2f5cba, FlexCore 274e5db, app 134938b + 6bc7058 | held until session 5c72594a was idle; review wf_f62e78d5-387 fixes shipped |
| 09-13 09:30 | 5d075fc | 6a92f8f | 7568d15 (0.1.99) | e21156b (0.2.41 published) | ca518c4 frozen | FlexKit ee0bc4d, FlexCore 2be495b, app 94202f2 | review landed AFTER the publish: 8 defects shipped, incl. a Multi-sort revert now permanent in 0.2.41 |
| 09-13 00:19 | 1f356cc | ee5aeb1 | 52f7e29 (0.1.97) | e4fc0aa (0.2.39) | ca518c4 frozen | app 80624d2, FlexKit 72a388f, FlexCore aeb1073 | second run of the night — first with the non-runtime strip (d7f6d95) |
| 09-13 00:02 | db81c50 | 6c4460d | 3abbebb (0.1.97) | d3f069c (0.2.39 published) | ca518c4 frozen | (same night) | first run with R1-UAT frozen by default (e1aae92) |
| 09-12 20:07 | ca3e8c9 | acbff0f | 676692e (0.1.96) | b684094 (0.2.38 published) | ca518c4 | app 144427c, FlexKit 4222c20, FlexCore 45a0f81, outer 78fb27d | first run of the login-skip guard; FlexKit packed once, byte-identical nupkg on both branches |
| 09-12 16:59 | 8f03fa3 | 9910e05 | 48eb73c (0.1.95) | 8cf89d7 (0.2.37 published) | 0519173 | app f50421f, FlexKit bfe8672, FlexCore f3ebcd2 | restored MR !33 that our 09-09 deploy had reverted; login-skip still shipped (guard not yet written) |
| 09-10 23:45 | 6de7a2b | 7c354ff | e7f6123 (0.1.94) | — | 5f23759 | app 63b16b4, FlexKit e7c8599 | FlexKit edited mid-run → R1-UAT got an unverified 0.1.94 |

nuget.org has no FlexCore 0.2.40 (skipped). FlexKit is not on nuget.org — `Deploy/local-packages` is its only ship route.

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

## Open as of 2026-09-16 13:30 (from the live preflight and refresh when this skill was written)

- **Teammates' work not yet in our tree.** hyphen-pb main moved 1313b21 → **432c842**: Irfan 40ef498
  (FDuplicateModelOptions, HHM-977/978/987/990) + merge 926f1a7 (HHM-991); Deepika 3581c61 (FCustomQuote,
  HHM-1061) + merge 5453b8e; Irfan 29f8287 (FOptions — BuildPro Settings, Community Standards); Irfan
  432c842 (FAddProperty.razor/.css, FOptions, new wwwroot/images/16/AddProperty.ico). Our app HEAD and
  working copy hold the PRE-commit blobs. **A push without `--pull` reverts all six.** None overlap the
  app's current uncommitted files.
- **FEstimateItems.razor has four divergent versions** — hyphen-pb main, `origin/wip/festimateitems-local-20260914`
  (d145722, "WIP snapshot before laptop migration", +236/−176), `origin/HHM-1057` (d9ae4c4, "seed and
  persist the gItems layout from the VB6 design-time columns", +283/−223, based on 29f8287, not merged),
  and the app working tree (being edited). Owner decision.
- **Another session was actively editing** the app (14 modified, 2 untracked), FlexKit (30 modified,
  11 untracked — Reports designer work) and FlexCore (25 modified, 11 untracked) at 13:11–13:24.
- **FlexKit → FlexCore port still owed** for FlexKit-only dirty files: DatePickerControl.razor,
  Grid/TreeGridControl.razor.cs, TextBoxControl.razor, wwwroot/legacy-scrollbar.js, wwwroot/textbox-control.js.
- **Next versions:** FlexKit ≥ 0.1.102 (0.1.101 burned), FlexCore ≥ 0.2.44 (0.2.43 published).
  FlexCore.Llm.csproj says 0.2.42 and FlexCore.Documents 0.2.35 (never published) — decide whether they
  track FlexCore.
- **Owner decisions still open:** Atlassian token rotation (App_Data/jira-settings.json on hyphen-pb
  main, R1-UAT, homefront main); the QA login page's live dev panel; the runtime-limitations panel on 14
  Estimating reports; designer saves to App_Data/report-designs that nothing reopens; session 5c72594a's
  HomeFrontPB FReportDesigner edit and its paste-image patch.
- **Housekeeping:** outer repo has uncommitted FlexKitTester Crystal bench work (another session) and
  stray `MobileSource/m/Program.cs` + `MobileSource/t_rec3.sh`; the deploy never pushes the outer repo.
