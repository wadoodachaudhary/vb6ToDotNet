# Update Repositories — history

Append one row per run, newest first. Heads are the pushed tips; "local" = commits made in the working
repos for that run. Git cannot tell which Claude account ran a deploy — every deploy commit carries the
same staging-clone identity — so say so in the hand-off note if it matters.

| When (EDT) | hyphen-pb main | homefront | flexkit (FlexKit) | flexcore GitHub (FlexCore → nuget.org) | R1-UAT | Local commits | Notes |
|---|---|---|---|---|---|---|---|
| 10-10 19:50 | 7a7d9ba | 7066e97 | 0b3a842 (0.1.117) | **NOT pushed**; GitHub main remains 28e409b, NuGet 0.2.48 held | R1-UAT 65cacba and R2-UAT f331f2d untouched | app 27b58e0, including b11230d, b4cf13a, 9e53386, cf40b3e, 28f2207, 584d6a0; FlexKit f00553b including a65dc60, 060ac15, d6367e4; outer 1787f6a | Owner requested GitLab/GitHub updates and explicitly approved merging/verifying both HHM-1243 branches (app 3e68ecd, Kit fc37b43). Integrated Irfan be5a1d2 (HHM-1245/1217) and Deepika 7c502cb (HHM-1231). Preserved newer formula, tree-rename, invoice/PO/native-report work. Review caught and fixed Save As ignoring refused saves and formula drafts missing the leave guard. Multi-PO report binding and parent/subreport projection regressions also fixed. Fresh 0.1.117 follows cached dry-run 0.1.116. Two normal 10-minute gates passed; package build 0 errors. Remote heads and protected config hashes verified. Generic verifier reports only the known non-runtime tests on unchanged, skipped FlexCore; eligible targets pass. Private GitHub mirror source tips: app 27b58e0, Kit f00553b; outer includes this record. |
| 10-08 11:32 | 7809259 | d1fe4b0 | ff5becf (0.1.115) | **NOT pushed** — FlexCore checkout was on `cursor/codeconvert-converter-shell-bda2` (PR #4); GitHub main stays 28e409b (Cursor PRs #1/#5 + ButtonControl/Sunburst commits made on GitHub) | **R2-UAT NOT updated** (owner: "not R2-UAT"); stays f331f2d (Irfan). R1-UAT 65cacba untouched | app c160780 over 1b254d0: a70bd78 + fcc2f83 + 2cada3e + 3493093 + 50a1424 (Homefront / task sessions: Feedback sync labels + sweep guard, HHM-1242 permissions on every FMain surface, Feedback Dashboard), dcc5323 (Codex 10-06: User Permissions no password reveal), de1d738 (Codex 10-08: HHM-1257 formula editor), 886ebc7 (merge of Irfan 1ce4ceb3 Sage Intacct + 8fd6789a HHM-1230/1229), c160780 (Baaria MR !47); FlexKit 66824f7 (TextAreaControl host commands, 0.1.115) | owner: "Update repositories per standing rules. Pull and merge and close any MRs (especially by Baaria Chaudhary). Do not update R2-UAT." **GitLab SSH had moved**: gitlab.innovatixinc.com now resolves to Cloudflare (web only); two hours lost testing NordVPN / FortiClient / a PAT before DevOps gave `git-ssh.innovatixinc.com` (same host keys) — staging clones now use it with `core.sshCommand = ssh -o HostKeyAlias=gitlab.innovatixinc.com`. MR !47 taken by safe merge vs merge base eeea330 (ours == base for all 6 files), commented and CLOSED; 0 open MRs. Two EstimateChecks pins updated for her changes (retainage keeps one %, comparison uses the job's community). Codex's uncommitted work reviewed (wf_1367ae3b, no blocker; fixes applied) and committed. New guard: deploy skips flexcore and the gate ignores it while the FlexCore checkout is off main. Round was briefly handed to the Homefront session and taken back on the owner's direct instruction. verify_deploy: 3 repos shipped, every check passed EXCEPT one about GitHub FlexCore main carrying tests/ (from Cursor's merges, not from this deploy). |
| 10-05 08:28 | da818da | e330e8b | 2ddf93e (unchanged) | 65af124 (unchanged) | **R2-UAT 2cda38b** (kept its own azure-pipelines-uat.yml; lockdown intact) | app 1b254d0 (owner: Custom Requests is not an R2 item — marker and green highlight removed from the Tasks menu row, both sidebar items and the PM workflow node) | owner: one mistake to undo, "no quiet window, no regression testing, as fast as you can" → no local build/sweep/dry run; the deploy's own staged builds (main and R2-UAT, 0 errors) were the compile check; gate with QUIET_MINUTES=0. Irfan had pushed cddf80a (pipelines only, FileTransform v2) → `--pull` to advance the clone, nothing to sync. His change to azure-pipelines-uat.yml on main CONFLICTED with R2-UAT's own copy (found with `git merge-tree` before the run): the R2 stage now keeps the branch's copy for azure-pipelines*.yml / appsettings*.json / NuGet.Config conflicts and still stops on any other conflict. verify_deploy: ALL CHECKS PASSED. |
| 10-05 07:54 | cd3368f | 497d890 | 2ddf93e (0.1.114, unchanged) | 65af124 (unchanged; 0.2.48 NOT published) | **R2-UAT 69e93bf** (main merged in over Irfan's b77dab0; lockdown + his new UAT pipeline intact). R1-UAT 65cacba untouched | app 7ca4ae0 (Homefront session: Custom Takeoff basic import wizard) + 6c3a740 (owner's menu change: (R2) marker on five wizard rows; _deploy ignored) + ca61551 (merge of Irfan 930f452 / MR HHM-1191: PO Price Update Wizard visibility, FMain option flags in one query, OptionsChanged) | owner: "Update all repositories per standing rules. Minor changes made to the menu system since last release." Irfan had pushed 930f452 + cc9643f + ec32355 (pipeline) to main and merged main into R2-UAT himself (4255fbf, b77dab0 UAT pipeline): `--pull` 3-way merged FMain cleanly (ours = his file + the five marker rows); azure-pipelines.yml / appsettings.json repo-owned, untouched. **An untracked `_deploy/` drop folder in the app root (HomeFront.zip 127 MB, FlexKit.zip, the owner's Windows-box zips) would have shipped** — now excluded in deploy_to_repos.sh and .gitignore. Harness sweep first showed 4 failures: 3 because the local SQL container was down (Docker not running), 1 stale EstimateChecks source pin after Irfan's refactor (updated; same rule). Owner waived the 10-minute quiet window ("nothing happening") → gate run with QUIET_MINUTES=1. No library change; FlexKit 0.1.114 package reused. verify_deploy: ALL CHECKS PASSED. |
| 10-01 14:18 | 77ba08f | 0a6fd00 | 2ddf93e (0.1.114) | 65af124 (0.2.48 NOT published — still held) | **R2-UAT f37822b** (main merged in; lockdown + pipeline intact). R1-UAT 65cacba untouched | app d05d68f (Homefront session: HHM-1122 / HHM-236 Custom Takeoff mapping precedence, every column in the mapping panel, Order Qty from the file, clean state per pick, VB6 number parsing) + 7fd1d18 (Book of Accounts sizable, size per user/caption; FItems New Phase message window offset by margin); FlexKit 05a7099 (DialogControl resize from the rendered box, drag + resize share one measured press, left/top instead of transform; 0.1.114); FlexCore 2c8b8b4 (mirror) | owner: "Fix both HHM-1122 and 236", "Dialog Size first", then "Deploy". The HHM-1122 regression shipped at 12:04 (a duplicate-heading column could not be mapped) is fixed on main and the client branch. DialogControl change reviewed three times (wf_f5b14959 / wf_48b54e65 / wf_fe1934b3): round 2 caught a blocker in a consumer (FItems New Phase window jumped ~430px because inline left replaces a stylesheet left) — fixed before commit; every check added for a finding was shown to fail without its fix (mutants on isolated copies). DbGridDialogChecks 183, R2QaChecks 814, CustomTakeoffChecks 194, sweep 30/0. Two sessions shared the app checkout: the Homefront session built in an isolated git-archive copy and committed only its paths. verify_deploy: ALL CHECKS PASSED. |
| 10-01 12:04 | ac916ab | 703307c | d9b6ffe (0.1.113) | b7ef9f6 (0.2.48 NOT published — still held) | **R2-UAT 5432dd1** = merge of main (first `--with-r2uat` round; lockdown + pipeline intact). R1-UAT 65cacba untouched | app da50493 over 8584ae3: f34efeb (Irfan cfcaa35 HHM-1182/1185 as upstream), 51dfc51 ((R2) markers — owner: a feature), 362752f (Book of Accounts remembers its position; FAssembly read-only combos), 6c0a2c2 (User Permissions blank non-viewable password), 65fec6e + da50493 (cloud HHM-1122 Custom Takeoff Excel import, merged and fixed by the Homefront session), 47a8165; FlexKit 739d952 (c34fc37 password policy; 72b0313 grid fix held back; 0.1.113); FlexCore 6559e66 (mirror) | owner: "Update repositories", "take care of existing items, close MR", "R1 is now R2 … releasing code to the client on R2 branch", later "as per standing rules ASAP". MR !43 commented + closed; 0 open MRs. `--with-r2uat` stage written and fixture-tested (tests/r2uat_fixture.sh, 22 checks) and used for the first time. Codex's TextBoxControl.ExistingPasswordViewable reviewed (wf_03805444: no blocker for User Permissions; hardened same-element re-bind + no AutoFocus on a recreated input). **Grid combo click-away fix written, reviewed (wf_3a9f292f: 2 blockers, 3 warns) and HELD BACK** on `wip/grid-combo-clickaway`. The Homefront session merged cloud branch jira/HHM-1122 mid-round; its own suite run found a heading-detection regression (unlabelled csv lost a row) and fixed it (da50493) before the push. Dry run #1 packed FlexKit 0.1.112 with the grid fix in the tree → burned → 0.1.113. The real run exited 128 AFTER every push succeeded (R2 stage RETURN trap: `merge --abort` with nothing to abort under set -e) — fixed. verify_deploy: ALL CHECKS PASSED, 4 repos + R2-UAT. New standing step 13: push the GitHub mirrors level with the Mac. |
| 10-01 01:40 | 8be0d36 | 250c041 | 5d393da (0.1.111, unchanged) | c03240b (unchanged; 0.2.48 NOT published) | ca518c4 frozen | app f7c6a65 (MR !43 taken) + 8584ae3 (Deepika 438cad9 Preview rule restored) | owner: "Shipmit" (= ship it) for MR !43. Took Baaria's HHM-1081 (Use_Timberline from System_Setup; safe merge vs a8fb8cb, clean; EstimateChecks 2498). Preflight then raised **LOST** on FPurchaseOrder.razor (Deepika 438cad9): her two attribute changes were dropped by our 09-26 merge (dff7e97 kept our side) and overwritten on main — NOT by today's pushes (absent from origin before and after). Resolved against VB6 FPurchaseOrder.frm's FieldsLocked setter: Preview `Enabled = Not RHS` (frm:2184) is hers → RESTORED (web had Disabled="true", Preview was dead); txtPODescription `Enabled = Not RHS` (frm:2202) locks with the fields → her removal of the lock deliberately NOT taken; both recorded in preflight-ack.txt. First guess ("superseded by our 0c7d060") was checked and was WRONG — 0c7d060 never touched those lines; the check caught it. MR !43 still shows OPEN in GitLab: the Chrome extension disconnected before I could close it — comment + close pending. |
| 10-01 00:55 | a50e6ff | 51f8640 | 5d393da (0.1.111) | c03240b (0.2.48 NOT published — still held) | ca518c4 frozen | app a6039f2 + d962075 (over 1f0e18d); FlexKit 035b196 (over 685ea54 which another session committed from my working tree, + its 1cf9e81 MdiHost); FlexCore 510d40a | owner: "wait 10 minutes and then update repositories — ship it". Shipped the three review-driven fixes (dms_documentclasses per-row delta in a tx, TBD raw captions, combo blur handoff) WITH the follow-ups their own adversarial review demanded (blanked loaded row never deleted + VB6 ValidateEdit mirrored; renames ordered; native blur yields to the browser module — Blazor's blur is a document-level capture listener so BOTH paths reached the server; no-pick forced blur fixes an AutoCompleteControl double-fire), plus a sibling session's page work (HHM-1087 movable Book of Accounts, HHM-955 Builder 1440 Style editors, Model Style spacing, System Settings save notification) after ITS review found a blocker — the dragged dialog snapped back on re-render (Position unbound; trap 2 of dialogcontrol_hosted_grid_traps) — fixed host-side before the push with CloseOnEscape off (VB6 parity) and data-fx-key-scope on the grid. `--pull` took Irfan a106829 (HHM-1181/1114; FOptions 3-way merged, 100/100 incoming + 120/120 local lines kept). MR !43 (Baaria, HHM-1081 Use_Timberline from System_Setup) is OPEN and NOT taken (standing rule: only when asked) — its one-line fix is not in our tree. 29/29 harnesses; GridLayoutSaveChecks 418 (real persister), TakeoffSettingsChecks 243 (client-path blur), R2QaChecks 792, DbGridDialogChecks 42→48 Chrome, DropDownOpeningChecks suite green. |
| 09-30 06:55 | f626515 | 1413f3e | 3ddf013 (0.1.110) | 2f305cb (0.2.48 NOT published — still held) | ca518c4 frozen | app 1f0e18d (merge of Irfan e9dd889 + Deepika c8e39e9, HHM-1175 Mass Change pane, password standard-input, HHM-1178 batch, HoldbackInvoiceChecks) over 63b2858; FlexKit 69a850c; FlexCore 2445ffb (mirror) | owner asked why HHM-1175 was not deployed — it had never been COMMITTED (uncommitted css). Round had been held since 09-29 on two blockers. (1) SECURITY: FUserPermissions decrypts every stored password (FLogin.MyDeCrypt) into the bound box with the reveal eye enabled, where VB6 frm:2509 uses a random Chr(1) sentinel — raised twice, **owner confirmed it is intended (standard input behaviour)** and it shipped; recorded in the app CLAUDE.md with the safe form if ever revisited. (2) DATA LOSS: the grid's new TextChanged staging made the dropdown's blur publish a no-op, so a typed custom combo value was staged and never written to the row (proved with an instrumented isolated harness copy: dropdownValue=new, rowValue=old). Fixed with an additive opt-in EditableCommitted callback in both libraries; TakeoffSettingsChecks 237 green again. PropertiesTreeChecks' panel-press assertion restated (guard now covers editable+hosted; a plain form list still gets none). `--pull` took Irfan e9dd889 + Deepika c8e39e9; FOptions 3-way merged; two of our hunks dropped in favour of theirs after checking VB6 (gStdItems picker button belongs on every row; empty OnMBAPILevelChanged replaced by a real handler). Gate REVERT flags on FOptions acked (one superseded by Irfan's own later commit, one the owner-approved password change). 29/29 harnesses, 2 clean builds. |
| 09-28 07:40 | 93cb86a | 103a4c9 | 3616b6d (0.1.109) | 8c65c30 (0.2.48 NOT published — still held) | ca518c4 frozen | app 9d36dd1 (AppGridLayout hidden columns keep their slots, VB6 DBPutGrid) + 3e8f067 (HHM-955/964 + review fixes); FlexKit 5a0d14b; FlexCore d8438bf (mirror) | owner asked to update the repositories, merge/pull and look for conflicts, then deploy once the review landed. NOTHING to pull: all four remotes still at the 09-27 push, zero open MRs, no conflict markers, every teammate change already in our trees. Shipped a sibling session's 09-27 work. Review wv_d9ace1f1-303 (6 lenses x 3 adversarial verifiers, 108 agents) returned **18 confirmed, 9 blockers**, 16 refuted — the push was HELD and every blocker fixed first: (1) `CellSelectablePredicate="@IsCellEditable"` made a POSTED worksheet completely inert (IsCellEditable starts `if (IsReadOnly) return false;`) → new `IsCellSelectableForGrid` exempts a read-only sheet; (2) `if (asmRows.Count == 0) return -1;` silently dropped every Global Option row (VB6 tests EOF on a UNION that also reads tblphaseitem, frm:4348-4384) → removed; (3) gating `HandleCellDblClick` killed `OnRecordDoubleClick`, which fires exactly on the blocked columns (worksheet Assembly Costs died) → ungated; (4) `HandleCellContextMenu` returned before updating the current record, so the host menu reached by bubbling acted on the PREVIOUS row (wrong-row Remove/Refresh Costs) → record follows the right-click; (5) the structural checkbox cell (-1) was gated → exempt; (6) an unresolvable item counted as blocked, so a virtualized grid pruned selections outside its provider window → permissive, matching IsCellEditableForItem; plus an O(columns) fast path for the per-render grid-edge scan. 7 new GridSelectionChecks cases pin fixes 3-6, each PROVEN to fail against an isolated pre-fix library copy (83 checks now, against both libraries). 28/28 harnesses, 5 clean builds, browser stages 142+109. |
| 09-27 00:08 | 161809c | f9ce0c2 | 99fe9a9 (0.1.108) | 7b0cfdd (0.2.48 NOT published — still held) | ca518c4 frozen | app 1fbb5e5 (notes; de17e8d HHM-1157 rename cap + dialog containment shipped from the same tree); FlexKit 51d62fa (HHM-1170 cell-mode selection stops covering authored column colours, 0.1.108) over 6b41bd9; FlexCore 6c22b4a (mirror) over 400d41b; outer 5cbc219 (skill extension) + 44b980b (FlexKitTester HHM-1170 benches, Crystal bench notes) | owner asked to finish the UpdateRepositories skill, then update the repositories and merge/close any open MRs. Skill grew four reference files (merge-requests, nuget-and-versions, flexkit-flexcore-sync, verification, handoff) + hard rules 10-13. Review wv4458p4v (3 lenses + adversarial verify) landed BEFORE the push: 9 confirmed, all info, 7 refuted — fixed the four that were cheap and clearly right (vb6-windows theme + four `var()` fallbacks still on the pre-fix #f5f5f5; `setRowPreview`'s cell branch never drained `paintedPreviewEls` on the preview-off path the keyboard callers use; `restoreSelectedLook` never cleared `data-fx-selection-muted`). 27/27 harnesses, 5 builds clean, bench re-run 0 failures at 0/400/800 ms RTT. No teammate commits to take (every one already in our trees). MR !41 CLOSED, not merged — all four of its tickets are on main and merging would have put back its unconditional HHM-1159 DELETE. Zero open MRs now. |
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

## Open as of 2026-10-10 19:50 (main-only repository update)

- **Held targets unchanged:** HomeFrontPB/POC were not edited, built or deployed. R2-UAT stays f331f2d; R1-UAT stays 65cacba. FlexCore remains on `cursor/codeconvert-converter-shell-bda2`; shared ports are retained uncommitted alongside existing work. GitHub Core main stays 28e409b and NuGet 0.2.48 remains held. Do not switch/reset/commit that checkout as part of this handoff.
- **Verification:** HomeFront.sln nonincremental build 171 warnings/0 errors; 32 offline suites/0 failures, including 3,597 EstimateChecks and 151 FormulaEditorChecks. SQL RecordLockChecks 242 passed in scratch databases, with source fingerprint unchanged. About/Table Locks/Logged In Users browser 136 checks each at 0ms and 150ms each-way latency; leave guard 85; InputDialog 300; formula browser 261; DbGrid browser rerun 183. All four scratch databases dropped and fixture servers stopped. No QA/Sample data writes or external email delivery.
- **Known drag race:** one rapid double-drag check initially failed, then the unchanged full suite passed. A separate baseline stress run served premerge GitHub dialog-control.js and reproduced the identical lost first delta on attempt 3 at 300ms each-way latency (expected +60/+25, actual +30/+15). Drag C#/CSS are unchanged. Existing intermittent behavior, not fixed in this repository update; assertion retained.
- **Build exception:** Core, Mutarjim and FlexKitTester pass; FlexCore.Showcase still fails its existing Sqlite 10.0.8/10.0.12 downgrade. Native report visual approval and application-wide live VB6 parity are not certified. Earlier audit gaps (HHM-1254 Takeoff title, HHM-1241 Yes/No buttons, HHM-1178 matching reset data) remain; unrelated live writes were not executed.
- **Security/config:** source added-line scans found no new secret candidates. Appsettings, pipelines, NuGet.Config and existing Jira settings blob hashes are unchanged. The already-tracked Jira token remains an owner rotation decision; it was not printed or modified. Local DevAutoLogin stays functional and is absent from GitLab deploy copies, as required; private GitHub mirrors retain source.
- **Next version:** FlexKit >= 0.1.118 if code/assets change after this run. No Jira moves or unrelated MR merges/closures were performed. Receipt logs: `/tmp/hhm1243-*` and `/tmp/repository-update-offline-checks/`.

## Open as of 2026-10-08 11:45 (after the 10-08 deploy — main only)

- **R2-UAT was NOT updated** on the owner's instruction. It is at f331f2d (Irfan); main 7809259 is ahead of it.
  Next client release: `--with-r2uat` when the owner says.
- **FlexCore is unresolved.** Local checkout is on `cursor/codeconvert-converter-shell-bda2` (open PR #4; PR #3
  artifact viewers also open) for the Mutarjim IDE. The FlexKit 0.1.115 TextAreaControl port sits UNCOMMITTED
  there (byte-identical) — never commit it on that branch. Local FlexCore `main` (2c8b8b4) lacks GitHub main's 9
  newer commits (Cursor PR #1 crystal charts, #5 samples DB, ButtonControl btn-* classes, Sunburst format). To
  ship FlexCore again: return the checkout to main, merge github/main, apply the TextArea port, build
  FlexCore.Showcase, then deploy. GitHub FlexCore main also now carries `tests/` (from Cursor's merges), which the
  "non-runtime files stripped" rule forbids — owner to decide.
- **Owner questions from the Codex review:** Setup > Project Managers still shows the same login password with a
  reveal eye (10-06 rule named User Permissions only); an existing user's password can no longer be cleared to
  blank on User Permissions. HHM-1257: live database save/reopen not exercised.
- **Grid combo in-grid click-away fix still NOT shipped** (`wip/grid-combo-clickaway`); the owner has the new plan
  (browser sends the draft ahead of the click) and has not yet chosen "fix only" vs "fix plus save once".
- **Manual PO browser suite stale since 09-26.** **FlexCore NuGet still held** at 0.2.48.
- **Next versions:** FlexKit >= 0.1.116 (0.1.115 shipped), FlexCore >= 0.2.48 (unpublished).

## Open as of 2026-10-05 08:30 (after the second 10-05 deploy)

- **R2-UAT's azure-pipelines-uat.yml still has Irfan's pre-FileTransform-v2 content** (the branch keeps its own
  copy; main's cddf80a version triggers on R1-UAT). If he wants v2 on the client branch he edits it there.

- **Irfan's HHM-1191 rule differs from VB6:** the PO Price Update Wizard now shows for AccountingSystem 1 (Sage 300),
  7 (Xero), 10 (QuickBooks Online), 11 (D365); VB6 FMain shows it for 1, 9 (Intacct), 11. Taken as pushed (merged
  MR); the owner was told.
- **Grid combo in-grid click-away fix still NOT shipped** — `wip/grid-combo-clickaway`.
- **HHM-1122 leftovers** (VB6 target columns, Excel displayed text, blank rows, Custom Requests reimagining) —
  Homefront session / owner.
- **Manual PO browser suite stale since 09-26.** **FlexCore NuGet still held** at 0.2.48.
- **Next versions:** FlexKit >= 0.1.115, FlexCore >= 0.2.48 (unpublished).

## Open as of 2026-10-01 14:20 (after the fourth 10-01 deploy)

- **Grid combo in-grid click-away fix still NOT shipped** — branch `wip/grid-combo-clickaway` (FlexKit d7ffed7,
  FlexCore 1ec58db, app 81628fa); see memory note grid_combo_clickaway_fix_held_back for the review findings to fix.
- **HHM-1122 left for the owner / next round (Homefront session):** VB6 target columns WBS1-40 / Assembly / Model /
  TaxGroup / part-number lookup, Excel cells as displayed text, blank rows, clearing the mapping section (needs a
  RemoveSection on IUserPreferencesService), and the Custom import wizard the owner asked for.
- **DialogControl, accepted residuals:** two presses inside one round trip on a very slow link are serialised, not
  merged; the non-modal `:has(:active)` fast path is Chromium/WebKit only (Gecko waits one round trip, as before).
- **Manual PO browser suite is stale since dff7e97 (09-26)** (PO Description moved, lost MaxLength=100). Not started.
- **FlexCore NuGet still held** at 0.2.48 (GitHub 65af124).
- **Next versions:** FlexKit >= 0.1.115 (0.1.114 shipped), FlexCore >= 0.2.48 (unpublished).

## Open as of 2026-10-01 12:10 (after the third 10-01 deploy — first with R2-UAT)

- **Grid combo in-grid click-away fix is NOT shipped** — branch `wip/grid-combo-clickaway` in FlexKit (d7ffed7),
  FlexCore (1ec58db) and the app (81628fa). Before it returns: release the commit lock when a waiter times out (or
  never re-raise OnCellSave for an edit whose save is in flight), keep the same press's click behind the mousedown
  leg, do not commit an untouched combo's display text, pin the staged draft with an EndEditAsync that causes no
  blur. Details: memory note grid_combo_clickaway_fix_held_back.
- **DialogControl resize** (owner asked 10-01): measure the rendered box at resize start, keep the opposite edge
  fixed, raise PositionChanged; then make the Book of Accounts window sizable and persist its size. Not started
  at deploy time.
- **Manual PO browser suite is stale since dff7e97 (09-26):** PO Description moved to the right header and lost
  MaxLength=100 + buffered typing. Chip offered, not started.
- **HHM-1122 (Custom Takeoff Excel import) shipped to main AND R2-UAT** on the owner's "as per standing rules ASAP";
  the owner had not tested it on the Mac. The Homefront session's static review of it was still running at push
  time — if it confirms anything, that is the next round.
- **Password box library contract (warn, not reachable from User Permissions):** with ExistingPasswordViewable=false
  a ValueChanged handler that awaits or normalises before storing the value wipes the replacement. Documented, not
  fixed. Harness gap: input recreation / listener re-bind are not asserted.
- **Text cells save 3× with an asynchronous host save** (blur, mousedown, click each commit while the first save is
  in flight) — pre-existing, measured 10-01, fixed only on the held-back branch.
- **FlexCore NuGet still held** at 0.2.48 (GitHub b7ef9f6).
- **Next versions:** FlexKit >= 0.1.114 (0.1.113 shipped, 0.1.112 burned), FlexCore >= 0.2.48 (unpublished).
- **GitHub mirrors:** app `github main`, FlexKit `github telerik-parity-20260904`, outer `origin main` are pushed
  level with the Mac at the end of every round (step 13).

## Open as of 2026-10-01 01:45 (after the second 10-01 deploy)

- **MR !43 (Baaria, HHM-1081) — its content SHIPPED in f7c6a65 → main 8be0d36, but the MR is still OPEN in GitLab:** the
  Chrome extension dropped before I could comment and close it. To finish: comment "Shipped in main 8be0d36 (taken
  against merge base a8fb8cb, clean); nothing left to merge" and close.
- **Deepika's 438cad9 description-box change deliberately not taken** (VB6 frm:2202 locks it with the fields) — acked.
  If she meant something else by it, she will say so; the Preview half is restored.
- **FlexCore NuGet still held** at 0.2.48 (GitHub c03240b). Everything else from "Open as of 2026-10-01 01:00" stands.
- **Next versions:** FlexKit >= 0.1.112, FlexCore >= 0.2.48 (unpublished).

## Open as of 2026-10-01 01:00 (after the 10-01 deploy)

- **MR !43 (Baaria, HHM-1081: read Use_Timberline from System_Setup, not AppOptions alone) is OPEN and untaken.** One
  line in FEstimateItems (`_useTimberline = await ResolveSystemSetupFlagAsync("Use_Timberline", ...)`); the helper
  already exists in our tree, our line 1909 still reads AppOptions only. Clean to take when the owner says so.
- **FlexCore NuGet still held** at 0.2.48, judged on GitHub main **c03240b**.
- **From the two reviews, deliberately not done:** VB6 FDBGrid is sizable and remembers size/position (IniGetForm) —
  the web Book of Accounts is Resizable=false and persists nothing; read-only FAssembly disables the Style combo where
  VB6 leaves cboStyle enabled; Builder1440StyleOptions.cs hardcodes Style choices (review asked where VB6 sources them —
  unanswered).
- **Combo handoff — one verification gap:** a fresh-page DOM probe shows click-away after typing (module attached)
  committing the live draft and closing the editor, and TakeoffSettingsChecks drives the client blur action; but a
  scripted check against the DropDownOpeningChecks fixture's `#client-state` key times out although the rendered cell is
  right. Suspect the fixture's grid-row state plumbing, not the control. The failing script was removed, not shipped.
- **A render landing inside the TBD save's few-millisecond window shows raw field-name headers for one frame** — accepted.
- **Next versions:** FlexKit >= 0.1.112 (0.1.111 burned), FlexCore >= 0.2.48 (unpublished).

## Open as of 2026-09-30 07:00 (after the 09-30 deploy)

- **FlexCore NuGet still held** at 0.2.48, judged on GitHub main **2f305cb**. Crystal-reader defects unchanged.
- **Owner decision on record (2026-09-30): the User Permissions password box uses standard input behaviour.**
  It decrypts every reversible stored credential into the bound box with the reveal eye enabled, so any user
  with EditSecurity can read other users' live passwords. Raised as a security concern twice before the push
  and reaffirmed by the owner. Do not silently revert it; the safe form, if it is ever revisited, is the `$h$`
  branch (sentinel + disabled eye until a replacement is typed).
- **From review wf_29a1ab56-aad (21 confirmed), NOT fixed this round** — worth a pass: FOptions Save still
  wipes all of `dms_documentclasses` before re-inserting outside a transaction; the Assign TBD grid writes
  web-invented captions into the shared AppGridLayout table; Department Approver and GL Account pickers are
  single-select where VB6 passes `MultiSelect:=True`; deleting an approver re-selects a row the user never
  picked and Edit then zeroes its limit; seven new SplitterControl conversions omit `MaxPrimary`; and three
  test-quality findings, including a worksheet check that proves "focus returns once per rejected commit"
  through a method that has never existed.
- **Unverified:** whether `preventDefault` on an EDITABLE dropdown panel affects scrollbar dragging. The guard
  now covers editable panels; a plain form list is unaffected and PropertiesTreeChecks pins that.
- **Next versions:** FlexKit >= 0.1.111 (0.1.110 burned), FlexCore >= 0.2.48 (unpublished).

## Open as of 2026-09-28 07:45 (after the 09-28 deploy)

- **FlexCore NuGet still held** at 0.2.48, judged on GitHub main **8c65c30**. Unchanged reasons (Crystal-reader
  defects: Latin-1 LoadString, CROSS JOIN for dropped joins, diagnostics not surfaced). Five rounds of content
  now sit under that one unpublished number.
- **HHM-955 / HHM-964 need the owner's live retest.** Both shipped with ship-review fixes on top of the sibling
  session's work, verified by harness and an isolated pre-fix mutation test but NOT in the owner's running app.
  The selection policy in particular changed shape: non-editable cells still take no cursor (owner directive
  2026-09-27), but a READ-ONLY (posted) worksheet is exempt and stays fully navigable, and right-click /
  double-click keep working on blocked cells. Worth a look before these move past Development Review.
- **Not fixed, deliberately — recorded from review wv_d9ace1f1-303 (all pre-existing, none introduced this round):**
  `AddAssemblyToGridAsync` INSERTs the detail row outside a transaction and then calls
  Purch_GetAssemblySalesSheetCost and CalcData; a throw in between leaves an orphan row that makes the next Add
  of the same model hit PK_tblSalesSheetDetails, and the HHM-1159 pre-delete will not run because its key was
  already consumed. Fixing it means a transaction or a compensating DELETE — a change to DB write semantics,
  not a thing to rush into a deploy.
- **Test gap left open:** WorksheetStyleChecks' browser.mjs needs a live app, a local SQL Server and a
  hand-made disposable worksheet, so it cannot run in run_harnesses.sh and its client-side guards are not
  pinned by anything that runs unattended. The C# side now is (GridSelectionChecks, 83).
- **MRs: none open**, and the stale branches (HHM-1053, HHM-1057, HHM-1057-persist-on-first-use,
  HHM-922-964-1159-1161, festimateitems-phase-dropdown, wip/festimateitems-local-20260914) still have no MR
  behind them — deleting them is the owner's call.
- **Next versions:** FlexKit >= 0.1.110 (0.1.109 burned by this round's pack), FlexCore >= 0.2.48 (unpublished).

## Open as of 2026-09-27 00:25 (after the 09-27 deploy)

- **FlexCore NuGet still held** at 0.2.48, judged on GitHub main **7b0cfdd**: the Crystal-reader defects that held
  0.2.48 on 09-26 were not worked on in this round, so the hold stands unchanged (TslvStreamReader.LoadString decodes
  UTF-8 where Latin-1 is needed; a dropped legacy join still renders as CROSS JOIN; ConversionDiagnostics are still not
  surfaced at runtime). A held publish is not resumable — when it is finally published, 0.2.48 will carry every change
  accumulated since 4566b0e, now four rounds of content.
- **MRs: none open.** !41 (HHM-922/964/1159/1161, Baaria) was closed with a comment naming main 161809c and explaining
  that HHM-1159 shipped as the narrower `_savedDetailKeys`/`_removedDetailKeys` variant (6272974); GitLab also reported
  it as conflicted. The stale branches HHM-1053, HHM-1057, HHM-1057-persist-on-first-use, HHM-922-964-1159-1161,
  festimateitems-phase-dropdown and wip/festimateitems-local-20260914 still exist with no MR behind them — deleting them
  is the owner's call.
- **HHM-1170 needs the owner's live retest.** The fix is verified on the FlexKitTester bench only (cell/row handoff
  316/460/606 frames; 30 colour cases; 0 failures at 0/400/800 ms added RTT). The owner's app was not restarted, so the
  worksheet check against the September 26 recording is still outstanding. The ticket stays in Development Review.
- **Info findings left unfixed from review wv4458p4v** (all verified as non-defects today): cell-mode mute is CSS-only,
  so it cannot suppress the inline `fx-selected` row paint the server writes — unreachable while cell-mode grids do not
  set HighlightSelectedRows, but it is a latent trap for the first one that does; and `GridControl.razor.cs:765`'s
  PersistenceKey doc comment names HomeFront (one of 14 such mentions in the library — a sweep, not a deploy fix).
- **Still open from earlier rounds:** whether GitHub FlexCore main should carry `docs/` and `tests/` (the strip removes
  them every deploy); Jira tickets left in Development Review (HHM-1157, HHM-1166, HHM-1138, HHM-1132); Atlassian token
  rotation; QA login dev panel; FlexCore.Llm / FlexCore.Documents versioning.
- **Next versions:** FlexKit >= 0.1.109 (0.1.108 burned by this round's pack), FlexCore >= 0.2.48 (unpublished).
- **Housekeeping:** the outer repo's bench work is now committed (44b980b); `MobileSource/m/` and `MobileSource/t_rec3.sh`
  are still untracked scratch. The deploy never pushes the outer repo — push it to GitHub by hand.

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
