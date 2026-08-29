# Start here — onboarding a new Claude session on vb6ToDotNet

Two parts. **Part A** is the prompt to paste. **Part B** is the context the
human operator needs to hand over correctly.

---

## PART A — paste this into the new Claude session

> Before doing any work on this codebase, read yourself in.
>
> **1. Establish which context you have.**
> Check whether a persistent memory store is loaded for this project. The live
> one is at
> `~/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/`
> and contains 126 files with a 125-line `MEMORY.md` index (counts as of
> 2026-08-29 — the store grows continuously, so treat this as a floor, not an
> exact match; the check that matters is 6 vs. many, below).
>
> - If you have it, read `MEMORY.md` first, then
>   `vb6todotnet_workstream_map.md`, `project_overview.md` and
>   `operating_rules.md`.
> - If you see a memory store with only **6 files** you are on a stale store
>   from May 2026 — say so and stop; every standing rule is missing from it.
> - If you have **no** memory at all, you are on a different account. Say so,
>   then rely entirely on the files in step 2.
>
> **2. Read the handoff set**, in
> `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-archive/`:
>
> - `AGENT-CONTEXT-vb6ToDotNet.md` — repo topology, GitLab access, deploy
>   sequence, standing rules. Read in full.
> - `vb6ToDotNet-Workstreams-20260823.html` — the four parallel workstreams and
>   the approval gate between the bench lane and the live apps.
> - `HomeFront-Handoff-20260823.html` — deploy detail, branches, open items.
>
> **3. Tell me, before touching anything:**
>
> - which memory store you loaded, and how many files it has;
> - which of the four workstreams my request belongs to, and whether it is in
>   the **bench lane** (FlexKitTester → FlexCore) or the **live lane**
>   (FlexKit → the apps);
> - whether the work is blocked behind the approval gate.
>
> **4. Rules to apply from the first edit, without being reminded:**
>
> - Never push from a source repo — their `origin` remotes are wrong and
>   HomeFrontPB's does not exist. Publishing happens only through
>   `tools/deploy_to_repos.sh`.
> - Never push local `appsettings*.json`. The remote owns environment config;
>   pushing local values breaks QA, UAT and UAT-Hyphen at once.
> - Pull `origin/main` from **both** `hyphen-pb` and `homefront` before any
>   deploy, or teammates' merged MRs get reverted by the rsync. Compare
>   `origin/main` against the sha we last pushed, **not** the clone's `HEAD` —
>   the `hyphen-pb` clone serves two branches and may be parked on `R1-UAT`.
> - Sync is **one-way: HomeFrontPB → HomeFront only** (owner directive
>   2026-08-28). HomeFrontPB is going away but is still being tested and now
>   feeds `hyphen-pb` `R1-UAT` directly, so never write into its tree — no
>   HF → PB flow of any kind, not even a fix to a form PB already has.
>   `FMain.razor` follows its own divergence rule, and it is far from the only
>   divergence: ~23 of the 127 shared page files differ, several permanently
>   (`FFeedback`, `Workflow`, `WorkflowEstimating`, `FInboxCustomQuote`,
>   `FSendingWizard`). Do not "reconcile" those. When comparing the two trees,
>   strip the shared `@namespace HomeFront.Components.Pages` line and
>   trailing-newline noise first, or the count inflates to 65–73 false
>   positives.
> - FlexKit and FlexCore are allowed to diverge and are **not** auto-mirrored.
>   Sync only when explicitly told.
> - Anything built in FlexKitTester → FlexCore stays there until the owner has
>   personally tested it and given an explicit go-ahead. Two approvals, never
>   one.
> - VB6 fidelity is the prime directive — read the matching `.frm`/`.bas` under
>   `HomeFrontVB6/` before changing a migrated `F*` page.
> - Grid columns come from AppGridLayout, never hardcoded. All UI primitives
>   come from FlexKit. Minimise JavaScript. Never interpolate user values into
>   SQL.
> - Run and runtime-test the apps yourself — start them through the Browser pane
>   preview / `launch.json` entries, never a raw Bash `dotnet run` — then ALWAYS
>   shut down every server you started. I test on the same ports (HomeFront on
>   5065, benches on 5299/5266) and a leftover server blocks me. Compile-only
>   verification is not "done" for a behaviour change when a runtime test is
>   feasible.
> - Do not change FAssembly `gItems`. Do not touch HomeFrontPOC. Crystal report
>   XMLs are read-only.
>
> Do not start work until you have reported step 3.

---

## PART B — what the operator needs to know

### If the new session runs under the SAME macOS account (`wadood`)

Memory loads automatically. Launch Claude from
`/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource` — the same directory
this work was done from — and the 126-file store (count as of 2026-08-29; it grows steadily, so treat it as a floor) resolves.

Have them confirm the file count before trusting it. A second, **stale** store
exists under the `…-VBToCSharp-HomeFront-MobileSource` slug with 6 files last
touched 7 May 2026, missing every rule added since.

### If the new session runs under a DIFFERENT macOS account

They get **none** of it. The memory store lives under `/Users/wadood/.claude/`,
and so do the credentials. Specifically they will be missing:

| What | Where it lives | Consequence |
|---|---|---|
| Memory store (~126 files) | `~/.claude/projects/…-HomeFront/memory/` | No standing rules, no incident history |
| Global instructions | `~/.claude/CLAUDE.md` | No project map |
| GitLab SSH key | `~/.ssh/id_ed25519` | Cannot fetch or push |
| NuGet key | Keychain `flexcore-nuget-key` | Cannot publish FlexCore |
| Jira token | `~/.jira_token` | Cannot sync the feedback board |
| SQL `sa` password | user-secrets | Cannot run the apps against the DB |

In that case `AGENT-CONTEXT-vb6ToDotNet.md` is the whole handover, and the
credentials have to be provisioned separately — do not copy them between
accounts casually, particularly the Jira token and the Cognito secret, both of
which are already overdue for rotation.

### Quickest way to verify the new session is correctly set up

```
ssh -T git@gitlab.innovatixinc.com     # → Welcome to GitLab, @wchaudhary!
gh auth status                          # → wadoodachaudhary
ls ~/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/*.md | wc -l   # → 120+ and growing
```

If the first two fail, that session cannot deploy. If the third returns 6, it is
on the stale memory store.

---

## PART C — one session per workstream

**Sessions are not created from folders.** Pointing Claude at a directory starts
*one* conversation. The four workstreams in the sidebar exist because someone
started four separate sessions. To reproduce them, launch a session per
workstream and open it with the matching prompt below.

### The working folder decides whether you get any memory at all

Verified 2026-08-23:

| Folder you launch from | What loads |
|---|---|
| `/Users/wadood/projects/VBToCSharp` | **Nothing.** No project store exists for it — blank project, zero memory, zero rules. |
| `/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource` | The live memory store — 126 files as of 2026-08-29, and growing. **Use this one.** |

**Launch every session from the same directory:**

```bash
cd /Users/wadood/projects/VBToCSharp/HomeFront/MobileSource
claude
```

Setting the folder to `VBToCSharp` (the parent) is the easy mistake — it looks
like the natural project root, and it silently yields a session with no memory
and no standing rules. Nothing warns you.

Session transcripts are keyed to that directory; memory resolves to the parent
project (`…-VBToCSharp-HomeFront`). Launching from anywhere else risks the stale
6-file memory store.

Paths used below, in full:

- Docs — `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-archive/`
- Memory — `/Users/wadood/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/`

Every prompt assumes **Part A has already been run in that session.**

---

### C1 · Update repositories for HomeFront/FlexKit

> This session owns the deploy pipeline. Read, in order:
>
> - `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-archive/AGENT-CONTEXT-vb6ToDotNet.md` — sections 3 to 6 in full
> - `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-archive/HomeFront-Handoff-20260823.html`
>
> Then these memory files, from
> `/Users/wadood/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/`:
> `vb6todotnet_workstream_map.md`, `operating_rules.md`, `deploy_pipeline.md`,
> `deploy_pull_main_first.md`, `deploy_never_push_env_config.md`,
> `deploy_update_r1uat.md`, `deploy_staging_clone_traps.md`,
> `feedback_bump_flexkit_before_deploy.md`, `flexcore_nuget_publish.md`
>
> Treat `tools/deploy_to_repos.sh` as the final authority wherever a memory file
> and the script disagree.
>
> Confirm back to me: the eight steps of the deploy sequence in order, which two
> steps are conditional and on what, what `ENV_CONFIG_EXCLUDE` protects, which
> source tree feeds `main` and which feeds `R1-UAT`, and why you must never merge
> main into R1-UAT. Then wait — do not deploy until I ask.

---

### C2 · HomeFrontPB / HomeFront App

> This session owns the migrated `F*` forms in both apps. Read:
>
> - `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-archive/vb6ToDotNet-Workstreams-20260823.html` — the "HomeFrontPB / HomeFront App" section
> - `/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/CLAUDE.md` — in full, especially the house style for migrated forms
>
> Then these memory files, from
> `/Users/wadood/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/`:
> `vb6todotnet_workstream_map.md`, `operating_rules.md`,
> `migrated_form_conventions.md`, `vb6_fidelity.md`, `flexkit_ui_mandatory.md`,
> `grid_layout_presenter.md`, `near_final_forms_dont_edit.md`
>
> VB6 originals are at
> `/Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/` (`HFEst/`,
> `HFSystem/Source/`). Read the matching `.frm` or `.bas` before changing any
> migrated page.
>
> Confirm back to me: which directory each app keeps its migrated forms in, why
> the two trees are no longer expected to match line-for-line, which files are
> permanently divergent and must never be reconciled, which direction a fix is
> allowed to travel between the pair, and which grid you must never touch.

---

### C3 · FlexKit Tester

> This session owns the bench lane. Read:
>
> - `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-archive/vb6ToDotNet-Workstreams-20260823.html` — the "FlexKit Tester" section
>
> Then these memory files, from
> `/Users/wadood/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/`:
> `vb6todotnet_workstream_map.md`, `operating_rules.md`,
> `flexkittester_only_sandbox.md`, `flexkit_port_needs_user_go_ahead.md`,
> `dev_server_stale_assembly_trap.md`, `row_selection_perf_bench.md`,
> `pm_entry_bench.md`
>
> The project is at `/Users/wadood/projects/VBToCSharp/HomeFront/FlexKitTester`
> and serves on port 5299. It references **FlexCore**, not FlexKit.
>
> Confirm back to me: what every new tester page must declare on its first line
> and what happens silently if it is missing; why a FlexKit-only bug cannot be
> reproduced here directly; and what to check first when I report that a fix made
> "no difference".

---

### C4 · Application-wide zoom facility

> This session owns the zoom feature, which is **gated**. Read:
>
> - `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-archive/vb6ToDotNet-Workstreams-20260823.html` — the "Application-wide zoom" section
>
> Then these memory files, from
> `/Users/wadood/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/`:
> `vb6todotnet_workstream_map.md`, `operating_rules.md`,
> `app_wide_text_zoom.md` — which carries the complete port recipe —
> `flexkit_port_needs_user_go_ahead.md`, `flexkittester_only_sandbox.md`,
> `dev_server_stale_assembly_trap.md`
>
> **Status: SHIPPED.** Built in the bench lane, gate passed, ported to FlexKit
> 2026-08-22 and wired into both apps 2026-08-23 (HHM-871). Live on main and
> R1-UAT. `app_wide_text_zoom.md` is now a record of what was done, not a plan —
> read it for the architecture and the layout traps, not as pending work.
>
> Confirm back to me: the owner's spec in one sentence, why `ZoomService` is
> registered `AddScoped` rather than singleton, why the zoom scope sets an em
> base as well as the CSS variable, and what has to happen before any of it may
> touch FlexKit again.
>
> The gate still governs any FURTHER zoom work: changes go bench-first, and I test
> them before they reach FlexKit or the apps.

---

### Why each prompt ends with a question

Reading is not evidence of understanding. Each prompt closes by asking the
session to state back the specific facts that are expensive to get wrong — the
conditional deploy steps, the exempt file, the silent SSR failure, the gate. If
the answer is vague, the session has not read enough to be trusted with an edit
yet.
