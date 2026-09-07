# Agent context — vb6ToDotNet

Point a new agent at this file. It is a portable snapshot of the machine-local
context that normally loads automatically. Prepared 2026-08-23.

---

## 1. Where the real context lives

| What | Path |
|---|---|
| **Live memory store** (126 files) | `~/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/` |
| Memory index (loaded each session) | `MEMORY.md` in that directory |
| Global instructions | `~/.claude/CLAUDE.md` |
| Living rule file | `~/projects/VBToCSharp/AGENTS.md` |
| HomeFront repo rules | `~/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/CLAUDE.md` |

> **Verify the memory store at session start.** A second, stale store exists at
> `…-VBToCSharp-HomeFront-MobileSource/memory/` with only 6 files, last touched
> 7 May 2026. The live one has 120+ files and a `MEMORY.md` with one line per file. If you
> see 6 files you are on the wrong store and every rule below is missing.

Start with `project_overview.md`, then `operating_rules.md`, then
`vb6todotnet_workstream_map.md`.

---

## 2. The four workstreams and the gate between them

Two lanes, separated by a hard approval gate.

- **Bench lane** — `FlexKitTester` → `FlexCore`. Nothing here reaches a user.
- **Live lane** — `FlexKit` → HomeFront / HomeFrontPB / GhostWriter /
  GhostWriterLLM. All consume FlexKit by **ProjectReference**, so an edit lands
  in every host app on the next build.

| Session | Lane | Owns |
|---|---|---|
| Update repositories for HomeFront/FlexKit | live | The deploy pipeline |
| HomeFrontPB/HomeFront App | live | Migrated `F*` forms — new work lands in **HomeFront** only; PB is frozen for UAT testing (sync is one-way, PB→HF) |
| FlexKit Tester | bench | Benches, repros, perf work |
| Application-wide zoom facility | bench | Text zoom — FlexCore/tester only |

**The gate (owner rule, 2026-08-19):** anything built in FlexKitTester →
FlexCore stays there until the owner has personally tested it on the bench and
given an explicit go-ahead. That is **two separate approvals**: build and verify,
STOP and report, wait — then port. If a port happened prematurely, reverse only
your own edits; never `git checkout` files that may carry unrelated work.

FlexKit and FlexCore are **allowed to diverge** and are NOT auto-mirrored
(owner directive 2026-07-25). A feature present in one and absent from the other
is often correct. Check which lane built it before assuming it is missing.

---

## 3. Repository topology — read before pushing

There are **two sets of clones**. You never push from the ones you edit.

**Source repos (edit here, commit locally, do NOT push):**

| Project | Path |
|---|---|
| HomeFront | `VBToCSharp/HomeFront/MobileSource/HomeFront` |
| HomeFrontPB | `VBToCSharp/HomeFront/HomeFrontPB` |
| FlexKit | `VBToCSharp/FlexKit` |
| FlexCore | `VBToCSharp/FlexCore` |
| FlexKitTester | `VBToCSharp/HomeFront/FlexKitTester` |

Their `origin` remotes are stale or nonexistent — HomeFrontPB's points at
`application-modernization/HomeFrontPB.git`, which does not exist.

**Deploy staging clones (what actually ships)** — under
`VBToCSharp/HomeFront/Deploy/repos/`:

| Clone | Fed from | Pushes to |
|---|---|---|
| `hyphen-pb` **(primary)** | `main` ← MobileSource/HomeFront; `R1-UAT` ← HomeFrontPB | gitlab · application-modernization/hyphen-pb.git |
| `homefront` | MobileSource/HomeFront | gitlab · application-modernization/homefront.git |
| `flexkit` | FlexKit | gitlab · application-modernization/flexkit.git |
| `flexcore` | FlexCore | github.com/wadoodachaudhary/FlexCore.git |

`hyphen-pb` is where team MRs land and where QA/UAT deploy from. It is the only
repo with an `R1-UAT` branch.

**FlexKit is consumed two ways.** Working tree = live `ProjectReference`. The
deploy script temporarily swaps to a `PackageReference` against a packed nupkg
in `local-packages/`, builds to verify, pushes that form — and never modifies
your working-tree csproj. Consequence: if an MR changes the app csproj —
`HomeFront.csproj` on `main`, `HomeFrontPB.csproj` on `R1-UAT` (pipelines and
csproj are branch-owned) — **hand-merge it**; copying the file wholesale
replaces your ProjectReference.

---

## 4. Connecting to GitLab / GitHub

Already configured and verified on this machine.

- **GitLab** `gitlab.innovatixinc.com` over SSH, default key `~/.ssh/id_ed25519`,
  account `@wchaudhary`. No `~/.ssh/config` entry.
  Check: `ssh -T git@gitlab.innovatixinc.com` → *Welcome to GitLab, @wchaudhary!*
- **GitHub** (FlexCore mirror only) via `gh`, account `wadoodachaudhary`, HTTPS.
  Check: `gh auth status`
- **Git identity** `wadood` / `wc2595@columbia.edu`

**Secrets — locations only, never echo them:**

| Secret | Where |
|---|---|
| NuGet API key (FlexCore publish) | macOS Keychain, service `flexcore-nuget-key` |
| Jira API token | `~/.jira_token` — reuse, never re-issue |
| SQL `sa` password | user-secrets (dev) / `Database__Password` env (prod) |

---

## 5. The deploy sequence

Run in order, every time. `tools/deploy_to_repos.sh` (`--dry-run`, `--pull`).

0. **Pull BOTH remotes FIRST — `hyphen-pb` and `homefront`** (owner directive
   2026-08-29). The deploy rsyncs the **working tree**, so anything upstream not
   also local gets reverted. Compare `origin/main` against the SHA this machine
   last pushed — **never** against the clone's `HEAD`: the `hyphen-pb` clone
   serves two branches (`main` ← HomeFront, `R1-UAT` ← HomeFrontPB) and may be
   parked on `R1-UAT`. Mirror incoming page changes **PB → HF only** (owner
   directive 2026-08-28); there is no HF → PB flow of any kind. Hand-merge the
   csproj.
0b. **ALSO check for OPEN merge requests** (added 2026-09-07). `origin/main`
   matching the last-pushed sha does **NOT** mean there is nothing to take. An
   open MR's commits live on its source branch, never on `main`, so the step-0
   sha check reports "no upstream work" while real fixes sit unmerged. Missed
   exactly this way on 2026-09-07 (hyphen-pb **!27**, open 18 hours, deployed
   straight past — the owner had to point it out):
   ```bash
   cd Deploy/repos/hyphen-pb && git fetch origin --prune
   for b in $(git branch -r --format='%(refname:short)' \
              | grep -vE 'HEAD|origin/(main|R1-UAT|R1-UAT-Hyphen|master|Dev)$'); do
       n=$(git log --oneline origin/main..$b | wc -l | tr -d ' ')
       [ "$n" != 0 ] && echo "$b  $n ahead — possible open MR"
   done
   ```
   Stale ticket branches (HHM-172/331/566/620, all July) sit ahead forever, so a
   hit is a *candidate*, not proof — confirm in the GitLab UI. Git cannot tell
   you who authored an MR: the commit author is this machine's identity
   (`wc2595@columbia.edu`) regardless of who wrote it, and only GitLab's MR
   author field names the real person. Claude-in-Chrome carries the owner's real
   session once they are signed in to GitLab **in Chrome** (a Safari session does
   not carry over, and the sandboxed Browser pane never authenticates).
   **Applying an open MR: diff against its MERGE BASE, never against `main`** —
   the branch is usually based on an older main, so `git diff origin/main..<branch>`
   shows it *deleting* everything since, a pure artifact. Use
   `git merge-base origin/main <branch>`, then 3-way merge each touched file
   (`git merge-file OURS BASE THEIRS`). On !27 that turned an alarming
   48-file/-6928-line diff into the real change: 4 files, +196/-128, zero
   conflicts — even though two of those files had been rewritten that same day.
   Apply multiple MRs in **base order** (oldest base first). **Never merge or
   close the MR from here** — that is the owner's or author's call on GitLab; say
   explicitly that the content reaches `main` via the source→main push while the
   MR stays open. (Verified 2026-09-07 on !28/!31: applying locally and letting
   the owner merge afterwards is safe — the two paths produced byte-identical
   content and the upstream merge was a clean no-op.)
1. Commit what is dirty in each source tree — HomeFront, HomeFrontPB, FlexKit,
   FlexCore. This is housekeeping, **not a gate**: the deploy rsyncs the WORKING
   TREE, not HEAD, so uncommitted edits ship either way and last-minute changes
   are expected (owner directive 2026-08-28). Never block a deploy on a clean tree.
   **Do not chase HomeFront ↔ HomeFrontPB parity.** Sync is one-way **PB→HF only**
   (owner directive 2026-08-28) and ~19-24 of ~127 shared page files differ *by
   design* — FFeedback(.css), Workflow / WorkflowEstimating(.css),
   FInboxCustomQuote, FSendingWizard, FMain. Any comparison must strip the shared
   `@namespace HomeFront.Components.Pages` line (**both** apps carry it — 51 of 67
   shared `.razor` files) and trailing-newline noise, or the diff count inflates
   with false positives.
2. Bump FlexKit + repack + refresh `local-packages/` nupkg — **only if FlexKit
   source changed.**
3. Build `HomeFront.sln` (+ `FlexCore.Showcase.sln` if FlexCore changed). All 0
   errors. Always build the `.sln`, never the csproj. **Do NOT build
   `HomeFrontPB.sln`** — `CLAUDE.md` (2026-08-31) says HomeFrontPB must not be
   modified, built, or synced without an explicit request naming it. The R1-UAT
   stage inside `deploy_to_repos.sh` builds its own staged copy; that is the
   pipeline's job, not an ad-hoc build on your part.
4. `bash tools/deploy_to_repos.sh`
5. Verify on `origin/main`: Azure token placeholders present,
   `appsettings.Development.json` still 97 lines, csproj publish-exclusion intact.
6. **R1-UAT is built directly from HomeFrontPB — there is no merge from main.**
   `deploy_r1uat` runs automatically at the end of step 4: it resets the
   `hyphen-pb` clone to `origin/R1-UAT`, rsyncs the HomeFrontPB tree with
   `Components/Pages/FLogin.razor` excluded (and re-checked-out after the
   package swap) so the five disabled security toggles survive, build-verifies,
   and refuses to push if the toggle count is not exactly 5. Nothing to do by
   hand — just confirm it reported `Pushed R1-UAT`.
7. Publish FlexCore to NuGet — **only if FlexCore changed.** Check the latest
   published version first and bump above it.

**Current shipped versions (2026-09-07):** FlexKit **0.1.86**, FlexCore **0.2.32**
(published to nuget.org). Both are bumped only when their own source changes; the
two version lines are independent and are allowed to diverge.

### Five traps in the deploy script — all fixed, all silent when they bite

Worth knowing because each one reported success while doing the wrong thing.

1. **R1-UAT silently stopped pushing** (found 2026-08-28, after two deploys had
   quietly skipped it). `wwwroot/resources/Estimating/Reports/xml_orig/*.xml` are
   CRLF, so `git status --porcelain` reports them modified on every run — but
   `git add -A` normalizes them to LF and the staged tree comes out **identical to
   HEAD**. `git commit` then exits non-zero with "nothing to commit", and
   `set -euo pipefail` kills the script *before* `push origin R1-UAT`. The fix is
   to stage first and test `git diff --cached --quiet`, which is what the main loop
   always did. **Tell:** a run that prints `Lockdown intact` but never
   `Pushed R1-UAT` / `no changes to push`, and exits non-zero.
2. **The `hyphen-pb` clone serves two branches.** Left parked on `R1-UAT`, the next
   run's `reset --hard origin/main` rewrites the UAT branch — that is exactly how
   the 2026-08-28 non-fast-forward incident happened. A function-level `RETURN`
   trap is not enough (it does not fire when `set -e` aborts mid-function), so
   there is now a script-level `trap … EXIT` that parks the clone on `main` on
   every exit path and shouts if it cannot.
3. **The deploy ships uncommitted working-tree edits — by design.** It rsyncs the
   working tree, not `HEAD`. The owner and other sessions edit these trees directly
   and a late edit is *expected* to ship (owner directive 2026-08-28). Never gate a
   deploy on a clean tree. The real safety net is the staged build verification,
   which runs before each push. The one judgement call: if a file's mtime is still
   moving, wait for it to settle before shipping — `wwwroot` JS gets no compile
   check, so a half-written file would ship silently.
4. **A stale FlexKit nupkg fails ONLY the staged build.** The apps consume FlexKit
   two ways: `ProjectReference` in the working tree, `PackageReference` against the
   packed nupkg in the staging clone. So a nupkg that predates a FlexKit source
   change produces a missing-member error that **cannot reproduce locally** — every
   local build stays green while the deploy dies. `pack_flexkit` used to repack only
   on a version change, so FlexKit advancing without a bump silently reused the old
   package. It now also repacks when any tracked FlexKit source file is newer than
   the nupkg, and says why it packed. **Still bump the version** whenever FlexKit
   source changed — that is what consumers pin. Seen 2026-08-29 on
   `GridControl.RequestScrollToOrigin` (two commits landed after the 0.1.74 bump
   without bumping again) and 2026-08-10 on `MaxLength`.
5. **Repacking the SAME version shadows itself via the GLOBAL NuGet cache** —
   a worse variant of #4 that *survives* the auto-repack. `dotnet restore`
   resolves a `PackageReference` against `~/.nuget/packages/flexkit/<version>/`
   first and treats an already-extracted version as immutable, so once a version
   has been resolved anywhere on this machine, restoring it again never re-reads
   the local feed — no matter how recently `pack_flexkit` rewrote the `.nupkg`.
   Hit 2026-09-06: FlexKit source changed without a version bump, `pack_flexkit`
   correctly detected source-newer-than-nupkg and repacked 0.1.85, and the staged
   build *still* failed on `GridControl.GetFilteredRecords` — a method that was
   in the fresh nupkg but not in what restore actually served. **There is no
   clean recovery except bumping the version**; never delete or edit
   `~/.nuget/packages/flexkit/<version>/`, it is machine-wide and other consumers
   (GhostWriter, GhostWriterLLM, FlexCore.Showcase) may have resolved it too.
   **Tell:** the staged error names a member that demonstrably *does* exist in
   current FlexKit source — grep to confirm before suspecting a bad merge.
   This is why "bump whenever FlexKit source changed" is not optional: the
   auto-repack is a safety net for a missed bump, not a substitute for one.
   **Fallout:** a failed staged build leaves the `hyphen-pb` clone dirty, so the
   R1-UAT stage then fails its checkout with "local changes would be overwritten".
   That is a consequence, not a second bug — `git reset --hard origin/main &&
   git clean -fd` in the clone, then re-run.

---

## 6. Rules that will bite you

**NEVER push local `appsettings*.json`.** `origin/main` carries Azure token
placeholders (`Server=#{Database.Server}#`) and a 97-line
`appsettings.Development.json` holding the auth cookie name, Serilog sinks, an
InternalTools allow-list and Cognito settings. This machine has `localhost,1433`
/ `sa` and a stripped stub. Pushing local over it breaks QA, UAT and UAT-Hyphen
at once. Enforced by `ENV_CONFIG_EXCLUDE` in the deploy script — do not remove it.

**Sync pairs.** HomeFront ↔ HomeFrontPB: the same migrated form lives in both
(`Components/Pages/Migrated/F*.razor` vs flat `Components/Pages/F*.razor`), but
sync is **ONE-WAY — PB → HF only** (owner directive 2026-08-28). There is **no
HF → PB flow of any kind**, not even a fix to a form PB already has: HomeFrontPB
is going away yet is still under test, so its tree must not be disturbed.
`FMain.razor` follows its own divergence rule (~1,600 lines of drift — hand-apply
there). Some divergences are permanent and must NOT be reconciled: `FFeedback`
(HF has label-scoped Jira sync via `JiraSettings.SyncLabel`; PB has a sort-by
dropdown + priority badges), `Workflow` / `WorkflowEstimating` (PB marks nodes
`enabled:false` and connectors `isBroken:true` for forms it lacks), and
`FInboxCustomQuote` / `FSendingWizard` (HF-only — they cannot compile in PB).
Baseline 2026-08-28: 19 of 127 shared page files differ. Comparisons MUST strip
the shared `@namespace HomeFront.Components.Pages` line (BOTH apps carry it) and
trailing-newline noise, or the count inflates to 65-73 false positives.

**Hands off.** FAssembly `gItems` (finished; a past change caused a major
regression). HomeFrontPOC (archived). Crystal report XMLs (read-only).

**Engineering.** VB6 fidelity is the prime directive — read the matching
`.frm`/`.bas` under `HomeFrontVB6/` first. Grid columns come from AppGridLayout,
never hardcoded. All UI primitives from FlexKit — a raw `<input>` or `<div>`
modal is a red flag. Minimise JavaScript (DOM-only need + tiny lazy import +
graceful fallback). Never interpolate user values into SQL.

**Run and test the apps — then shut them down.** The owner *wants* runtime
testing: start the app and exercise the change (especially DB-writing flows)
before reporting done; compile-only verification is not "done" for a behaviour
change when a runtime test is feasible. Start it through the Browser pane's
preview / `launch.json` entries, never a raw Bash `dotnet run`. **Always kill
every server you started when testing finishes** — the owner tests on the same
ports and a stray process blocks them. The owner's own HomeFront instance runs
on **:5065**; for an environment-specific bug, attach to *their* server
(`preview_start` with `{url: 'http://localhost:5065/'}`) rather than your own
preview port.

---

## 7. Bench lane specifics

- **All repros and benches go in FlexKitTester.** A repro placed in
  GhostWriterLLM was reverted immediately.
- Every new FlexKitTester page **must** start with `@rendermode InteractiveServer`
  — without it the page renders static SSR: rows appear, every click is dead, and
  there is no error anywhere.
- FlexKitTester references **FlexCore**, not FlexKit. A FlexKit-only behaviour
  difference must be reproduced by applying the change under test to FlexCore too,
  or tested through the real apps — run and test them yourself via the
  `.claude/launch.json` entries (never a raw Bash `dotnet run`), then ALWAYS shut
  down every server you started, since the owner tests on the same ports.
- Existing pages include `/row-selection`, `/virtualization`, `/pm-entry`,
  `/edit-items`, `/userperms-lists`.

**Dev servers — restart after ANY library rebuild.** A running server holds the
assembly it loaded at process start, and the launch configs use `--no-build`, so
nothing reloads. Symptom: the owner reports "no difference" for hours. Verify the
*served* build marker, not the built one.

| Port | Serves |
|---|---|
| 5299 | FlexKitTester (FlexCore) |
| 5266 | HomeFrontPB |
| 5065 | HomeFront |
| 5298 | optional latency proxy in front of 5299 |

---

## 8. Open items

- **Two unrotated secrets in git history** on hyphen-pb: the Jira `ApiToken` in
  `App_Data/jira-settings.json` and the Cognito `AppClientSecret` in
  `appsettings.Development.json`. Deleting the files will not scrub history.
- **App-wide text zoom** — gate passed and **shipped**: FlexKit 2026-08-22, both
  apps 2026-08-23 (HHM-871). Live on main and R1-UAT. `app_wide_text_zoom.md`
  now records what was done rather than a pending plan.
- **HHM-829 PO Indexes column sizing** unresolved. `FPOFormats.razor` hardcodes
  every width; VB6 `FPOFormats.frm` uses `ColWidth(0)=240` + `AutoSize(1,.Cols-1)`
  and never reads AppGridLayout. `GridColumn.AllowAutoFit` exists as groundwork.
- **R1-UAT carries branch-only changes** — five hard-disabled security toggles in
  `FLogin.razor` plus `Sec.EncryptDbConnection = false` (same file). Nothing is
  merged: `deploy_r1uat` rsyncs HomeFrontPB straight onto the branch with
  `R1UAT_PRESERVE_EXCLUDE` (`--exclude='Components/Pages/FLogin.razor'`), then
  re-runs `git checkout -- Components/Pages/FLogin.razor` after the package swap,
  and refuses to push unless all five toggles are still present. Set the toggle
  *before* any early return.
- **Docs live outside the repos**, in `HomeFront/Docs-archive/`. Both repos
  gitignore `Docs/`. Put new documentation in the archive, not a repo.
