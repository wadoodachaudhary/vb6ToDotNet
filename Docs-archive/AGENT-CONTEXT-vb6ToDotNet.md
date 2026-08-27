# Agent context — vb6ToDotNet

Point a new agent at this file. It is a portable snapshot of the machine-local
context that normally loads automatically. Prepared 2026-08-23.

---

## 1. Where the real context lives

| What | Path |
|---|---|
| **Live memory store** (101 files) | `~/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/` |
| Memory index (loaded each session) | `MEMORY.md` in that directory |
| Global instructions | `~/.claude/CLAUDE.md` |
| Living rule file | `~/projects/VBToCSharp/AGENTS.md` |
| HomeFront repo rules | `~/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/CLAUDE.md` |

> **Verify the memory store at session start.** A second, stale store exists at
> `…-VBToCSharp-HomeFront-MobileSource/memory/` with only 6 files, last touched
> 7 May 2026. The live one has ~101 files and a ~101-line `MEMORY.md`. If you
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
| HomeFrontPB/HomeFront App | live | Migrated `F*` forms in both apps |
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
| `hyphen-pb` **(primary)** | HomeFrontPB | gitlab · application-modernization/hyphen-pb.git |
| `homefront` | MobileSource/HomeFront | gitlab · application-modernization/homefront.git |
| `flexkit` | FlexKit | gitlab · application-modernization/flexkit.git |
| `flexcore` | FlexCore | github.com/wadoodachaudhary/FlexCore.git |

`hyphen-pb` is where team MRs land and where QA/UAT deploy from. It is the only
repo with an `R1-UAT` branch.

**FlexKit is consumed two ways.** Working tree = live `ProjectReference`. The
deploy script temporarily swaps to a `PackageReference` against a packed nupkg
in `local-packages/`, builds to verify, pushes that form — and never modifies
your working-tree csproj. Consequence: if an MR changes `HomeFrontPB.csproj`,
**hand-merge it**; copying the file wholesale replaces your ProjectReference.

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

0. **Pull `hyphen-pb` `origin/main` FIRST.** The deploy is an rsync — anything
   upstream not also local gets reverted. Mirror page changes to the other app;
   hand-merge the csproj.
1. Verify sync-pair parity, then commit all four repos.
2. Bump FlexKit + repack + refresh `local-packages/` nupkg — **only if FlexKit
   source changed.**
3. Build `HomeFront.sln` and `HomeFrontPB.sln` (+ `FlexCore.Showcase.sln` if
   FlexCore changed). All 0 errors. Always build the `.sln`, never the csproj.
4. `bash tools/deploy_to_repos.sh`
5. Verify on `origin/main`: Azure token placeholders present,
   `appsettings.Development.json` still 97 lines, csproj publish-exclusion intact.
6. Merge main into **R1-UAT**, confirm the five disabled security toggles and env
   config, build-verify, push.
7. Publish FlexCore to NuGet — **only if FlexCore changed.** Check the latest
   published version first and bump above it.

---

## 6. Rules that will bite you

**NEVER push local `appsettings*.json`.** `origin/main` carries Azure token
placeholders (`Server=#{Database.Server}#`) and a 97-line
`appsettings.Development.json` holding the auth cookie name, Serilog sinks, an
InternalTools allow-list and Cognito settings. This machine has `localhost,1433`
/ `sa` and a stripped stub. Pushing local over it breaks QA, UAT and UAT-Hyphen
at once. Enforced by `ENV_CONFIG_EXCLUDE` in the deploy script — do not remove it.

**Sync pairs.** HomeFront ↔ HomeFrontPB: same migrated form in both
(`Components/Pages/Migrated/F*.razor` vs flat `Components/Pages/F*.razor`); a fix
in one belongs in the other. Only expected difference is HomeFront's extra
`@namespace` line. `FMain.razor` carries ~1,700 lines of drift — hand-apply there.

**Hands off.** FAssembly `gItems` (finished; a past change caused a major
regression). HomeFrontPOC (archived). Crystal report XMLs (read-only).

**Engineering.** VB6 fidelity is the prime directive — read the matching
`.frm`/`.bas` under `HomeFrontVB6/` first. Grid columns come from AppGridLayout,
never hardcoded. All UI primitives from FlexKit — a raw `<input>` or `<div>`
modal is a red flag. Minimise JavaScript (DOM-only need + tiny lazy import +
graceful fallback). Never interpolate user values into SQL.

**The owner runs the apps.** Do not auto-start `dotnet run`.

---

## 7. Bench lane specifics

- **All repros and benches go in FlexKitTester.** A repro placed in
  GhostWriterLLM was reverted immediately.
- Every new FlexKitTester page **must** start with `@rendermode InteractiveServer`
  — without it the page renders static SSR: rows appear, every click is dead, and
  there is no error anywhere.
- FlexKitTester references **FlexCore**, not FlexKit. A FlexKit-only behaviour
  difference must be reproduced by applying the change under test to FlexCore too,
  or tested through the real apps (which the owner runs).
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
  `FLogin.razor` plus `Sec.EncryptDbConnection = false`. Keep both sides on
  conflict; set the toggle *before* any early return.
- **Docs live outside the repos**, in `HomeFront/Docs-archive/`. Both repos
  gitignore `Docs/`. Put new documentation in the archive, not a repo.
