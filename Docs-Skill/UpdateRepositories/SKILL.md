---
name: update-repositories
description: Deploy HomeFront, FlexKit and FlexCore to GitLab/GitHub and publish FlexCore to NuGet — the "Update Repositories" routine for this Mac. Covers the standing owner rules, taking and acking teammates' merge requests, the FlexKit↔FlexCore mirror, version burning and NuGet publishing (including held publishes), the verification matrix, and the two-account hand-off. Use when the user asks to update, deploy, push or sync the repositories, check whether they are in sync, pull teammates' merge requests, publish FlexCore, or hand the work over between the Wadood and Innovatix Claude accounts.
---

# Update Repositories

Ships this machine's working trees to four remote repositories — after merging teammates' work and
proving nothing gets reverted — then proves what went out. `tools/deploy_to_repos.sh` does the pushing;
the checks before and the proof after are this skill.

Canonical, git-tracked source: `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories/`.
Installed READ-ONLY for every Claude account on this Mac at `~/.claude/skills/update-repositories/` (both
accounts share `~/.claude`). Never edit the installed copy — edit the canonical one, commit it in the
HomeFront repo, then `bash /Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories/install.sh`.

| Source working tree | Remote | Notes |
|---|---|---|
| `HomeFront/MobileSource/HomeFront` (nested git repo) | GitLab `hyphen-pb` **main** + `homefront` main | main is the QA-deployed app; `homefront` is a source mirror |
| (the `main` this run pushed) | GitLab `hyphen-pb` **R2-UAT** | the **client release branch** (owner 2026-10-01). `--with-r2uat` MERGES `origin/main` into it; its login lockdown and UAT pipeline file are branch-owned and preserved |
| `HomeFront/HomeFrontPB` | — | **FROZEN tree, feeds nothing.** `R1-UAT` was renamed to R2; `--with-r1uat` is retired — it would push this tree over R2 code |
| `FlexKit` on branch **`telerik-parity-20260904`** | GitLab `flexkit` main | never `git checkout main` in FlexKit; not on NuGet |
| `FlexCore` (main) | GitHub `wadoodachaudhary/FlexCore` + **nuget.org** | GitHub main is deploy-snapshot history; `FlexCore.Llm/` goes to GitHub only |

Staging clones: `HomeFront/Deploy/repos/{hyphen-pb,homefront,flexkit,flexcore}`. FlexKit feed:
`HomeFront/Deploy/local-packages`. The outer git repo is `HomeFront/` itself (owns `tools/` and this
skill); `VBToCSharp/` is not a repo. `HomeFront/Deploy/` is git-ignored and machine-local; the skill keeps
three files there: `.update-repositories-gate.state`, `pull-conflicts.txt`, `preflight-ack.txt`.

## Only when asked

Deploy and publish ONLY when the owner explicitly asks to update/deploy the repositories in this request.
"Check whether the repositories are in sync — do not update" means: run `preflight.sh` (step 1) and
report; change nothing. A message asking for details or fixes does not authorise a push or a publish.

## Cold start

New account, or first run of the week? Read [reference/handoff.md](reference/handoff.md) first: both
accounts share `~/.claude`, the memory store is pinned per repo root (and three roots load nothing without
that pin), and it carries the day-one / end-of-week checklists, the Jira conventions and the machine-local
prerequisites git does not carry (keychain key, FlexKit feed, Playwright, the owner's own ports).

## Hard rules

1. **Working trees ship, uncommitted edits included — by design** (owner 2026-08-28). Never gate on a
   clean tree; still commit everything first so history matches what shipped.
2. **HomeFrontPB is frozen**: never write, commit, revert or build it (`HomeFrontPB.sln` never).
   **R2-UAT is the client release branch** (owner 2026-10-01: "R1 is no more frozen as it has been renamed to R2
   and we are releasing code to the client on R2 branch"): every round passes `--with-r2uat` to the gate AND the
   deploy. **Never run `--with-r1uat`** — it rsyncs the frozen PB tree over the R2 code now on `R1-UAT`.
3. **Teammates' work comes in before any push-mode run.** A teammate change our tree lacks is silently
   REVERTED by the push. Every push-mode run — **even `--dry-run`** — resets the staging clone's `main`
   to `origin/main`, after which `--pull` reports nothing new. The gate judges unshipped teammate
   content directly, so it still blocks after such a reset; `--pull` itself no longer can.
4. **Burned versions never ship.** A FlexKit version already in `~/.nuget/packages/flexkit/` is shadowed
   in the staged build; a FlexCore version on nuget.org can never be replaced. `--dry-run` really packs
   and restores FlexKit, so bump BEFORE the dry run and again if FlexKit changes after it. The script
   repacks FlexKit only on `.cs/.razor/.css/.js/.csproj` changes — an asset-only change also needs a bump
   (the gate says so).
5. **FlexKit → FlexCore always** (owner 2026-09-10): a FlexKit change ships with its FlexCore port;
   FlexCore → FlexKit only when asked. Never merge FlexCore (or FlexKit's working branch) into a main.
6. **A NuGet publish is permanent.** Publish only through `publish_flexcore.sh`, after `verify_deploy.sh`
   passes. Since 09-13 every round has also held the publish (and the push) until any review of the
   shipping diff landed — FlexCore 0.2.41 shipped a defect because it didn't; the owner has not formally
   confirmed that rule.
7. **Type flags exactly.** `deploy_to_repos.sh` silently ignores unknown arguments: `--dryrun` or
   `--with-r1-uat` runs a REAL push, and `--pull --dryrun` a real pull. The gate rejects unknown arguments;
   nothing guards `--pull`.
8. Never force-push. Never push `appsettings*.json`. Never print or commit a secret. Never merge or close
   someone else's merge request unless asked. Never commit scratch files. Never leave a conflict marker in
   a source tree — it ships as-is.
9. The dev login-skip (`Services/DevAutoLogin.cs`) stays **working locally** and **absent from every
   pushed branch** — the script strips it from staging clones only. Never delete it locally.
10. **An MR's state comes only from the GitLab UI.** `refs/merge-requests/N/head` exists for open, closed
    and merged alike. Reading "open" from git shipped closed MR !39 on 09-22; the owner had closed it on
    09-21 and it was reverted on 09-23. Say "not in main", not "open", unless the UI said so — and an MR's
    content only survives in main once it exists in OUR source tree (a teammate's merge we do not hold is
    overwritten by our own rsync). [reference/merge-requests.md](reference/merge-requests.md).
11. **Bench work needs the owner's own go-ahead.** Anything developed FlexKitTester → FlexCore stays there
    until the owner has tested it on the bench and says go; only then may FlexKit change and ship. FlexKit
    feeds every host app on ProjectReference, so an unapproved port lands everywhere on the next build.
12. **Scan for secrets before the push, and never trust a silent scan.** macOS `grep` aborts on a long
    alternation and prints nothing, so `HITS=$(grep …)` reads empty as clean — that happened on 09-14. Scan
    with Python, one pattern per compile, and treat any tool error as FAIL. A secret already tracked on a
    pushed branch (the live Atlassian token in `App_Data/jira-settings.json`) is an owner decision, not
    something to fix mid-round: report it and do not make it worse.
13. **A held publish stays held until the owner lifts it**, and it is not resumable: `publish_flexcore.sh`
    refuses a gate state older than 720 minutes, and a held version keeps absorbing later rounds' content
    under the same number. Record the GitHub sha the hold was judged on.
    [reference/nuget-and-versions.md](reference/nuget-and-versions.md).

Reference pages (edit the canonical copy — the installed one is read-only):

| Page | What it answers |
|---|---|
| [rules-and-traps.md](reference/rules-and-traps.md) | every trap with its tell, and the dated owner rules |
| [history.md](reference/history.md) | every past run, the lessons, and the current "Open as of" list |
| [merge-requests.md](reference/merge-requests.md) | MR state, authorship, taking one, acking, who closes |
| [flexkit-flexcore-sync.md](reference/flexkit-flexcore-sync.md) | the mirror: identity, direction, port-not-copy, parity, build matrix |
| [nuget-and-versions.md](reference/nuget-and-versions.md) | burning, the FlexKit feed, the package, publishing, held publishes |
| [verification.md](reference/verification.md) | the evidence matrix: harnesses, browser suites, report suites, consumer builds |
| [handoff.md](reference/handoff.md) | cold start, memory pinning, day-one / end-of-week, Jira, machine-local setup |

## Procedure

Run from `/Users/wadood/projects/VBToCSharp/HomeFront`. The scripts are bash — run them with `bash`, never
`source` them into zsh.

### 1. Preflight (read-only)

```bash
bash Docs-Skill/UpdateRepositories/scripts/preflight.sh
```

Reports active sessions, working trees, step 0 per clone, MR branches, **every teammate change and
whether it is in our trees**, edits that undo a recent commit, versions, install drift — then an
**ATTENTION** list. Per-file states:

| State | Meaning | What to do |
|---|---|---|
| `REVERT` | unshipped teammate change missing from our tree | step 3 (`--pull`, or hand-merge) |
| `ABSENT` | file on origin/main missing here — the push would delete it | step 3 |
| `RESURRECT` | teammate deleted/renamed it away, we still have it — the push would re-add it | `git rm` it in the source after checking |
| `HANDMERGE` | `Program.cs`/`HomeFront.csproj`/`HomeFront.sln` hunk missing (the deploy rewrites these) | merge the hunk by hand, keep our ProjectReference |
| `LOST` | our earlier deploy already overwrote a teammate change we never had (the MR !33 case) | restore it, or record a deliberate rejection in `Deploy/preflight-ack.txt` |
| `NOTE` | lines gone, but our own history had them and changed them later — superseded | informational |

Also watch for `upstream commit(s) past our last push`, `UNPUSHED commit(s) from a failed push`, and
`UNRESOLVED --pull conflicts`. Resolve every ATTENTION line or tell the owner why not.

### 2. Is anyone else mid-work?

If another session saved source in the last ~15 minutes, or is waiting on a background agent or workflow,
**hold** — the push would freeze half-finished work into a package version. "Idle" means its transcript
AND every file under its session directory (`subagents/`, `subagents/workflows/<id>/`) are quiet. Or ask.

### 3. Bring in teammates' work

**a. A failed earlier push first.** If preflight reports UNPUSHED commits on a clone, a push was rejected
or aborted and left a `Deploy …` commit as the clone's `main`; `--pull` would take it as the base and copy
GitLab's OLDER files over our own changes. Drop it (a deploy snapshot only — source trees untouched):
`git -C Deploy/repos/<repo> reset --hard $(git -C Deploy/repos/<repo> merge-base main origin/main)`.

**b. Merged work on hyphen-pb main.** Record the base, look, then pull:
```bash
BASE=$(git -C Deploy/repos/hyphen-pb rev-parse main)        # keep this: renames are found against it
bash tools/deploy_to_repos.sh --pull --dry-run              # incoming commits + files; changes nothing
bash tools/deploy_to_repos.sh --pull                        # 3-way merges into MobileSource/HomeFront
git -C Deploy/repos/hyphen-pb diff -M --name-status "$BASE" origin/main | grep '^R'   # renames
```
Handle every red line — only the first is harmless:
- `✗ HomeFrontPB is frozen — nothing was written` — expected.
- `✗ Branch-owned files changed upstream — HAND-MERGE these, they are NOT synced:` — for `HomeFront.csproj`
  and `HomeFront.sln` this is **an action**: the push rsyncs OUR copies over them, so merge the teammate's
  hunks into the source by hand (keep our FlexKit ProjectReference). `appsettings*.json`,
  `azure-pipelines*.yml`, `NuGet.Config` stay branch-owned — leave them.
- `✗ Deleted upstream — remove by hand if intended:` — `git rm` each in the app after checking; otherwise the
  push re-adds it. Same for the OLD path of every rename listed by the `grep '^R'` line.
- `✗ N file(s) CONFLICT — left UNCHANGED, merge by hand:` — write that list to `Deploy/pull-conflicts.txt`
  at once. **Do not `--pull` again or push until every file is resolved**: `--pull` has already merged
  `origin/main` into the clone, so the next `--pull` uses the teammate's version as the base and silently
  drops their side. Resolve each with the safe merge below, then delete the file.

Then **commit the app** (`Merge hyphen-pb main: <teammate commits>`), rebuild, and re-run preflight.

**c. Other repos.** `--pull` covers hyphen-pb main only. If `homefront`, `flexkit` or `flexcore` reports
upstream commits, merge each touched file into the owning source tree (app / FlexKit / FlexCore) with the
safe merge, base `main`, then `git -C Deploy/repos/<repo> reset --hard origin/main` once the merged source
is committed.

**d. Open (unmerged) MRs.** Find them in the GitLab UI (the Chrome extension carries the owner's session):
`https://gitlab.innovatixinc.com/groups/application-modernization/-/merge_requests/?state=opened` — the MR
author is the real person; every git commit on this machine reads `wadood`. Take one only when the owner
asks. Diff it against its **merge base**, never against main (that would also "revert" everything merged
into main since it branched):
`git -C Deploy/repos/hyphen-pb fetch origin <branch>` →
`MB=$(git -C Deploy/repos/hyphen-pb merge-base origin/main origin/<branch>)` → safe merge each file from
`git diff --name-only $MB origin/<branch>` with base `$MB`, theirs `origin/<branch>`.

**e. Deliberately not taking a teammate hunk** — record it so preflight stops flagging it:
`echo "<sha> <path> <reason>" >> Deploy/preflight-ack.txt`.

**f. Cloud sessions' GitHub mirrors** (owner 2026-10-01: the mirrors are always the same as the Mac). Claude Code
cloud sessions work from private GitHub mirrors and push `jira/<KEY>` branches there. Before deploying, fetch them
and look for anything not in the local lines — a branch is taken only when the owner asks (as for an MR):
```bash
git -C MobileSource/HomeFront fetch github --prune && git -C ../FlexKit fetch github --prune && git fetch origin --prune
git -C MobileSource/HomeFront branch -r --no-merged main | grep github/
git -C ../FlexKit branch -r --no-merged telerik-parity-20260904 | grep github/
git branch -r --no-merged main
```
Step 13 pushes the mirrors back level after the deploy.

**Safe merge** (never writes conflict markers into a source tree):
```bash
T=$(mktemp -d); C=Deploy/repos/<repo>; F=<path>; SRC=<source tree>
git -C $C show <base>:$F > $T/base; git -C $C show <theirs>:$F > $T/theirs; cp "$SRC/$F" $T/ours
git merge-file $T/ours $T/base $T/theirs && cp $T/ours "$SRC/$F"
# rc ≠ 0: conflicts are in $T/ours only — resolve there by hand, then copy it into the source
```

Done when preflight shows no `REVERT`/`ABSENT`/`RESURRECT`/`HANDMERGE`/`LOST` line, no
`upstream commit(s) past our last push`, no UNPUSHED commits and no unresolved conflicts.

### 4. Read what is about to ship

`git diff` in the app, FlexKit and FlexCore. Look for: half-finished edits; an unknown Razor component
attribute (compiles, then throws at runtime); a FlexKit change without its FlexCore port (compare the two
libraries' diffs file by file, and grep every ported file for `_content/FlexKit/` — assembly-scoped paths
compile and fail at runtime); library code that names HomeFront; a bench port the owner has not approved
(rule 11); secrets (rule 12). For a large diff run a parallel review first — and apply rule 6 while it runs.
Mirror mechanics and the parity command: [reference/flexkit-flexcore-sync.md](reference/flexkit-flexcore-sync.md).

### 5. Bump versions (before any dry run)

- `FlexKit/FlexKit.csproj` `<Version>` — bump if FlexKit changed (code or asset) and its version is burned.
- `FlexCore/FlexCore.csproj` `<Version>` — bump above nuget.org's latest if FlexCore changed.
  (`FlexCore.Llm.csproj` and `FlexCore.Documents.csproj` carry their own versions; neither is published.)

Re-check both numbers against reality here, not from memory — the gate does not check FlexCore's at all:
[reference/nuget-and-versions.md](reference/nuget-and-versions.md).

### 6. Build — serially

```bash
dotnet build MobileSource/HomeFront/HomeFront.sln --no-incremental -v q -nologo
dotnet build ../FlexCore.Showcase/FlexCore.Showcase.sln --no-incremental -v q -nologo
```

Each must print `0 Error(s)`. If FlexCore changed, also build its other consumers one at a time (owner
2026-09-14): Mutarjim, GhostWriter, DotNetCCM. Concurrent builds of one library race on its `obj/`. Never
build `HomeFrontPB.sln`. Full matrix: [reference/flexkit-flexcore-sync.md](reference/flexkit-flexcore-sync.md).

### 7. Run the harnesses

```bash
bash Docs-Skill/UpdateRepositories/scripts/run_harnesses.sh
```

All must pass — a green build proves nothing about them (csproj-excluded, reflection-based). Browser-driven
fixture hosts (`Sdk="Microsoft.NET.Sdk.Web"`, e.g. InputDialogBrowserChecks) are skipped and listed; run one
by hand per its README when its area changed. If a harness fails, decide whether it is stale against an
intended change or the code regressed, and say which.

`run_harnesses.sh` runs the C# half only. The browser suites, the ReportDesigner suites (run these when
`Reports/*` changed in either library), the Crystal bench pack, FlexCore's own suites and the consumer
builds are all in [reference/verification.md](reference/verification.md), with the exact commands — several
of those suites drive the OWNER'S live instance and need his say-so first.

### 8. Commit each repository

The app is a **nested** repo (`MobileSource/HomeFront`); FlexKit on `telerik-parity-20260904`; FlexCore on
`main`; the outer `HomeFront/` repo for `tools/` and `Docs-Skill/`. Describe other sessions' work from their
own CLAUDE.md sections and results files, not guesses. Never commit HomeFrontPB; leave scratch untracked.
If another session is editing a file you must commit, stage only your own hunk (write the patch against
`git show HEAD:<path>` and `git apply --cached`) rather than committing their work in progress.

### 9. Gate, dry run, gate, deploy

```bash
bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh --with-r2uat && bash tools/deploy_to_repos.sh --dry-run --with-r2uat
bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh --with-r2uat && bash tools/deploy_to_repos.sh --with-r2uat
```

The gate blocks on: upstream commits or UNPUSHED leftovers on any clone; any unshipped teammate change not in
our trees (judged by content, so the dry run's reset cannot hide it); unresolved `--pull` conflicts; conflict
markers; source saved in the last 10 minutes; FlexKit off its branch; a FlexKit repack under a burned version
or an asset change the script would not repack; unknown arguments. On PASS it records the baseline that
`verify_deploy.sh` and `publish_flexcore.sh` need.

Per repo the script runs: fetch → reset if local `main` ≠ `origin/main` (`✓ Reset to origin/main (remote had new
commits)` — also printed for an unpushed local commit, so it is not proof upstream moved) → rsync
(`✓ Synced source`) → login-skip strip (`✓ <repo>: no login skip in the pushed copy — password check intact`,
hyphen-pb and homefront only) → `▸ <repo>: removing files not needed to run` → hyphen-pb only: `Packing
FlexKit X` / `nupkg is current`, swap, staged build (`0 Error(s)`, `✓ Build succeeded`) → index gate →
commit → `main -> main`. Then `▸ Done.`, `▸ R1-UAT: FROZEN — not deployed`, and the
**R2-UAT stage**: `▸ R2-UAT: merging origin/main (N commit(s) ahead)` → `✓ Lockdown intact (4 disabled toggles, …)`
→ `✓ UAT pipeline still triggers on R2-UAT` → `✓ R2-UAT: no login skip` → build → `✓ Pushed R2-UAT`. It stops
unpushed (exit 1, main and the other repos already shipped) on a merge conflict, a lost lockdown, a retargeted
pipeline, a leaked login-skip or a failed build — resolve on the branch and re-run. In a dry run main is not
pushed, so the stage only checks `origin/main` as it stands.

**What failure looks like.** A login-guard, staged-build or index-gate failure skips only that repo
(`✗ Skipping <repo> due to …` or `✗ <repo>: … — not pushing`); the others still push and the run still ends
`▸ Done.` with exit 0. **A rejected `git push` aborts the whole run** (`set -e`): later repos never ship, no
`Done.`, and the clone keeps an unpushed commit — never force-push; go to step 3a. With `--with-r1uat`, R1-UAT
runs AFTER `Done.`, and a failure there exits 1. `✓ Packed`, `✓ Swapped` and `✓ Restored branch-only FLogin
lockdown` print whether or not the work happened. Step 10 is the proof.

### 10. Prove it

```bash
bash Docs-Skill/UpdateRepositories/scripts/verify_deploy.sh
```

Compares every clone with the gate's baseline: `shipped`, `unchanged` (fine only where the run printed
`✓ No changes to push`) or `UNPUSHED` (failed push). Then: login-skip gone, non-runtime files stripped,
password check ships, hyphen-pb's FlexKit PackageReference matches its package (a silent pack failure breaks
this), R2-UAT contains main with its lockdown and pipeline trigger intact, R1-UAT untouched, clones clean, local
login-skip intact. Must end
`ALL CHECKS PASSED — N repo(s) shipped since the gate`.

### 11. Publish FlexCore to NuGet — only if FlexCore changed, after step 10, and rule 6

```bash
bash Docs-Skill/UpdateRepositories/scripts/publish_flexcore.sh --check   # packs + inspects, no push
bash Docs-Skill/UpdateRepositories/scripts/publish_flexcore.sh
```

It packs the snapshot that was PUSHED to GitHub (a worktree of the flexcore staging clone at `origin/main`),
never the live FlexCore tree; refuses a version that is already published or not above the latest; checks
the package has FlexCore.dll + README and no FlexCore.Llm, agent notes, docs or tests; treats a duplicate or
a push without "Your package was pushed" as failure; then waits for nuget.org to list the version. The key
comes from the keychain inside the command and is never printed.

If the publish is **held** (rule 13), say so in the record with the version, the GitHub sha the hold was
judged on, and what must be fixed to lift it. Indexing can outlast the script's 6-minute wait — confirm the
listing before recording it as published. [reference/nuget-and-versions.md](reference/nuget-and-versions.md).

### 12. Record it for the other account

Two Claude accounts (Wadood, Innovatix) alternate weeks on this Mac and share one memory store.
- Update the first line of `~/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/MEMORY.md`
  and `handoff_pending_2026_09_12.md` beside it: remote heads, versions shipped/published, local commit ids,
  teammate commits merged, what is open or deliberately not done.
- Append a row to
  `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories/reference/history.md`
  and refresh its "Open as of" block (the canonical copy — the installed one is read-only), commit it in the
  HomeFront repo **and push that repo if it has a remote**, then run
  `bash /Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories/install.sh`.
- Give each library's `CLAUDE.md` a dated section for this round's own ship-review fixes — both, since the
  port is byte-identical.
- Jira moves belong to the hand-off, not the deploy, and only when the owner asks: see
  [reference/handoff.md](reference/handoff.md).

### 13. Bring the GitHub mirrors level with the Mac — last step, every round

Owner 2026-10-01: "everything should be same and current". After the records of step 12 are committed:
```bash
git -C MobileSource/HomeFront push github main
git -C ../FlexKit push github telerik-parity-20260904
git push origin main                                   # outer repo (vb6ToDotNet)
```
Plain pushes only. A non-fast-forward rejection means a cloud session pushed to that line: stop, fetch, show the
owner what is there (step 3f) — never force. FlexCore's GitHub `main` is the deploy snapshot the script pushes, so
it needs nothing here. These mirrors are private and carry the full working source (login skip included); they are
not deploy targets and nothing is stripped.

## Stop and ask the owner when

- another session is mid-work and the owner has not said to ship anyway;
- a `--pull` conflict cannot be resolved by reading both sides, or a merged teammate change deliberately
  conflicts with something we changed on purpose (record the decision in `preflight-ack.txt`);
- a harness fails and it is unclear whether the code or the harness is wrong;
- a review of the shipping diff is still running when everything else is ready;
- the R2-UAT stage stops on a conflict or a lost lockdown (that branch is the client release — never force it);
- the request would touch the HomeFrontPB tree, `--with-r1uat`, a secret, or someone else's merge request;
- a FlexKit diff turns out to be an unapproved bench port (rule 11) — holding costs a day, shipping it
  costs every host app;
- an MR's state cannot be read in the GitLab UI: list the candidates and ask, never decide from git refs;
- a publish is held and the hold's conditions are not demonstrably met.

## Maintaining this skill

`preflight.sh`, `pre_deploy_gate.sh` and `verify_deploy.sh` share `scripts/_common.sh`; `run_harnesses.sh`
and `install.sh` stand alone. After changing `_common.sh` or the gate, run `bash tests/classify_fixture.sh`
and `bash tests/gate_fixture.sh` (throwaway repos; they touch nothing real).

Keep this skill current in the same round as the lesson: a trap that lives only in a transcript is lost at
the hand-off. Every claim here carries its evidence — a commit, a script line, a dated owner directive or a
memory note — so the next reader can check it rather than trust it.
