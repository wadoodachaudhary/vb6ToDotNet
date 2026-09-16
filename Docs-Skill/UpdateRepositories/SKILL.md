---
name: update-repositories
description: Deploy HomeFront, FlexKit and FlexCore to GitLab/GitHub and publish FlexCore to NuGet — the "Update Repositories" routine for this Mac. Use when the user asks to update, deploy, push or sync the repositories, check whether they are in sync, pull teammates' merge requests, or hand the work over between the Wadood and Innovatix Claude accounts.
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
| `HomeFront/HomeFrontPB` | GitLab `hyphen-pb` **R1-UAT** | **FROZEN** — skipped unless `--with-r1uat` |
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

## Hard rules

1. **Working trees ship, uncommitted edits included — by design** (owner 2026-08-28). Never gate on a
   clean tree; still commit everything first so history matches what shipped.
2. **HomeFrontPB is frozen**: never write, commit, revert or build it (`HomeFrontPB.sln` never).
   **R1-UAT is frozen**: `--with-r1uat` only when the owner names R1-UAT in this request.
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

Every trap with its tell: [reference/rules-and-traps.md](reference/rules-and-traps.md). Past runs and what
is open: [reference/history.md](reference/history.md). When updating either, edit the canonical copy
(absolute path above) — the installed copy is read-only.

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
libraries' diffs file by file); library code that names HomeFront; secrets. For a large diff run a parallel
review first — and apply rule 6 while it runs.

### 5. Bump versions (before any dry run)

- `FlexKit/FlexKit.csproj` `<Version>` — bump if FlexKit changed (code or asset) and its version is burned.
- `FlexCore/FlexCore.csproj` `<Version>` — bump above nuget.org's latest if FlexCore changed.
  (`FlexCore.Llm.csproj` and `FlexCore.Documents.csproj` carry their own versions; neither is published.)

### 6. Build — serially

```bash
dotnet build MobileSource/HomeFront/HomeFront.sln --no-incremental -v q -nologo
dotnet build ../FlexCore.Showcase/FlexCore.Showcase.sln --no-incremental -v q -nologo
```

Each must print `0 Error(s)`. If FlexCore changed, also build its other consumers one at a time (owner
2026-09-14): Mutarjim, GhostWriter, DotNetCCM. Concurrent builds of one library race on its `obj/`. Never
build `HomeFrontPB.sln`.

### 7. Run the harnesses

```bash
bash Docs-Skill/UpdateRepositories/scripts/run_harnesses.sh
```

All must pass — a green build proves nothing about them (csproj-excluded, reflection-based). Browser-driven
fixture hosts (`Sdk="Microsoft.NET.Sdk.Web"`, e.g. InputDialogBrowserChecks) are skipped and listed; run one
by hand per its README when its area changed. If a harness fails, decide whether it is stale against an
intended change or the code regressed, and say which. If `Reports/*` changed in FlexKit or FlexCore, also run
the suites in `/Users/wadood/projects/JavaToCSharp/tools/ReportDesigner.*` (their READMEs give the
arguments; FlexCore via `-p:ReportSourceDir=/Users/wadood/projects/VBToCSharp/FlexCore/Reports`).

### 8. Commit each repository

The app is a **nested** repo (`MobileSource/HomeFront`); FlexKit on `telerik-parity-20260904`; FlexCore on
`main`; the outer `HomeFront/` repo for `tools/` and `Docs-Skill/`. Describe other sessions' work from their
own CLAUDE.md sections and results files, not guesses. Never commit HomeFrontPB; leave scratch untracked.

### 9. Gate, dry run, gate, deploy

```bash
bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh && bash tools/deploy_to_repos.sh --dry-run
bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh && bash tools/deploy_to_repos.sh
# only when the owner names R1-UAT:
bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh --with-r1uat && bash tools/deploy_to_repos.sh --with-r1uat
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
commit → `main -> main`. Then `▸ Done.` and `▸ R1-UAT: FROZEN — not deployed`.

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
this), R1-UAT untouched unless deployed, clones clean, local login-skip intact. Must end
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

### 12. Record it for the other account

Two Claude accounts (Wadood, Innovatix) alternate weeks on this Mac and share one memory store.
- Update the first line of `~/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/MEMORY.md`
  and `handoff_pending_2026_09_12.md` beside it: remote heads, versions shipped/published, local commit ids,
  teammate commits merged, what is open or deliberately not done.
- Append a row to
  `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories/reference/history.md`
  (the canonical copy — the installed one is read-only), commit it in the HomeFront repo, then run
  `bash /Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories/install.sh`.

## Stop and ask the owner when

- another session is mid-work and the owner has not said to ship anyway;
- a `--pull` conflict cannot be resolved by reading both sides, or a merged teammate change deliberately
  conflicts with something we changed on purpose (record the decision in `preflight-ack.txt`);
- a harness fails and it is unclear whether the code or the harness is wrong;
- a review of the shipping diff is still running when everything else is ready;
- the request would touch R1-UAT, HomeFrontPB, a secret, or someone else's merge request.

## Maintaining this skill

Scripts share `scripts/_common.sh`. After changing it or the gate, run
`bash tests/classify_fixture.sh` and `bash tests/gate_fixture.sh` (throwaway repos; touch nothing real).
