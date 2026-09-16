---
name: update-repositories
description: Deploy HomeFront, FlexKit and FlexCore to GitLab/GitHub and publish FlexCore to NuGet — the "Update Repositories" routine for this Mac. Use when the user asks to update, deploy, push or sync the repositories, check whether they are in sync, pull teammates' merge requests, or hand the work over between the Wadood and Innovatix Claude accounts.
---

# Update Repositories

Ships this machine's working trees to four remote repositories — after merging teammates' work and
proving nothing gets reverted — then proves what went out. `tools/deploy_to_repos.sh` does the
pushing; the checks before and the proof after are this skill.

Canonical, git-tracked source: `/Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories/`.
Installed for every Claude account on this Mac at `~/.claude/skills/update-repositories/` (both
accounts share `~/.claude`). Edit the canonical source, then `bash install.sh` there.

| Source working tree | Remote | Notes |
|---|---|---|
| `HomeFront/MobileSource/HomeFront` (nested git repo) | GitLab `hyphen-pb` **main** + `homefront` main | main is the QA-deployed app; `homefront` is a source mirror |
| `HomeFront/HomeFrontPB` | GitLab `hyphen-pb` **R1-UAT** | **FROZEN** — skipped unless `--with-r1uat` |
| `FlexKit` on branch **`telerik-parity-20260904`** | GitLab `flexkit` main | never `git checkout main` in FlexKit; not on NuGet |
| `FlexCore` (main) | GitHub `wadoodachaudhary/FlexCore` + **nuget.org** | GitHub main is deploy-snapshot history; `FlexCore.Llm/` goes to GitHub only |

Staging clones: `HomeFront/Deploy/repos/{hyphen-pb,homefront,flexkit,flexcore}`. FlexKit feed:
`HomeFront/Deploy/local-packages`. The outer git repo is `HomeFront/` itself (owns `tools/` and this
skill); `VBToCSharp/` is not a repo.

## Only when asked

Deploy and publish ONLY when the owner explicitly asks to update/deploy the repositories in this
request. "Check whether the repositories are in sync — do not update" means: run steps 1 and 10
(`preflight.sh`, `verify_deploy.sh`), report, change nothing. A message asking for details or fixes
does not authorise a push or a publish.

## Hard rules

1. **Working trees ship, uncommitted edits included — by design** (owner 2026-08-28). Never gate on a
   clean tree; still commit everything first so history matches what shipped.
2. **HomeFrontPB is frozen**: never write, commit, revert or build it (`HomeFrontPB.sln` never).
   **R1-UAT is frozen**: `--with-r1uat` only when the owner names R1-UAT in this request.
3. **`--pull` before ANY push-mode run.** A teammate commit on hyphen-pb main that our tree lacks is
   silently REVERTED by the push. Worse: every push-mode run — **even `--dry-run`** — resets the
   staging clone's `main` to `origin/main` when upstream moved, after which `--pull` says "No new
   changes" and the teammates' work is lost on the next push. `pre_deploy_gate.sh` enforces this.
4. **Burned versions never ship.** A FlexKit version already in `~/.nuget/packages/flexkit/` is
   shadowed in the staged build; a FlexCore version on nuget.org can never be replaced. `--dry-run`
   really packs and restores FlexKit, so it burns the version too: bump BEFORE the dry run, and bump
   again if FlexKit changes after it.
5. **FlexKit → FlexCore always** (owner 2026-09-10): a FlexKit change ships with its FlexCore port;
   FlexCore → FlexKit only when asked. Never merge FlexCore into FlexKit.
6. **A NuGet publish is permanent.** Publish only after `verify_deploy.sh` passes. Since 09-13 every
   round has also held the publish (and the push) until any review of the shipping diff has landed —
   FlexCore 0.2.41 shipped a defect because it didn't; the owner has not formally confirmed this rule.
7. **Type flags exactly.** `deploy_to_repos.sh` silently ignores unknown arguments: `--dryrun` or
   `--with-r1-uat` runs a REAL push.
8. Never force-push. Never push `appsettings*.json`. Never print or commit a secret. Never merge or close
   someone else's merge request unless asked. Never commit scratch files.
9. The dev login-skip (`Services/DevAutoLogin.cs`) stays **working locally** and **absent from every
   pushed branch** — the script strips it from staging clones only. Never delete it locally.

Every trap with its tell: [reference/rules-and-traps.md](reference/rules-and-traps.md).
Past runs, lessons, and what is open: [reference/history.md](reference/history.md).

## Procedure

Run from `/Users/wadood/projects/VBToCSharp/HomeFront`. The helper scripts are bash; call them with
`bash`, never `source` them into zsh. `$S` below means `Docs-Skill/UpdateRepositories/scripts`.

### 1. Preflight (read-only)

```bash
bash Docs-Skill/UpdateRepositories/scripts/preflight.sh
```

Reports active agent sessions, working-tree state, step 0 (each remote vs the sha we last pushed),
branches that may carry merge requests, **teammate files the push would revert** (blob proof), edits
that exactly undo a recent commit, burned library versions, and skill-install drift, then an
**ATTENTION** list. Resolve every line or tell the owner why not. Section 5's blob check still works
when an earlier push-mode run already hid upstream commits from step 0 — trust it over step 0.

### 2. Is anyone else mid-work?

If another session saved source in the last ~15 minutes, or is waiting on a background agent or
workflow, **hold** — the push would freeze half-finished work into a package version. "Idle" means its
transcript AND every file under its session directory (`subagents/`, `subagents/workflows/<id>/`) are
quiet; a main transcript can sit silent while its agents still edit. Or ask the owner.

### 3. Bring in teammates' work

- **Open MRs:** GitLab UI (the Chrome extension carries the owner's session):
  `https://gitlab.innovatixinc.com/groups/application-modernization/-/merge_requests/?state=opened`.
  The MR author is the real person; every git commit on this machine reads `wadood`.
- **First, a failed earlier push?** If preflight says a staging clone has UNPUSHED commits, a previous
  push was rejected or aborted and left a `Deploy …` commit as the clone's `main`. `--pull` would take
  that as the base and copy GitLab's OLDER files over our own changes. Drop it first (it is only a
  deploy snapshot; our source trees are untouched):
  `git -C Deploy/repos/<repo> reset --hard $(git -C Deploy/repos/<repo> merge-base main origin/main)`.
- **Merged work on hyphen-pb** (preflight shows `origin/main moved`, `REVERT:`, `LOST:`, `ABSENT:` or
  `DIVERGED:`):
  ```bash
  bash tools/deploy_to_repos.sh --pull --dry-run   # lists incoming commits and files, changes nothing
  bash tools/deploy_to_repos.sh --pull             # 3-way merges them into MobileSource/HomeFront
  ```
  Expect `✓ Fast-forwarded N, 3-way merged M file(s) into HomeFront`. Two red lines are **normal**:
  `✗ HomeFrontPB is frozen — nothing was written`, and `✗ Branch-owned files changed upstream —
  HAND-MERGE these` (csproj, sln, appsettings, pipelines, NuGet.Config are listed, never synced).
  Read the `--pull --dry-run` file list first: it must name only the teammates' files. **Save the
  `✗ N file(s) CONFLICT — left UNCHANGED` list** — `--pull` merges `origin/main` into the clone, so a
  second run prints `No new changes` and never shows those conflicts again. Hand-merge each one with the
  `git -C … diff <base>..origin/main -- <file>` it prints. Then **commit the app**
  (`Merge hyphen-pb main: <teammate commits>`), rebuild, and re-run preflight until it has no `REVERT:`,
  `LOST:`, `ABSENT:` or `DIVERGED:` line and no `origin/main moved` line for ANY repo.
- **`--pull` covers hyphen-pb main only.** If `homefront`, `flexkit` or `flexcore` moved upstream, merge
  by hand before any push-mode run: `git -C Deploy/repos/<repo> log main..origin/main --stat`, then for
  each file 3-way merge into the owning source tree (app / FlexKit / FlexCore) with
  `git merge-file <ours> <base = main:<file>> <theirs = origin/main:<file>>`, and fast-forward that clone's
  `main` only after the merged source is committed.

### 4. Read what is about to ship

`git diff` in the app, FlexKit and FlexCore. Look for: half-finished edits; an unknown Razor component
attribute (compiles, then throws at runtime); a FlexKit change without its FlexCore port (compare the
two libraries' diffs file by file); library code that names HomeFront; secrets.
For a large diff run a parallel review first — and apply rule 6 while it runs.

### 5. Bump versions (before any dry run)

- `FlexKit/FlexKit.csproj` `<Version>` — bump if FlexKit changed and its version is burned.
- `FlexCore/FlexCore.csproj` `<Version>` — bump if FlexCore changed and its version is on nuget.org.
  (`FlexCore.Llm.csproj` and `FlexCore.Documents.csproj` carry their own versions; neither is published.)

### 6. Build — serially

```bash
dotnet build MobileSource/HomeFront/HomeFront.sln --no-incremental -v q -nologo
dotnet build ../FlexCore.Showcase/FlexCore.Showcase.sln --no-incremental -v q -nologo
```

Each must print `0 Error(s)`. If FlexCore changed, also build its other consumers one at a time
(owner 2026-09-14 moved them to FlexCore): Mutarjim, GhostWriter, DotNetCCM. Concurrent builds of one
library race on its `obj/` and print phantom CS0101. Never build `HomeFrontPB.sln`.

### 7. Run the harnesses

```bash
bash Docs-Skill/UpdateRepositories/scripts/run_harnesses.sh
```

All must pass — a green build proves nothing about them (csproj-excluded, reflection-based). If one
fails, decide whether the harness is stale against an intended change or the code regressed, and say
which. If `Reports/*` changed in FlexKit or FlexCore, also run the report-designer suites in
`/Users/wadood/projects/JavaToCSharp/tools/ReportDesigner.*` (their READMEs give the arguments); they
compile against FlexKit by default and against FlexCore with
`-p:ReportSourceDir=/Users/wadood/projects/VBToCSharp/FlexCore/Reports`.

### 8. Commit each repository

The app is a **nested** repo (`MobileSource/HomeFront`); FlexKit on `telerik-parity-20260904`; FlexCore on
`main`; the outer `HomeFront/` repo for `tools/` and `Docs-Skill/`. Describe other sessions' work from
their own CLAUDE.md sections and results files, not guesses. Never commit HomeFrontPB; leave scratch
files untracked.

### 9. Gate, dry run, deploy

```bash
bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh && bash tools/deploy_to_repos.sh --dry-run
bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh && bash tools/deploy_to_repos.sh
# only when the owner names R1-UAT:
bash Docs-Skill/UpdateRepositories/scripts/pre_deploy_gate.sh --with-r1uat && bash tools/deploy_to_repos.sh --with-r1uat
```

The gate blocks when upstream moved past our last push, when source changed in the last 10 minutes,
when FlexKit is off its branch, or when FlexKit would be repacked under a burned version.

Per repo the script runs: fetch → reset (only if origin moved) → rsync (`✓ Synced source`) → login-skip
strip (`✓ <repo>: no login skip in the pushed copy — password check intact`) → `▸ <repo>: removing
files not needed to run` → hyphen-pb only: `Packing FlexKit X` / `nupkg is current`, swap, staged build
(`0 Error(s)`, `✓ Build succeeded`) → index gate → commit → `main -> main`. Then `▸ Done.` and
`▸ R1-UAT: FROZEN — not deployed`.

**A rejected `git push` aborts the WHOLE run** (`set -e`): later repos never ship, there is no `Done.`,
and the clone keeps an unpushed `Deploy …` commit — never force-push; go back to step 3's recovery.
**Green lines are not proof.** A login-guard, staged-build or index-gate failure only skips that repo
(`✗ … Skipping`) — the others still push and the run exits 0 with `Done.`; `✓ Packed`, `✓ Swapped` and `✓ Restored branch-only FLogin lockdown` print whether or
not the work happened, and a failed `dotnet pack` does not stop the run. A missing `main -> main` means
that repo did not ship. Step 10 is the proof.

### 10. Prove it

```bash
bash Docs-Skill/UpdateRepositories/scripts/verify_deploy.sh       # add --with-r1uat if used
```

Must end `ALL CHECKS PASSED`: login-skip gone from every pushed app branch; non-runtime files
stripped; the password check still ships; hyphen-pb's FlexKit PackageReference matches the package on
the branch (catches a silent pack failure); clones clean on main; the local login-skip intact.

### 11. Publish FlexCore to NuGet — only if FlexCore changed, after step 10, and rule 6

```bash
S=$(mktemp -d)
cd /Users/wadood/projects/VBToCSharp/FlexCore && dotnet pack FlexCore.csproj -c Release -o "$S"
unzip -l "$S"/FlexCore.*.nupkg | grep -ciE 'llm|CLAUDE|AGENTS|/docs/|PUBLISHING|/tests/'   # must print 0
dotnet nuget push "$S"/FlexCore.<version>.nupkg \
  --api-key "$(security find-generic-password -s flexcore-nuget-key -w)" \
  --source https://api.nuget.org/v3/index.json --skip-duplicate 2>&1 \
  | grep -vE 'api-key|ApiKey' | sed -E 's/[A-Za-z0-9]{40,}/[redacted]/g'
rm -rf "$S"
```

The owner's 08-11 note says publish after every deploy; the 09-07 context says only when FlexCore
changed — publishing unchanged code under a new version helps nobody, so publish when it changed.

### 12. Record it for the other account

Two Claude accounts (Wadood, Innovatix) alternate weeks on this Mac and share one memory store. Update
the first line of `~/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/MEMORY.md`
and the hand-off note `handoff_pending_2026_09_12.md` beside it: remote heads, versions
shipped/published, local commit ids, teammate commits merged, and what is open or deliberately not
done. Append a row to [reference/history.md](reference/history.md), then `bash install.sh`.

## Stop and ask the owner when

- another session is mid-work and the owner has not said to ship anyway;
- `--pull` leaves a conflict you cannot resolve by reading both sides, or a merged MR deliberately
  conflicts with a change we made on purpose;
- a harness fails and it is unclear whether the code or the harness is wrong;
- a review of the shipping diff is still running when everything else is ready;
- the request would touch R1-UAT, HomeFrontPB, a secret, or someone else's merge request.
