# Update Repositories — rules and traps

Everything here was learned the hard way on this project between 2026-08-24 and 2026-09-16.
Each trap has a **tell** (how it shows up) so you can recognise it mid-run.

## Owner rules (binding, with dates)

| Rule | Since |
|---|---|
| Uncommitted working-tree edits ship by design — "changes can happen at the last minute, outside Claude as well". Never gate on a clean tree; still commit first. | 2026-08-28 |
| Pull BOTH remotes before pushing; compare `origin/main` against the sha we last pushed (the clone's local `main`), never against the clone's HEAD. | 2026-08-30 |
| Before deploying, check GitLab for open merge requests (step 0b) and bring them in. | 2026-09-07 |
| HomeFrontPB feeds R1-UAT only and nothing writes into it. Never build `HomeFrontPB.sln`. | 2026-08-31 / 09-08 |
| "No changes to HomeFrontPB on the local device and R1-UAT is frozen." The deploy skips R1-UAT unless `--with-r1uat`; the FlexKit feed moved to `Deploy/local-packages`. | 2026-09-12/13 |
| FlexKit → FlexCore always, as a port (keep FlexCore's own code and branding), then rebuild FlexCore.Showcase. FlexCore → FlexKit only on request. | 2026-09-10 |
| Never merge FlexCore to main of FlexKit; FlexKit ships from `telerik-parity-20260904`. | 2026-08-22 |
| Publish FlexCore to nuget.org after every deploy that changed it. NuGet key only via `security find-generic-password -s flexcore-nuget-key -w`. | 2026-08-11 |
| Bump FlexKit before every deploy that changes it. | 2026-08-29 |
| Never push local `appsettings*.json` — the remote owns environment config. | 2026-08-24 |
| No comment stripping — code ships as written. | 2026-08-02 |
| The dev login-skip stays functional on this machine and is removed from every pushed copy. | 2026-09-12 |
| "Delete codex, claude files and other files not needed for HomeFront to run when deploying (keep them locally)." | 2026-09-13 |
| FlexCore.Llm: "push it to Git but not to NuGet". | 2026-09-14 |
| Never force-push; reconcile a "fetch first" rejection. Never merge/close someone's MR unless asked. Never enter credentials into a login form. | standing |

## What the deploy strips from every pushed branch (and keeps locally)

Two layers, because `rsync --exclude` neither copies NOR deletes — a file already committed on the
branch survives every later deploy unless it is actively removed:

- **Login-skip guard** `strip_dev_login_bypass()`: deletes code files containing `DevAutoLogin`, edits
  the registration out of `Program.cs` (anchored; asserts `UseAuthentication`, `UseAuthorization`,
  `UseAntiforgery`, `/auth/signin` survive), REFUSES the deploy if a config/CI/script file holds the
  marker, leaves `.md` prose alone, asserts `FLogin.razor` still consults `Sec.EnforcePasswordCheck`,
  then a second gate greps the staged INDEX.
- **Non-runtime strip** `NON_RUNTIME_EXCLUDE` + `strip_non_runtime_files()`: CLAUDE.md, AGENTS.md,
  GEMINI.md, .agents/, .claude/, .codex/, .codex-backups/, verification/, tests/, docs/, Docs/,
  PUBLISHING.md, Data/DataControl.API.md, *.vb6_gap.txt.
  **Kept on purpose:** README.md (FlexKit/FlexCore `PackageReadmeFile` — pack FAILS without it),
  LICENSE files, .gitignore, build/CI files, and `wwwroot/resources/**.zip` (runtime download
  templates, not backups).
- **Repo-owned, never touched:** `azure-pipelines*.yml`, `NuGet.Config`, `local-packages/`, `certs/`,
  `appsettings*.json`, `App_Data/jira-settings.json`.

Only hyphen-pb main gets the FlexKit pack + PackageReference swap and a staged build. `homefront` main
keeps `<ProjectReference Include="..\..\..\FlexKit\FlexKit.csproj" />` and no local-packages, so that repo
cannot be built on its own; `homefront`, `flexkit` and `flexcore` would all take a bad removal silently —
dry-run and pack from a stripped scratch copy before changing a strip list.

## Traps

### A merged MR is silently reverted by our own deploy
The push rsyncs our tree over hyphen-pb main. A teammate's commit merged between two of our deploys
is overwritten by our older copy; afterwards `origin/main == our tree`, the MR still shows Merged, its
commit is still in history, and step 0 looks clean forever.
**Tell:** none, unless you look. **Check:** blob hashes per file — ours == `commit^:path` means
reverted (preflight §5). **Fix:** `deploy_to_repos.sh --pull`, or `git show <sha> -- <files> | git apply`
when our copy still hashes to the MR base. Happened to MR !33 (07faa53, 2026-09-09, restored 09-12);
caught in time for Irfan b23d656 + Deepika aede435 (09-15) and Irfan 432c842 / 29f8287 / 40ef498 +
Deepika 3581c61 (pending as of 09-16).

### Any push-mode run hides upstream commits from `--pull`
`--pull` compares `origin/main` with the staging clone's LOCAL `main` (= what we last pushed). Every
push-mode run — including `--dry-run` — resets local `main` to `origin/main` whenever upstream moved.
From then on `--pull` prints `✓ No new changes on remote`, step 0 looks clean, and the next push
rsyncs our tree over the teammates' commits. **Tell:** a `✓ Reset to origin/main (remote had new
commits)` line in a push log. **Fix:** run `--pull` FIRST; if it already happened, preflight §5's blob
check still sees the lost files — merge them from the commits it names.

### `--dry-run` burns the FlexKit version
A dry run skips only commit/push. It still runs a real `dotnet pack` into `Deploy/local-packages` and
the full staged build, which restores that version into `~/.nuget/packages`. Change FlexKit after a dry
run and the real run repacks the same version against a cache holding the dry run's bytes.
**Fix:** bump before the dry run; bump again if FlexKit changes after it.

### The script reports success that did not happen
- Unknown arguments are ignored: `--dryrun`, `--with-r1-uat` → a REAL push. (`pre_deploy_gate.sh`
  rejects unknown arguments.)
- A skipped repo (`✗ … Skipping`) still ends `▸ Done.` with exit 0; the other repos still push.
- `pack_flexkit` runs inside `$(…)`, where bash 3.2's errexit does not apply: a failed `dotnet pack`
  still prints `✓ Packed FlexKit X`, empties the clone's local-packages, and — if X is in the NuGet
  cache — the staged build passes and main is pushed referencing a package it does not carry.
  `verify_deploy.sh` checks PackageReference == package on the branch.
- `✓ Swapped …`, `✓ Removed FlexKit from .sln` and `✓ Restored branch-only FLogin lockdown` print
  unconditionally.
- The push loop's `git fetch origin 2>/dev/null || true` hides network/auth errors, and the reset is
  conditional, so a clone left dirty by an aborted run is not cleaned.
- The login-skip strip treats `.md` AND `.txt` as prose, but the staged-index gate exempts only `*.md`:
  a `.txt` mentioning DevAutoLogin passes the strip and then blocks the push (fails safe, confusingly).

### An uncommitted edit exactly undoes a recent commit
A stale-copy edit restores a file to the blob BEFORE a fix. Shipped a removed "Multi-sort" caption
back into FlexKit 0.1.99 and **published FlexCore 0.2.41** (09-13).
**Tell:** `git diff`'s `index A..B` is the reverse of a recent commit. **Check:** preflight §6.

### Burned package versions (NuGet cache shadowing)
`dotnet restore` resolves `PackageReference FlexKit X` from `~/.nuget/packages/flexkit/X/` first and
treats it as immutable. Repacking the same version refreshes only the feed; the staged build verifies
the OLD bytes while the branch ships the NEW nupkg.
**Tell:** the staged build names a member that exists in source; or the cache dir predates the pack.
**Fix:** bump. Never delete from `~/.nuget/packages` (other consumers share it).

### FlexKit edited during the run
The script packs once per run, but if FlexKit source changes between the main and R1-UAT passes, the
second pass repacks the same version and the staged build restores the first pack from the cache.
**Tell:** two `Packing FlexKit X` lines for the same X in one log. **Fix:** don't deploy while a
session is editing FlexKit; re-run with a bumped version.

### Deploying while another session is mid-work
Froze half-finished GridControl work into R1-UAT's package on 09-10 and would have frozen Crystal
designer work into FlexKit 0.1.100 on 09-14 (held).
**Tell:** preflight §1/§2. **Idle check:** the session transcript AND every file under its session
directory, recursively — background workflows write to `subagents/workflows/<id>/`. A turn that
"ended" with text saying it is waiting on a workflow is NOT done.

### Publishing FlexCore before a review lands
Git pushes can be followed up; nuget.org versions can only be unlisted. FlexCore 0.2.41 shipped a
defect because the review finished after the publish (09-13). Hold the publish until reviews land.

### Silent success in shell
- `grep -c … || echo 0` is **fail-open**: zero matches prints "0" AND exits 1, so the fallback appends
  a second "0", the integer test errors, and `if` treats it as false. The R1-UAT lockdown check did
  this for weeks. Count with `grep -o … | wc -l`.
- `--dry-run` used to push R1-UAT for real (fixed 09-12). Still verify remote heads after any run.
- A failed hyphen-pb build skips only that repo; the others still push and the run ends "Done."
  Read the log for `✗` and missing `main -> main` lines.
- `checkout -- .` restores the worktree FROM the index, so staged strips survive a dry run — the
  script now resets `--hard`.

### FlexCore's other consumers
Since 2026-09-14 Mutarjim, GhostWriter and DotNetCCM ProjectReference **FlexCore** (GhostWriterLLM,
HomeFront and HomeFrontPB stay on FlexKit). After a FlexCore change build them too — one at a time;
concurrent builds race on FlexCore's `obj/`.

### FlexCore on GitHub is a deploy mirror
GitHub `main` holds `Deploy flexcore — <date>` snapshot commits, so the local FlexCore `main` diverges
from `github/main` and a direct `git push github main` is rejected as non-fast-forward. Never force it —
FlexCore reaches GitHub only through the staging clone. `FlexCore.Llm/` goes to GitHub, never NuGet
(owner 2026-09-14: "push it to Git but not to NuGet"; FlexCore.csproj excludes it).

### zsh (the interactive shell here)
- `$ref:local-packages/x` — zsh reads `:l` as the lowercase modifier. Write `${ref}:path`.
- No GNU `timeout` on macOS — use `perl -e 'alarm shift; exec @ARGV' 420 cmd …`.
- An unmatched glob aborts the whole command (`no matches found`). Quote globs or use `(N)`.
- Empty output is not proof of "clean" — check the command actually ran.
- zsh does not word-split unquoted variables: `set -- $x` gives ONE argument. Use `set -- ${=x}` in zsh,
  or run the loop under bash.

### Reflection harnesses
`verification/**` is excluded from `HomeFront.csproj`, so API/signature changes build clean and then
throw at runtime. Run them. A failing harness is often stale against an intended behaviour change —
e.g. WizardChecks needed `_gridReady` seeding (09-10) and a debounce wait for filter-row input.

### Published FlexCore SourceLink points at commits GitHub does not have
The nuspec of each published FlexCore names the LOCAL FlexCore commit (e.g. 0.2.43 → 6d8d7bf), but
GitHub `main` is deploy-snapshot history (5be387c), so the SourceLink commit does not exist there. Local
tracking refs also look diverged for the same reason (FlexCore `main…github/main`, app `main…origin/main`
on the personal wchaudhary remote) — that is not drift; compare against the staging clones instead.

### Secrets that are already pushed
`App_Data/jira-settings.json` carries a live Atlassian API token and is tracked on hyphen-pb main,
R1-UAT and homefront main (the script used to `git add -f` it; stopped 09-12). Rotation is a human
action in Atlassian — still OPEN. Never print the value.

### Security finding still open (owner decision)
The QA login page shows a live "Security Options (Dev)" panel; `ISecurityOptionsService` is a
singleton, so one anonymous "Deselect All" disables password checks for every user until the pool
recycles. `Security__HideDevPanel=true` on the app pools closes it (and the R1-UAT "Deselect All"
gap) but ends the owner's ADMIN sign-in shortcut, so it has not been done.
