# Update Repositories — rules and traps

Everything here was learned the hard way on this project between 2026-08-24 and 2026-09-16.
Each trap has a **tell** (how it shows up) so you can recognise it mid-run.

## Owner rules (binding, with dates)

| Rule | Since |
|---|---|
| Uncommitted working-tree edits ship by design — "changes can happen at the last minute, outside Claude as well". Never gate on a clean tree; still commit first. | 2026-08-28 |
| Pull BOTH remotes before pushing; compare `origin/main` against the sha we last pushed (the clone's local `main`), never against the clone's HEAD. | 2026-08-17 (hyphen-pb) / 08-28 (homefront) |
| Before deploying, check GitLab for open merge requests (step 0b) and bring them in. | 2026-09-07 |
| HomeFrontPB feeds R1-UAT only and nothing writes into it. Never build `HomeFrontPB.sln`. | 2026-08-31 / 09-08 |
| "No changes to HomeFrontPB on the local device and R1-UAT is frozen." The deploy skips R1-UAT unless `--with-r1uat`; the FlexKit feed moved to `Deploy/local-packages`. | 2026-09-12/13 |
| FlexKit → FlexCore always, as a port (keep FlexCore's own code and branding), then rebuild FlexCore.Showcase. FlexCore → FlexKit only on request. | 2026-09-10 |
| Never merge FlexCore (or FlexKit's working branch) into a GitLab main; FlexKit ships from `telerik-parity-20260904`. | 2026-09-01 |
| Publish FlexCore to nuget.org after every deploy that changed it. NuGet key only via `security find-generic-password -s flexcore-nuget-key -w`. | 2026-08-11 |
| Bump FlexKit before every deploy that changes it. | 2026-08-29 |
| Never push local `appsettings*.json` — the remote owns environment config. | 2026-08-20 |
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
- **Repo-owned, never touched:** `azure-pipelines*.yml`, `NuGet.Config`, `certs/`, `appsettings*.json`,
  `App_Data/jira-settings.json`.
- **`local-packages/`**: rsync never touches it, but on hyphen-pb (main, and R1-UAT when deployed) the
  script replaces `FlexKit.*.nupkg/.snupkg` with the freshly packed version and commits it.

Only hyphen-pb main gets the FlexKit pack + PackageReference swap and a staged build. `homefront` main
keeps `<ProjectReference Include="..\..\..\FlexKit\FlexKit.csproj" />` and no local-packages, so that repo
cannot be built on its own; `homefront`, `flexkit` and `flexcore` would all take a bad removal silently —
dry-run and pack from a stripped scratch copy before changing a strip list.

## Traps

### A merged MR is silently reverted by our own deploy
The push rsyncs our tree over hyphen-pb main. A teammate's commit merged between two of our deploys
is overwritten by our older copy; afterwards `origin/main == our tree`, the MR still shows Merged, its
commit is still in history, and step 0 looks clean forever.
**Tell:** none, unless you look. **Check:** preflight §5 and the gate judge every teammate change by its
distinctive lines (`REVERT`, `ABSENT`, `RESURRECT`, `HANDMERGE`; `LOST` when our deploy already overwrote it). **Fix:** `deploy_to_repos.sh --pull`, or `git show <sha> -- <files> | git apply`
when our copy still hashes to the MR base. Happened to MR !33 (07faa53, 2026-09-09, restored 09-12);
caught in time for Irfan b23d656 + Deepika aede435 (09-15) and Irfan 432c842 / 29f8287 / 40ef498 +
Deepika 3581c61 (pending as of 09-16).

### Any push-mode run hides upstream commits from `--pull`
`--pull` compares `origin/main` with the staging clone's LOCAL `main` (= what we last pushed). Every
push-mode run — including `--dry-run` — resets local `main` to `origin/main` whenever upstream moved.
From then on `--pull` prints `✓ No new changes on remote`, step 0 looks clean, and the next push
rsyncs our tree over the teammates' commits. **Tell:** a `✓ Reset to origin/main (remote had new
commits)` line in a push log. **Fix:** run `--pull` FIRST; if it already happened, preflight §5 and the gate still see the
missing content (`REVERT`) — merge it from the commits they name with the safe merge in SKILL.md.

### A rejected push, then `--pull`, reverts OUR OWN changes
A push rejected because a teammate merged between the fetch and the push aborts the run and leaves the
unpushed `Deploy <repo> — …` commit as the clone's local `main`. `--pull` uses local `main` as
LAST_PUSHED, so its file list includes every file WE changed; for each, our source equals that "base",
and the fast-forward branch copies origin/main's OLDER copy over our working tree — uncommitted edits
included. Reproduced with a fixture on 2026-09-16. **Tell:** preflight/gate report UNPUSHED commits on a
clone. **Fix:** `git -C <clone> reset --hard $(git -C <clone> merge-base main origin/main)` BEFORE
`--pull`, then check that `--pull --dry-run` lists only the teammates' files. Never force-push.

### `--pull` conflicts vanish — and the next `--pull` silently drops the teammate's side
`--pull` runs `git merge origin/main` in the clone before the per-file 3-way merges, so afterwards local
`main == origin/main`. The `✗ N file(s) CONFLICT — left UNCHANGED` list is printed once. A second `--pull`
with no new upstream prints `No new changes`; worse, if a teammate merges again, the next `--pull` uses the
FIRST teammate's version as the base, our unmerged copy looks like a deliberate removal of their change, the
merge comes out "clean", and their hunk is gone (reproduced 2026-09-16). **Fix:** write the list to
`Deploy/pull-conflicts.txt` immediately — the gate blocks while it exists — resolve every file with the safe
merge, then delete it. Never `--pull` again or push in between.

### `--pull` output that is NOT harmless
- `✗ Branch-owned files changed upstream — HAND-MERGE these, they are NOT synced:` — `HomeFront.csproj` and
  `HomeFront.sln` are NOT excluded from the push rsync, so our copies overwrite a teammate's edits (Irfan's
  67b3a03 and 2e740db both edited the csproj). Merge their hunks into the source; the swap only rewrites the
  FlexKit reference. Only `appsettings*.json`, `azure-pipelines*.yml` and `NuGet.Config` are truly safe to leave.
- `✗ Deleted upstream — remove by hand if intended:` — `--pull` never deletes; the next push re-adds the file.
- **Renames are not reported at all:** the new path is synced, the old path is neither synced nor listed, and
  the push restores it on main. Find them with `git diff -M --name-status <base> origin/main | grep '^R'`.
  Preflight/gate classify with `--no-renames`, so a leftover old path shows as `RESURRECT`.

### `--pull` covers hyphen-pb main only
Pull mode is hard-wired to the hyphen-pb clone and `MobileSource/HomeFront` (the `[repo]` argument in the
script header is ignored). Upstream commits on `homefront`, `flexkit` or GitHub `flexcore` must be merged
by hand; a push-mode run would otherwise reset those clones and rsync over them.

### `--dry-run` burns the FlexKit version
A dry run skips only commit/push. It still runs a real `dotnet pack` into `Deploy/local-packages` and
the full staged build, which restores that version into `~/.nuget/packages`. Change FlexKit after a dry
run and the real run repacks the same version against a cache holding the dry run's bytes.
**Fix:** bump before the dry run; bump again if FlexKit changes after it.

### The script reports success that did not happen
- Unknown arguments are ignored: `--dryrun`, `--with-r1-uat` → a REAL push. (`pre_deploy_gate.sh`
  rejects unknown arguments.)
- A login-guard, staged-build or index-gate failure skips only that repo (`✗ Skipping <repo> due to …` or
  `✗ <repo>: … — not pushing`); the others still push and the run ends `▸ Done.` with exit 0. A rejected
  `git push`, by contrast, aborts the whole run under `set -e`. With `--with-r1uat`, R1-UAT runs AFTER
  `▸ Done.`, and any failure there exits 1.
- `✓ Reset to origin/main (remote had new commits)` fires whenever local `main` ≠ `origin/main` — also for an
  unpushed local commit, so it is not proof that upstream moved.
- FlexKit is repacked only when a `.cs/.razor/.css/.js/.csproj` file is newer than the feed nupkg: a changed
  image, json or README under the same version ships the STALE package (the gate blocks and asks for a bump).
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
`pack_flexkit` runs on each pass — main, then R1-UAT with `--with-r1uat` — and repacks the same version if
FlexKit source got newer in between, while the staged build restores the first pack from the cache.
Pack-once-per-run is proposed, not applied.
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
The hold creates its own trap: packing the LIVE FlexCore tree hours later publishes whatever other
sessions changed meanwhile. `publish_flexcore.sh` packs the snapshot actually pushed to GitHub instead.

### Silent success in shell
- `grep -c … || echo 0` is **fail-open**: zero matches prints "0" AND exits 1, so the fallback appends
  a second "0", the integer test errors, and `if` treats it as false. The R1-UAT lockdown check did
  this for weeks. Count with `grep -o … | wc -l`.
- `--dry-run` used to push R1-UAT for real (fixed 09-12). Still verify remote heads after any run.
- A failed hyphen-pb build skips only that repo; the others still push and the run ends "Done."
  Read the log for `✗` and missing `main -> main` lines.
- `checkout -- .` restores the worktree FROM the index, so staged strips survive a dry run — the
  script now resets `--hard`.
- `producer | grep -q …` under `set -o pipefail` fails at RANDOM on large output: `grep -q` exits at the first
  match, the producer gets SIGPIPE, and the pipeline's status is 141. Capture into a variable or use a
  here-string (`grep -q … <<< "$x"`).

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
- zsh ties the names `path`, `fpath`, `status`, `argv` to shell internals: `local path=…` inside a function
  sourced into zsh wipes `$PATH` and every command in it becomes "command not found". Run scripts with `bash`.
- macOS `/var` is a symlink to `/private/var`. A .NET project under `mktemp -d` (/var/folders/…) makes the
  Razor generator see two roots, fall back to bare file names, and fail on duplicates
  (`Editor/TextAreaControl.razor` vs `TextAreaControl.razor`). Use `$(cd "$(mktemp -d)" && pwd -P)`.

### `git rev-parse` echoes what it cannot resolve
`git rev-parse <sha>:<missing path>` prints the argument itself on stdout, so `x=$(git rev-parse …)` is
never empty and every "does this file exist at that commit?" test silently passes. Use
`git rev-parse --verify --quiet …`, which prints nothing on failure.

### Reflection harnesses
`verification/**` is excluded from `HomeFront.csproj`, so API/signature changes build clean and then
throw at runtime. Run them. A failing harness is often stale against an intended behaviour change —
e.g. WizardChecks needed `_gridReady` seeding (09-10) and a debounce wait for filter-row input; on 09-17
three were stale at once — Irfan's Duplicate Model intro step (HHM-977), our own HHM-1062 vendor box
(committed 09-16 without its harness) and MR !35's design-time gItems columns. Check `git log` on the
page: a teammate's or our own recent commit that changed the behaviour the harness asserts means stale.
A harness that reads `appsettings.json` via `Directory.GetCurrentDirectory()` (PricingWorksheetChecks)
must run from the app directory — `run_harnesses.sh` does that; by hand it throws FileNotFoundException.
`Sdk="Microsoft.NET.Sdk.Web"` projects under `verification/` (InputDialogBrowserChecks) are browser-driven
fixture HOSTS, not console harnesses: started bare they never exit or fail to bind :5000 (AirPlay owns it).
`run_harnesses.sh` skips and lists them; run one by hand per its README.

### Published FlexCore SourceLink points at commits GitHub does not have
The nuspec of each published FlexCore names the LOCAL FlexCore commit (e.g. 0.2.43 → 6d8d7bf), but
GitHub `main` is deploy-snapshot history (5be387c), so the SourceLink commit does not exist there. Local
tracking refs also look diverged for the same reason (FlexCore `main…github/main`, app `main…origin/main`
on the personal wchaudhary remote) — that is not drift; compare against the staging clones instead.
Packages published with `publish_flexcore.sh` are built from the staging clone, so their SourceLink commit is
the GitHub snapshot and does exist there.

### Secrets that are already pushed
`App_Data/jira-settings.json` carries a live Atlassian API token and is tracked on hyphen-pb main,
R1-UAT and homefront main (the script used to `git add -f` it; stopped 09-12). Rotation is a human
action in Atlassian — still OPEN. Never print the value.

### Security finding still open (owner decision)
The QA login page shows a live "Security Options (Dev)" panel; `ISecurityOptionsService` is a
singleton, so one anonymous "Deselect All" disables password checks for every user until the pool
recycles. `Security__HideDevPanel=true` on the app pools closes it (and the R1-UAT "Deselect All"
gap) but ends the owner's ADMIN sign-in shortcut, so it has not been done.

<!-- APPEND to reference/rules-and-traps.md. Also amend the page header (line 3): the range is now
     2026-08-24 to 2026-09-26. -->

## Traps learned 2026-09-17 … 2026-09-26

### A held NuGet publish keeps absorbing later work
FlexCore 0.2.48 went to GitHub on 09-23 17:20 (`4566b0e`) with the publish held until five
Crystal-reader defects were fixed (history.md:10, :76-86). FlexCore then changed three more times
(09-25 `344af31`, `2587807`, `0a473d8`; 09-26 `fa8011c`) and was pushed again on 09-26 under the **same
version number** (`f4de318`). `publish_flexcore.sh` packs the pushed snapshot rather than the live tree,
but the snapshot moved: publishing 0.2.48 now ships the 09-26 content, not the content the hold was
judged on. Two mechanics keep that invisible and make the hold unresumable:
- `scripts/publish_flexcore.sh:40` requires the snapshot's `<Version>` to equal the live
  `FlexCore.csproj` version — so the held number must not be bumped while held.
- `scripts/publish_flexcore.sh:32` refuses a gate state older than 720 minutes
  (`gate state is $age min old — run the full procedure (gate, deploy, verify) first`). A hold that
  outlasts 12 hours can only be released by a whole fresh gate → deploy → verify round.

**Tell:** a history row reading `(0.2.48 NOT published — still held)` with a later deploy row under the
same version. **Fix:** record in the history row the GitHub sha the hold was judged on; at release time
either re-review `git -C /Users/wadood/projects/VBToCSharp/HomeFront/Deploy/repos/flexcore log --oneline
<hold-sha>..origin/main`, or bump and publish the current snapshot as a new version. Read the script's
last line too: step 6 is fail-open — no nuget.org listing within 6 minutes prints `! pushed, but
nuget.org did not list …` and still exits 0 (`:90`; hit 09-23 08:56, listed 09:04).

### FlexCore's source repo can push straight to GitHub, so GitHub main is not only deploy snapshots
The existing "FlexCore on GitHub is a deploy mirror" entry no longer holds on its own.
`/Users/wadood/projects/VBToCSharp/FlexCore` has its own `github` remote and its `main` tracks
`github/main`; `git -C /Users/wadood/projects/VBToCSharp/FlexCore reflog show github/main --date=iso`
records direct pushes on 2026-09-09 (`0450fc3`) and 2026-09-25 (`2587807` 12:24, `0a473d8` 22:12), each
preceded by a `Merge github/main into main` — which is what turns the otherwise rejected push into a
fast-forward. `2587807`'s body: it "Restored docs/tests that deploy commits had removed on github/main
so cloud clones keep the full tree". The 09-26 deploy stripped `docs/` and `tests/` again. GitHub main
now interleaves both kinds of commit:

```
f4de318 Deploy flexcore — 2026-09-26 02:18
0a473d8 fix: clip page-header subreports and soft-fail print formulas
2587807 Merge github/main into main; keep local Crystal/R2-QA and tests
344af31 Crystal formula engine rewrite, native RPT recovery, R2-QA grid/dropdown mirrors
4566b0e Deploy flexcore — 2026-09-23 17:20
```

`homefront` and `flexkit` origin/main are still pure `Deploy <repo>` history (last ten commits of each).
**Tell:** the flexcore staging clone will not fast-forward; `git -C …/Deploy/repos/flexcore reflog main
--date=iso | grep reset:` shows a manual reset after each direct push (2026-09-06, 09-09, and 2026-09-26
01:15 → `0a473d8`). **Fix:** confirm the commits are ours, then `git -C
/Users/wadood/projects/VBToCSharp/HomeFront/Deploy/repos/flexcore reset --hard origin/main` before the
run. Whether GitHub main should carry `docs/` and `tests/` is an open owner decision (history.md:69-70).

### classify_file reads FlexCore's own pushed commits as teammate changes
`unshipped_commits()` treats every non-merge commit after the last `Deploy <repo>` commit as somebody
else's work, filtering on the subject alone — `grep -v '|Deploy '` (`scripts/_common.sh:100-104`;
`last_deploy()` at `:96`). Since flexcore's GitHub main also carries our own source commits, preflight §5
and gate check 2 report `REVERT`/`LOST` on lines those commits themselves superseded.

**Tell:** `upstream commit(s) past our last push` on **flexcore** only, authored `wadood`, with ordinary
feature subjects. **Fix:** verify each sha is in the source repo's own history
(`git -C /Users/wadood/projects/VBToCSharp/FlexCore merge-base --is-ancestor <sha> HEAD`), then ack it.
The 09-26 run did that: 12 of the 17 non-comment lines in
`/Users/wadood/projects/VBToCSharp/HomeFront/Deploy/preflight-ack.txt` now read "FlexCore's own history
(ancestor of source HEAD 0a473d8 == GitHub origin/main), superseded by later commits". It is a
workaround — an acked sha can never flag again. The fix belongs in `_common.sh`: for flexkit/flexcore,
a commit the source repo can reach is ours.

### The running app writes files into the trees the deploy rsyncs
`RUNTIME_DATA_EXCLUDE` (`tools/deploy_to_repos.sh:47-52`) covers `wwwroot/tickets/`, `wwwroot/feedback/`,
`Logs/`, `*.log`; `ENV_CONFIG_EXCLUDE` (`:38-41`) covers `appsettings*.json` and
`App_Data/jira-settings.json`. **`User/` is in neither.** The app writes per-user preference JSON and
imported report definitions there — 17 files now, including other people's `User/Reports/<hash>/…/*.rpt`
and `converted/*.xml`. They are git-ignored locally (`MobileSource/HomeFront/.gitignore:63-64`), rsync
does not read `.gitignore`, so all 17 sit in the working trees of **both** `Deploy/repos/hyphen-pb` and
`Deploy/repos/homefront`, off the branch only because the clone's copy of our `.gitignore` ignores
`User/` and `git add -A` skips them — one `.gitignore` edit away from shipping customer data. Not
hypothetical: `wwwroot/tickets/**` and `wwwroot/feedback/debug_capture.png` were tracked on both app
mains before the 2026-08-24 exclusion.

**Tell:** none — `verify_deploy.sh:97-101` uses `git status --porcelain`, which hides ignored files.
**Check:**

```bash
for r in hyphen-pb homefront; do
  git -C /Users/wadood/projects/VBToCSharp/HomeFront/Deploy/repos/$r ls-tree -r --name-only origin/main \
    | grep -E '^(User/|App_Data/|wwwroot/tickets/|wwwroot/feedback/|Logs/)' \
    | grep -v '^App_Data/jira-settings.json$'
done
```

Only the repo-owned `App_Data/jira-settings.json` may come back. Adding `--exclude='User/'` to
`RUNTIME_DATA_EXCLUDE` and this assertion to `verify_deploy.sh` are both still to do.

### A secret is already on the pushed branches, and a grep secret-gate can pass without scanning
`App_Data/jira-settings.json` is tracked on `hyphen-pb origin/main`, `hyphen-pb origin/R1-UAT` and
`homefront origin/main`, and its `ApiToken` is a 192-character live Atlassian token (verified read-only:
field names and lengths only). Rotation is a human action in Atlassian, still open. No script looks for
it — the file appears in `scripts/` only as repo-owned (`_common.sh:23`) and in a comment in
`publish_flexcore.sh`. The local copy is rewritten by the running app (it stores `LastSyncTime`), so it
always looks edited without being a source change; the gate prunes `*/HomeFront/App_Data`
(`pre_deploy_gate.sh:76-78`), which is why it does not false-block.

The scan you would reach for is itself unreliable: macOS `grep`/ugrep aborts on a long alternation regex
("exceeds complexity limits") and prints **nothing**, so `HITS=$(grep …)` read empty as clean and a push
went ahead unscanned on 2026-09-14 (re-scanned with Python: clean — memory note
`zsh_shell_silent_failure_traps`:50). **Fix:** scan with Python, one pattern per compile, and treat any
tool error as FAIL, never as empty. Never print a matched value — report file, field name, length. Two
committed secrets are outstanding and both need rotation: this `ApiToken`, and the Cognito
`AppClientSecret` in the remote's `appsettings.Development.json` (memory note
`deploy_never_push_env_config`) — one more reason the `appsettings*.json` exclusion never comes off.

### "Not in main" is not "open" — MR state exists only in GitLab
`git ls-remote origin 'refs/merge-requests/*/head'` lists every MR ever opened; open, closed, merged and
merged-into-another-branch look identical. MR !39 was read as open from the refs and taken into main on
09-22 (`afaf1b2`); the owner had closed it on 09-21 because its description asked for a check on a real
division first, and it was reverted on 09-23 (`8e10e6d`). !35, !36 and !38 were closed too and were *not*
reverted, so "closed MR content kept by omission" is an unrecorded state on hyphen-pb main
(history.md:54-57, :105-107; memory note `mr_state_only_from_gitlab`). **Fix:** confirm OPEN in the GitLab
UI first — the owner's Chrome session, or the browser pane once signed in; if neither is reachable, list
the candidates and ask. Report "not in main", not "open", unless the UI said so, and record in the history
row what the UI said — and, if it later proves closed, whether the owner kept or reverted it.

### The outer repo's history row is never pushed
The deploy does not push the outer `HomeFront` repo and step 12 stops at committing the row. That repo
has a remote (`https://github.com/wadoodachaudhary/vb6ToDotNet.git`) and currently reads
`* main cf379af [origin/main: ahead 1]` — the 09-26 history row has never left this Mac. The previous
push was `d17b297` on 2026-09-25 12:24, made by the Crystal session bundling its bench commit, not by a
deploy round; before that, 2026-08-09. **Fix:** push that one branch by hand after recording the row —
`git -C /Users/wadood/projects/VBToCSharp/HomeFront push origin main` — and nothing else: the repo also
holds 17 `claude/*` worktree branches, so `git push --all` is wrong. Whether this belongs in the routine
is an owner decision; write the answer down either way.

### One run can ship different app content to hyphen-pb and homefront
The repos are processed hyphen-pb → homefront → flexkit → flexcore (`tools/deploy_to_repos.sh:510-517`),
each rsyncs the **live** source at its own turn (`:710`), and only hyphen-pb packs FlexKit and runs the
staged build (`:722-729`). That build took 16 minutes on 09-22 (history.md:12) while the gate's quiet
window is 10 (`pre_deploy_gate.sh:20`, `QUIET="${QUIET_MINUTES:-10}"`). A session that resumes editing
after the gate passes therefore lands on homefront, flexkit and flexcore but not on hyphen-pb: the two
app mains diverge for that run, and FlexKit/FlexCore can ship code the staged build never compiled.
`verify_deploy.sh` compares each clone against the gate baseline and never compares the two app branches
with each other. **Fix:** on a long run re-check the source trees' mtimes before `▸ Done.`, and raise the
window (`QUIET_MINUTES=20 bash …/pre_deploy_gate.sh`). A cross-repo app-tree diff — excluding
`local-packages/`, `HomeFront.csproj`, `HomeFront.sln`, `NuGet.Config`, `appsettings*.json` — would be a
real added check.

### A harness or fixture pack encodes a premise the code later disproved
The Crystal bench pack recorded per-report verdicts taken when 52 reports were execution blockers, and
`JavaToCSharp/Reports/blocked-52` exists because of that premise. The 09-24/25 engine work disproved it,
so the pack was rebuilt (`CrystalSamples.Seed build`, then `verify`) and reinstalled at
`/Users/wadood/projects/VBToCSharp/HomeFront/FlexKitTester/Data/CrystalSamples.db` on 09-25 23:45:
530/530 paginate — 220 Rendered, 310 Review, **0 Blocked** (`FlexKitTester/CRYSTAL-REPORTS.md:78-85`;
memory note `crystal_blocked52_resolution`). Until that rebuild every bench number from the old pack
described a world the library no longer lived in. Same shape in C#: on 09-23 a review found
`EstimateChecks` asserting a constant for the disabled-Forecasting guard, so the check could not fail
(history.md:11). **Tell:** expectations generated before the fix they are now used to judge; an assertion
comparing a literal the test itself just set. **Fix:** regenerate the fixture against the library it will
judge, record date and sha in its README, keep the previous pack (the 09-23 one is beside the regenerated
pack in the regenerating session's `crystal-pack` folder), and make assertions read rendered output or
source, never their own input.

### Committing a shared file another session is editing
FlexKit and FlexCore are normally being edited by several sessions during a round, and the deploy must
commit them anyway. Never stage such a file wholesale and never swap one aside: on 2026-09-21 FlexKit
`Grid/GridControl.razor.cs` was copied to HEAD for about 9 seconds while the HHM-1149 session applied its
autofit patch to the same file; nothing was lost only because the snapshot happened to be taken 1 s after
that patch landed (memory note `mutation_test_never_swap_shared_files`). **Fix:** commit an own-hunk
patch — take the HEAD content, replace only your region, `git diff --no-index` (fix the `a/` `b/` paths),
then `git apply --cached --check` and `git apply --cached`. Build HEAD-plus-that-patch in an isolated
copy (`git archive HEAD | tar -x -C "$SP/iso"`) first. If a peer session holds the same file, message it
with the hunk you are committing.

### The scratchpad is wiped on restart
Holds last hours, which is the exposure: the session scratchpad
(`/private/tmp/claude-501/<project>/<session>/scratchpad`) comes back **empty** after a session restart —
on 2026-09-23→24 a multi-agent workflow's private FlexKit workspaces, patches and baseline logs vanished
mid-run (memory note `scratchpad_wiped_on_restart`). Agent transcripts under
`~/.claude/projects/<project>/<session>/subagents/workflows/<run>/agent-*.jsonl` survive; a replay from
them recovered the work byte-exact. **Fix:** while a held round has a long workflow editing files in
scratch, copy patches and tarballs to `~/.claude/projects/<project>/<session>/<name>-backup/` every few
minutes — no `git add` in that loop, it races the agents' index. The skill's own state files are safe:
`.update-repositories-gate.state`, `pull-conflicts.txt` and `preflight-ack.txt` live in
`/Users/wadood/projects/VBToCSharp/HomeFront/Deploy/`, not in scratch.

### Two VB6 copies, same file name, different line numbers
Two live VB6 trees exist and a citation is worthless unless it says which one:

| Copy | Path | Role |
|---|---|---|
| New VB6 | `/Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/{HFEst,HFSystem}/Source` | current desktop source; has features the port does not |
| As-migrated | `/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFrontVB6/{HFEst,HFSystem/Source}` | the snapshot the Blazor migration was built from |

They are not the same file with a different header. `FEstimateItems.frm` is 14263 lines in the new copy
and 14111 in the as-migrated one; `gItems` begins at `:1336` versus `:1327`, so
`FEstimateItems.frm:1389` is `ScrollTrack = 0 'False` in one copy and `MultiTotals = -1 'True` in the
other. Content differs too: the new `FAssembly.frm` filters `tblcategories` on `isnull(inactive,0)=0`
where the as-migrated copy does not. The app's `CLAUDE.md:790` points at the `MobileSource/HomeFrontVB6`
copy; `~/.claude/CLAUDE.md` points at the project-root one. **Tell:** a quoted line whose text does not
match the citation; two notes citing different lines for one behaviour. **Fix:** quote the line's text
with its full path, not a bare `frm:NNNN`, and name the copy. Behaviour the port must match is the
as-migrated copy unless the work is explicitly re-migration (memory note `vb6_source_original_vs_new` has
the diff scope; its `HomeFrontVB6/Original/…` paths no longer exist on disk).

### Design-time `.frm` / `.frx` values are real VB6 state, not placeholders
Design-time control definitions may be reproduced literally; runtime-derived values may not
(AGENTS.md §1). The sharpest case: a **blank** grid caption is not a missing caption, it is what hides
the column — `FColumns.frm:367` lists a column only when `.ColKey(i) <> "" And .TextMatrix(0, i) <> ""`.
Inventing captions for `RoundTo` and `Seq` on FEstimateItems.gItems put two columns legacy never offers
into Choose Columns and wrote those captions into four `AppGridLayout` rows (memory note
`full_parity_is_the_default`; owner 2026-09-20: "We are trying to achieve 100% VB6 parity nothing less or
more … Remove HF inventions"). The same rule removed the invented `[gBillingItems FALLBACK]` marker from
FCreateContract (VB6 `gBillingItems` is design-time only) and set gItems' `ScrollTrack="false"` from the
`.frm` (memory note `r2qa_batch_2026_09_21`). **Citing `.frx`:** it is binary and the `.frm` references
it by hex offset — `FormatString = $"FEstimateItems.frx":433F` at
`HomeFrontVB6/HFEst/Source/FEstimateItems.frm:1388`. An `frx:3665` citation is a byte offset, not a line;
quote the extracted text beside it.

## Owner rules (binding, with dates) — addendum 2026-09-17 … 2026-09-26

| Rule | Since |
|---|---|
| Grid filter / search boxes: search-as-you-type is the DEFAULT; commit-only (Enter/Tab) is opt-in via `FilterSettings.SearchAsYouType = false`. Reverses the 09-17 commit-only default. | 2026-09-18 |
| Limit dialog boxes on the web: wizards keep their steps and pickers inside the wizard page, never a nested dialog; Work Reassignment starts on its inline vendor list. (AGENTS.md §2) | 2026-09-19 |
| HomeFrontPB is frozen, released and deployed: no edits, no sync into it, no build, no redeploy. Supersedes the older "keep it synchronised" instructions. (AGENTS.md §4, §5) | 2026-09-19 |
| Mass Change's Choose-Action page carries no pickers, so Next stays enabled where VB6 disables it for Add/Remove/Substitute. | 2026-09-22 |
| Cost Forecasting stays disabled, Field PO Requests stays descoped, and the TBD wizard keeps Ctrl+Delete. VB6 does the opposite in all three; do NOT "restore parity" — EstimateChecks and WizardChecks pin them, so a failure there means someone re-added it. | 2026-09-23 |
| Reports → Manage ships with its known defects: "just ship it" — report work is not QA's focus this round. | 2026-09-23 |
| Hold the FlexCore NuGet publish until the Crystal-reader defects are fixed. 0.2.48 is on GitHub, unpublished. | 2026-09-23 |
| Take MR !41 in the next update, reviewed; do not merge it in GitLab first — main is QA. | 2026-09-23 |
| The FlexCore Crystal fix `0a473d8` was ported INTO FlexKit on owner request, confirmed 09-26. FlexCore → FlexKit still needs that explicit ask every time. | 2026-09-26 |
