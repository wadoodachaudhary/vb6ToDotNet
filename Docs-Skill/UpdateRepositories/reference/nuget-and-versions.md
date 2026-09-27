# Versions and NuGet

FlexKit and FlexCore both carry a `<Version>`, and in both cases a number can be spent — but by two
different mechanisms with two different remedies. Every number below was verified on 2026-09-26 and comes
with the command that re-checks it, because they move every round.

### Today's numbers

| | Value | Re-check |
|---|---|---|
| FlexKit `<Version>` | 0.1.107 (FlexKit.csproj:14), HEAD `36885e8` on `telerik-parity-20260904` | `grep '<Version>' /Users/wadood/projects/VBToCSharp/FlexKit/FlexKit.csproj` |
| FlexKit burned versions | 100 extracted in `~/.nuget/packages/flexkit`, highest **0.1.107** (dir dated 2026-09-26 02:15) → next is **≥ 0.1.108** | `ls ~/.nuget/packages/flexkit \| sort -V \| tail -3` |
| FlexKit on nuget.org | not published — HTTP 404. `Deploy/local-packages` is its only ship route (history.md:24) | `curl -s -o /dev/null -w '%{http_code}\n' https://api.nuget.org/v3-flatcontainer/flexkit/index.json` |
| FlexCore `<Version>` | 0.2.48 (FlexCore.csproj:21), HEAD `fa8011c` on `main` | `grep '<Version>' /Users/wadood/projects/VBToCSharp/FlexCore/FlexCore.csproj` |
| FlexCore on nuget.org | 48 versions, latest **0.2.47**. 0.2.48 NOT published — held since 09-23 for the Crystal-reader defects (history.md:9-10) | `curl -s https://api.nuget.org/v3-flatcontainer/flexcore/index.json` |
| Satellites | FlexCore.Documents 0.2.35, FlexCore.Llm 0.2.42 — both 404 on nuget.org, both on GitHub main | same URL with `flexcore.documents` / `flexcore.llm` |
| Feed | `HomeFront/Deploy/local-packages/FlexKit.0.1.107.{nupkg,snupkg}`, 2026-09-26 02:15 | `ls -l /Users/wadood/projects/VBToCSharp/HomeFront/Deploy/local-packages` |

nuget.org has no FlexCore 0.2.34 or 0.2.40 (never published; 0.2.34 exists only in the local cache). The
two libraries' versions are independent and may diverge (owner 2026-07-25, memory `flexcore_nuget_publish`).

### FlexKit is burned by the local NuGet cache

`dotnet restore` resolves `PackageReference FlexKit X` from `~/.nuget/packages/flexkit/X/` and treats it as
immutable. Repacking X refreshes only the feed and the branch, while the staged build in the clone
(`verify_build`, deploy_to_repos.sh:492-504 — `rm -rf bin obj`, then `--no-incremental`) verifies the OLD
extracted bytes.

```bash
[ -d ~/.nuget/packages/flexkit/0.1.107 ] && echo BURNED || echo fresh
```

Never delete from `~/.nuget/packages` — GhostWriter, GhostWriterLLM and other consumers share it. The
remedy is always a bump. `--dry-run` really packs and really restores, so it burns the number too: bump
before the dry run, and again if FlexKit changes after it.

### FlexCore is burned by nuget.org

A published version can be unlisted but never replaced, and FlexCore's consumers are strangers. Nothing
local records this — nuget.org is the only authority:

```bash
curl -s https://api.nuget.org/v3-flatcontainer/flexcore/index.json | python3 -m json.tool | tail -8
```

Only `preflight.sh` asks that question before a push (preflight.sh:161-185, verdicts `PUBLISHED` /
`LOWER` / `OK`), and it runs at step 1 — possibly hours earlier. `pre_deploy_gate.sh` never validates the
FlexCore version; it only copies it into the gate state (pre_deploy_gate.sh:118). So a run can push a
`Deploy flexcore` snapshot whose `<Version>` is already on nuget.org, pass `verify_deploy.sh`, and fail
only at `publish_flexcore.sh:53` — after which the whole deploy must be redone with a bumped version.
Re-run the preflight version check immediately before the gate.

### When a bump is required

- **Any FlexKit code change** (`.cs .razor .css .js .csproj`) while the current number is burned.
- **An asset-only FlexKit change** (image, json, README) under an unchanged version: the script repacks
  only on those five extensions (deploy_to_repos.sh:203-207), so the STALE package ships. The gate blocks
  (pre_deploy_gate.sh:105, "the script only repacks on code changes — bump FlexKit.csproj to force a repack").
- **A repack under a burned number**: the gate blocks (pre_deploy_gate.sh:99, "would be REPACKED but …
  already extracted in ~/.nuget/packages").
- **Whenever FlexKit's public API grew**, even if the mtime guard says the feed is current. The guard is
  `find -newer` against the nupkg and loses races with a concurrent session's repack: on 2026-09-01 it
  printed "0.1.79 nupkg is current" while `WrapRowsUntilEdge` existed only in source, the staged build
  failed with two errors, and hyphen-pb plus R1-UAT were skipped (memory
  `feedback_bump_flexkit_before_deploy`). A version mismatch forces the repack deterministically.
- **FlexCore**: bump above nuget.org's latest if FlexCore changed and you intend to publish.
  `FlexCore.Documents` and `FlexCore.Llm` carry their own versions, are never published, and are not
  bumped by a FlexCore bump.

### How FlexKit reaches the deployed app

The working tree uses a ProjectReference; only the hyphen-pb branch uses a package. The swap happens in the
staging clone, never in a source tree, and only for hyphen-pb main (plus R1-UAT under `--with-r1uat`,
deploy_to_repos.sh:801-802):

1. **Pack** — `pack_flexkit` (deploy_to_repos.sh:175-234) reads `<Version>` from FlexKit.csproj and
   repacks into `HomeFront/Deploy/local-packages/` when no nupkg is present, when the feed's version
   differs, or when a tracked code file is newer than the nupkg; then copies `.nupkg` + `.snupkg` into
   `<clone>/local-packages/`. The feed is git-ignored and machine-local — it moved out of
   `HomeFrontPB/local-packages` on 2026-09-13 (`e1aae92`), so nothing is committed in the frozen PB tree.
2. **Swap** — `swap_to_package_ref` (deploy_to_repos.sh:237-281) replaces the FlexKit
   `<ProjectReference …/>` with `<PackageReference Include="FlexKit" Version="$ver" />` (:249-250), drops
   `wwwroot/images/32/*` icons FlexKit also ships **only** where the app has no copy of its own (deleting
   an app-owned copy silently blanks the icon — fixed 2026-08-27), and deletes the FlexKit project block
   from `HomeFront.sln`.
3. **Staged build** — `verify_build` must succeed or that repo is skipped and not pushed
   (deploy_to_repos.sh:725-728).
4. **Restore source** — the branch-owned `NuGet.Config` clears all sources and pins
   `<package pattern="FlexKit" />` to `./local-packages`. The rsync excludes both `NuGet.Config` and
   `local-packages/` (deploy_to_repos.sh:72-73), so that file exists only on the branch and the local app
   tree has none; lose it there and the staged build asks nuget.org for FlexKit and gets a 404. Its comment
   block still talks about HomeFrontPB and calls the local source "idle" — stale prose on a repo-owned
   file; leave it alone.

What the pushed branch looks like today (hyphen-pb `origin/main`, tip `4a6d61f`):

```bash
cd /Users/wadood/projects/VBToCSharp/HomeFront/Deploy/repos/hyphen-pb
git show origin/main:HomeFront.csproj | grep -n FlexKit   # 61: <PackageReference Include="FlexKit" Version="0.1.107" />
git ls-tree --name-only origin/main local-packages/       # FlexKit.0.1.107.nupkg + .snupkg
git show origin/main:HomeFront.sln | grep -c FlexKit      # 0
```

`homefront` main is different by design: `HomeFront.csproj:61` keeps
`<ProjectReference Include="..\..\..\FlexKit\FlexKit.csproj" />`, with no `local-packages` and no
`NuGet.Config`, so that repo cannot be built on its own. `flexkit` and `flexcore` get no pack, no swap and
no staged build. `verify_deploy.sh:78-81` proves the branch's PackageReference version equals the nupkg
filename on the branch — a version-string match, which catches a silent pack failure but not stale bytes
inside a correctly named package. That is what the bump rule is for.

### How FlexCore reaches GitHub and nuget.org

Both routes go through the staging clone `HomeFront/Deploy/repos/flexcore`, whose `origin` is
`https://github.com/wadoodachaudhary/FlexCore.git`:

- **GitHub** — the deploy rsyncs the working tree into the clone and commits `Deploy flexcore — <date>`.
  Verified on `origin/main` today: `FlexCore.Llm/`, `FlexCore.Documents/` and `README.md` are there;
  `tests/`, `docs/`, `CLAUDE.md`, `AGENTS.md` and `PUBLISHING.md` are not (stripped every run —
  deploy_to_repos.sh:97,311). Owner 2026-09-14 on FlexCore.Llm: "push it to Git but not to NuGet."
- **nuget.org** — `publish_flexcore.sh` packs a detached worktree of that clone at `origin/main`
  (publish_flexcore.sh:57-62), never the live FlexCore tree, so the published SourceLink commit exists on
  GitHub.

The FlexCore working repo's own `github/main` tracking ref is not the authority — it reads `0a473d8` today
while the clone's `origin/main` is `f4de318`. The publish also requires the GitHub tip's subject to start
with `Deploy flexcore` (publish_flexcore.sh:29), so if a session pushes the FlexCore source repo straight
to GitHub after a deploy — as happened before the 09-26 round — the publish dies until a fresh deploy
re-snapshots it.

### What the FlexCore package contains

`FlexCore.csproj` excludes three trees from its own globs, each with the reason in a comment, because the
project sits at the repo root and the SDK globs would otherwise pull the satellites and their dependencies
back in: `tests/**` (:5), `FlexCore.Documents/**` (:10), `FlexCore.Llm/**` (:13). `publish_flexcore.sh`
re-checks the built package independently: `lib/net10.0/FlexCore.dll` and `README.md` must be present
(:67-68), and any match of `FlexCore.Llm|CLAUDE.md|AGENTS.md|docs/|PUBLISHING.md|tests/` fails the publish
(:69-70).

The dependency split is deliberate: ClosedXML lives in FlexCore.Documents, "not in FlexCore … FlexCore's
own export path uses the dependency-free `Fx.ControlKit.Excel.XlsxWriter` instead"
(FlexCore.Documents.csproj:38-42), and FlexCore grants it `InternalsVisibleTo` (FlexCore.csproj:57) rather
than widening its public API. FlexCore references no ClosedXML at all (its package refs are Ical.Net,
ZXing.Net, AngleSharp and two Microsoft ones, :61-65); FlexKit needs no split because it carries ClosedXML
itself (FlexKit.csproj:50).

⚠ FlexCore's shipped `<Description>` (FlexCore.csproj:25) still advertises "native ClosedXML Excel export",
untrue since the Documents split. It is package metadata on nuget.org, so the correction only reaches
users with the next publish.

### "Permanent" also means the public API

A git push can be followed up; a nuget.org version can only be unlisted. Three consumer-facing things the
push does not cover:

- **Binary breaks have shipped in patch bumps.** Recorded identically in both libraries' CLAUDE.md
  (FlexCore/CLAUDE.md:173, FlexKit/CLAUDE.md:182): "ReportVectorShape's positional constructor gained a
  parameter (binary break vs 0.2.44)". Nothing enforces SemVer here — versions have only ever moved
  0.2.x → 0.2.x+1. Say in the history row whether a release breaks binary or source compatibility.
- **Published FlexCore is incomplete on purpose.** "FlexCore NuGet users need FlexCore.Documents
  (unpublished) for the Crystal PDF preview" and "trimmed hosts should pass `PdfViewerComponentType`
  explicitly" (CLAUDE.md:173/182 and :181/190) — an external consumer cannot obtain Documents at all.
  Blazor Server's 32 KB hub limit is another asymmetry: "HomeFront raises it; FlexCore.Showcase,
  FlexKitTester and NuGet consumers do not" (FlexCore/CLAUDE.md:206, FlexKit/CLAUDE.md:215).
- **A published defect is permanent.** FlexCore 0.2.41 (09-13) shipped a reverted "Multi-sort" caption
  because the review landed after the publish (history.md:17, lesson 4). Every round since has held the
  publish until reviews land.

### The publish command and what it refuses

The key is the macOS login-Keychain item `flexcore-nuget-key`, read inside the command substitution so it
is never echoed or written to disk (owner 2026-08-11; memory `flexcore_nuget_publish`). publish_flexcore.sh:77-78
runs exactly this:

```bash
dotnet nuget push "$nupkg" \
  --api-key "$(security find-generic-password -s flexcore-nuget-key -w)" \
  --source https://api.nuget.org/v3/index.json
```

If the key is lost, regenerate it at <https://www.nuget.org/account/apikeys> with scope "Push new packages
and package versions", glob `FlexCore`.

```bash
bash Docs-Skill/UpdateRepositories/scripts/publish_flexcore.sh --check   # every step except the push
bash Docs-Skill/UpdateRepositories/scripts/publish_flexcore.sh
```

It fails closed on: the flexcore clone not on `main` (:26); unpushed commits in it (:27); a GitHub tip
whose subject is not `Deploy flexcore…` (:29); gate state missing or older than 720 min (:30-34); a
snapshot `<Version>` differing from the live FlexCore.csproj (:37-40); a version already published or not
above the latest (:41-54); a pack failure (:58-60); a missing `FlexCore.dll` or `README.md`, or a forbidden
file in the package (:67-70); a push whose output lacks "Your package was pushed" (:80). `--check` stops
before the push (:73) and downgrades the version verdict to a `!` warning; it is the only accepted flag —
anything else exits 2 (:13). There is no `--skip-duplicate`: a duplicate means the version was not bumped,
which is a failure. One rough edge: when nuget.org is unreachable the version check returns `UNKNOWN` and
the script dies with the misleading "bump FlexCore.csproj" wording — check the network before believing it.

### A held publish is not resumable

`publish_flexcore.sh:32` dies when `Deploy/.update-repositories-gate.state` is older than 720 minutes ("run
the full procedure (gate, deploy, verify) first"); `verify_deploy.sh:25` already fails past 180. FlexCore
0.2.48 is in that position now: the state file reads `time=1790403426` = 2026-09-26 02:17 EDT, 1218 minutes
old as this page was written, so the held publish can only be completed by re-running the gate, the deploy
and verify.

A held version also keeps accumulating content. 0.2.48 was pushed to GitHub as `4566b0e` on 09-23 17:20
with the publish held; FlexCore changed three more times and was pushed again as `f4de318` on 09-26 02:18
under the **same** version number — `git diff --shortstat 4566b0e f4de318` in the flexcore clone is 38
files, 4969 insertions, 691 deletions. Publishing 0.2.48 now would publish the 09-26 content, not the
09-23 content the hold was judged on. The number cannot be bumped while held either, because :37-40
requires the snapshot's version to equal the live csproj's — which is what keeps the drift invisible.

So: record the GitHub sha a hold was judged on. At release time either re-review the delta from that sha to
the current snapshot, or bump and publish the current snapshot as a new version. Decide before the push
whether the hold is likely to outlast 12 hours.

### Indexing can outlast the script's wait

After a successful push the script polls nuget.org 12 × 30 s, then prints `! pushed, but nuget.org did not
list <ver> within 6 minutes — check https://www.nuget.org/packages/FlexCore before recording it` and still
**exits 0** (publish_flexcore.sh:84-91). That is not a failure — it fired for real on 09-23, when 0.2.47
was pushed at 08:56 and listed at 09:04 (history.md:11). Re-check the flat-container index a few minutes
later and only then write "published, listed HH:MM" in the history row; a second small commit for that is
normal (outer repo `fe40c9d`, 09-23 09:04).

### Stale documents on disk

`FlexCore/PUBLISHING.md` (last touched `bea21b5`, 2026-08-13) contradicts every current rule: it packs the
LIVE tree into `./nupkg` (:60-61), pushes with `--skip-duplicate` (:78), reads the key from
`~/.flexcore-nuget-key` (:16,:76), and claims HomeFrontPB "now references FlexCore via NuGet — no longer a
`<ProjectReference>`" (:17) when HomeFrontPB.csproj:70 references **FlexKit** by ProjectReference.
`FlexKit/PUBLISHING.md` (`b6594ee`, 2026-06-21) is the same document, still titled "Publishing FlexCore",
although FlexKit is never published to nuget.org. Both are stripped from every push and blocked from the
package, so they mislead only local sessions — this skill's own audience. **`publish_flexcore.sh` and this
page are canonical; both PUBLISHING.md files are historical and wrong.**

⚠ Correction to memory: `flexcore_nuget_publish` records the plaintext key as removed, but what was removed
is `~/.nuget-api-key`. `~/.flexcore-nuget-key` still exists (46 bytes, mode 600, dated 2026-08-12), which
is why the stale document's `--api-key "$(cat …)"` line would still work. Never `cat` it; the Keychain item
is canonical. Removing the leftover is an owner decision, not a deploy step.
