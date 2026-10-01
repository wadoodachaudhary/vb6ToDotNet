#!/bin/bash
# deploy_to_repos.sh — Push source to deploy repos / pull MR changes back.
#
# PUSH (default):
#   Source directories → staging repos → commit & push.
#   hyphen-pb gets FlexKit packed as nupkg + ProjectReference swapped.
#
# PULL (--pull [repo]):
#   Fetch remote changes (from other devs' MRs) into staging,
#   then rsync code changes back to source. Only hyphen-pb is
#   bi-directional; the other repos are push-only.
#
# Usage:
#   ./tools/deploy_to_repos.sh              # push all repos
#   ./tools/deploy_to_repos.sh --dry-run    # push dry run
#   ./tools/deploy_to_repos.sh --with-r1uat # ALSO deploy R1-UAT (frozen by default)
#   ./tools/deploy_to_repos.sh --pull       # pull MR changes from hyphen-pb
set -euo pipefail

ROOT="/Users/wadood/projects/VBToCSharp"
HF_ROOT="$ROOT/HomeFront"
DEPLOY="$HF_ROOT/Deploy/repos"
# The packed FlexKit feed lives beside the staging clones (outer .gitignore: /Deploy/).
# It used to be HomeFrontPB/local-packages, so every FlexKit change rewrote files in the
# frozen HomeFrontPB tree (AGENTS.md: never modify HomeFrontPB unless asked by name).
# Each staged clone still receives its own copy in <clone>/local-packages.
PKG_DIR="$HF_ROOT/Deploy/local-packages"
FLEXKIT_SRC="$ROOT/FlexKit"
FLEXKIT_CSPROJ="$FLEXKIT_SRC/FlexKit.csproj"

# Environment config is owned by the REMOTE, never by this machine. main carries
# Azure token placeholders (#{Database.Server}# etc.), the per-environment
# Auth.CookieName, Serilog sinks, InternalTools allow-list and Cognito settings;
# a local dev tree carries localhost/sa and a stripped logging stub. Pushing the
# local copies would silently break every pipeline deployment, so they are held
# back and whatever `git reset --hard origin/main` restored is what ships.
# Owner directive 2026-08-20.
ENV_CONFIG_EXCLUDE=(
    --exclude='appsettings*.json'
    --exclude='App_Data/jira-settings.json'
)

# Runtime data must never ship in a repo. Tickets are per-environment feedback
# data the app writes itself (~430MB across both apps); logs are Serilog output.
# The Azure pipelines carry a matching -skip rule for wwwroot/tickets so a deploy
# does not delete the server's copies. Owner directive 2026-08-24.
RUNTIME_DATA_EXCLUDE=(
    --exclude='wwwroot/tickets/'
    --exclude='wwwroot/feedback/'
    --exclude='Logs/'
    --exclude='*.log'
)

# Not needed by a Windows IIS deployment: AI/editor tooling, static-analysis
# config, internal notes, and macOS cruft. Owner directive 2026-08-24.
NON_WINDOWS_EXCLUDE=(
    --exclude='.agents/'
    --exclude='qodana.yaml'
    --exclude='HomeFrontDemoTalkingPoints.md'
    --exclude='dev-review-tickets.md'
    --exclude='__MACOSX/'
    --exclude='._*'
    --exclude='*.command'
    --exclude='*.mobileconfig'
)

# Deployment plumbing that lives in the REPO, not in the HomeFront source tree.
# HomeFront has none of these, so without the exclusion rsync --delete would strip
# CI/CD and the FlexKit package feed on the first push. Owner directive 2026-08-27.
REPO_OWNED_EXCLUDE=(
    --exclude='azure-pipelines*.yml'
    --exclude='NuGet.Config'
    --exclude='local-packages/'
    --exclude='certs/'
)

# R1-UAT carries a deliberate UAT lockdown that HomeFrontPB source does not:
# five hard-disabled security toggles in FLogin plus its own open.ico. Excluding
# them means the rsync leaves the branch's copies alone. Owner directive 2026-08-27.
R1UAT_PRESERVE_EXCLUDE=(
    --exclude='Components/Pages/FLogin.razor'
)

# Not needed to build or run the apps / libraries: AI-agent instruction files,
# test and verification harnesses, internal notes. Owner 2026-09-13: "delete codex,
# claude files and other files not needed for HomeFront to run when deploying (keep
# them locally if needed) -- remove them from repositories as well."
# Leading "/" anchors each pattern to the repo root, so nothing deeper can match.
# Two layers, because an rsync --exclude neither copies NOR deletes: this list keeps
# new copies out, strip_non_runtime_files() deletes copies already on the branch.
# KEPT on purpose: README.md (packed into the FlexKit/FlexCore nupkgs — pack fails
# without it), LICENSE files (vendored third-party code), .gitignore, build/CI files.
NON_RUNTIME_EXCLUDE=(
    --exclude='/CLAUDE.md'
    --exclude='/AGENTS.md'
    --exclude='/GEMINI.md'
    --exclude='/PUBLISHING.md'
    --exclude='/Data/DataControl.API.md'
    --exclude='/.agents/'
    --exclude='/.claude/'
    --exclude='/.codex/'
    --exclude='/.codex-backups/'
    --exclude='/verification/'
    --exclude='/tests/'
    --exclude='/docs/'
    --exclude='/Docs/'
    --exclude='*.vb6_gap.txt'
)
RSYNC_EXCLUDE=(
    "${NON_RUNTIME_EXCLUDE[@]}"
    "${ENV_CONFIG_EXCLUDE[@]}"
    "${RUNTIME_DATA_EXCLUDE[@]}"
    "${NON_WINDOWS_EXCLUDE[@]}"
    "${REPO_OWNED_EXCLUDE[@]}"
    --exclude='.git'
    --exclude='.claude'
    --exclude='.codex-backups'
    --exclude='.idea'
    --exclude='.DS_Store'
    --exclude='bin/'
    --exclude='obj/'
    --exclude='Logs/'
    --exclude='.vs/'
    --exclude='.vscode/'
)

# PULL_EXCLUDE_PB retired 2026-08-31. Pull mode no longer targets HomeFrontPB —
# that tree is FROZEN and is a source only, feeding the R1-UAT branch. Incoming
# MRs on main now come back to MobileSource/HomeFront, and pull mode selects the
# files the incoming commits touched (see PULL_PATHSPEC in PULL MODE below)
# rather than rsyncing a whole tree.

MODE="push"
DRY_RUN=false
# R1-UAT is FROZEN (owner 2026-09-12: "R1-UAT is frozen"). A push run no longer
# deploys it; --with-r1uat deploys it for that one run, only when the owner asks.
WITH_R1UAT=false
# R2-UAT is the live UAT branch (owner 2026-10-01). --with-r2uat merges the main this run pushed
# into it; see deploy_r2uat. Opt-in per run, like --with-r1uat.
WITH_R2UAT=false
for arg in "$@"; do
    case "$arg" in
        --pull)   MODE="pull" ;;
        --dry-run) DRY_RUN=true ;;
        --with-r1uat) WITH_R1UAT=true ;;
        --with-r2uat) WITH_R2UAT=true ;;
    esac
done

log() { printf "\n\033[1;36m▸ %s\033[0m\n" "$1"; }
ok()  { printf "  \033[32m✓ %s\033[0m\n" "$1"; }
err() { printf "  \033[31m✗ %s\033[0m\n" "$1"; }

# ── Always hand the hyphen-pb clone back on main ────────────────────
# That clone serves TWO branches (main ← HomeFront, R1-UAT ← HomeFrontPB).
# If it is left parked on R1-UAT, the next run's `reset --hard origin/main`
# rewrites the UAT branch instead — which is exactly how the 2026-08-28
# non-fast-forward incident happened. An EXIT trap covers every termination
# path, including a `set -e` abort mid-stage (the RETURN trap inside
# deploy_r1uat does not).
park_hyphenpb_on_main() {
    local rc=$?
    [ -d "$DEPLOY/hyphen-pb/.git" ] || return $rc
    local b
    b=$(git -C "$DEPLOY/hyphen-pb" rev-parse --abbrev-ref HEAD 2>/dev/null || echo "")
    if [ -n "$b" ] && [ "$b" != "main" ]; then
        if git -C "$DEPLOY/hyphen-pb" checkout main --quiet 2>/dev/null; then
            ok "Parked hyphen-pb clone back on main (was $b)"
        else
            err "COULD NOT park hyphen-pb clone on main — it is on '$b'."
            err "Fix before the next deploy, or the main push will rewrite that branch."
        fi
    fi
    return $rc
}
trap park_hyphenpb_on_main EXIT

# ── Pack FlexKit ────────────────────────────────────────────────────
pack_flexkit() {
    local target_dir="$1"
    log "FlexKit: check version and pack" >&2

    local cur_ver
    cur_ver=$(grep '<Version>' "$FLEXKIT_CSPROJ" | sed 's/.*<Version>\(.*\)<\/Version>.*/\1/')
    mkdir -p "$PKG_DIR"
    local existing_nupkg
    existing_nupkg=$(ls "$PKG_DIR/FlexKit."*.nupkg 2>/dev/null | head -1 || true)
    local existing_ver=""
    if [ -n "$existing_nupkg" ]; then
        existing_ver=$(basename "$existing_nupkg" | sed 's/FlexKit\.\(.*\)\.nupkg/\1/')
    fi

    local need_pack=false
    local pack_reason=""
    if [ -z "$existing_nupkg" ]; then
        need_pack=true; pack_reason="no nupkg present"
    elif [ "$existing_ver" != "$cur_ver" ]; then
        need_pack=true; pack_reason="version $existing_ver -> $cur_ver"
    else
        # A version match is NOT proof the package is current. FlexKit source can
        # advance without a version bump (2026-08-29: c9b726b added
        # GridControl.RequestScrollToOrigin and 00a37db ported grid perf work, both
        # after the 0.1.74 bump). The stale nupkg then fails the staged build with a
        # missing-member error that never reproduces locally, because the working
        # tree builds FlexKit by ProjectReference. Repack whenever any tracked source
        # file is newer than the nupkg.
        local newer
        newer=$(find "$FLEXKIT_SRC" -newer "$existing_nupkg" \
                    -not -path '*/obj/*' -not -path '*/bin/*' -not -path '*/.git/*' \
                    \( -name '*.cs' -o -name '*.razor' -o -name '*.css' -o -name '*.js' -o -name '*.csproj' \) \
                    -print -quit 2>/dev/null)
        if [ -n "$newer" ]; then
            need_pack=true
            pack_reason="source newer than nupkg ($(basename "$newer"))"
        fi
    fi

    if $need_pack; then
        log "  Packing FlexKit $cur_ver — $pack_reason" >&2
        (dotnet pack "$FLEXKIT_CSPROJ" -c Release -o /tmp/flexkit-pack) >&2
        rm -f "$PKG_DIR/FlexKit."*.nupkg
        rm -f "$PKG_DIR/FlexKit."*.snupkg
        cp /tmp/flexkit-pack/FlexKit."$cur_ver".nupkg "$PKG_DIR/"
        cp /tmp/flexkit-pack/FlexKit."$cur_ver".snupkg "$PKG_DIR/" 2>/dev/null || true
        rm -rf /tmp/flexkit-pack
        ok "Packed FlexKit $cur_ver" >&2
    else
        ok "FlexKit $cur_ver nupkg is current" >&2
    fi

    mkdir -p "$target_dir/local-packages"
    rm -f "$target_dir/local-packages/FlexKit."*.nupkg
    rm -f "$target_dir/local-packages/FlexKit."*.snupkg
    cp "$PKG_DIR/FlexKit."*.nupkg "$target_dir/local-packages/"
    cp "$PKG_DIR/FlexKit."*.snupkg "$target_dir/local-packages/" 2>/dev/null || true

    echo "$cur_ver"
}

# ── Swap ProjectReference → PackageReference ────────────────────────
swap_to_package_ref() {
    local deploy_dir="$1"
    local ver="$2"
    local proj="${4:-HomeFront}"
    local csproj="$deploy_dir/$proj.csproj"
    local sln="$deploy_dir/$proj.sln"

    python3 - "$csproj" "$ver" <<'PYEOF'
import re, sys
csproj_path, ver = sys.argv[1], sys.argv[2]
text = open(csproj_path, encoding='utf-8').read()
text = re.sub(
    r'<ProjectReference\s+Include="[^"]*FlexKit[^"]*"\s*/>',
    f'<PackageReference Include="FlexKit" Version="{ver}" />',
    text
)
open(csproj_path, 'w', encoding='utf-8').write(text)
PYEOF
    ok "Swapped ProjectReference → PackageReference FlexKit $ver"

    # Only drop an overlapping asset when the app does NOT ship its own copy.
    # Pages reference these app-relatively ("images/32/open.ico"), NOT via
    # _content/, so deleting a copy the app owns silently blanks the icon —
    # every onerror handler just hides it. Fixed 2026-08-27.
    local src_assets="${3:-}"
    local kept=0 dropped=0
    for f in $(ls "$FLEXKIT_SRC/wwwroot/images/32/" 2>/dev/null); do
        if [ -n "$src_assets" ] && [ -f "$src_assets/wwwroot/images/32/$f" ]; then
            kept=$((kept+1)); continue
        fi
        rm -f "$deploy_dir/wwwroot/images/32/$f"; dropped=$((dropped+1))
    done
    ok "Overlapping static assets: $dropped dropped, $kept kept (app ships its own)"

    if [ -f "$sln" ]; then
        python3 - "$sln" <<'PYEOF'
import re, sys
sln_path = sys.argv[1]
text = open(sln_path, encoding='utf-8').read()
text = re.sub(r'Project\([^)]*\)\s*=\s*"FlexKit".*?EndProject\r?\n?', '', text, flags=re.DOTALL)
open(sln_path, 'w', encoding='utf-8').write(text)
PYEOF
        ok "Removed FlexKit from .sln"
    fi
}

# ── The development login-skip must never ship ──────────────────────
# Owner directive 2026-09-12: "make sure the Login page skip is removed and
# taken back, so users login properly" — AND "Login Page skip should remain
# functional on this development device". So the working tree KEEPS
# Services/DevAutoLogin.cs + the app.UseDevAutoLogin() line, and the deploy
# removes them from the STAGED CLONE, before the build check so the build
# proves the stripped tree still compiles.
#
# An rsync --exclude cannot do this job: --exclude means rsync neither copies
# NOR deletes, so the copy already committed on origin/main (it went up in
# 8f03fa3, 2026-09-12) would simply survive untouched in the clone. The strip
# has to delete it, every run.
#
# Files are matched on CONTENT, not on a hard-coded path, so renaming, moving
# or splitting the middleware cannot walk it past the guard. After stripping,
# the guard re-scans and refuses to push on any surviving trace, and asserts
# POSITIVELY that the real password check is still wired up — an empty tree
# would otherwise pass an absence-only check.
# ── Files not needed to build or run never stay on a branch ─────────
# Deletes tracked copies of NON_RUNTIME_EXCLUDE from the staged clone (the rsync
# exclude alone would leave them committed forever), removes untracked leftovers
# of those root folders, and drops .sln project entries that point into them so
# the staged build cannot reference a project that is no longer there.
strip_non_runtime_files() {
    local git_dir="$1"
    python3 - "$git_dir" <<'PYEOF'
import os, re, shutil, subprocess, sys
root = sys.argv[1]
ROOT_FILES = {"CLAUDE.md", "AGENTS.md", "GEMINI.md", "PUBLISHING.md", "Data/DataControl.API.md"}
ROOT_DIRS = [".agents", ".claude", ".codex", ".codex-backups", "verification", "tests", "docs", "Docs"]
SUFFIXES = (".vb6_gap.txt",)
def doomed(rel):
    return rel in ROOT_FILES or rel.split("/")[0] in ROOT_DIRS or rel.endswith(SUFFIXES)
tracked = subprocess.run(["git", "-C", root, "ls-files", "-z"], capture_output=True).stdout.decode("utf-8", "surrogateescape").split("\0")
removed = {}
for rel in tracked:
    if rel and doomed(rel):
        p = os.path.join(root, rel)
        if os.path.lexists(p):
            os.remove(p)
        key = rel.split("/")[0] if "/" in rel else rel
        removed[key] = removed.get(key, 0) + 1
for d in ROOT_DIRS:
    if os.path.isdir(os.path.join(root, d)):
        shutil.rmtree(os.path.join(root, d))
for f in ROOT_FILES:
    if os.path.isfile(os.path.join(root, f)):
        os.remove(os.path.join(root, f))
dirs = "|".join(re.escape(d) for d in ROOT_DIRS)
proj = re.compile(r'Project\([^)]*\)\s*=\s*"[^"]*",\s*"(?:' + dirs + r')[\\/][^"]*"[^\n]*\n.*?EndProject\r?\n', re.S)
for name in os.listdir(root):
    if name.endswith(".sln"):
        path = os.path.join(root, name)
        with open(path, encoding="utf-8", newline="") as fh:
            text = fh.read()
        text2, n = proj.subn("", text)
        if n:
            with open(path, "w", encoding="utf-8", newline="") as fh:
                fh.write(text2)
            removed[name + " project entries"] = n
print("  " + ("; ".join(f"{k} ({v})" for k, v in sorted(removed.items())) if removed else "nothing to strip"))
PYEOF
}

strip_dev_login_bypass() {
    local git_dir="$1"
    local label="${2:-repo}"

    # Not an app tree (flexkit / flexcore) — nothing to strip or assert.
    [ -f "$git_dir/Program.cs" ] || return 0

    if ! python3 - "$git_dir" "$label" <<'PYEOF'
import os, re, sys

staged, label = sys.argv[1], sys.argv[2]
# Every spelling of the bypass: the class, the extension method, the
# configuration key, and the environment-variable form. IConfiguration keys are
# case-INSENSITIVE, so the haystack is lowercased before matching — a
# "devautologin:user" key arms it exactly like the canonical casing.
MARKERS = ("devautologin", "usedevautologin", "devautologin__")
SKIP_DIRS = {".git", "bin", "obj", "node_modules", "local-packages", ".vs", ".idea",
             ".codex-backups", "__MACOSX"}
# Binaries and media cannot be a bypass, and the tree carries large assets
# (screen recordings, .nupkg) that are pointless to read.
BINARY_EXT = {".dll", ".pdb", ".exe", ".nupkg", ".snupkg", ".zip", ".ico", ".png",
              ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".woff", ".woff2", ".ttf",
              ".eot", ".mp4", ".mov", ".avi", ".pdf", ".rpt", ".xls", ".xlsx",
              ".doc", ".docx", ".cache", ".bin", ".so", ".dylib"}
MAX_BYTES = 2 * 1024 * 1024
CODE_EXT = {".cs", ".razor", ".cshtml"}      # the bypass itself — delete
DOC_EXT = {".md", ".txt"}                    # prose about a dev feature — harmless
# Anything else that holds a marker (config, script, CI, extensionless) is
# treated as "could ARM the bypass" and refuses the deploy rather than being
# deleted on a guess: deleting the wrong appsettings or pipeline file would
# break the deployment, and that decision belongs to a human.

def scan():
    hits = {}
    for root, dirs, files in os.walk(staged):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for name in files:
            ext = os.path.splitext(name)[1].lower()
            if ext in BINARY_EXT:
                continue
            full = os.path.join(root, name)
            try:
                if os.path.getsize(full) > MAX_BYTES:
                    continue
                text = open(full, encoding="utf-8", errors="ignore").read().lower()
            except OSError:
                continue
            if any(m in text for m in MARKERS):
                hits[os.path.relpath(full, staged).replace(os.sep, "/")] = ext
    return hits

removed, edited, docs = [], [], []
for rel, ext in sorted(scan().items()):
    if ext in DOC_EXT:
        docs.append(rel)
    elif ext not in CODE_EXT:
        print(f"  CONFIGURATION COULD ARM THE BYPASS: {rel} — refusing to deploy; "
              f"remove the key by hand (this file is never deleted automatically)")
        sys.exit(2)
    elif rel != "Program.cs":
        os.remove(os.path.join(staged, rel))
        removed.append(rel)

# Program.cs keeps the whole startup, so edit rather than delete: drop the
# registration line plus the comment block that introduces it. The upward walk
# stops at a blank line so it cannot eat a comment belonging to the PREVIOUS
# statement, and the anchor is the exact call, so no neighbouring middleware
# line can be swallowed. `using HomeFront.Auth;` deliberately STAYS —
# HomeFront.Auth.Cognito keeps that namespace alive, so it still compiles, and a
# strip that touches more than it must is how a staged build breaks.
program = os.path.join(staged, "Program.cs")
if os.path.exists(program):
    lines = open(program, encoding="utf-8").read().split("\n")
    call = re.compile(r"^\s*app\.UseDevAutoLogin\(\)\s*;\s*$")
    comment = re.compile(r"^\s*//")
    drop = set()
    for i, line in enumerate(lines):
        if call.match(line):
            drop.add(i)
            j = i - 1
            while j >= 0 and comment.match(lines[j]) and lines[j].strip():
                drop.add(j)
                j -= 1
    for i, line in enumerate(lines):
        if i in drop or not any(m in line.lower() for m in MARKERS):
            continue
        if comment.match(line):
            drop.add(i)
        else:
            print(f"  Program.cs:{i + 1} uses the bypass outside its registration — "
                  f"refusing to guess: {line.strip()[:90]}")
            sys.exit(2)
    if drop:
        open(program, "w", encoding="utf-8").write(
            "\n".join(l for i, l in enumerate(lines) if i not in drop))
        edited.append(f"Program.cs (-{len(drop)} lines)")

survivors = {r: e for r, e in scan().items() if e not in DOC_EXT}
if survivors:
    for rel in sorted(survivors):
        print(f"  STILL PRESENT: {rel}")
    sys.exit(2)

# Absence is not enough — prove the real login path still ships, and that the
# strip did not take a neighbouring middleware line with it.
problems = []
flogin = []
for root, dirs, files in os.walk(staged):
    dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
    flogin += [os.path.join(root, f) for f in files if f == "FLogin.razor"]
if not flogin:
    problems.append("no FLogin.razor in the staged tree")
elif not any("Sec.EnforcePasswordCheck" in open(p, encoding="utf-8", errors="ignore").read()
             for p in flogin):
    problems.append("FLogin.razor no longer consults Sec.EnforcePasswordCheck")
if not os.path.exists(program):
    problems.append("no Program.cs in the staged tree")
else:
    startup = open(program, encoding="utf-8", errors="ignore").read()
    for required in ("app.UseAuthentication();", "app.UseAuthorization();",
                     "app.UseAntiforgery();", "/auth/signin"):
        if required not in startup:
            problems.append(f"Program.cs no longer has {required} — the strip was too greedy")
if problems:
    for p in problems:
        print(f"  LOGIN PATH BROKEN: {p}")
    sys.exit(3)

summary = []
if removed:
    summary.append(f"removed {', '.join(removed)}")
if edited:
    summary.append(", ".join(edited))
if docs:
    summary.append(f"{len(docs)} doc mention(s) left as prose ({', '.join(docs)})")
print("  " + ("; ".join(summary) if summary else "nothing to strip"))
PYEOF
    then
        err "$label: development login-skip could not be removed from the staged tree"
        return 1
    fi
    ok "$label: no login skip in the pushed copy — password check intact"
}

# ── Build verification ──────────────────────────────────────────────
verify_build() {
    local deploy_dir="$1"
    local proj="${2:-HomeFront}"
    log "  Build verification"
    (cd "$deploy_dir" && rm -rf bin obj && dotnet build "$proj.sln" --no-incremental -v q 2>&1) | tail -3
    if [ "${PIPESTATUS[0]}" -eq 0 ]; then
        ok "Build succeeded"
        return 0
    else
        err "Build FAILED — check output above"
        return 1
    fi
}

mkdir -p "$DEPLOY"

# ── Project definitions ─────────────────────────────────────────────
# FORMAT: name|source_dir|deploy_git_dir|remote_url
PROJECTS=(
    # hyphen-pb main is fed from HomeFront (owner directive 2026-08-27) — HomeFrontPB
    # is untouched on this machine and now feeds the R1-UAT branch instead.
    "hyphen-pb|$HF_ROOT/MobileSource/HomeFront|$DEPLOY/hyphen-pb|git@gitlab.innovatixinc.com:application-modernization/hyphen-pb.git"
    "homefront|$HF_ROOT/MobileSource/HomeFront|$DEPLOY/homefront|git@gitlab.innovatixinc.com:application-modernization/homefront.git"
    "flexkit|$ROOT/FlexKit|$DEPLOY/flexkit|git@gitlab.innovatixinc.com:application-modernization/flexkit.git"
    "flexcore|$ROOT/FlexCore|$DEPLOY/flexcore|https://github.com/wadoodachaudhary/FlexCore.git"
)

# ════════════════════════════════════════════════════════════════════
#  PULL MODE — fetch MR changes from hyphen-pb main → HomeFront
# ════════════════════════════════════════════════════════════════════
# main is fed from MobileSource/HomeFront (owner directive 2026-08-27), so an
# MR merged on main comes back to HomeFront. HomeFrontPB is FROZEN: it is a
# SOURCE ONLY, feeding the R1-UAT branch, and nothing writes into it ever again
# (owner directive 2026-08-31). The old pull path rsynced main onto HomeFrontPB
# — written when main was still fed from PB — which would now drop HomeFront's
# tree onto the wrong app. The PB→HF page mirror and its @namespace injection
# went with it: main already carries HomeFront's Migrated/ layout.
if [ "$MODE" = "pull" ]; then
    log "PULL — hyphen-pb main → HomeFront"
    GIT_DIR="$DEPLOY/hyphen-pb"
    SRC_DIR="$HF_ROOT/MobileSource/HomeFront"

    if [ ! -d "$GIT_DIR/.git" ]; then
        err "No staging repo at $GIT_DIR — run a push first to initialize"
        exit 1
    fi

    git -C "$GIT_DIR" fetch origin

    # The clone serves two branches; park it on main before reading any sha.
    if git -C "$GIT_DIR" rev-parse --verify main >/dev/null 2>&1; then
        git -C "$GIT_DIR" checkout main --quiet
    else
        git -C "$GIT_DIR" checkout -b main origin/main --quiet
    fi

    # Compare origin/main to the sha WE last pushed — local main — never to HEAD.
    # HEAD can be parked on R1-UAT, which makes our own main commits look upstream.
    LAST_PUSHED=$(git -C "$GIT_DIR" rev-parse main)
    REMOTE_HEAD=$(git -C "$GIT_DIR" rev-parse origin/main)

    if [ "$LAST_PUSHED" = "$REMOTE_HEAD" ]; then
        ok "No new changes on remote"
        exit 0
    fi

    log "Incoming changes:"
    git -C "$GIT_DIR" log --format='  %h %an | %s' "$LAST_PUSHED..$REMOTE_HEAD"
    echo ""

    # Branch-owned / environment files never come back to source: the csproj and
    # sln carry the deploy's PackageReference form (source keeps ProjectReference),
    # and appsettings/pipelines/certs belong to the remote. Deletions are reported,
    # not applied — removing a local file is not something to do unattended.
    PULL_PATHSPEC=(
        ':!HomeFront.csproj' ':!HomeFront.sln'
        ':!local-packages/' ':!NuGet.Config'
        ':!appsettings*.json' ':!App_Data/jira-settings.json'
        ':!azure-pipelines*.yml' ':!certs/'
        ':!wwwroot/tickets/' ':!wwwroot/feedback/' ':!Logs/'
    )

    CHANGED_LIST=$(mktemp)
    git -C "$GIT_DIR" diff --name-only --diff-filter=d "$LAST_PUSHED..$REMOTE_HEAD" \
        -- "${PULL_PATHSPEC[@]}" > "$CHANGED_LIST"
    DELETED=$(git -C "$GIT_DIR" diff --name-only --diff-filter=D "$LAST_PUSHED..$REMOTE_HEAD" \
        -- "${PULL_PATHSPEC[@]}")
    SKIPPED=$(git -C "$GIT_DIR" diff --name-only "$LAST_PUSHED..$REMOTE_HEAD" \
        -- HomeFront.csproj HomeFront.sln 'appsettings*.json' 'azure-pipelines*.yml' NuGet.Config)

    echo "  Files to sync into HomeFront:"
    if [ -s "$CHANGED_LIST" ]; then sed 's/^/    /' "$CHANGED_LIST"; else echo "    (none)"; fi
    if [ -n "$SKIPPED" ]; then
        echo ""
        err "Branch-owned files changed upstream — HAND-MERGE these, they are NOT synced:"
        echo "$SKIPPED" | sed 's/^/    /'
    fi
    if [ -n "$DELETED" ]; then
        echo ""
        err "Deleted upstream — remove by hand if intended:"
        echo "$DELETED" | sed 's/^/    /'
    fi
    echo ""

    if $DRY_RUN; then
        ok "DRY RUN — would merge and sync the above into $SRC_DIR"
        rm -f "$CHANGED_LIST"
        exit 0
    fi

    # Capture the BASE (what we last pushed) BEFORE merging — the 3-way merge
    # below needs the common ancestor of "our local edits" and "their upstream
    # change", and after the merge local main no longer points at it.
    MERGE_BASE_SHA="$LAST_PUSHED"

    git -C "$GIT_DIR" merge origin/main --no-edit
    ok "Merged origin/main into staging"

    # Bring in ONLY the files the incoming commits touched, and 3-WAY MERGE each
    # one rather than copying over it.
    #
    # A plain copy is safe only when the source file is untouched since the last
    # push. These trees are edited live by the owner and other sessions, so a file
    # can be BOTH incoming and locally dirty — on 2026-09-02 that was FPriceList
    # and FPricingWorkSheet, and a copy would have silently destroyed uncommitted
    # toolbar work. Merging with the last-pushed content as the base keeps both
    # sides when the edits are in different regions, and stops with the file
    # untouched when they genuinely collide.
    MERGED=0; COPIED=0; CONFLICTED=0
    CONFLICT_LIST=$(mktemp)
    while IFS= read -r f; do
        [ -n "$f" ] || continue
        mkdir -p "$SRC_DIR/$(dirname "$f")"
        if [ ! -f "$SRC_DIR/$f" ]; then
            git -C "$GIT_DIR" show "origin/main:$f" > "$SRC_DIR/$f" 2>/dev/null && COPIED=$((COPIED+1))
            continue
        fi
        BASE_TMP=$(mktemp); THEIRS_TMP=$(mktemp); OURS_TMP=$(mktemp)
        if ! git -C "$GIT_DIR" show "$MERGE_BASE_SHA:$f" > "$BASE_TMP" 2>/dev/null; then
            : > "$BASE_TMP"          # new upstream file: empty base
        fi
        git -C "$GIT_DIR" show "origin/main:$f" > "$THEIRS_TMP"
        cp "$SRC_DIR/$f" "$OURS_TMP"

        if cmp -s "$OURS_TMP" "$BASE_TMP"; then
            # Source untouched since the last push — fast-forward, no merge needed.
            cp "$THEIRS_TMP" "$SRC_DIR/$f"; COPIED=$((COPIED+1))
        else
            set +e
            git merge-file -L "HomeFront (local)" -L "last pushed" -L "hyphen-pb main" \
                "$OURS_TMP" "$BASE_TMP" "$THEIRS_TMP" >/dev/null 2>&1
            rc=$?
            set -e
            if [ "$rc" -eq 0 ]; then
                cp "$OURS_TMP" "$SRC_DIR/$f"; MERGED=$((MERGED+1))
            else
                echo "$f" >> "$CONFLICT_LIST"; CONFLICTED=$((CONFLICTED+1))
            fi
        fi
        rm -f "$BASE_TMP" "$THEIRS_TMP" "$OURS_TMP"
    done < "$CHANGED_LIST"

    ok "Fast-forwarded $COPIED, 3-way merged $MERGED file(s) into HomeFront"
    if [ "$CONFLICTED" -gt 0 ]; then
        echo ""
        err "$CONFLICTED file(s) CONFLICT — left UNCHANGED, merge by hand:"
        sed 's/^/    /' "$CONFLICT_LIST"
        err "Compare: git -C $GIT_DIR diff $MERGE_BASE_SHA..origin/main -- <file>"
    fi
    rm -f "$CONFLICT_LIST" "$CHANGED_LIST"

    echo ""
    err "HomeFrontPB is frozen — nothing was written to it (owner directive 2026-08-31)."
    log "Pull complete. Build to verify:"
    echo "  dotnet build $SRC_DIR/HomeFront.sln --no-incremental"
    exit 0
fi

# ════════════════════════════════════════════════════════════════════
#  PUSH MODE (default) — source → staging → commit & push
# ════════════════════════════════════════════════════════════════════
for entry in "${PROJECTS[@]}"; do
    IFS='|' read -r name src_dir git_dir remote_url <<< "$entry"
    log "$name: $src_dir → $git_dir"

    # ── Ensure deploy git repo exists ────────────────────────────────
    if [ ! -d "$git_dir/.git" ]; then
        log "  Init deploy repo at $git_dir"
        mkdir -p "$git_dir"
        git -C "$git_dir" init
        git -C "$git_dir" remote add origin "$remote_url"
        git -C "$git_dir" fetch origin 2>/dev/null || true
        if git -C "$git_dir" rev-parse origin/main >/dev/null 2>&1; then
            git -C "$git_dir" checkout -b main origin/main
        fi
    fi

    # ── Fetch remote first to avoid non-fast-forward ─────────────────
    git -C "$git_dir" fetch origin 2>/dev/null || true
    # Park on main first. deploy_r1uat leaves the clone on R1-UAT if it exits
    # early, and a reset --hard here would then rewrite that branch instead —
    # the main content ends up committed onto R1-UAT. Fixed 2026-08-28.
    if git -C "$git_dir" rev-parse --verify main >/dev/null 2>&1; then
        git -C "$git_dir" checkout main --quiet
    elif git -C "$git_dir" rev-parse --verify origin/main >/dev/null 2>&1; then
        git -C "$git_dir" checkout -b main origin/main --quiet
    fi

    if git -C "$git_dir" rev-parse origin/main >/dev/null 2>&1; then
        LOCAL_HEAD=$(git -C "$git_dir" rev-parse HEAD 2>/dev/null || echo "none")
        REMOTE_HEAD=$(git -C "$git_dir" rev-parse origin/main 2>/dev/null || echo "none")
        if [ "$LOCAL_HEAD" != "$REMOTE_HEAD" ] && [ "$LOCAL_HEAD" != "none" ]; then
            git -C "$git_dir" reset --hard origin/main 2>/dev/null
            ok "Reset to origin/main (remote had new commits)"
        fi
    fi

    # ── Sync source → deploy dir (excluding .git and junk) ──────────
    rsync -a --delete "${RSYNC_EXCLUDE[@]}" "$src_dir/" "$git_dir/"
    ok "Synced source"

    # No pushed copy may be able to skip the login page (owner 2026-09-12).
    if ! strip_dev_login_bypass "$git_dir" "$name"; then
        err "Skipping $name due to the login-skip guard"
        continue
    fi
    log "  $name: removing files not needed to run (agent files, harnesses, notes)"
    strip_non_runtime_files "$git_dir"

    # ── hyphen-pb: pack FlexKit, swap ref, verify build ─────────────
    if [ "$name" = "hyphen-pb" ]; then
        FLEXKIT_VER=$(pack_flexkit "$git_dir")
        swap_to_package_ref "$git_dir" "$FLEXKIT_VER" "$src_dir" "HomeFront"
        if ! verify_build "$git_dir"; then
            err "Skipping $name due to build failure"
            continue
        fi
    fi

    # ── Check for tracked junk that slipped through ──────────────────
    git -C "$git_dir" rm -r --cached --quiet .claude .codex-backups .idea 2>/dev/null || true

    # ── Commit & push ────────────────────────────────────────────────
    git -C "$git_dir" add -A
    # Final gate on the BYTES that would be pushed, not just the worktree the
    # strip edited: anything the file walk missed still has to pass through the
    # index to reach GitLab. Docs are exempt — prose describing a dev-only
    # feature is not a bypass. Added 2026-09-12.
    if git -C "$git_dir" grep --cached -I -i -q -e devautologin -- ':!*.md' 2>/dev/null; then
        err "$name: the staged index still carries the login-skip — not pushing"
        continue
    fi

    # App_Data/jira-settings.json is NOT force-added any more (2026-09-12). It
    # holds a live Atlassian API token, .gitignore says "never commit/ship", and
    # ENV_CONFIG_EXCLUDE already treats it as remote-owned — yet the force-add
    # re-committed it on every run, so the exposure could never be cleaned up.
    # The rsync excludes the path, so a copy already on the branch is left alone
    # and the server keeps its settings; nothing new is written.
    if git -C "$git_dir" diff --cached --quiet 2>/dev/null; then
        ok "No changes to push"
        continue
    fi

    TIMESTAMP=$(date +"%Y-%m-%d %H:%M")
    if $DRY_RUN; then
        ok "DRY RUN — would commit and push ($TIMESTAMP)"
        git -C "$git_dir" diff --cached --stat
        # `checkout -- .` restores the worktree FROM the index, so the staged
        # strip (a deleted file, an edited Program.cs) stayed staged and the
        # clone was left dirty — and the reset at the top of the loop is
        # conditional, so nothing cleaned it up. Fixed 2026-09-12.
        git -C "$git_dir" reset -q --hard HEAD 2>/dev/null || true
    else
        git -C "$git_dir" commit -m "Deploy $name — $TIMESTAMP"
        git -C "$git_dir" push origin main
        ok "Pushed to $remote_url"
    fi
done

log "Done."

# ════════════════════════════════════════════════════════════════════
#  R1-UAT — fed from HomeFrontPB (main is fed from HomeFront)
# ════════════════════════════════════════════════════════════════════
deploy_r1uat() {
    local git_dir="$DEPLOY/hyphen-pb"
    local src_dir="$HF_ROOT/HomeFrontPB"

    # Always hand the clone back on main, on every exit path.
    trap 'git -C "'"$DEPLOY"'/hyphen-pb" checkout main --quiet 2>/dev/null || true' RETURN

    # --dry-run only ever gated the main loop; this stage ignored it and pushed
    # for real — a "dry run" deployed to the branch QA pulls. Fixed 2026-09-12.
    if $DRY_RUN; then
        log "R1-UAT: DRY RUN — skipping (source $src_dir)"
        return 0
    fi

    log "R1-UAT: $src_dir → hyphen-pb R1-UAT"
    git -C "$git_dir" fetch origin --quiet
    git -C "$git_dir" checkout R1-UAT --quiet
    git -C "$git_dir" reset --hard origin/R1-UAT --quiet
    ok "Reset to origin/R1-UAT"

    rsync -a --delete "${RSYNC_EXCLUDE[@]}" "${R1UAT_PRESERVE_EXCLUDE[@]}" "$src_dir/" "$git_dir/"
    ok "Synced HomeFrontPB source"

    local ver
    ver=$(pack_flexkit "$git_dir")
    swap_to_package_ref "$git_dir" "$ver" "$src_dir" "HomeFrontPB"

    # Restore branch-only files AFTER the swap. Excluding them from the rsync is
    # not enough: swap_to_package_ref deletes app-side copies of FlexKit-shipped
    # icons, and the pages reference them app-relatively ("images/32/open.ico"),
    # not via _content/ — so the removal would blank the toolbar icon.
    git -C "$git_dir" checkout -- Components/Pages/FLogin.razor 2>/dev/null || true
    ok "Restored branch-only FLogin lockdown"

    if ! strip_dev_login_bypass "$git_dir" "R1-UAT"; then
        err "R1-UAT login-skip guard failed — not pushing"
        git -C "$git_dir" checkout main --quiet
        return 1
    fi
    log "  R1-UAT: removing files not needed to run (agent files, harnesses, notes)"
    strip_non_runtime_files "$git_dir"

    if ! verify_build "$git_dir" "HomeFrontPB"; then
        err "R1-UAT build failed — not pushing"
        git -C "$git_dir" checkout main --quiet
        return 1
    fi

    # The lockdown must survive every run; refuse to push if it did not.
    local toggles
    # `grep -c` with ZERO matches prints "0" and EXITS 1, so a `|| echo 0`
    # fallback appended a SECOND line: toggles became "0\n0", `[ -ne 5 ]` then
    # died with "integer expression expected" (status 2), and `if` read that as
    # false — so a completely missing lockdown printed "intact" and PUSHED an
    # FLogin with all five security toggles enabled. Counting occurrences with
    # `grep -o | wc -l` always yields one integer (0 for a missing file, which
    # blocks), and it no longer under-counts two toggles on one line.
    # Fixed 2026-09-12.
    toggles=$(grep -o 'Disabled="true"' "$git_dir/Components/Pages/FLogin.razor" 2>/dev/null | wc -l | tr -d ' ')
    if [ "$toggles" -ne 5 ]; then
        err "R1-UAT security lockdown lost ($toggles/5 toggles) — not pushing"
        git -C "$git_dir" checkout main --quiet
        return 1
    fi
    ok "Lockdown intact ($toggles/5 toggles)"

    # Stage FIRST, then compare the INDEX to HEAD. `status --porcelain` reports the
    # xml_orig/*.xml files as modified on every run (they are CRLF in the source tree
    # and git normalizes them to LF on add), but the staged tree is identical to HEAD —
    # so `git commit` exited non-zero with "nothing to commit" and `set -e` killed the
    # script BEFORE the push. R1-UAT silently stopped deploying. Fixed 2026-08-28;
    # this now matches the main loop's guard.
    git -C "$git_dir" add -A
    # Final gate on the BYTES that would be pushed, not just the worktree the
    # strip edited: anything the file walk missed still has to pass through the
    # index to reach GitLab. Docs are exempt — prose describing a dev-only
    # feature is not a bypass. Added 2026-09-12.
    if git -C "$git_dir" grep --cached -I -i -q -e devautologin -- ':!*.md' 2>/dev/null; then
        err "R1-UAT: the staged index still carries the login-skip — not pushing"
        git -C "$git_dir" checkout main --quiet
        return 1
    fi

    if git -C "$git_dir" diff --cached --quiet 2>/dev/null; then
        ok "R1-UAT: no changes to push"
    else
        git -C "$git_dir" commit -q -m "Deploy R1-UAT from HomeFrontPB — $(date '+%Y-%m-%d %H:%M')"
        git -C "$git_dir" push origin R1-UAT
        ok "Pushed R1-UAT"
    fi
    git -C "$git_dir" checkout main --quiet
}

# R2-UAT (owner 2026-10-01: "R1 is now R2 -- so we have to do what we were doing with R1 there
# as well"). The branch is main plus two branch-owned edits a teammate maintains: the UAT login
# lockdown in Components/Pages/Migrated/FLogin.razor and azure-pipelines-uat.yml's trigger. It is
# fed from the SAME snapshot this run pushed to main -- never from HomeFrontPB -- by MERGING
# origin/main into it, so a teammate's direct push to R2-UAT is kept and nothing is overwritten.
# Any conflict, a lost lockdown, a retargeted pipeline or a failed build stops the stage unpushed.
deploy_r2uat() {
    local git_dir="$DEPLOY/hyphen-pb"

    # Always hand the clone back on main with no half-finished merge, on every exit path.
    trap 'git -C "'"$DEPLOY"'/hyphen-pb" merge --abort 2>/dev/null; git -C "'"$DEPLOY"'/hyphen-pb" checkout main --quiet 2>/dev/null || true' RETURN

    git -C "$git_dir" fetch origin --quiet
    if ! git -C "$git_dir" rev-parse --verify --quiet origin/R2-UAT >/dev/null; then
        err "R2-UAT: origin/R2-UAT does not exist -- not deployed"
        return 1
    fi
    local ahead
    ahead=$(git -C "$git_dir" rev-list --count origin/R2-UAT..origin/main)
    log "R2-UAT: merging origin/main ($ahead commit(s) ahead) into hyphen-pb R2-UAT"
    if $DRY_RUN; then
        ok "DRY RUN -- main's new snapshot is not pushed in a dry run, so this only checks origin/main as it stands"
    fi
    if [ "$ahead" -eq 0 ]; then
        ok "R2-UAT already contains main -- nothing to push"
        return 0
    fi

    if ! git -C "$git_dir" checkout -B R2-UAT origin/R2-UAT --quiet; then
        err "R2-UAT: could not check the branch out -- not deployed"
        return 1
    fi
    if ! git -C "$git_dir" merge --no-ff --no-commit origin/main >/dev/null 2>&1; then
        err "R2-UAT: merging main CONFLICTS -- not pushing. Resolve on the branch, then re-run. Conflicted:"
        git -C "$git_dir" diff --name-only --diff-filter=U | sed 's/^/      /'
        return 1
    fi

    local flogin="$git_dir/Components/Pages/Migrated/FLogin.razor" t missing=""
    for t in EnforceAuth EncryptDbConnection EnforcePasswordCheck GetPasswordFromSystemSecrets; do
        grep -q "@bind-Checked=\"Sec.$t\" Disabled=\"true\"" "$flogin" 2>/dev/null || missing="$missing $t"
    done
    grep -q 'Sec.EncryptDbConnection = false;' "$flogin" 2>/dev/null || missing="$missing EncryptDbConnection=false"
    if [ -n "$missing" ]; then
        err "R2-UAT security lockdown lost (${missing# }) -- not pushing"
        return 1
    fi
    ok "Lockdown intact (4 disabled toggles, EncryptDbConnection forced off)"

    if ! grep -q -- '- R2-UAT' "$git_dir/azure-pipelines-uat.yml" 2>/dev/null; then
        err "R2-UAT: azure-pipelines-uat.yml no longer triggers on R2-UAT -- not pushing"
        return 1
    fi
    ok "UAT pipeline still triggers on R2-UAT"

    if git -C "$git_dir" grep --cached -I -i -q -e devautologin -- ':!*.md' 2>/dev/null; then
        err "R2-UAT: the merged index carries the login-skip -- not pushing"
        return 1
    fi
    ok "R2-UAT: no login skip in the merged copy"

    if $DRY_RUN; then
        ok "DRY RUN -- R2-UAT would take $ahead commit(s) from main; the merge is clean and the lockdown holds"
        return 0
    fi

    if ! verify_build "$git_dir" "HomeFront"; then
        err "R2-UAT build failed -- not pushing"
        return 1
    fi
    git -C "$git_dir" commit -q -m "Merge main into R2-UAT -- $(date '+%Y-%m-%d %H:%M')"
    git -C "$git_dir" push origin R2-UAT
    ok "Pushed R2-UAT"
}

if [ "$MODE" = "push" ]; then
    if $WITH_R1UAT; then
        deploy_r1uat
    else
        log "R1-UAT: FROZEN — not deployed (pass --with-r1uat to deploy it)"
    fi
    if $WITH_R2UAT; then
        deploy_r2uat || { err "R2-UAT stage failed — main and the other repositories are unaffected"; exit 1; }
    else
        log "R2-UAT: not deployed (pass --with-r2uat to merge main into it)"
    fi
fi
