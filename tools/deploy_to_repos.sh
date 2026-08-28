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
#   ./tools/deploy_to_repos.sh --pull       # pull MR changes from hyphen-pb
set -euo pipefail

ROOT="/Users/wadood/projects/VBToCSharp"
HF_ROOT="$ROOT/HomeFront"
DEPLOY="$HF_ROOT/Deploy/repos"
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

RSYNC_EXCLUDE=(
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

# Files that differ between source (ProjectReference) and deploy (PackageReference).
# These are NEVER synced back from the remote — source versions are authoritative.
PULL_EXCLUDE_PB=(
    --exclude='HomeFrontPB.csproj'
    --exclude='HomeFrontPB.sln'
    --exclude='local-packages/'
    --exclude='NuGet.Config'
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

MODE="push"
DRY_RUN=false
for arg in "$@"; do
    case "$arg" in
        --pull)   MODE="pull" ;;
        --dry-run) DRY_RUN=true ;;
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
    local existing_nupkg
    existing_nupkg=$(ls "$HF_ROOT/HomeFrontPB/local-packages/FlexKit."*.nupkg 2>/dev/null | head -1 || true)
    local existing_ver=""
    if [ -n "$existing_nupkg" ]; then
        existing_ver=$(basename "$existing_nupkg" | sed 's/FlexKit\.\(.*\)\.nupkg/\1/')
    fi

    local need_pack=false
    if [ -z "$existing_nupkg" ]; then
        need_pack=true
    elif [ "$existing_ver" != "$cur_ver" ]; then
        need_pack=true
    fi

    if $need_pack; then
        log "  Packing FlexKit $cur_ver" >&2
        (dotnet pack "$FLEXKIT_CSPROJ" -c Release -o /tmp/flexkit-pack) >&2
        rm -f "$HF_ROOT/HomeFrontPB/local-packages/FlexKit."*.nupkg
        rm -f "$HF_ROOT/HomeFrontPB/local-packages/FlexKit."*.snupkg
        cp /tmp/flexkit-pack/FlexKit."$cur_ver".nupkg "$HF_ROOT/HomeFrontPB/local-packages/"
        cp /tmp/flexkit-pack/FlexKit."$cur_ver".snupkg "$HF_ROOT/HomeFrontPB/local-packages/" 2>/dev/null || true
        rm -rf /tmp/flexkit-pack
        ok "Packed FlexKit $cur_ver" >&2
    else
        ok "FlexKit $cur_ver nupkg is current" >&2
    fi

    mkdir -p "$target_dir/local-packages"
    rm -f "$target_dir/local-packages/FlexKit."*.nupkg
    rm -f "$target_dir/local-packages/FlexKit."*.snupkg
    cp "$HF_ROOT/HomeFrontPB/local-packages/FlexKit."*.nupkg "$target_dir/local-packages/"
    cp "$HF_ROOT/HomeFrontPB/local-packages/FlexKit."*.snupkg "$target_dir/local-packages/" 2>/dev/null || true

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
#  PULL MODE — fetch MR changes from remote → source
# ════════════════════════════════════════════════════════════════════
if [ "$MODE" = "pull" ]; then
    log "PULL — fetching MR changes from hyphen-pb"
    GIT_DIR="$DEPLOY/hyphen-pb"
    SRC_DIR="$HF_ROOT/HomeFrontPB"
    REMOTE_URL="git@gitlab.innovatixinc.com:application-modernization/hyphen-pb.git"

    if [ ! -d "$GIT_DIR/.git" ]; then
        err "No staging repo at $GIT_DIR — run a push first to initialize"
        exit 1
    fi

    # Fetch and show what's new
    git -C "$GIT_DIR" fetch origin
    LOCAL_HEAD=$(git -C "$GIT_DIR" rev-parse HEAD)
    REMOTE_HEAD=$(git -C "$GIT_DIR" rev-parse origin/main)

    if [ "$LOCAL_HEAD" = "$REMOTE_HEAD" ]; then
        ok "No new changes on remote"
        exit 0
    fi

    # Show incoming changes
    log "Incoming changes:"
    git -C "$GIT_DIR" log --oneline "$LOCAL_HEAD..$REMOTE_HEAD"
    echo ""
    git -C "$GIT_DIR" diff --stat "$LOCAL_HEAD..$REMOTE_HEAD" -- \
        ':!HomeFrontPB.csproj' ':!HomeFrontPB.sln' ':!local-packages/' ':!NuGet.Config'
    echo ""

    if $DRY_RUN; then
        ok "DRY RUN — would merge and sync to $SRC_DIR"
        exit 0
    fi

    # Merge remote into staging
    git -C "$GIT_DIR" merge origin/main --no-edit
    ok "Merged origin/main into staging"

    # Sync staging → source (excluding build artifacts and deploy-only files)
    rsync -a "${PULL_EXCLUDE_PB[@]}" "$GIT_DIR/" "$SRC_DIR/"
    ok "Synced changes to HomeFrontPB source"

    # Also mirror to HomeFront (HF) for the sync pair
    HF_DIR="$HF_ROOT/MobileSource/HomeFront"
    log "Mirroring to HomeFront (sync pair)"
    echo "  NOTE: Only razor/cs/css files in Components/Pages/ and Services/"
    echo "        are auto-mirrored. FMain.razor is excluded (known drift)."
    echo "        Review manually if the MR touched other files."
    echo ""

    # Mirror page files (PB flat → HF Migrated/)
    if [ -d "$GIT_DIR/Components/Pages" ]; then
        for f in "$GIT_DIR"/Components/Pages/F*.razor "$GIT_DIR"/Components/Pages/F*.razor.css; do
            [ -f "$f" ] || continue
            base=$(basename "$f")
            # Skip FMain — known drift
            [[ "$base" == FMain.razor* ]] && continue
            hf_target="$HF_DIR/Components/Pages/Migrated/$base"
            if [ -f "$hf_target" ]; then
                cp "$f" "$hf_target"
            fi
        done
        ok "Mirrored page files to HF Migrated/"

        # HF Migrated/ files need @namespace HomeFront.Components.Pages
        # (PB's flat Pages/ auto-gets the right namespace; Migrated/ doesn't).
        # Inject the directive if missing — affects files referenced as types
        # by other pages (FAttachments, FTakeoff, etc.).
        for razor in "$HF_DIR"/Components/Pages/Migrated/F*.razor; do
            [ -f "$razor" ] || continue
            [[ "$razor" == *.razor.css ]] && continue
            if ! grep -q '@namespace' "$razor"; then
                # Insert after the FIRST closing comment block, or as line 1
                if grep -qn '───── \*@' "$razor"; then
                    line_num=$(grep -n '───── \*@' "$razor" | head -1 | cut -d: -f1)
                    sed -i '' "${line_num}a\\
@namespace HomeFront.Components.Pages" "$razor"
                else
                    sed -i '' '1i\
@namespace HomeFront.Components.Pages
' "$razor"
                fi
            fi
        done
        ok "Ensured @namespace on HF Migrated/ files"
    fi

    # Mirror service files
    if [ -d "$GIT_DIR/Services" ]; then
        rsync -a --existing "$GIT_DIR/Services/" "$HF_DIR/Services/"
        ok "Mirrored service files to HF"
    fi

    log "Pull complete. Build both solutions to verify:"
    echo "  dotnet build $SRC_DIR/HomeFrontPB.sln --no-incremental"
    echo "  dotnet build $HF_DIR/HomeFront.sln --no-incremental"
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
    if [ -f "$git_dir/App_Data/jira-settings.json" ]; then
        git -C "$git_dir" add -f App_Data/jira-settings.json 2>/dev/null || true
    fi
    if git -C "$git_dir" diff --cached --quiet 2>/dev/null; then
        ok "No changes to push"
        continue
    fi

    TIMESTAMP=$(date +"%Y-%m-%d %H:%M")
    if $DRY_RUN; then
        ok "DRY RUN — would commit and push ($TIMESTAMP)"
        git -C "$git_dir" diff --cached --stat
        git -C "$git_dir" checkout -- . 2>/dev/null || true
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

    if ! verify_build "$git_dir" "HomeFrontPB"; then
        err "R1-UAT build failed — not pushing"
        git -C "$git_dir" checkout main --quiet
        return 1
    fi

    # The lockdown must survive every run; refuse to push if it did not.
    local toggles
    toggles=$(grep -c 'Disabled="true"' "$git_dir/Components/Pages/FLogin.razor" || echo 0)
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
    if git -C "$git_dir" diff --cached --quiet 2>/dev/null; then
        ok "R1-UAT: no changes to push"
    else
        git -C "$git_dir" commit -q -m "Deploy R1-UAT from HomeFrontPB — $(date '+%Y-%m-%d %H:%M')"
        git -C "$git_dir" push origin R1-UAT
        ok "Pushed R1-UAT"
    fi
    git -C "$git_dir" checkout main --quiet
}

if [ "$MODE" = "push" ]; then
    deploy_r1uat
fi
