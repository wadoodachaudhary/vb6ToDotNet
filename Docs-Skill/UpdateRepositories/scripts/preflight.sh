#!/bin/bash
# preflight.sh — READ-ONLY checks before an Update Repositories run.
# Changes nothing except `git fetch` in the staging clones. Every check runs and reports;
# the ATTENTION list at the end is what needs a decision before deploying.
#
#   bash preflight.sh            # default: upstream commits from the last 10 days
#   bash preflight.sh 21         # look back 21 days for teammates' merged work
set -uo pipefail   # deliberately NOT -e: a failing check must not hide the ones after it

DAYS="${1:-10}"
V=/Users/wadood/projects/VBToCSharp
H=$V/HomeFront
HF=$H/MobileSource/HomeFront
PB=$H/HomeFrontPB
D=$H/Deploy/repos
SKILL_SRC=$H/Docs-Skill/UpdateRepositories
SKILL_DST=$HOME/.claude/skills/update-repositories
ATTN=()
attn() { ATTN+=("$1"); }
hdr()  { printf '\n== %s\n' "$1"; }
now=$(date +%s)

# ── 1. Other agent sessions ─────────────────────────────────────────────────────
# A session that dispatched a background agent/workflow can be silent in its main
# transcript for many minutes while its agents still edit files, so the check walks
# every file under each session directory (subagents/, subagents/workflows/<id>/).
hdr "1. Agent sessions active in the last 30 minutes (your own session is listed too)"
found=0
for t in "$HOME"/.claude/projects/*VBToCSharp*/*.jsonl; do
  [ -e "$t" ] || continue
  newest=$(stat -f %m "$t")
  base=${t%.jsonl}
  if [ -d "$base" ]; then
    sub=$(find "$base" -type f -exec stat -f %m {} + 2>/dev/null | sort -rn | head -1)
    [ -n "$sub" ] && [ "$sub" -gt "$newest" ] && newest=$sub
  fi
  age=$(( (now - newest) / 60 ))
  if [ "$age" -lt 30 ]; then
    found=1
    printf '   %3s min ago  %s\n' "$age" "$(basename "$(dirname "$t")" | sed 's/-Users-wadood-projects-VBToCSharp-//')/$(basename "$base" | cut -c1-8)"
  fi
done
[ $found = 0 ] && echo "   none"

# ── 2. Working trees ────────────────────────────────────────────────────────────
hdr "2. Working trees (the deploy ships WORKING TREES — uncommitted edits go out)"
for spec in "app|$HF" "FlexKit|$V/FlexKit" "FlexCore|$V/FlexCore" "outer|$H" "HomeFrontPB|$PB"; do
  name=${spec%%|*}; dir=${spec#*|}
  br=$(git -C "$dir" branch --show-current 2>/dev/null)
  n=$(git -C "$dir" status --porcelain 2>/dev/null | wc -l | tr -d ' ')
  newest=$(git -C "$dir" status --porcelain 2>/dev/null | awk '{print $NF}' | while read -r f; do [ -e "$dir/$f" ] && stat -f %m "$dir/$f"; done | sort -rn | head -1)
  when=""; [ -n "$newest" ] && when="newest edit $(( (now - newest) / 60 )) min ago"
  printf '   %-12s %-26s %3s dirty  %s\n' "$name" "$br" "$n" "$when"
  if [ "$name" = FlexKit ] && [ "$br" != "telerik-parity-20260904" ]; then
    attn "FlexKit is on '$br' — it ships from telerik-parity-20260904; never check out main"
  fi
  if [ "$name" != HomeFrontPB ] && [ "$name" != outer ] && [ "$n" -gt 0 ] && [ -n "$newest" ] && [ $(( (now - newest) / 60 )) -lt 15 ]; then
    attn "$name was edited less than 15 min ago — another session may be mid-edit"
  fi
done
echo "   (HomeFrontPB is FROZEN: its local entries are expected — never commit, revert or build it)"

# ── 3. Step 0: both remotes vs what we last pushed ──────────────────────────────
hdr "3. Step 0 — fetch every staging clone; origin vs the sha we LAST PUSHED (local main)"
for r in hyphen-pb homefront flexkit flexcore; do
  c=$D/$r
  git -C "$c" fetch -q --all --prune 2>/dev/null || attn "$r: git fetch failed"
  cur=$(git -C "$c" branch --show-current)
  om=$(git -C "$c" rev-parse --short origin/main 2>/dev/null)
  lm=$(git -C "$c" rev-parse --short main 2>/dev/null)
  extra=""
  [ "$r" = hyphen-pb ] && extra="  R1-UAT=$(git -C "$c" rev-parse --short origin/R1-UAT 2>/dev/null)"
  printf '   %-10s on=%-8s origin/main=%-9s last-pushed=%-9s%s\n' "$r" "$cur" "$om" "$lm" "$extra"
  [ "$cur" != main ] && attn "$r staging clone is parked on '$cur', not main"
  dirty=$(git -C "$c" status --porcelain | wc -l | tr -d ' ')
  [ "$dirty" -gt 0 ] && attn "$r staging clone has $dirty dirty entries"
  if [ "$om" != "$lm" ]; then
    new=$(git -C "$c" log --oneline "main..origin/main" 2>/dev/null | grep -v 'Deploy ' | wc -l | tr -d ' ')
    attn "$r: origin/main moved past our last push ($new non-deploy commit(s)) — run --pull before pushing"
    git -C "$c" log --format='      %h %cs %an: %s' "main..origin/main" 2>/dev/null | grep -v 'Deploy ' | head -10
  fi
done

# ── 4. Step 0b: branches that may carry open merge requests ─────────────────────
hdr "4. Step 0b — hyphen-pb branches ahead of main (confirm open MRs in the GitLab UI too)"
c=$D/hyphen-pb
for b in $(git -C "$c" for-each-ref --format='%(refname:short)' refs/remotes/origin | grep -vE '^origin/(HEAD|main|R1-UAT)$|^origin$'); do
  n=$(git -C "$c" rev-list --count "origin/main..$b" 2>/dev/null || echo 0)
  [ "${n:-0}" -gt 0 ] || continue
  last=$(git -C "$c" log -1 --format=%ct "$b")
  age=$(( (now - last) / 86400 ))
  printf '   %-44s ahead=%-3s last commit %s days ago (%s)\n' "$b" "$n" "$age" "$(git -C "$c" log -1 --format=%an "$b")"
  # parked/* are snapshots pushed from another machine, not merge requests — listed, never flagged.
  case "$b" in origin/parked/*) continue ;; esac
  [ "$age" -le 14 ] && attn "branch $b has commits from the last 14 days — check for an open MR"
done
echo "   GitLab UI: https://gitlab.innovatixinc.com/groups/application-modernization/-/merge_requests/?state=opened"

# ── 5. Would the push REVERT a teammate's merged work? ──────────────────────────
# The deploy rsyncs our tree over hyphen-pb main. A teammate's commit that is on
# origin/main but whose content is not in our tree gets silently undone — and once
# undone, origin/main == our tree, so step 0 looks clean forever after.
# Blob proof per file: ours == commit^ blob -> REVERTED; ours == commit blob -> present.
hdr "5. Teammate commits on origin/main in the last $DAYS days — is their content in OUR tree?"
reverted=0
while IFS='|' read -r sha date author subj; do
  [ -z "$sha" ] && continue
  files=$(git -C "$c" show --name-only --format='' "$sha" 2>/dev/null | grep -v '^$')
  [ -z "$files" ] && continue          # merge commits list no files
  printf '   %s %s %s: %s\n' "$sha" "$date" "$author" "$(echo "$subj" | cut -c1-60)"
  while read -r f; do
    [ -z "$f" ] && continue
    # Repo-owned / environment files never live in our tree by design (the deploy's
    # REPO_OWNED_EXCLUDE + ENV_CONFIG_EXCLUDE leave the branch's copies alone).
    case "$f" in
      azure-pipelines*.yml|NuGet.Config|local-packages/*|certs/*|appsettings*.json|App_Data/jira-settings.json)
        printf '      %-58s %s\n' "$f" "repo-owned (never in our tree — not a revert)"; continue ;;
    esac
    base=$(git -C "$c" rev-parse "${sha}^:${f}" 2>/dev/null)
    post=$(git -C "$c" rev-parse "${sha}:${f}" 2>/dev/null)
    if [ -e "$HF/$f" ]; then ours=$(git -C "$HF" hash-object "$HF/$f"); else ours=MISSING; fi
    if   [ "$ours" = "$post" ]; then st="present"
    elif [ -n "$base" ] && [ "$ours" = "$base" ]; then st="REVERTED — ours is the pre-commit blob"; reverted=1
         attn "REVERT: $f from $sha ($author) is NOT in our tree — merge it before pushing"
    elif [ "$ours" = MISSING ] && [ -z "$post" ]; then st="deleted upstream"
    elif [ "$ours" = MISSING ]; then st="ABSENT in our tree"; attn "ABSENT: $f from $sha ($author)"
    else st="diverged — read both versions"; fi
    printf '      %-58s %s\n' "$f" "$st"
  done <<< "$files"
done < <(git -C "$c" log --since="$DAYS days ago" --no-merges --format='%h|%cs|%an|%s' origin/main 2>/dev/null | grep -v '|Deploy ')
[ $reverted = 0 ] && echo "   no teammate file is at its pre-commit blob in our tree"

# ── 6. Does a working-tree edit UNDO a recent commit? ───────────────────────────
# Seen 2026-09-13: a stale-copy edit restored a file byte-for-byte to the blob BEFORE
# a fix, re-shipping the bug. Proof: working blob == <recent commit>^:<path>.
hdr "6. Uncommitted edits that exactly undo a recent commit"
undo=0
for spec in "app|$HF" "FlexKit|$V/FlexKit" "FlexCore|$V/FlexCore"; do
  name=${spec%%|*}; dir=${spec#*|}
  while read -r f; do
    [ -z "$f" ] || [ ! -f "$dir/$f" ] && continue
    ours=$(git -C "$dir" hash-object "$dir/$f")
    for sha in $(git -C "$dir" log -8 --format=%h -- "$f" 2>/dev/null); do
      parent=$(git -C "$dir" rev-parse "${sha}^:${f}" 2>/dev/null) || continue
      mine=$(git -C "$dir" rev-parse "${sha}:${f}" 2>/dev/null)
      if [ "$ours" = "$parent" ] && [ "$ours" != "$mine" ]; then
        printf '   %-9s %-50s UNDOES %s (%s)\n' "$name" "$f" "$sha" "$(git -C "$dir" log -1 --format=%s "$sha" | cut -c1-50)"
        attn "$name/$f exactly undoes $sha — likely a stale copy; check before shipping"; undo=1; break
      fi
    done
  done < <(git -C "$dir" diff --name-only 2>/dev/null)
done
[ $undo = 0 ] && echo "   none"

# ── 7. Library versions: burned or not? ─────────────────────────────────────────
# NuGet treats an extracted version as immutable: a repack of a version already in
# ~/.nuget/packages is SHADOWED in the staged build, which then verifies stale bytes.
# A version published to nuget.org can be unlisted but never replaced.
hdr "7. Library versions"
fk=$(grep -o '<Version>[^<]*' "$V/FlexKit/FlexKit.csproj" | head -1 | cut -c10-)
fc=$(grep -o '<Version>[^<]*' "$V/FlexCore/FlexCore.csproj" | head -1 | cut -c10-)
fk_changed=$(( $(git -C "$V/FlexKit" status --porcelain | wc -l) ))
fc_changed=$(( $(git -C "$V/FlexCore" status --porcelain | wc -l) ))
if [ -d "$HOME/.nuget/packages/flexkit/$fk" ]; then fk_state="BURNED (already extracted)"; else fk_state="fresh"; fi
published=$(curl -s --max-time 15 https://api.nuget.org/v3-flatcontainer/flexcore/index.json | python3 -c 'import sys,json; v=json.load(sys.stdin)["versions"]; print(v[-1])' 2>/dev/null || echo "?")
printf '   FlexKit  csproj %-9s %s   (uncommitted entries: %s)\n' "$fk" "$fk_state" "$fk_changed"
printf '   FlexCore csproj %-9s nuget.org latest: %s   (uncommitted entries: %s)\n' "$fc" "$published" "$fc_changed"
printf '   feed: %s\n' "$(ls "$H/Deploy/local-packages" 2>/dev/null | tr '\n' ' ')"
[ "$fk_state" != fresh ] && attn "FlexKit $fk is burned — bump FlexKit.csproj if ANY FlexKit change ships this run"
[ "$published" = "$fc" ] && attn "FlexCore $fc is already on nuget.org — bump FlexCore.csproj before the next publish"

# ── 8. The installed skill matches its canonical source? ────────────────────────
hdr "8. Skill install"
if [ -d "$SKILL_DST" ]; then
  if diff -rq "$SKILL_SRC" "$SKILL_DST" >/dev/null 2>&1; then echo "   ~/.claude/skills/update-repositories matches Docs-Skill"
  else echo "   installed copy DIFFERS from Docs-Skill"; attn "skill install is stale — run: bash $SKILL_SRC/install.sh"; fi
else echo "   not installed"; attn "skill not installed — run: bash $SKILL_SRC/install.sh"; fi

# ── Summary ─────────────────────────────────────────────────────────────────────
hdr "ATTENTION (${#ATTN[@]})"
if [ ${#ATTN[@]} -eq 0 ]; then echo "   nothing — clear to proceed to commit / build / harnesses"
else for a in "${ATTN[@]}"; do echo "   - $a"; done; fi
