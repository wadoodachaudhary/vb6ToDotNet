#!/bin/bash
# preflight.sh — READ-ONLY checks before an Update Repositories run.
# Changes nothing except `git fetch` in the staging clones. Every check runs and reports; the ATTENTION
# list at the end is what needs a decision before deploying.
#
#   bash preflight.sh        # historic check looks back 10 days (unshipped commits are always checked)
#   bash preflight.sh 21
set -uo pipefail   # deliberately NOT -e: a failing check must not hide the ones after it

DAYS="${1:-10}"
source "$(cd "$(dirname "$0")" && pwd)/_common.sh"
ATTN=()
attn() { ATTN+=("$1"); }
hdr()  { printf '\n== %s\n' "$1"; }
now=$(date +%s)

# ── 1. Other agent sessions ─────────────────────────────────────────────────────
# A session that dispatched a background agent/workflow can be silent in its main transcript for many
# minutes while its agents still edit files, so this walks every file under each session directory.
hdr "1. Agent sessions active in the last 30 minutes (your own session is listed too)"
found=0
for t in "$HOME"/.claude/projects/*VBToCSharp*/*.jsonl; do
  [ -e "$t" ] || continue
  newest=$(stat -f %m "$t"); base=${t%.jsonl}
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
if [ -s "$PULL_CONFLICTS" ]; then
  attn "UNRESOLVED --pull conflicts in $PULL_CONFLICTS — hand-merge each, then delete the file; do not --pull again or push until then"
  sed 's/^/   conflict: /' "$PULL_CONFLICTS"
fi

# ── 3. Step 0: every remote vs what we last pushed ──────────────────────────────
hdr "3. Step 0 — fetch every staging clone; origin vs the sha we LAST PUSHED (local main)"
for r in hyphen-pb homefront flexkit flexcore; do
  c=$D/$r
  git -C "$c" fetch -q --all --prune 2>/dev/null || attn "$r: git fetch failed — step 0 cannot be trusted"
  cur=$(git -C "$c" branch --show-current)
  om=$(git -C "$c" rev-parse --short origin/main 2>/dev/null)
  lm=$(git -C "$c" rev-parse --short main 2>/dev/null)
  extra=""; [ "$r" = hyphen-pb ] && extra="  R1-UAT=$(git -C "$c" rev-parse --short origin/R1-UAT 2>/dev/null)  R2-UAT=$(git -C "$c" rev-parse --short origin/R2-UAT 2>/dev/null) (main is +$(git -C "$c" rev-list --count origin/R2-UAT..origin/main 2>/dev/null) ahead of it)"
  printf '   %-10s on=%-8s origin/main=%-9s last-pushed=%-9s%s\n' "$r" "$cur" "$om" "$lm" "$extra"
  [ "$cur" != main ] && attn "$r staging clone is parked on '$cur', not main"
  dirty=$(git -C "$c" status --porcelain | wc -l | tr -d ' ')
  [ "$dirty" -gt 0 ] && attn "$r staging clone has $dirty dirty entries"
  # A REJECTED or aborted push leaves an unpushed "Deploy …" commit as the clone's main; --pull would
  # take it as the base and copy GitLab's OLDER files over our own changes.
  ahead=$(git -C "$c" rev-list --count "origin/main..main" 2>/dev/null || echo 0)
  if [ "${ahead:-0}" -gt 0 ]; then
    attn "$r staging clone main has $ahead UNPUSHED commit(s) from a failed push — before --pull or any push run: git -C $c reset --hard \$(git -C $c merge-base main origin/main)"
    git -C "$c" log --format='      unpushed: %h %cs %s' "origin/main..main" 2>/dev/null | head -3
  fi
  behind=$(git -C "$c" rev-list --count "main..origin/main" 2>/dev/null || echo 0)
  if [ "${behind:-0}" -gt 0 ]; then
    if [ "$r" = hyphen-pb ]; then
      attn "hyphen-pb: $behind upstream commit(s) past our last push — run 'bash tools/deploy_to_repos.sh --pull' BEFORE any push-mode run (even --dry-run)"
    else
      attn "$r: $behind upstream commit(s) past our last push — --pull does NOT cover $r: merge by hand BEFORE any push-mode run (SKILL.md step 3)"
    fi
    git -C "$c" log --format='      %h %cs %an: %s' "main..origin/main" 2>/dev/null | grep -v 'Deploy ' | head -10
  fi
done

# ── 4. Step 0b: branches that may carry open merge requests ─────────────────────
hdr "4. Step 0b — hyphen-pb branches ahead of main (confirm open MRs in the GitLab UI too)"
c=$D/hyphen-pb
for b in $(git -C "$c" for-each-ref --format='%(refname:short)' refs/remotes/origin | grep -vE '^origin/(HEAD|main|R1-UAT)$|^origin$'); do
  n=$(git -C "$c" rev-list --count "origin/main..$b" 2>/dev/null || echo 0)
  [ "${n:-0}" -gt 0 ] || continue
  last=$(git -C "$c" log -1 --format=%ct "$b"); age=$(( (now - last) / 86400 ))
  printf '   %-50s ahead=%-3s last commit %s days ago (%s)\n' "$b" "$n" "$age" "$(git -C "$c" log -1 --format=%an "$b")"
  case "$b" in origin/parked/*) continue ;; esac   # snapshots pushed from another machine, not MRs
  [ "$age" -le 14 ] && attn "branch $b has commits from the last 14 days — check for an open MR"
done
echo "   GitLab UI: https://gitlab.innovatixinc.com/groups/application-modernization/-/merge_requests/?state=opened"

# ── 5. Would the push undo a teammate's work? ───────────────────────────────────
# The deploy rsyncs our trees over each remote, so a teammate change missing from our tree is silently
# undone — and once undone, origin/main == our tree and step 0 looks clean forever. Checked for EVERY
# commit after our last "Deploy <repo>" commit (unshipped, regardless of date) plus the last $DAYS days
# (to catch changes an earlier deploy already overwrote). Judged by whether the teammate's distinctive
# lines are in our file, so our own edits to the same file and the deploy's transforms don't false-alarm.
# --no-renames: a rename is a delete + add, so the old path is checked too.
hdr "5. Teammate changes on each origin/main — are they in OUR tree?"
problems=0
for r in hyphen-pb homefront flexkit flexcore; do
  c=$D/$r
  unshipped=$(unshipped_commits "$r")
  # Only hyphen-pb takes teammates' merges; on the push-only repos every non-deploy commit is our own
  # history, so only commits after our last deploy (someone else pushed there) are worth checking.
  historic=""
  [ "$r" = hyphen-pb ] && historic=$(git -C "$c" log --since="$DAYS days ago" --no-merges --format='%h|%cs|%an|%s' origin/main 2>/dev/null | grep -v '|Deploy ')
  commits=$(printf '%s\n%s\n' "$unshipped" "$historic" | awk 'NF && !seen[$0]++')
  [ -z "$commits" ] && continue
  echo "   [$r -> $(src_for "$r" | sed "s|$V/||")]"
  while IFS='|' read -r sha date author subj; do
    [ -z "$sha" ] && continue
    files=$(git -C "$c" show --name-only --no-renames --format='' "$sha" 2>/dev/null | grep -v '^$')
    [ -z "$files" ] && continue
    tag=""; printf '%s\n' "$unshipped" | grep -q "^$sha|" && tag="  [NOT YET SHIPPED FROM HERE]"
    printf '   %s %s %s: %s%s\n' "$sha" "$date" "$author" "$(echo "$subj" | cut -c1-56)" "$tag"
    while IFS= read -r f; do
      [ -z "$f" ] && continue
      res=$(classify_file "$r" "$sha" "$f"); state=${res%%|*}; text=${res#*|}
      printf '      %-58s %s%s\n' "$f" "$([ "$state" = ok ] || echo "$state — ")" "$text"
      if [ "$state" != ok ] && [ "$state" != NOTE ]; then
        problems=1
        attn "$state: $r $f ($sha, $author) — $text"
      fi
    done <<< "$files"
  done <<< "$commits"
done
[ $problems = 0 ] && echo "   every teammate change checked is present in our trees"

# ── 6. Does a working-tree edit UNDO a recent commit? ───────────────────────────
# Seen 2026-09-13: a stale-copy edit restored a file byte-for-byte to the blob BEFORE a fix.
hdr "6. Uncommitted edits that exactly undo a recent commit"
undo=0
for spec in "app|$HF" "FlexKit|$V/FlexKit" "FlexCore|$V/FlexCore"; do
  name=${spec%%|*}; dir=${spec#*|}
  while IFS= read -r f; do
    { [ -z "$f" ] || [ ! -f "$dir/$f" ]; } && continue
    ours=$(git -C "$dir" hash-object "$dir/$f")
    for sha in $(git -C "$dir" log -8 --format=%h -- "$f" 2>/dev/null); do
      parent=$(blob_at "$dir" "${sha}^" "$f"); mine=$(blob_at "$dir" "$sha" "$f")
      if [ -n "$parent" ] && [ "$ours" = "$parent" ] && [ "$ours" != "$mine" ]; then
        printf '   %-9s %-50s UNDOES %s (%s)\n' "$name" "$f" "$sha" "$(git -C "$dir" log -1 --format=%s "$sha" | cut -c1-50)"
        attn "$name/$f exactly undoes $sha — likely a stale copy; check before shipping"; undo=1; break
      fi
    done
  done < <(git -C "$dir" diff --name-only 2>/dev/null)
done
[ $undo = 0 ] && echo "   none"

# ── 7. Library versions ─────────────────────────────────────────────────────────
# NuGet treats an extracted version as immutable: a repack of a version already in ~/.nuget/packages is
# shadowed in the staged build. A version on nuget.org can be unlisted but never replaced.
hdr "7. Library versions"
fk=$(grep -o '<Version>[^<]*' "$V/FlexKit/FlexKit.csproj" | head -1 | cut -c10-)
fc=$(grep -o '<Version>[^<]*' "$V/FlexCore/FlexCore.csproj" | head -1 | cut -c10-)
if [ -d "$HOME/.nuget/packages/flexkit/$fk" ]; then fk_state="BURNED (already extracted)"; else fk_state="fresh"; fi
fc_check=$(curl -s --max-time 15 https://api.nuget.org/v3-flatcontainer/flexcore/index.json | python3 -c '
import sys, json
want = sys.argv[1]
try: vs = json.load(sys.stdin)["versions"]
except Exception: print("UNKNOWN|nuget.org unreachable"); sys.exit()
key = lambda v: tuple(int(x) if x.isdigit() else 0 for x in v.split("-")[0].split("."))
latest = max(vs, key=key)
if want in vs: print(f"PUBLISHED|{want} is already on nuget.org (latest {latest})")
elif key(want) <= key(latest): print(f"LOWER|{want} is not above the latest published {latest}")
else: print(f"OK|{want} is above the latest published {latest}")' "$fc" 2>/dev/null)
printf '   FlexKit  csproj %-9s %s   (uncommitted entries: %s)\n' "$fk" "$fk_state" "$(git -C "$V/FlexKit" status --porcelain | wc -l | tr -d ' ')"
printf '   FlexCore csproj %-9s %s   (uncommitted entries: %s)\n' "$fc" "${fc_check#*|}" "$(git -C "$V/FlexCore" status --porcelain | wc -l | tr -d ' ')"
printf '   feed: %s\n' "$(ls "$H/Deploy/local-packages" 2>/dev/null | tr '\n' ' ')"
[ "$fk_state" != fresh ] && attn "FlexKit $fk is burned — bump FlexKit.csproj if ANY FlexKit change (code or asset) ships this run"
case "${fc_check%%|*}" in
  PUBLISHED|LOWER) attn "FlexCore: ${fc_check#*|} — bump FlexCore.csproj above it before the next publish" ;;
  UNKNOWN) attn "FlexCore version could not be checked against nuget.org" ;;
esac

# ── 8. The installed skill matches its canonical source? ────────────────────────
hdr "8. Skill install"
if [ -d "$INSTALLED" ]; then
  if diff -rq -x .DS_Store "$CANON" "$INSTALLED" >/dev/null 2>&1; then echo "   $INSTALLED matches Docs-Skill"
  else echo "   installed copy DIFFERS from Docs-Skill"; attn "skill install is stale — run: bash $CANON/install.sh"; fi
else echo "   not installed"; attn "skill not installed — run: bash $CANON/install.sh"; fi

# ── Summary ─────────────────────────────────────────────────────────────────────
hdr "ATTENTION (${#ATTN[@]})"
if [ ${#ATTN[@]} -eq 0 ]; then echo "   nothing — clear to proceed to commit / build / harnesses"
else for a in "${ATTN[@]}"; do echo "   - $a"; done; fi
