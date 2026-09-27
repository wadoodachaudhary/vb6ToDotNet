# Merge requests

Teammates' work reaches us as GitLab merge requests on **`hyphen-pb`** only. Verified 2026-09-26:
`git ls-remote origin 'refs/merge-requests/*/head'` returns 8 refs (!34–!41) in the hyphen-pb clone and
**zero** in the `homefront` and `flexkit` clones (`flexcore` is GitHub, so it has no such refs at all).
All MR work to date is app work.

## git cannot tell you an MR's state

`refs/merge-requests/N/head` exists for every MR ever opened — open, closed and merged alike, and
nothing distinguishes them. "Its head is not in main" does **not** mean "open": a closed MR, or one
merged into another MR's branch, looks identical (memory note `mr_state_only_from_gitlab`).

The cost, in full: !39 (`89c784a`, branch `HHM-1057-persist-on-first-use`) was read as open from the
refs and taken into main on 09-22 (`afaf1b2`). The owner had **closed** it on 09-21, and its
description asked for a check on a real division before merging. It shipped, and was reverted on
09-23 on the owner's instruction (`8e10e6d`). The same misreading then told the owner that
!36/!38/!39 were "still open"; all three were closed (`reference/history.md:105-107`).

Nor is a branch an MR. hyphen-pb carries 26 branches, several of which never had one —
`wip/festimateitems-local-20260914` (`d145722`) is recorded as "Not an MR, not applied"
(`reference/history.md:177`).

**Rule.** When reporting MR status, say "not in main" rather than "open" unless the GitLab UI said so.
If the UI is unreachable, list the candidate MRs to the owner and ask; never decide from git refs.

### The one question git does answer, and what the answer means

```bash
git -C /Users/wadood/projects/VBToCSharp/HomeFront/Deploy/repos/hyphen-pb \
    merge-base --is-ancestor <mr head sha> origin/main
```

Ancestor = the MR's own commits are in main, i.e. **someone merged it in GitLab**. Today only !34
(`40ef498`) and !40 (`1ddea12`, merged by the owner as `9a3f78a`) pass. !35–!39 and !41 all fail —
and their content is on main anyway, because we take MR content by re-applying it into the source
tree, never by merging the MR branch. So a failing `--is-ancestor` says nothing about whether the
content is present. Judge that per file, by the distinctive lines (below).

## Reading the list

```
https://gitlab.innovatixinc.com/groups/application-modernization/-/merge_requests/?state=opened
```

Swap `state=merged` or `state=all` as needed. Use **Claude-in-Chrome** (`mcp__claude-in-chrome__*`):
it carries the owner's already-authenticated Chrome session. The sandboxed Browser pane
(`mcp__Claude_Browser__*`) is a separate, isolated browser and always hits the sign-in wall, even
after the owner has signed in elsewhere (memory note `gitlab_mr_author_vs_commit_author`:33-39).

If Chrome is not signed in either, ask the owner to sign in — never enter credentials into the login
form (`reference/rules-and-traps.md:24`). Say plainly that the extension was unreachable; twice the
round ended with MR work simply not done and the history row recording it: "GitLab MR close NOT done
(not signed in, Chrome extension disconnected)" (`reference/history.md:12`, and again at `:134-136`).

## Authorship: the MR author is never in git

Every commit made on this machine is authored `wadood <wc2595@columbia.edu>`. The only other identity
in the whole history is `Wadood Chaudhary <wchaudhary@innovatixinc.com>`, and that appears only on
bare merge-button commits — zero real content has ever landed under it. A teammate's name appears
nowhere in git. This is structural, not a search gap: GitLab records the account that opened the MR,
git records this machine's `user.*` (memory note `gitlab_mr_author_vs_commit_author`:8-25, checked
exhaustively 2026-09-05).

So: do not run `git log --author`, `--grep`, or a branches-ahead-of-main scan to find a person — a
full archaeology pass on 09-05 found nothing because there was nothing at the git layer to find. Read
the author badges in the MR list. Baaria Chaudhary's authorship of !19/!21/!22/!23/!24/!26 was
established that way, not from git.

## Taking an MR

Take one **only when the owner asks for it in this request**.

```bash
cd /Users/wadood/projects/VBToCSharp/HomeFront
C=Deploy/repos/hyphen-pb
git -C $C fetch origin HHM-922-964-1159-1161            # the branch, while it still exists…
git -C $C fetch origin refs/merge-requests/41/head      # …or the MR ref (works either way, tested 09-26)
MB=$(git -C $C merge-base origin/main FETCH_HEAD)
git -C $C diff --name-only $MB FETCH_HEAD
```

Diff against `$MB`, **never against main** — a main-relative diff also "reverts" everything merged
into main since the branch forked. Then safe-merge each file with base `$MB` and theirs
`FETCH_HEAD`/`origin/<branch>`, using the safe-merge block in SKILL.md step 3 (it never writes
conflict markers into a source tree).

### Verify each file before believing it landed

The same test preflight uses: every added line with 8+ non-space characters must be present in our
file (`scripts/_common.sh:35-52`, `hunks_present`). Run it yourself per file:

```bash
cd /Users/wadood/projects/VBToCSharp/HomeFront
C=Deploy/repos/hyphen-pb; SRC=MobileSource/HomeFront; SHA=<mr tip>; F=<repo-relative path>
git -C $C diff --no-renames "$SHA^" "$SHA" -- "$F" | sed -e '/^+++ /d' | sed -n 's/^+//p' \
  | awk '{t=$0; gsub(/[[:space:]]/,"",t); if (length(t)>=8) print}' \
  | while IFS= read -r l; do grep -Fxq -- "$l" "$SRC/$F" || echo "MISSING: $l"; done
```

Every `MISSING` line is either a hunk you still owe or an ack you still owe (next section). Silence is
the pass. Run 2026-09-26 as a sanity check: `f3ea82c` (!38) on `FEstimateItems.razor` prints nothing,
and `7e4face` on `FAttachments.razor` prints exactly the two `AvailableColumns`/`DefaultColumns` lines
that `preflight-ack.txt` already explains.

### Say how you applied it

The three shapes that have actually occurred, all recorded in the commit body:

- **Byte-identical** — !38 / HHM-1053: "Applied byte-identical to the MR" (`cffe624`).
- **3-way merge against the merge base** — !41: "Applied byte-for-byte by 3-way merge against its
  merge base" (`b021eb1`).
- **By intent** — !39, written against a pre-`622851a` version of the page: "applied by intent rather
  than by patch" (`afaf1b2`), with each behaviour restated and the VB6 line cited.

Also record in the deploy's history row what the **UI** said (OPEN / merged / closed), not what the
refs implied — and, if it later turns out to have been closed, whether the owner kept or reverted it.
That entry is what lesson 10 (`reference/history.md:54-57`) exists to prevent repeating.

### Taking it does not merge it

HomeFront feeds `hyphen-pb` main, so the content reaches main at the next push while the MR stays
**open and divergent** in GitLab (memory note `gitlab_mr_author_vs_commit_author`:50-53). Say that
explicitly in the hand-back and name the MR, so the owner can merge or close it.

## When to ack instead

`Deploy/preflight-ack.txt` stops preflight and the gate re-flagging a teammate change forever. One
line per entry, `<sha> <path> <reason>`; `is_acked` matches on sha + path only
(`scripts/_common.sh:86-93`), so paths are **repo-relative and unqualified** — the live file mixes app
paths (`Components/Pages/Migrated/FOptions.razor`) with FlexCore ones (`Grid/GridControl.razor`) and
nothing distinguishes them. Grouping or dating sections is safe.

Three classes of entry exist, not one:

1. **Deliberately not taken.** The HHM-1075 precedent: `935b451` (MR !37's tip, merged into !36's
   branch) added tree/money hover text; our own `3b45a5e` already titles the tree text via the VB6
   `PrettyName` port (frm:8147-8190), so taking both would double the tooltip. Acked
   2026-09-17 — with the rest of that commit (HHM-1078/1092 + System_Setup job rules) explicitly
   recorded as taken, so the ack cannot be read as rejecting the whole commit. Before writing one of
   these, diff our file against `origin/main` and confirm the rejected hunk is the *only* missing part
   (`reference/history.md:44-46`).
2. **Merged, but your version supersedes the literal lines.** `classify_file` falls through
   `hunks_present` to `REVERT` whenever the teammate's literal lines are absent
   (`scripts/_common.sh:140-146`), so a reworded or reindented merge still flags. Two live entries are
   this class, both merged in `2eb1836` on 09-21: `7e4face` on `FAttachments.razor` (re-indent only;
   its unfiltered `AvailableColumns`/`DefaultColumns` lines are superseded by our
   `IsAttachmentDataField` filters) and `71497be` on `FOptions.razor` (Irfan's line differs only by
   indentation, ours stays inside the `EmailProtocol==2` guard). Write "merged <date> in <sha>" in the
   reason so the entry is not mistaken for a rejection.
3. **Classifier noise from our own out-of-band history.** 12 of the file's 17 non-comment lines are
   one 09-25 event: FlexCore's own pushed source commits read as teammate work. Those are a
   workaround, prune-able once the flexcore clone is back on deploy-snapshot history — see the
   FlexCore/GitHub trap in `reference/rules-and-traps.md`.

## A merged MR we do not hold locally is reverted by our own push

The push rsyncs our tree over hyphen-pb main, so a teammate's merge that we never took is overwritten
by our older copy; afterwards `origin/main == our tree`, the MR still reads Merged, and step 0 looks
clean forever (`reference/rules-and-traps.md:54-60`). It happened to MR !33 (`07faa53`, merged
09-09, reverted by our 09-09 deploy, restored 09-12 in `f50421f` — `reference/history.md:21`).

**The rule that follows: an MR's content only survives once it exists in the local source tree.**
Merged in GitLab is not enough, and neither is "it is on main today". Take it, or ack it; there is no
third state that survives a push.

## Who closes MRs

Never merge or close someone's MR unless asked — standing owner rule
(`reference/rules-and-traps.md:24`). Two things follow:

- The owner may **ask** us to close them, and then it is ours to do: on 09-22 he asked for !35–!39 to
  be closed once their content was on main `34f41b9`; it was not done because GitLab was not signed
  in, and it is still recorded as outstanding (`reference/history.md:134-137`).
- The owner may want an MR **taken but not merged**. For !41 the instruction on 09-23 was explicit:
  take it in that update and do not merge it in GitLab first, because main is the QA-deployed branch
  (`reference/history.md:110-111`, `b021eb1`).

In the hand-back, per MR: what the UI said, how the content was applied (or why not), and the exact
action left for the owner — merge, close, or nothing.

## Standing per-MR decisions

Last confirmed in the GitLab UI on **2026-09-23**. `reference/history.md:71-72` flags that the
branches `HHM-1053`, `HHM-1057`, `HHM-1057-persist-on-first-use`, `HHM-922-964-1159-1161`,
`festimateitems-phase-dropdown` and `wip/festimateitems-local-20260914` have recent commits and that
**no MR has been re-checked in GitLab since** — treat every state below as stale until the UI is read
again.

| MR | State (09-23 UI) | Decision on record |
|---|---|---|
| !33 | merged, then reverted by our own 09-09 push | Content restored 09-12 (`f50421f`). Its PO-vendor plain text is **not** re-applied (HomeFront opens the vendor editor instead) — acked. |
| !34 | UI state never read | Its head `40ef498` is an ancestor of origin/main, so it was merged upstream; pulled 09-17 with Irfan's other commits. |
| !35 (HHM-1057) | closed | Content shipped; not reverted. Still refuses Project in Quote mode — owner 09-13: literal VB6 (`history.md:170`). |
| !36 (festimateitems-phase-dropdown) | closed | Content shipped (`aac9d99`, plus `23bcb3a` for its `935b451` update); not reverted. |
| !37 | merged into !36's branch | Taken via !36. HHM-1075 hover-text half deliberately **not** taken (ack, `935b451`). |
| !38 (HHM-1053) | closed | Content shipped (`cffe624`, byte-identical); not reverted. The commit records that the root cause is a FlexKit hole (`HandleCellDblClick` skips the edit-button check) and was **not** fixed there. |
| !39 (HHM-1057-persist-on-first-use) | closed (by owner 09-21) | Taken 09-22 in error, **reverted** 09-23 (`8e10e6d`) on the owner's instruction. Do not re-take. |
| !40 (HHM-1100) | merged by the owner (`9a3f78a`, no file change) | Content also taken locally (`3d6f63b`). |
| !41 (HHM-922/964/1159/1161) | open | Taken 09-23 (`b021eb1`, + review fix `6272974`); content on main `316afba`. Owner: **close it in GitLab** — still outstanding. Open owner question: HHM-964's `CellEditablePredicate` skips multi-row edits on rows VB6 writes to (VB6 gates only the active cell, frm:3195-3199). |

!35, !36 and !38 are the unrecorded state worth naming: **closed MRs whose content is on main by
omission** — nobody reverted them, and nobody has asked for that.
