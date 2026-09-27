# Two accounts, one Mac — cold start, hand-off, Jira

Two Claude Code accounts work these repos a week at a time: **Wadood** (wadood@gmail.com) and
**Innovatix** (wchaudhary@innovatixinc.com). Everything below is machine-local; git does not carry it.
Source: memory note `agent_weekly_handoff`, owner setup 2026-09-12.

## What both accounts already share

Claude Code keys its state to the macOS user (`wadood`) and the project path, not to the Anthropic
account — `~/.claude.json` holds ONE `oauthAccount` record, replaced on switch. So both accounts get the
same `~/.claude/CLAUDE.md`, settings, plugins, installed skills, the memory store, and every session
transcript (either account can read or `--resume` the other's sessions). Nothing needs exporting.

## Confirm the memory store before trusting a note

The store is derived from the **git repo root** of the launch directory, and this work spans four roots.
Three of them used to load zero notes, so each is pinned in `.claude/settings.local.json`:

```json
{ "autoMemoryDirectory": "/Users/wadood/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory",
  "autoMemoryEnabled": true }
```

It must be `settings.local.json` — `autoMemoryDirectory` is ignored in a checked-in `settings.json`.
The four pinned roots are the outer `HomeFront`, the nested app, `FlexKit` and `FlexCore`. HomeFrontPB is
deliberately not pinned (frozen; no work starts there). If a session reports a different or empty store,
the pin is missing: read the store directly and say so.

```bash
for r in /Users/wadood/projects/VBToCSharp/HomeFront /Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront \
         /Users/wadood/projects/VBToCSharp/FlexKit /Users/wadood/projects/VBToCSharp/FlexCore; do
  echo "== $r"; cat "$r/.claude/settings.local.json" 2>/dev/null || echo "  NOT PINNED"; done
```

## Day one, incoming agent

1. Confirm the memory store (above), then read `MEMORY.md` line 1 — it points at the hand-off note and
   carries the last deployed heads and versions.
2. Read `VBToCSharp/AGENTS.md` (the binding project rules) and the app's `CLAUDE.md`.
3. Establish the deployed state before touching anything: `bash Docs-Skill/UpdateRepositories/scripts/preflight.sh`.
4. `git log` since the last hand-off in all four repos. FlexKit is on `telerik-parity-20260904`, never main.
5. Check version burn: FlexKit against `~/.nuget/packages/flexkit/`, FlexCore against nuget.org — see
   [nuget-and-versions.md](nuget-and-versions.md).
6. Read the open owner decisions: the newest "Open as of" block in [history.md](history.md) and the top
   section of the hand-off note. Those are the things you must not silently re-decide.

## End of week, outgoing agent

- Nothing important may exist only in a transcript. It goes in a memory note, this skill, or a committed doc.
- Commit every working tree. The deploy rsyncs **working trees**, so handing over mid-edit means the next
  round ships someone's half-finished work.
- Leave the four staging clones parked on `main` and clean (`git -C Deploy/repos/<repo> status`).
- State the machine-local things git does not carry: the memory pins, the NuGet API key in the keychain,
  the local packages feed, Playwright's location, which ports the owner's own instances use.
- Correct or delete notes the week proved wrong. A stale note costs the next agent more than a missing one.
- Say plainly what was deliberately NOT done, and why (owner decision, held publish, unverified area).

## Recording a round (step 12 in full)

1. `MEMORY.md` line 1 — remote heads, versions shipped or held, local commit ids, what is open. It is the
   first thing the next session reads.
2. The hand-off note beside it (`handoff_pending_2026_09_12.md`) — a new dated UPDATE section at the top,
   marking the previous one superseded.
3. A row in [history.md](history.md) plus a refreshed "Open as of" block, in the **canonical** skill copy.
4. Each library's `CLAUDE.md` gets a dated section for that round's own ship-review fixes — FlexKit and
   FlexCore both, since the port is byte-identical.
5. Commit the outer repo **and push it** if it has a remote; a history row left uncommitted or unpushed is
   invisible to the other account.
6. Re-install the skill: `bash /Users/wadood/projects/VBToCSharp/HomeFront/Docs-Skill/UpdateRepositories/install.sh`.

## Jira, at the end of a round

Ticket moves are part of the hand-off, not part of the deploy — do them only when the owner asks, and
report what you moved.

- Fixes go to **Development Review**; a ticket only reaches **Ready for Test** on evidence from a live
  check plus VB6 comparison. A ticket moved on a round's own evidence has been reopened by the owner
  before now — say what the evidence was.
- Attribute completion comments to the ticket's assignee unless the owner says otherwise, through
  HomeFront's `Author: text` relay. Never attribute to ADMIN, and never change a ticket's reporter or
  assignee to set a comment author.
- One short UI-facing line per fixed ticket. No intermediate progress comments.
- Baaria Chaudhary's comments are short fragments without "please retest"; Wadood Chaudhary's are terse,
  at most two lines. (Owner directive 2026-09-19; memory notes `jira_sync`, `jira_workflow_statuses`.)

## Machine-local prerequisites a cold run depends on

| Thing | Where | Notes |
|---|---|---|
| NuGet API key | macOS keychain, `security find-generic-password -s flexcore-nuget-key -w` | never printed or written to disk |
| FlexKit feed | `HomeFront/Deploy/local-packages` | the staged build restores from here |
| Staging clones | `HomeFront/Deploy/repos/{hyphen-pb,homefront,flexkit,flexcore}` | git-ignored, machine-local |
| Skill state files | `HomeFront/Deploy/.update-repositories-gate.state`, `pull-conflicts.txt`, `preflight-ack.txt` | machine-local |
| Playwright | `/Users/wadood/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules` | browser suites read it via `NODE_PATH` or their own default |
| The owner's own instances | ports 5065 / 5071 (and 5270 for QA runs) | never restart or stop them without being asked |

## The scratchpad is wiped on restart

Session scratch under `/private/tmp/claude-501/...` comes back **empty** after a restart, so a round held
overnight can lose its workspace (review output, generated patches, harness fixtures). Back anything a
held round depends on up under `~/.claude`, or regenerate it. Memory note `scratchpad_wiped_on_restart`.
