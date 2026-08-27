# User guide — setting up a new Claude user on vb6ToDotNet

For the **person** doing the handover, not for Claude. Six steps, about ten
minutes. Follow them in order; step 1 is the one everything else depends on.

Companion files, all in
`/Users/wadood/projects/VBToCSharp/HomeFront/Docs-archive/`:

| File | What it is |
|---|---|
| `START-HERE-new-agent.md` | The prompts you will paste (Parts A, B, C) |
| `AGENT-CONTEXT-vb6ToDotNet.md` | Portable project context |
| `vb6ToDotNet-Workstreams-20260823.html` | The four workstreams and the approval gate |
| `HomeFront-Handoff-20260823.html` | Repos, GitLab, deploy sequence, open items |

---

## Step 1 — set the working folder correctly

**`/Users/wadood/projects/VBToCSharp` is wrong.** No project store exists for it,
so the session starts blank: no memory, no standing rules, and **no warning that
anything is missing**. It is the natural-looking choice, which is exactly what
makes it dangerous — everything appears to work while every project rule is absent.

Set the folder to:

```
/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource
```

Verified 2026-08-23:

| Folder | What loads |
|---|---|
| `…/VBToCSharp` | nothing — blank project |
| `…/VBToCSharp/HomeFront/MobileSource` | the live **102-file** memory store ✅ |

---

## Step 2 — verify before doing anything

```bash
ls ~/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/*.md | wc -l
```

| Result | Meaning |
|---|---|
| **102** | Correct. Continue. |
| **6** | The stale May 2026 store — missing every rule added since. Stop. |
| **0 or an error** | Folder is still wrong. Return to step 1. |

Then confirm GitLab, or nothing can be deployed:

```bash
ssh -T git@gitlab.innovatixinc.com
```

Expect `Welcome to GitLab, @wchaudhary!`. If this fails, that session can read
the code but cannot publish anything.

---

## Step 3 — prime the first session

Open a session and paste **Part A** from `START-HERE-new-agent.md`.

Part A makes Claude report back which memory store it loaded and how many files
it contains. **If it does not say roughly 102, stop and fix the folder.**
Everything downstream assumes those rules are loaded — a session that skipped
this will look fine and quietly ignore every standing rule.

---

## Step 4 — create one session per workstream

Nothing here is automatic. **Pointing Claude at a folder starts one
conversation, not four.** The sidebar grouping appears *because* four sessions
exist, not the other way round — so the four workstreams have to be created
deliberately.

Start a **new session for each**, all launched from that same `MobileSource`
folder:

1. Update repositories for HomeFront/FlexKit
2. HomeFrontPB / HomeFront App
3. FlexKit Tester
4. Application-wide zoom facility

---

## Step 5 — give each session its reading list

Open `START-HERE-new-agent.md`, scroll to **Part C**, and paste the block whose
heading matches the session you just created:

| Session | Paste | Covers |
|---|---|---|
| Update repositories for HomeFront/FlexKit | **C1** | the deploy pipeline |
| HomeFrontPB / HomeFront App | **C2** | migrated forms |
| FlexKit Tester | **C3** | the bench lane |
| Application-wide zoom facility | **C4** | the gated feature |

Each block names its documents and its memory files by exact filename and full
path, so the session pulls the right context instead of whatever happens to
surface.

To read one block without opening the whole file:

```bash
sed -n '156,175p' /Users/wadood/projects/VBToCSharp/HomeFront/Docs-archive/START-HERE-new-agent.md   # C1
sed -n '176,199p' …   # C2
sed -n '200,222p' …   # C3
sed -n '223,247p' …   # C4
```

*(C1–C4 are labels for this guide's convenience, not a Claude Code concept. The
new user only needs to find the heading matching their workstream.)*

---

## Step 6 — read the answers back

Every Part C block ends by asking the session to state specific facts back: the
conditional deploy steps, the file exempt from wholesale copying, the silent SSR
failure, why `ZoomService` is scoped rather than singleton.

This is the checkpoint, not a formality. **If those answers are vague, the
handoff has not taken** — have the session re-read before trusting it with an
edit. Every one of those facts maps to something that has already cost real time
on this project.

---

## The short version

1. Folder → `…/HomeFront/MobileSource` — **not** `VBToCSharp`
2. Verify memory returns **102**
3. Paste Part A; check it reports ~102
4. Create four sessions, one per workstream
5. Paste C1 / C2 / C3 / C4 into the matching one
6. Read the answers back — vague means re-read, not proceed

---

## If the new user is on a different macOS account

Everything above assumes the **same account (`wadood`)**. A different account
inherits none of it — the memory store, the GitLab key, the NuGet key and the
Jira token all live under `/Users/wadood/`. In that case
`AGENT-CONTEXT-vb6ToDotNet.md` is the entire handover and credentials must be
provisioned separately. See Part B of `START-HERE-new-agent.md` for the full
list of what goes missing.

Two of those credentials — the Jira token and the Cognito client secret — are
already committed in git history and overdue for rotation. Rotate rather than
copy them.
