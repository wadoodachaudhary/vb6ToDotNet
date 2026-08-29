# vb6ToDotNet handoff set — 2026-08-23 (revised 2026-08-29)

Read in this order.

| # | File | What it covers |
|---|---|---|
| 1 | `AGENT-CONTEXT-vb6ToDotNet.md` | **Start here.** Portable context: memory locations, repo topology, GitLab access, the deploy sequence, standing rules, bench lane. Plain markdown — point a new agent at this file. Current as of 2026-08-29: `main` ← HomeFront and `R1-UAT` ← HomeFrontPB (the old "merge main into R1-UAT" step is retired), step 0 pulls **both** remotes, and sync rule 4b is **PB → HF only**. |
| 2 | `vb6ToDotNet-Workstreams-20260823.html` | The four parallel sessions, the bench/live lane split, and the approval gate. FlexKitTester, the zoom facility, the app workstream. |
| 3 | `HomeFront-Handoff-20260823.html` | Repos, GitLab/GitHub auth, secrets, the 8-step deploy sequence, branches, open items. |

Published (private until shared):
- Workstreams — https://claude.ai/code/artifact/eaec0d6c-d5f6-45e7-8531-47642d2aec32
- Handoff — https://claude.ai/code/artifact/5309095a-ce48-4c4d-acdf-b920322ba8b4

## Machine-local context these files describe

- Live memory store: `~/.claude/projects/-Users-wadood-projects-VBToCSharp-HomeFront/memory/` (126 files as of 2026-08-29; 101 when this handoff was written — the count grows every session, so use it only to tell this store apart from the stale 6-file one below, never as a checksum)
- New memory entry added: `vb6todotnet_workstream_map.md`, indexed near the top of `MEMORY.md` (line 5 as of 2026-08-29 — the index grows downward from line 1, so this position shifts as newer entries are added above it)
- A **stale** 6-file memory store also exists under the `…-HomeFront-MobileSource` slug — verify which one loads at session start.
