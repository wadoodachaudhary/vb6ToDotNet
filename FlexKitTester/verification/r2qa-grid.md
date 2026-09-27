# R2-QA grid regression bench

`/r2qa-grid` uses only synthetic data and FlexCore controls. It never opens a database or writes to Jira.

Build FlexKitTester, run it on port 5299 with owner permission, then run these scripts serially:

```sh
node verification/r2qa-batch-browser.mjs
node verification/r2qa-grid-edit.mjs
node verification/r2qa-sort-dropdown.mjs
node verification/keyboard-selection-colors.mjs
LONG_NAV=1 CASE=cell-grey node verification/keyboard-selection-colors.mjs
```

The scripts use Chrome and Playwright. Set `PLAYWRIGHT_MODULE` to an absolute Playwright module path if the bundled Codex runtime is unavailable; `BENCH_URL` and `OUTPUT_DIR` override the URL and artifact directory. Source-JavaScript routing avoids stale precompressed assets. Stop the bench before rebuilding either library.

- Batch browser: compares 20 versus 5 deferred-scroll buffer rows with 5,000 rows and 129 columns, locally and with 300 ms added round-trip latency; checks read-only drag selection, single-cell keyboard feedback, and first-frame dropdown visibility.
- Grid edit: checks single-cell feedback and three-row client-buffered edits committed by Enter, Tab, and clicking away, at 300 ms added round-trip latency. `BASELINE=1` reads the committed grid JavaScript to demonstrate the stale-highlight regression.
- Sort/dropdown: checks descending primary sorting with an ascending tie-breaker, ordinary header sorting replacing custom sort levels, early dropdown selections and Escape at 500 ms added round-trip latency, and checkbox geometry at desktop and narrow widths.
- Keyboard colors: checks computed cell backgrounds and focus borders on every animation frame at 0/400/800 ms added RTT. Covers default/grey/green/blue cell selection, grey/blue row selection, disabled row highlighting and disabled selection. Long navigation crosses the 20-step silent server checkpoints. `CASE` selects one case; `LIBRARY_ROOT` selects FlexKit or FlexCore source JavaScript; `BASELINE_MODULE` can point to a pre-fix module to record the old failure without failing the run. Build the tester with `-p:UseFlexKit=true` for a direct FlexKit check.

HHM-1170 follow-up: the original test checked duplicate borders only. The old module painted the departed cell `rgb(182, 200, 221)` while the committed cell-selection shade was `rgb(245, 245, 245)`. The corrected code passes 24 colour/latency cases over 4,646 frames, long traversal, and direct FlexKit default-colour tests. ClientBuffered three-row edits still pass Enter, Tab and click-away. Nonincremental active HomeFront and FlexCore.Showcase builds pass; no QA database or owner-app restart was involved.

The bench exposes `WindowColumns=true` for investigation only. Column virtualization is not enabled by this HomeFront fix: its keyboard feedback failed the navigation check. The retained optimization is `DeferredScrollOverscanRows=5`.
