# R2-QA grid regression bench

`/r2qa-grid` uses only synthetic data and FlexCore controls. It never opens a database or writes to Jira.

Build FlexKitTester, run it on port 5299 with owner permission, then run these scripts serially:

```sh
node verification/r2qa-batch-browser.mjs
node verification/r2qa-grid-edit.mjs
node verification/r2qa-sort-dropdown.mjs
```

The scripts use Chrome and Playwright. Set `PLAYWRIGHT_MODULE` to an absolute Playwright module path if the bundled Codex runtime is unavailable; `BENCH_URL` and `OUTPUT_DIR` override the URL and artifact directory. Source-JavaScript routing avoids stale precompressed assets. Stop the bench before rebuilding either library.

- Batch browser: compares 20 versus 5 deferred-scroll buffer rows with 5,000 rows and 129 columns, locally and with 300 ms added round-trip latency; checks read-only drag selection, single-cell keyboard feedback, and first-frame dropdown visibility.
- Grid edit: checks single-cell feedback and three-row client-buffered edits committed by Enter, Tab, and clicking away, at 300 ms added round-trip latency. `BASELINE=1` reads the committed grid JavaScript to demonstrate the stale-highlight regression.
- Sort/dropdown: checks descending primary sorting with an ascending tie-breaker, ordinary header sorting replacing custom sort levels, early dropdown selections and Escape at 500 ms added round-trip latency, and checkbox geometry at desktop and narrow widths.

The bench exposes `WindowColumns=true` for investigation only. Column virtualization is not enabled by this HomeFront fix: its keyboard feedback failed the navigation check. The retained optimization is `DeferredScrollOverscanRows=5`.
