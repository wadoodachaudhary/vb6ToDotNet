# Choose Columns Popup Checks

The `/choose-columns-popup` bench uses an in-memory grid with a host
right-click counter, an optional enclosing DialogControl and a layout-commit
counter. It does not open a database or save application layouts.

Build FlexKitTester normally for FlexCore, or with `-p:UseFlexKit=true` for
FlexKit. Run only one target at a time; stop the tester before rebuilding.
The corresponding DLLs are in `bin/Debug/net10.0` and
`bin/Debug/flexkit/net10.0`.

Run against the local tester:

```sh
BENCH_URL=http://127.0.0.1:5301 node verification/choose-columns-popup.mjs
```

`PLAYWRIGHT_MODULE` may point to an installed `playwright/test` module.
`RESULTS_DIR` overrides the default `/tmp/choose-columns-popup` screenshots.

Checks cover:

- Bold black caption rendered as normal, non-editable, non-disabled title text.
- Header dragging without selecting text, with and without 150 ms each-way WebSocket latency.
- Right-click isolation on the header, list, instructions, buttons and backdrop.
- Cancellation without committing working column changes; OK applying order and visibility.
- Close, overlay dismissal, immediate Escape, arrow/Space navigation and reopening.
- Nested-dialog Escape leaving the parent open and the underlying grid usable afterward.
- 1400x900, 1024x600 and 390x740 viewport bounds, screenshots and browser errors.

The bench remains available for manual testing after the automated run. No
HomeFront server needs to be restarted for this isolated verification.

Verified 2026-09-21: 336 Chrome checks per library, plus 300 existing
InputDialogBrowserChecks against FlexKit. Active HomeFront and FlexCore.Showcase
non-incremental builds pass. Temporary test servers were stopped after testing.

Caption follow-up 2026-09-21: 352 Chrome checks pass against FlexKit after adding
caption assertions. The mirrored FlexCore stylesheet matches exactly, and both
required non-incremental host builds pass. Desktop/narrow screenshots reviewed;
temporary bench stopped without restarting the owner's apps.
