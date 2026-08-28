# Recipe 07 — Visual-parity sweeps (screen + popup screenshots, both lanes)

Capture every migrated screen AND its popups/dialogs in both apps, at the
same resolution, for side-by-side parity review.

## Desktop layout convention

On a 4K desktop, each quadrant is exactly HD (1920x1080):

- **VB6 lane** (source app, e.g. remote Windows session in a browser tab or
  the Windows App): TOP-LEFT quadrant → region `0,0,1920,1080`.
- **Blazor lane** (migrated app in a browser window): TOP-RIGHT quadrant →
  region `1920,0,1920,1080`.

Same pixel size on both sides means screenshots pair 1:1 with no scaling.

## Tooling (Mutarjim app root)

- `parity_sweep.py` — plan-driven runner: executes a JSON click-through plan
  against a screen region and captures named PNGs along the way. Steps:
  `capture`, `click`, `dblclick`, `rclick`, `key`, `hotkey`, `type`, `wait`,
  `scroll`. Coordinates are REGION-RELATIVE.
- `ScreenAutomationService.RunParitySweepAsync(projectDir, lane, plan, region)`
  — the in-app entry point; shells the script.
- Plans live in `{project}/sweeps/*.json`; screenshots land in
  `{project}/screenshots/{lane}/{Form}__{state}.png`.
- Permissions (macOS): the host app needs **Accessibility** (synthetic
  clicks/keys — pyautogui and Quartz both gate on it) and **Screen
  Recording** (`screencapture`). Grant both in System Settings →
  Privacy & Security before the first run.

## Finding the popups to capture (don't guess — read the source)

The dialog inventory for each screen comes from the VB6 code, not from
clicking around blindly:

```bash
grep -oE "F[A-Za-z]+\.(Show|ShowForm|Choose|ShowDialog)[A-Za-z]*" FScreen.frm | sort | uniq -c
grep -oE "(FGenerate|FTakeoff[A-Za-z]*|FCustomQuote)\.[A-Za-z]+" FScreen.frm | sort -u
```

Purchasing Tasks inventory (from HFEst source, 2026-08-27):

| Screen | Form | Dialogs worth capturing |
|---|---|---|
| Inbox | FInboxJobs | (pushes into FEstimateItems; column menu) |
| Custom Requests | FInboxCustomQuote | FCustomQuote (BuildQuote), FPickList, column menu |
| TBD Assignments | FInboxTBDAssignments | vendor FPickList, column menu |
| Prepare Job Quote | FEstimateItems (quote) | FPickList lookups (38 call sites!), FFind (Ctrl+F), FAddons |
| Issue budgets | FEstimateItems (budget) | FGenerate.ValidateData/VerifyBudgets, FEstimateItemsRefreshCosts, FEstimateItemsFormatting (bid rules) |
| Issue PO's | FEstimateItems (PO) | FGenerate, FEstimateItemsFormatting (purchasing rules), FRptViewer (Preview), FPOSendingWizard, FCancelPO |
| Create Manual PO's | FPurchaseOrder | FPickList (14 sites: vendor/JC), FFind, FSplitItems, FPriceComparison |
| Field PO Requests | FFieldPOs | approve/deny MsgBoxes, column menu |

The Blazor lane mirrors the same inventory (FlexKit DialogControl / FPickList /
MessageBoxControl equivalents) — same names, same states.

## Naming

`{Form}__{state}.png` — e.g. `FEstimateItems_budget__main.png`,
`FEstimateItems_budget__FGenerate_validate.png`, `FInboxCustomQuote__FCustomQuote.png`.
Identical names in `screenshots/vb6/` and `screenshots/blazor/` form a pair.

## Sweep discipline

- One plan per screen per lane; start every plan with two `key: escape`
  steps (dismiss strays) and end by closing what it opened (Escape or the
  form's close button) so plans stay order-independent.
- After any click that opens a dialog, `capture` BEFORE interacting with it.
- MsgBox-class popups are cheap to skip; form-class dialogs are the parity
  surface that matters.
- Remote-desktop lag: keep `settle` ≥ 1.2s for the VB6 lane; the Blazor
  lane can run at 0.6s.
