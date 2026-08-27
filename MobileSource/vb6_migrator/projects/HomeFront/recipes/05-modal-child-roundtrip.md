# RECIPE: Modal child form with a result round-trip (references: FCustomQuote hosted by FInboxCustomQuote; FAddPricelist hosted by FPriceList; FModelDimensions)

Use when VB6 shows a child form `vbModal` from inside another form and reads values back after `.Show vbModal` returns (ByRef args / public members). Per the app rules, vbModal fidelity is KEPT for child-dialog launches from inside a form (menu/taskbar launches are non-modal MDI panes instead).

## Child form (the migrated F* page)
1. Dual-shell markup: define the form body once as a RenderFragment (`@{ RenderFragment formBody = @<text>…</text>; }`), then render it inside a `DialogControl` when `IsModal` (Header, sized, `ShowCloseIcon`, `CloseOnEscape`, `OnClose="ClosePage"`, one-way `Visible="true"` so a cancelled close keeps the window open) and as the bare page otherwise. `PageTitle` only when `!IsModal`.
2. API: `[Parameter] public bool IsModal`, `[Parameter] public EventCallback<TResult> OnClose`, and a result record carrying the VB6 ByRef quartet plus the saved flag, e.g. `public record CustomQuoteResult(bool Saved, decimal Cost, decimal Sell, decimal Qty, string Uom);`
3. `Saved` means "SaveData actually wrote" (a dirty save executed) — VB6 `mSaved`. Closing clean returns Saved=false. The out-values are the LIVE state at close time.
4. `ClosePage()` = VB6 Form_Unload: run the Yes/No/Cancel save prompt (`MessageBoxControl`; **Cancel aborts the close**), then `OnClose.InvokeAsync(result)`. Never navigate when `IsModal` — disable or guard any toolbar action that would route away (Preview/Takeoff), since it would tear down the host pane.

## Host form
1. `@if (_showChildModal && _childRow != null) { <FChild IsModal="true" ...params marshalled from the current row... OnClose="OnChildClosed" /> }` — the `@if` re-instantiates the child fresh each open.
2. Marshal the VB6 argument list verbatim (including quirks like Qty==0 → pass 1, ReadOnly = view filter condition).
3. Triggers: the VB6 gestures (double-click via the grid events bundle's `OnRecordDoubleClick`, Enter via `OnHostKeyDown` with `EditOnEnterKey=false` on that grid so Enter means "open child" not "edit cell") plus any explicit button. Add a re-entry guard.
4. `OnChildClosed(result)`: hide the modal; when `result.Saved`, write back into the SAME row exactly as VB6 did (watch for extended-vs-unit values — e.g. VB6 writes `cost * qty` into the Cost cell but unit `price` into Price), mark the row + form dirty, `StateHasChanged()`.

## Traps
- A modal child navigating with NavigationManager/ParentNavigate destroys the host tab silently — always suppress navigation in modal mode.
- `Dialogs.ConfirmAsync` is OK/Cancel only — the VB6 Yes/No/Cancel close prompt needs `MessageBoxControl` (`MessageBoxButtons.YesNoCancel`), with Cancel aborting the close.
- Pre-edit state capture for validation flows (e.g. restoring Status when a reason prompt is cancelled) uses the grid events bundle's `OnCellEdit` (fires at edit start) — `OnCellSave` is post-commit and too late.
- PB scope rule: if the child form does not exist in HomeFrontPB, the host's PB twin gets only the dependency-free subset of the change — never add the child form to PB (operating rule 4b).
