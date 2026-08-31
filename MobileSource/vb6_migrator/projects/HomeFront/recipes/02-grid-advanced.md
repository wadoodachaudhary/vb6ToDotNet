# RECIPE: Full-feature editing grid + form chrome (references: FAssembly.razor, FItems.razor; FPriceList.razor is the worked application)

Use for item-database-style screens: editable grid, toolbar, views, splitter, keyboard, context menu, per-user state. Everything is FlexKit (`Fx.ControlKit.*`) — hand-rolled `<button>`/popup/fixed-div menus are migration bugs.

## Events bundle (never inline per-cell lambdas)
One `GridControlEvents<TRow>` instance bound via `EventsRef="@_gridEvents"`, populated once in OnInitializedAsync with `EventCallback.Factory.Create`:
- `RowSelected`/`RowDeselected` → maintain a `HashSet<TRow> _selectedRows` (feeds bulk operations) + `_selectedRow`.
- `CellSelected` → track the active cell (`CellSelectEventArgs` has RowIndex/CellIndex, NOT a column name — map CellIndex through the rendered field order).
- `OnCellSave` (`CellSaveArgs` has `ColumnName`) → set `row.RowState = "DIRTY"`, `IsDirty = true`, apply form-specific side effects (e.g. FPriceList price-history shift, price-group fan-out).
- `OnEditButtonClick` → dispatch pickers by `args.ColumnName` (FPickList pattern).
- `OnHostKeyDown = NonRenderingEventHandler.Create<KeyboardEventArgs>((Func<KeyboardEventArgs, Task>)(e => OnGridKeyDown(e)))` — the VB6 gXxx_KeyDown hook. TRAP: method groups / this-only lambdas defeat the non-rendering receiver — keep this exact closure shape.

## Keyboard (VB6 gItems_KeyDown parity)
- Ctrl+Delete → soft-delete the SELECTION (see the database-save recipe: rows leave the DataSource into a pending-delete list immediately — HHM-259, "delete does nothing" otherwise).
- Ctrl+F → FFind dialog on the active cell's column (`GridData` = single-column string lists, `OnFound` → highlight via `RowCssClassSelector` returning "find-match" + scrollIntoView).
- Accept `e.CtrlKey || e.MetaKey` and show both shortcut forms ("Ctrl+S / ⌘S") in any UI text.

## Mass edit + fill-down (VB6 ValidateEdit applies to the whole selection)
- Grid: `BatchEditBehavior="GridBatchEditBehavior.SingleCell"` + `AllowSingleCellColumnMassEdit="true"`.
- Every data column: `AllowCellDragSelection="true"` (FItems opts ALL columns in — it also arms the Enter fill-down).
- Pickers (edit-button dialogs) apply the picked value to every row in `_selectedRows`, not just the clicked row.

## Toolbar / menus / splitter / context menu (FItems chrome)
- Buttons: `ToolbarButtonControl` (`Vertical="true" IconSrc="images/32/x.ico" Label="..." OnClick=...`, `Disabled=` for gating). File pickers: `FilePickerControl` wrapping a ToolbarButtonControl.
- Dropdown menus (e.g. View picker with ✓ on active): `MenuDropDownControl` `SplitButton="true"` with `<PrimaryContent>` (icon+label) + `<ChildContent>` of `MenuActionControl` items. NOTE: PrimaryContent renders ONLY in SplitButton mode.
- Splitter: `SplitterControl` with `@bind-PrimarySize="_paneWidth"` + `OnResizeCommitted` → persist via `IUserPreferencesService` (drag-end only, never per-tick). PrimaryContent/SecondaryContent slots.
- Context menu: `ContextMenuControl Visible= X= Y= MinWidth= OnClose=` with `ContextMenuItemControl`/`ContextMenuSeparatorControl`; open from a wrapper div's `@oncontextmenu` (+`:preventDefault`). No Windows-only Cut/Copy/Paste menus — the browser owns the clipboard.

## Per-user state (VB6 IniGetForm/IniPut → IUserPreferencesService)
- `@inject IUserPreferencesService Prefs`; section = the VB6 form name. Load in OnInitializedAsync (`Prefs.Get("FForm","Key")` with VB6's defaults), save at the interaction point (`Prefs.Put`).
- What VB6 kept per-user stays PER-USER: view index, decimals, splitter px. Writing user prefs to the division-wide AppOptions table is a cross-user leak (FPriceList had exactly this bug). AppOptions is for division-level options only, read `WHERE DivisionID IN (0, @Div) ORDER BY DivisionID` (division row wins).

## View switching (FItems / FPriceList)
Views change the tree/grid queries AND the grid layout: per-view AppGridLayout identity (FItems: instanceKey = view index; FPriceList: per-view grid tag "gItems.{view}" probed with the presenter's candidate-list `LoadAsync(svc, candidates...)` — 3-call fallback ending at the bare tag). Save layout changes under the SAME per-view identity that loaded them. Dirty-gate every view switch (prompt Save first).

## Cardinal rules
- VB6 form/grid/option/colkey NAMES are config keys — never rename them in lookups.
- Save starts disabled, enabled only when dirty; dirty-gate navigation and destructive flows.
- No JavaScript except DOM-only capabilities (scrollIntoView, caret geometry) — tiny, flagged, with C# fallback.

## Cell edit-button ("…") convention (owner directive 2026-08-27)

VB6's ComboButton appears on the ACTIVE cell only — mirror that: use plain
`ShowEditButton="true"` (+ `ShowEditButtonPredicate` for per-row gating) and
wire `EventsRef.OnEditButtonClick`; the button then renders only when the cell
is clicked or reached by arrow/Tab. Do NOT default to `AlwaysShowEditButton`
(reserve it for deliberate exceptions). The button's subtle chip styling lives
in each app's `wwwroot/css/fx-shared.css` under
`.fx-grid .fx-cell .fx-cell-action-btn` (with !important — a bundled stylesheet
otherwise strips the border/background). The chip sits FLUSH RIGHT: the cell's
`padding: 0 4px` would leave a 4px gap, so the chip carries
`margin-right: -3px` → ~1px of daylight to the cell border (owner directive).
Vertically it is `height: 15px; padding: 0 3px 3px` — 16px spills the ~17px
row, and the bottom padding lifts the baseline-hugging "…" glyph to center.

## Runtime column-parameter changes need a @key bump

Changing a `GridColumn` parameter at runtime (`Format` for VB6
SetDecimalPlaces-style More/Less buttons, captions, etc.) does NOT repaint
already-rendered cells — the grid's render caching keeps the old markup. Put a
version value in `<GridColumnsBase @key="...">` and bump it (e.g.
`GridLayoutPresenter.BumpVersion()`) whenever such a parameter changes; the
columns re-create and cells re-render. Dirty state and edited values survive
the re-key. Display-only changes (decimals) must NOT touch the form's dirty
flag — VB6's SetDecimalPlaces never sets mDirty.

## VB6 apply-and-clear numeric header fields

VB6 forms often have numeric textboxes that APPLY on Enter or focus-out
(KeyDown vbKeyReturn → Validate) and then clear themselves (Price Change /
RoundTo on FPricingWorksheet). Port them as
`NumericTextBoxControl TValue="decimal?"` + `ValueChange` handler that applies
to all rows and sets the bound value back to null. The control commits on
Enter/NumpadEnter, Tab, and blur (it buffers input and commits Enter itself —
browsers do not fire `change` on Enter for number inputs), and it parses
nullable TValue via the underlying type (blank ⇒ null). Guard handlers with
`!x.HasValue` and never mark dirty unless rows actually changed.

## Editable grid hosted in a DialogControl (list-entry dialogs)

Pattern (FPricingWorksheetLayouts): a VB6 multiline-textbox list becomes a
single-column FlexKit grid — hidden PK column + one editable column, header
hidden (`::deep thead.fx-grid-header { display: none }`), Batch mode with
EditOnEnterKey + EditOnActiveCellClick, RowHeight/MinRowHeight ~18, a trailing
blank ENTRY row (OnCellSave on it appends the next blank and re-opens the
editor there when the active cell stayed put), Delete removes the selected row.
Two traps when the editor must open the moment the dialog opens:
1. Set `AutoFocus="false"` on DialogControl — its default focus pass lands
   AFTER your BeginEditCellAsync and the focusout tears the batch editor down.
2. BeginEditCellAsync silently no-ops during the dialog's first renders (grid
   columns/view not built). Retry on REAL time (bounded ~10×40ms loop in
   OnAfterRenderAsync) until a wired OnCellEdit confirms the editor opened;
   StateHasChanged-chained retries share one synchronous cascade and all fail.

## VB6 ColorizeItems-style conditional row formatting (added 2026-08-30)

VB6 forms often repaint grid rows from user-editable `Format_*` AppOptions
(e.g. FEstimateItems.ColorizeItems frm:7486: negative qty/rate, zero pretax,
overridden rate, missing-data + warning icon; defaults in Options.cls:1044).
Port pattern (donor: FEstimateItems):
- `RowCssClassSelector` returns ONE fmt class following VB6's overwrite order
  (LAST matching rule wins; a locked row short-circuits all conditionals).
- Colors/weights are NEVER hardcoded: read the `Format_{rule}_{ForeColor,
  BackColor,FontStyle}` AppOptions (division-aware `IN (0,@div)`), convert OLE
  colors (value is BGR; `&H80000000` flags a system index — map 0x05→#FFF,
  0x08→#000, 0x11→#808080), and emit CSS custom properties in a `style=` on a
  page-scoped wrapper; the static `.razor.css` rules consume `var(--...)`.
- Guard rule backgrounds with `:not(.fx-selected)` so row selection stays
  visible; fore/font apply unconditionally.
- Missing-data messages recompute at row load AND in OnCellSave (VB6 re-runs
  ColorizeItems per edit); the warning column is a GridColumn Template
  rendering `images/16/Warn.ico` with the message list as tooltip.

## Single-cell mass edit needs FOUR opt-ins together (2026-08-30)

`EditMode.Batch` alone gives in-cell editing but NOT the VB6 drag-a-band-and-
type mass edit. The full FItems/FPricingWorkSheet stack is:
1. `BatchEditBehavior="GridBatchEditBehavior.SingleCell"` on the GridControl,
2. `AllowSingleCellColumnMassEdit="true"` (the actual gate),
3. `SelectionSettings.Mode = SelectionMode.Cell` (+ Type=Multiple),
4. `AllowCellDragSelection="true"` on every column (also enables Enter
   fill-down — the flag is overloaded),
plus donor-parity `TextEditorTypingBehavior="TextBoxTypingBehavior.ClientBuffered"`.
Missing any one of 1-4 degrades silently: typing lands in a single-cell editor
and only the active row commits. Symptom to recognize: "multi-edit lets me type
but only the first row commits".

## Per-row/per-cell edit locks (VB6 BeforeEdit parity, added 2026-08-30)

VB6 grids gate edits per ROW in `g*_BeforeEdit` (locked once money is
committed). Port pattern: GridControl's `CellEditablePredicate` —
`Func<TValue,string,bool>` (item, field) — evaluated in ADDITION to
column-level AllowEditing at every edit start, all mass-edit/fill-down
fan-out targets (locked rows silently skipped, matching VB6's per-row
ValidateEdit re-check), checkbox toggles (locked boxes render disabled),
the OnTypeAheadCommit handoff lists, and the fx-cell-editable cue.
Write the page predicate as a verbatim port of the VB6 BeforeEdit Cancel
logic (row flags first, RowType short-circuit, blanket row gate, then the
per-column Select Case — keep VB6's Case Else=locked whitelist and any
latent clobber bugs deliberately, with comments). Keep it O(1) — it runs
per rendered cell. Donor: FEstimateItems.CanEditEstimateCell.

## VB6 ComboList cell buttons + EditMaxLength + value lists (added 2026-08-31)

VB6 `BeforeEdit` side state maps to three GridColumn parameter groups on the
layout-driven column loop (donors: FItems `_lookupPickerFields`,
FEstimateItems ComboList batch):

1. `.EditMaxLength = n` → `MaxLength="@n"` (keep VB6's cap even when the DB
   column is longer — e.g. Location varchar(200) but VB6 caps 50).
2. `.ComboList = "|..."` (typed + button) → keep `AllowEditing` and add
   `ShowEditButton="true"`. `.ComboList = "..."` (picker-only) → same, with
   typing already blocked (`AllowEditing=false`); dblclick opens the picker
   (`OpenEditButtonOnDoubleClick` default). Gate per-row button visibility
   with `ShowEditButtonPredicate` = the page's BeforeEdit predicate — VB6's
   Cancel meant no edit ⇒ no button. A conditional button (VB6
   `IIf(option, "|...", "")`) is just a bool in the helper (JCExtra/
   Use_Timberline).
3. `.ComboList = GetComboList(field)` (leading-pipe editable combo) →
   `EditOptions="@values" AllowCustomEditOptionValue="true"`; load the
   distinct-values lists per data load (VB6 re-queried per BeforeEdit); an
   empty list must yield `EditOptions=null` (plain text cell, VB6 `""`).

The `gItems_CellButtonClick` body becomes ONE `OnEditButtonClick` handler on
GridControlEvents dispatching per `args.ColumnName`: capture
`(row, field, targets)` where targets = `GetSelectedRecordsForColumn(field)`
if the clicked row belongs to a >1 selection else the clicked row (the web
analog of `For Row = Min(.Row,.RowSel) To Max`), then `FPickList.Open` with
the VB6-exact SQL and a `Cell*` context; the `OnPickListSelected` branch
fans the ONE pick over targets, re-gating EVERY row through the BeforeEdit
predicate (VB6 re-calls `gItems_BeforeEdit` per row). Watch for VB6's
ungated exceptions (FEI: Assembly range-writes ignore the gate; Formula
writes only the CURRENT cell — `FillStyle=flexFillSingle` makes the loop's
`.Text` writes hit the anchor N times, so port it clicked-row-only).
Code+desc column pairs share one picker context; check per-case whether VB6
swaps the SQL column order for the Desc variant (JCCostCodeDesc does,
POIndexDescription does NOT). Long-text cells (`FComments.Edit`) get a
dedicated FComments instance; Formula gets the real FFormulaEditor.
Multi-step VB6 flows that prompt per row (variance category) batch into ONE
picker after the fan-out — chain pickers via context handoff and wire
FPickList `OnCancel` to clear the pending state (do NOT keep VB6's
dirty-on-cancel AfterEdit quirk).
