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
