# Recipe 10 — Page-level Tab navigation (VB6 TabIndex → PageNavigationGraph)

VB6 forms get keyboard flow for free from per-control `TabIndex` plus the
grids' `TabBehavior` (`flexTabCells` = Tab moves cell-to-cell and leaves only
at the terminal cell; `flexTabControls` = Tab leaves immediately). The web
equivalent is FlexKit's PageControl graph.

## Pattern

1. Wrap the page root in `<PageControl NavigationGraph="@MyGraph">`.
2. Declare one `PageNavigationNode` per VB6 control, ordered by the VB6
   TabIndexes (or by an owner-directed sequence — comment any deviation):

```csharp
private static readonly PageNavigationGraph MyGraph = new(
    new PageNavigationNode[]
    {
        new("Toolbar",  ".toolbar-cntr",  -100, Mode: PageNavigationNodeMode.Descendants),
        new("Tree",     ".tree-panel .fx-treegrid", -50, Mode: PageNavigationNodeMode.Element),
        new("FieldA",   ".my-field-a", 0, Mode: PageNavigationNodeMode.Element),
        // …fields in VB6 TabIndex order…
        new("ItemsGrid", ".my-grid", 36, Mode: PageNavigationNodeMode.Element)
    });
```

Nodes whose selector matches nothing (a detail panel for a different node
kind) are skipped automatically — declare every panel's fields once.

3. Grid participation — pick the `GridTabNavigationMode`:
   - `PageControl` = VB6 `flexTabControls`: the grid is ONE tab stop, every
     Tab leaves it.
   - **`WrapRowsUntilEdge` = VB6 `flexTabCells`** (added 2026-09-01): Tab and
     arrows walk cells and wrap rows INSIDE the grid; only Shift+Tab on the
     first navigable cell of the first row / Tab on the last cell of the last
     row hand off to the page graph. Pair it with
     `SeedActiveCellOnHostFocus="true"` so tabbing INTO the grid seeds an
     active cell (without a seed Tab/arrows have no anchor and go dead).

```razor
<GridControl TabNavigationMode="GridTabNavigationMode.WrapRowsUntilEdge"
             SeedActiveCellOnHostFocus="true" ... />
```

Mechanics (do not re-implement): the grid renders
`data-fx-grid-tab-edge="none|first|last|both"` from its active cell and
page-control.js takes Tab only when the edge matches the direction; both = an
empty grid, which Tab passes straight through.

## Traps

- **TextAreaControl / wrapper controls**: `CssClass` lands on a WRAPPER, not
  the textarea — `PageNavigationNodeMode.Element` silently skips the node.
  Use the default `Auto` mode (descends to focusable descendants).
  TextBoxControl puts the class on the `<input>` itself, so `Element` is fine
  there. DatePickerControl: target the inner input explicitly
  (`".my-date input"`).
- **Lazy IEnumerable DataSource** (`ItemsData.Where(...)` computed property):
  before 2026-09-01 `ResolveRowIndex` returned -1 for every item on such a
  grid and ALL keyboard cell navigation silently no-oped (activation
  "succeeded" without moving). Fixed in FlexKit by threading the caller's
  display index into `SelectProgrammaticCellAsync` — but if navigation goes
  dead on a grid, a non-IList DataSource is the first suspect.
- Disabled/read-only inputs are skipped by the graph automatically (the
  focusable filter) — no per-field conditions needed.
- Testing: dispatch real `KeyboardEvent`s via JS on `document.activeElement`
  (harness `computer key` does not reach these paths), and wait ~400-500ms
  after each dispatch for the server round-trip before reading
  `td.fx-cell-active` / the edge attribute.

## Load-time focus

VB6 leaves focus on the lowest-TabIndex visible control and never selects a
tree row on load. If the owner wants "tree first node selected on open"
(FEstimateItems, owner directive 2026-09-01): set a pending flag after the
tree data loads, and in `OnAfterRenderAsync` call
`tree.SelectRecordAsync(firstRoot)` (it raises RowSelected, driving the same
handler as a user click) + `tree.FocusAsync()` — deferred to after render so
the TreeGridControl has rebuilt its flattened node list; focus only on a
fresh job open so reloads don't steal the keyboard.

## Tree key spec (owner 2026-09-01, all screens)

- ENTER = double-click (activate/load) — it must NEVER expand/collapse the
  node. TreeGridControl's built-in Enter fires RowActivated (falling back to
  RowDoubleClicked); wire one of them to the form's load action and you are
  done. Space is the toggle key (VB6 parity).
- Right/Left = Windows-Explorer folders: Right expands, then moves to first
  child; Left moves to parent, then collapses. Built into TreeGridControl —
  do not add page handlers for these.
- Per-form Tab behavior comes from the VB6 designer's `TabBehavior`:
  0/flexTabControls → the grid is ONE Tab stop (`GridTabNavigationMode.PageControl`);
  1/flexTabCells → Tab walks cells until the terminal cell
  (`WrapRowsUntilEdge`). Check the .frx designer block, don't guess.

## Dropdown / picker no-trap rules (library-enforced since 2026-09-01)

- DropDownListControl: Escape always closes; a keyboard/pick close hands
  focus back to the trigger (non-editable mode); Tab on an open panel
  commits the highlighted option; Escape on a mounted-but-CLOSED dropdown is
  forwarded to OnKeyDown so an inline-cell host can dismiss the editor
  (wire `OnKeyDown="@(e => { if (e.Key == "Escape") ClearActiveCell(); })"`
  and refocus the tree/grid in ClearActiveCell — FInboxJobs pattern).
- MenuDropDownControl and DialogControl already close on Escape and restore
  focus; FPickList is fully keyboard-operable (arrows + Enter pick).

## Wrapping an EXISTING sized root — use the layout-neutral form

`<PageControl>` renders a real `<div>`. If the page already has a sized root
(`height:100%; display:flex`), wrapping it in a bare PageControl inserts a box
into the height chain and collapses every `height:100%` grid beneath it (the
2026-09-03 FAssembly blank-grid regression). Either make the PageControl the
sized root itself (give it the page's root `CssClass`) or, when wrapping, use:

```razor
<PageControl NavigationGraph="@MyGraph"
             Style="display:contents" ScrollMode="PageScrollMode.None">
```

## Load-time focus via the graph

`FocusFirstNodeOnLoad="true"` focuses the graph's first target on load (VB6:
keyboard lands on the lowest TabIndex). When that node is a grid, add
`SeedActiveCellOnHostFocus="true"` to the grid so the focus also seeds an
active cell. Leave it off on pages that seed their own focus (FEstimateItems).
