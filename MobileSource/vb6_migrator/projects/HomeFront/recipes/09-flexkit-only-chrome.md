# Recipe 09 — FlexKit-only chrome (no raw HTML controls, no hand-rolled widgets)

**Standing owner rule (2026-08-30, reaffirmed on the FEstimateItems purge):**
every migrated form's UI must be FlexKit controls. Raw `<button>`, `<input>`,
`<select>`, `<textarea>`, hand-rolled dropdown menus, hand-rolled splitters,
and div-based modals are RED FLAGS. **A FlexKit-purity pass is part of every
form audit** — when auditing a migrated form against VB6, list raw-HTML
controls as findings even if nobody asked.

## Mapping table (VB6 → FlexKit)

| VB6 construct | FlexKit control | Notes |
|---|---|---|
| Toolbar (`Toolbar` OCX) | `ToolbarControl` + `ToolbarButtonControl` (or `ButtonControl BareStyle` with app CSS) | icon via `IconSrc="images/32/x.ico"` or an `<img>` child (sanctioned) |
| Toolbar `ButtonMenu` / split button | `MenuDropDownControl SplitButton="true"` + `MenuActionControl` items | `PrimaryClick` = the button's default action; items auto-close the menu (`ParentMenu.RequestClose`) |
| `PopupMenu` off a toolbar button (e.g. a View menu) | `MenuDropDownControl` (non-split) + `MenuActionControl` / `MenuSeparatorControl` | active item: `Icon="✓"` + `CssClass="is-selected"` (FItems idiom). NEVER hand-roll a positioned menu div + backdrop div |
| VB6 `Slider`/pane splitter | `SplitterControl` (`PrimaryContent`/`SecondaryContent`, `@bind-PrimarySize`, `OnResizeCommitted` → per-user prefs) | never a pointer-drag div + fullscreen overlay |
| Label acting as a link (`lblJobNumber`) | `ButtonControl BareStyle="true"` + link CSS | no raw `<label @onclick>` |
| Modal | `DialogControl` / MessageBox service | no `.xxx-modal` div stacks |

## Sanctioned raw HTML (display-only, established across 20+ forms)
- `<label>`/`<span>` field captions and section titles (bold-blue `job-section-title` style).
- `<img src="images/32|16/*.ico" onerror="this.style.display='none'">` icon glyphs
  (inside FlexKit buttons or grid cell templates).
- `tbar-separator` divs, `loading-overlay`, detail-panel `span.detail-label/value` rows.
- Grid/TreeGrid column `<Template>` wrapper divs/spans whose only handlers live on
  FlexKit controls inside them (the FJob gProperties value-cell idiom).

## CSS gotchas when converting
- A FlexKit component's root carries the LIBRARY scope attribute — page CSS must
  reach it via `::deep` from a page-scoped ancestor (e.g. `.toolbar-host ::deep .tbar-btn`).
  Content INSIDE `PrimaryContent`/`ChildContent`/`Template` fragments keeps the PAGE
  scope, so plain selectors still match there.
- The library sizes `.fx-menu-primary`/`.fx-menu-trigger` as 22px menu-bar buttons;
  VB6 icon-over-label toolbar buttons need `height:auto !important` overrides
  (donor: FItems `.fitems-view-primary`).
- Panel look: restyle `.fx-menu-action`/`.fx-menu-separator` under a `PanelCssClass`
  to the VB6 white popup-menu look (donor blocks in FItems.razor.css:92-215 and
  FEstimateItems.razor.css `tbar-menu-panel`).

## Grid checkbox columns select their row (GridControl fix 2026-08-30)
`ColumnType.CheckBox` cells render a `CheckBoxControl` with click-propagation
stopped, so `HandleCellClick` (the row-selection site) never fires.
`HandleCheckboxMouseDown` therefore ALSO calls `SelectRow` (plain VB6 semantics:
clicking a boolean cell moves the row cursor; Ctrl/Shift fall through to
range/toggle; FullMultiSelect grids preserve an armed multi-selection).
When porting a form with selectable checkbox columns, no page-side wiring is
needed — but test multi-select + mass-edit FIRST after touching this path.
