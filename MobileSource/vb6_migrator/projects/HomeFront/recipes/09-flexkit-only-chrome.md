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

## VB6 Toolbar_ButtonDropDown → MenuDropDownControl OnOpening (added 2026-08-31)

VB6 dropdown toolbar buttons (Style=5) rebuild their popup's items at OPEN
time (`Toolbar_ButtonDropDown` — fresh DB reads, checked flags, stamped
captions). The web analog is the `MenuDropDownControl.OnOpening`
EventCallback (additive FlexKit param, 2026-08-31): fired on the
closed→open transition from both mouse Toggle and keyboard Open,
fire-and-forget so the panel opens instantly and an async handler
re-renders it when the refresh lands. Wire `OnOpening="RefreshXxxAsync"`
and ALSO load the same state when the host record loads, so the first
open never paints stale. Checked VB6 menu items = `Icon="✓"` +
`CssClass="is-selected"` (FItems idiom). Donor: FEstimateItems Snap Shot
(3-slot checked menu, captions carry "(userid  dd-MMM-yy)" — VB6
Format "medium date").

## Report launches: seed conventions (Filter:/Multi:, updated 2026-08-31)

Report XMLs are READ-ONLY; all VB6-side parameterization happens at launch
via `IReportSeedParameters` (`ReportSeeds.Set(...)` then navigate to
`/report-viewer/{escaped xml path}` — query strings never survive FMain).
Three seed shapes, resolved by ReportWriterControl.ShowReportAsync:

| Seed key | Meaning | VB6 analog |
|---|---|---|
| `Name` | single parameter value; a seeded param never prompts | `c.ParameterValue("Name", v)` once |
| `Filter:Table.Field` | host-level row filter injected into the SQL (only for tables provably in the FROM) — for XMLs with no such parameter | report selection formulas |
| `Multi:Name` | value = a U+0002-separated list -> string[] in the param pipeline: no prompt + `IN (...)` expansion via MultiValueParameterRewriter | repeated `c.ParameterValue("Name", v)` (multi-value Crystal param) |

Donor: FEstimateItems OnPreviewPOsClick (Multi:PONumber per POFormat group)
and OnPoHistoryReportClick (plain seeds). Traps: the PO print formats live
in `wwwroot/resources/Estimating/PO Formats/xml/*.xml` (NOT next to the
.rpt files one level up); and navigating the viewer to the SAME route while
its tab is open re-uses the component without re-running the report — close
the viewer tab (or vary the route) for a fresh render.
