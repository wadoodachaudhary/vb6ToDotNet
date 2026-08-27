# RECIPE: Basic AppGridLayout-driven grid (reference: FIntersection.razor, worked example FAssemblyCosts.razor)

Use for every GridControl whose VB6 form calls IniGetGrid/IniPutGrid (AppGridLayout table). If VB6 hardcodes the columns (designer FormatString, no IniGetGrid), keep them hardcoded — fidelity rule.

## Wiring (exact shape)
1. Usings/injects: `@using HomeFront.Models`, `@using HomeFront.Services`, `@inject IGridLayoutService GridLayoutSvc`.
2. Presenter field with the EXACT VB6 form + grid control names (they are the AppGridLayout keys — never invent names):
   `private readonly GridLayoutPresenter<TRow> _xxxGrid = new("FVb6FormName", "gVb6GridName");`
   Optional ctor args: `gridIndex` (control-array index), `instanceKey` (VB6 IniGetGrid InstanceKey arg — e.g. FExportEstimates passes `Rewrite` → "True"/"False"; runtime keys mean the field must be nullable and constructed at load time), `grouped: true` when VB6 passed Grouped:=True, `aliases` (normalized colkey → row property).
3. Load first in OnInitializedAsync:
   `try { await _xxxGrid.LoadAsync(GridLayoutSvc); } catch (Exception ex) { Logger.LogWarning(ex, "..."); }`
   then `LoadDefaultAsync` in its own try/catch, then `if (!_xxxGrid.HasDefault) _xxxGrid.SetDefaultLayout(BuildXxxFactoryDefault());`
4. Factory default = the form's known-good column set as `GridLayoutColumn { Field, Caption, Width(twips), Hidden, ColIndex }`. twips ≈ (px − 16) × 15 (GridLayoutPresenter.WidthPx converts back with /15 + 16, clamped 32..2000).
5. GridControl parameters (add, keep existing ones):
   `AvailableColumns="@_xxxGrid.GetChooseColumnsSchema()" DefaultColumns="@_xxxGrid.GetDefaultColumnsSchema()" OnColumnsChosen="@OnXxxColumnsChosenAsync" OnLayoutChanged="@OnXxxLayoutChangedAsync"`
6. Markup: `<GridColumnsBase @key="XxxGridLayoutVersion">` (Version property from the presenter — forces column re-render). Utility columns (constant Visible=false ids, RowState, selection checkboxes, button columns) render OUTSIDE the layout @if/else, are EXCLUDED from the factory default, and never appear in the chooser. Then:
   ```
   @if (HasXxxGridLayout) {
       foreach (var entry in _xxxGrid.GetOrdered()) {
           var layout = entry.Layout; var fieldName = entry.Field;
           var headerText = string.IsNullOrWhiteSpace(layout.Caption) ? layout.Field : layout.Caption;
           var width = layout.Width is > 0 ? $"{GridLayoutPresenter.WidthPx(layout.Width.Value)}px" : null;
           // special-case templated/PK/checkbox/format columns by fieldName, else:
           <GridColumn Field="@fieldName" HeaderText="@headerText" Width="@width"
                       Type="@entry.ColumnType" TextAlign="@entry.TextAlign"
                       AllowEditing="..." ClipMode="ClipMode.Ellipsis" />
       }
   } else { /* the EXACT previous hardcoded columns — the no-layout fallback */ }
   ```
7. Save-back handlers (per grid, exact VB6 names + same instanceKey as the ctor):
   - Chooser OK: `GridLayoutPersister.MergeAndSaveChooseColumnsAsync(GridLayoutSvc, "Form", "gGrid", instanceKey, _xxxGrid.Layout, result, _xxxGrid.ResolveField); await GridLayoutSvc.FlushPendingGridLayoutsAsync();`
   - OnLayoutChanged (resize/reorder/hide/rename): `GridLayoutPersister.MergeSaveAndFlushAsync(...)` (NOT MergeAndSaveAsync alone — that only stages).
   Both: guard on `Has`, then `BumpVersion()` + `InvokeAsync(StateHasChanged)`.
8. Surface the fallback: append `[gXxx FALLBACK]` to the PageTitle (or the dialog Header for modal components) when the layout didn't load. No silent fallbacks.

## Traps (all hit in production)
- **ResolveField prefers a REAL row property over aliases.** An alias `{"poindex" → "POIndexDisplay"}` is inert if the row also has a `POIndex` property — the layout column then binds/persists against the wrong field silently. Fix: a page-local resolver used consistently in the render loop AND both persister calls AND the chooser schemas, or bind the raw property.
- **`HeaderText="@layout.Caption"` breaks Razor** — `@layout` parses as the layout DIRECTIVE. Write `@(layout.Caption)` or hoist to a local.
- All-hidden saved layouts are treated as absent by the presenter (HHM-233 recovery) — don't re-implement.
- Caption-less layout rows are internal/spare columns; `GetChooseColumnsSchema` already filters them (VB6 FColumns parity) — don't list them manually.
- Blazor `@layout`-loop specifics: inside a C# `else` block use plain `for/foreach` (no `@`), comments as `@* *@`.
