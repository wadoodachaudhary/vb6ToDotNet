# Recipe 06 — Sequential modal pick-list chains (VB6 FPickList.Choose loops)

VB6 tools often chain several MODAL picks in one procedure:

```vb
If Not FPickList.Choose(db, "Model and Option", sql1, ...) Then Exit Sub   ' pick source
If Not FPickList.Choose(db, "Model and Option", sql2, , , , , hide, True) Then Exit Sub ' multi-pick options
If Not FPickList.Choose(db, "Model and Option", sql3, , , , , , True) Then Exit Sub     ' multi-pick dests
bCopy = vbYes = MsgBox("...?", vbYesNo)
For each combination: HFApp.SqlExec "exec SomeProc ..."
```

Blazor's FPickList is CALLBACK-based (`OnSelected`/`OnCancel` parameters), so a
straight port would need a state machine. Instead, wrap ONE shared FPickList
instance in a TaskCompletionSource so the chain reads top-to-bottom like VB6
(reference implementation: FMain.razor `ToolsPickAsync` + `DuplicateModelOptionsAsync`):

```razor
<FPickList @ref="_toolsPickList"
           OnSelected="OnToolsPickSelected"
           OnCancel="OnToolsPickCancelled" />
<Fx.ControlKit.Dialogs.MessageBoxControl @ref="_toolsMessageBox" />
```

```csharp
private TaskCompletionSource<List<IDictionary<string, object>>?>? _toolsPickTcs;

private async Task<List<IDictionary<string, object>>?> ToolsPickAsync(
    string itemTitle, string sql, string? hideColumns, bool multiSelect,
    string chooseCaption, IDictionary<string, object>? parameters)
{
    if (_toolsPickList == null) return null;
    _toolsPickTcs = new TaskCompletionSource<List<IDictionary<string, object>>?>();
    await _toolsPickList.Open(
        title: itemTitle, queries: sql, hideColumns: hideColumns,
        multiSelect: multiSelect, dialogCaption: chooseCaption,
        windowTitle: $"Precision Builder - {itemTitle} List",
        parameters: parameters);          // FPickList passes these to the SQL — parameterize, never interpolate values
    return await _toolsPickTcs.Task;      // null = cancelled (VB6 "Exit Sub")
}

private Task OnToolsPickSelected(List<IDictionary<string, object>> rows)
{ _toolsPickTcs?.TrySetResult(rows); return Task.CompletedTask; }
private Task OnToolsPickCancelled()
{ _toolsPickTcs?.TrySetResult(null); return Task.CompletedTask; }
```

Rules learned porting FMain's Duplicate Model Options:
- **windowTitle vs dialogCaption**: `Open`'s `_windowTitle` falls back to
  `dialogCaption` when `windowTitle` is omitted. VB6 keeps them separate
  (window = "Precision Builder - {ItemTitle} List", caption label = the
  ChooseCaption argument), so pass BOTH explicitly.
- **MessageBoxControl.ShowAsync is already awaitable** (returns
  `Task<MessageBoxResult>`) — use it for every VB6 MsgBox in the chain;
  no wrapper needed. `MessageBoxButtons.{OkCancel,YesNo,YesNoCancel}`.
- **exec loops stay sequential and untransacted when VB6 ran independent
  SqlExec calls** — wrapping several stored procs in one client transaction
  can break procs that manage their own transactions, and VB6's semantics
  were per-call anyway. (Multi-STATEMENT saves the app owns still get one
  DbTransactionSession — recipe 04.)
- **Proc calls are positional**: `exec Purch_DuplicateModelOption @DivisionID,
  @AssemblyID, @Model, @CopyQuotes` — the client parameter names need not
  match the proc's parameter names; order does.
- Hidden plumbing columns (`hideColumns:` "areaid,modelid,...") mirror VB6's
  HideColumns argument verbatim; the hidden values are still present in the
  returned row dictionaries.

## Companion trap — in-grid dropdown editors (cell popup clipping)

When a grid cell hosts a DropDownListControl (FInboxJobs pattern), the grid
builder writes `overflow:hidden` inline on every `td`, which clips the open
panel invisibly (panel IS in the DOM at z-index 9999 but zero visible pixels).
Fix in the page's scoped CSS (reference: FVendorChange.razor.css,
FPOFormats.razor.css):

```css
.work-grid ::deep .fx-cell-active {
    overflow: visible !important;   /* inline style needs !important */
    position: relative;
    z-index: 30;                    /* paint above following rows */
}
```
