using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using System.Globalization;
using System.Reflection;

namespace HomeFront.Components.Grid;

public partial class HfGrid<TValue> : IHfGridOwner
{
    void IHfGridOwner.RegisterColumnsContainer(HfGridColumnsBase container)
    {
        _columnsContainer = container;
    }

    // ── Parameters ───────────────────────────────────────────────────────

    [Parameter] public IEnumerable<TValue>? DataSource { get; set; }
    [Parameter] public RenderFragment? ChildContent { get; set; }
    [Parameter] public string? Height { get; set; }
    [Parameter] public string? Width { get; set; }
    [Parameter] public int RowHeight { get; set; }

    // Feature flags
    [Parameter] public bool AllowSorting { get; set; }
    [Parameter] public bool AllowMultiSorting { get; set; }
    [Parameter] public bool AllowFiltering { get; set; }
    [Parameter] public bool AllowPaging { get; set; }
    [Parameter] public bool AllowSelection { get; set; } = true;
    [Parameter] public bool AllowGrouping { get; set; }
    [Parameter] public bool AllowResizing { get; set; }
    [Parameter] public bool EnableHover { get; set; } = true;
    [Parameter] public bool EnableAltRow { get; set; } = true;
    [Parameter] public bool ShowSearchBar { get; set; }
    [Parameter] public GridLines GridLines { get; set; } = GridLines.Default;
    [Parameter] public List<string>? Toolbar { get; set; }

    /// <summary>Initial group columns (field names).</summary>
    [Parameter] public List<string>? GroupColumns { get; set; }

    /// <summary>Aggregate row definitions for group footers and grid footer.</summary>
    [Parameter] public List<HfAggregateRow>? AggregateRows { get; set; }

    /// <summary>Show an Expand All / Collapse All toggle button in the group drop area.</summary>
    [Parameter] public bool ShowGroupExpandCollapse { get; set; } = true;

    // Child settings components
    [Parameter] public HfFilterSettings? FilterSettingsRef { get; set; }
    [Parameter] public HfPageSettings? PageSettingsRef { get; set; }
    [Parameter] public HfEditSettings? EditSettingsRef { get; set; }
    [Parameter] public HfSelectionSettings? SelectionSettingsRef { get; set; }
    [Parameter] public HfGridEvents<TValue>? EventsRef { get; set; }

    // Direct event callbacks (shorthand without EventsRef)
    [Parameter] public EventCallback<string> OnToolbarItemClick { get; set; }

    // ── Internal State ───────────────────────────────────────────────────

    internal HfGridColumnsBase? _columnsContainer;
    private readonly Dictionary<string, ColumnState> _columnStates = new();
    private readonly PageState _pageState = new();
    private readonly HashSet<TValue> _selectedItems = new();
    private readonly List<(int RowIndex, int CellIndex)> _selectedCells = new();
    private bool _expandAllGroups;
    private bool _allGroupsCollapsed;
    private (int RowIndex, int CellIndex)? _lastSelectedCell;
    private int? _lastSelectedRowIndex;

    // Editing
    private bool _isEditing;
    private int _editingRowIndex = -1;
    private TValue? _editItem;

    // Filtering popup
    private string? _filterPopupField;

    // Search
    private string? SearchText;
    private CancellationTokenSource? _searchCts;

    // ── Grouping State ───────────────────────────────────────────────────
    private readonly List<GroupDescriptor> _groupDescriptors = new();
    private string? _draggingColumnField;
    private string? _draggingGroupChipField;
    private bool _dragOverGroupArea;

    // ── Resize State ─────────────────────────────────────────────────────
    private HfGridColumn? _resizingCol;
    private double _resizeStartX;
    private double _resizeStartWidth;

    // ── Computed Properties ──────────────────────────────────────────────

    private bool ShowCheckboxColumn =>
        SelectionSettingsRef?.CheckboxOnly == true ||
        VisibleColumns.Any(c => c.Type == ColumnType.CheckBox && string.IsNullOrEmpty(c.Field));

    private FilterType ResolvedFilterType =>
        FilterSettingsRef?.Type ?? FilterType.FilterBar;

    private int[] ResolvedPageSizes =>
        PageSettingsRef?.PageSizes ?? [5, 10, 20, 50];

    private IEnumerable<HfGridColumn> VisibleColumns =>
        _columnsContainer?.Columns.Where(c => c.Visible) ?? Enumerable.Empty<HfGridColumn>();

    public IReadOnlyList<HfGridColumn> Columns =>
        _columnsContainer?.Columns ?? Array.Empty<HfGridColumn>();

    private bool AllRowsSelected =>
        PagedData.Any() && PagedData.All(item => _selectedItems.Contains(item));

    private int TotalColumnCount =>
        VisibleColumns.Count()
        + (ShowCheckboxColumn ? 1 : 0)
        + (AllowGrouping ? _groupDescriptors.Count : 0);

    private bool HasAnyData =>
        AllowGrouping && _groupDescriptors.Count > 0
            ? GroupedData.Any()
            : PagedData.Any();

    // ── Data Pipeline: DataSource → Filter → Search → Sort → Page ────────

    private IEnumerable<TValue> FilteredData
    {
        get
        {
            var data = DataSource ?? Enumerable.Empty<TValue>();

            // Apply column filters
            foreach (var kvp in _columnStates)
            {
                var colField = kvp.Key;
                var state = kvp.Value;

                if (!string.IsNullOrEmpty(state.FilterValue))
                {
                    var filterVal = state.FilterValue;
                    data = data.Where(item =>
                    {
                        var val = GetPropertyValue(item, colField)?.ToString() ?? "";
                        return val.Contains(filterVal, StringComparison.OrdinalIgnoreCase);
                    });
                }

                if (state.CheckedFilterValues.Count > 0)
                {
                    var checkedVals = state.CheckedFilterValues;
                    data = data.Where(item =>
                    {
                        var val = GetPropertyValue(item, colField)?.ToString() ?? "";
                        return checkedVals.Contains(val);
                    });
                }
            }

            // Apply global search
            if (!string.IsNullOrEmpty(SearchText))
            {
                var searchLower = SearchText.ToLower();
                data = data.Where(item =>
                    VisibleColumns.Any(col =>
                    {
                        var val = GetPropertyValue(item, col.Field)?.ToString() ?? "";
                        return val.Contains(searchLower, StringComparison.OrdinalIgnoreCase);
                    }));
            }

            return data;
        }
    }

    private IEnumerable<TValue> SortedData
    {
        get
        {
            var data = FilteredData;
            var sortedCol = _columnStates.FirstOrDefault(kvp => kvp.Value.SortDirection.HasValue);
            if (sortedCol.Key != null)
            {
                var sortField = sortedCol.Key;
                if (sortedCol.Value.SortDirection == SortDirection.Ascending)
                    data = data.OrderBy(item => GetPropertyValue(item, sortField));
                else
                    data = data.OrderByDescending(item => GetPropertyValue(item, sortField));
            }
            return data;
        }
    }

    private IEnumerable<TValue> PagedData
    {
        get
        {
            var data = SortedData;
            _pageState.TotalRecords = data.Count();

            if (!AllowPaging)
                return data;

            return data
                .Skip((_pageState.CurrentPage - 1) * _pageState.PageSize)
                .Take(_pageState.PageSize);
        }
    }

    // ── Grouped Data ─────────────────────────────────────────────────────

    private IEnumerable<GroupResult<TValue>> GroupedData
    {
        get
        {
            if (_groupDescriptors.Count == 0)
                return Enumerable.Empty<GroupResult<TValue>>();

            var data = SortedData.ToList();
            _pageState.TotalRecords = data.Count;
            return BuildGroups(data, 0);
        }
    }

    private IEnumerable<GroupResult<TValue>> BuildGroups(IEnumerable<TValue> data, int level)
    {
        if (level >= _groupDescriptors.Count)
            return Enumerable.Empty<GroupResult<TValue>>();

        var gd = _groupDescriptors[level];
        var groups = data
            .GroupBy(item => GetPropertyValue(item, gd.Field)?.ToString() ?? "(empty)")
            .Select(g =>
            {
                var allItems = g.ToList();
                var group = new GroupResult<TValue>
                {
                    Field = gd.Field,
                    HeaderText = gd.HeaderText,
                    Key = g.Key,
                    Count = allItems.Count,
                    Items = level == _groupDescriptors.Count - 1 ? allItems : Enumerable.Empty<TValue>(),
                    SubGroups = level < _groupDescriptors.Count - 1
                        ? BuildGroups(allItems, level + 1)
                        : Enumerable.Empty<GroupResult<TValue>>()
                };

                // Compute aggregates for this group across all its items (including sub-groups)
                if (AggregateRows is { Count: > 0 })
                {
                    group.Aggregates = ComputeAggregates(allItems);
                }

                return group;
            })
            .ToList();

        if (_expandAllGroups)
        {
            foreach (var group in groups)
                group.IsCollapsed = false;
        }

        if (_allGroupsCollapsed)
        {
            foreach (var group in groups)
                group.IsCollapsed = true;
        }

        return groups;
    }

    /// <summary>
    /// Compute aggregate values for a list of items based on AggregateRows config.
    /// </summary>
    private Dictionary<string, object?> ComputeAggregates(IEnumerable<TValue> items)
    {
        var result = new Dictionary<string, object?>();
        if (AggregateRows == null) return result;

        var itemsList = items.ToList();
        if (itemsList.Count == 0) return result;

        foreach (var aggRow in AggregateRows)
        {
            foreach (var aggCol in aggRow.Columns)
            {
                var key = $"{aggCol.Field}_{aggCol.Type}";
                if (result.ContainsKey(key)) continue;

                var values = itemsList
                    .Select(item => GetPropertyValue(item, aggCol.Field))
                    .Where(v => v != null)
                    .Select(v =>
                    {
                        if (v is double d) return d;
                        if (v is int i) return (double)i;
                        if (v is decimal dec) return (double)dec;
                        if (v is float f) return (double)f;
                        if (v is long l) return (double)l;
                        if (double.TryParse(v?.ToString(), System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)) return parsed;
                        return (double?)null;
                    })
                    .Where(v => v.HasValue)
                    .Select(v => v!.Value)
                    .ToList();

                if (values.Count == 0)
                {
                    result[key] = null;
                    continue;
                }

                double computed = aggCol.Type switch
                {
                    AggregateType.Sum => values.Sum(),
                    AggregateType.Average => values.Average(),
                    AggregateType.Count => values.Count,
                    AggregateType.Min => values.Min(),
                    AggregateType.Max => values.Max(),
                    _ => 0
                };

                result[key] = computed;
            }
        }

        return result;
    }

    /// <summary>
    /// Format an aggregate value using the column's format and template.
    /// </summary>
    internal string FormatAggregateValue(HfAggregateColumn aggCol, object? value, string? templateOverride = null)
    {
        if (value == null) return "";

        string formatted;
        if (!string.IsNullOrEmpty(aggCol.Format) && value is IFormattable formattable)
            formatted = formattable.ToString(aggCol.Format, CultureInfo.CurrentCulture);
        else
            formatted = value.ToString() ?? "";

        var template = templateOverride ?? aggCol.FooterTemplate ?? "{value}";
        return template.Replace("{value}", formatted);
    }

    /// <summary>
    /// Toggle expand/collapse all groups.
    /// </summary>
    public void ToggleExpandCollapseAll()
    {
        if (_allGroupsCollapsed)
        {
            _expandAllGroups = true;
            _allGroupsCollapsed = false;
        }
        else
        {
            _expandAllGroups = false;
            _allGroupsCollapsed = true;
        }
        StateHasChanged();
    }

    public Task CollapseAllGroupAsync()
    {
        _allGroupsCollapsed = true;
        _expandAllGroups = false;
        return InvokeAsync(StateHasChanged);
    }

    // ── Lifecycle ────────────────────────────────────────────────────────

    protected override void OnInitialized()
    {
        _pageState.PageSize = PageSettingsRef?.PageSize ?? 10;

        // Apply initial group columns
        if (GroupColumns is { Count: > 0 })
        {
            foreach (var colField in GroupColumns)
            {
                if (!_groupDescriptors.Any(g => g.Field == colField))
                {
                    var col = VisibleColumns.FirstOrDefault(c => c.Field == colField);
                    _groupDescriptors.Add(new GroupDescriptor
                    {
                        Field = colField,
                        HeaderText = col?.DisplayHeader ?? colField
                    });
                }
            }
        }
    }

    protected override void OnAfterRender(bool firstRender)
    {
        // Capture the columns container from child content
        if (firstRender)
        {
            // Re-apply initial group columns now that columns are loaded
            if (GroupColumns is { Count: > 0 } && _groupDescriptors.Count == 0)
            {
                foreach (var colField in GroupColumns)
                {
                    var col = VisibleColumns.FirstOrDefault(c => c.Field == colField);
                    if (col != null && !_groupDescriptors.Any(g => g.Field == colField))
                    {
                        _groupDescriptors.Add(new GroupDescriptor
                        {
                            Field = colField,
                            HeaderText = col.DisplayHeader
                        });
                    }
                }
            }
            StateHasChanged();
        }
    }

    protected override void OnParametersSet()
    {
        _pageState.PageSize = PageSettingsRef?.PageSize ?? _pageState.PageSize;
    }

    // ── Sorting ──────────────────────────────────────────────────────────

    private async Task HandleSort(HfGridColumn col)
    {
        if (!AllowSorting || !col.AllowSorting || string.IsNullOrEmpty(col.Field))
            return;

        var state = GetColumnState(col.Field);

        // Fire Sorting event
        if (EventsRef?.Sorting.HasDelegate == true)
        {
            var args = new SortEventArgs
            {
                Field = col.Field,
                Direction = state.SortDirection == SortDirection.Ascending
                    ? SortDirection.Descending : SortDirection.Ascending
            };
            await EventsRef.Sorting.InvokeAsync(args);
            if (args.Cancel) return;
        }

        // Clear other sorts unless multi-sort
        if (!AllowMultiSorting)
        {
            foreach (var kvp in _columnStates)
                if (kvp.Key != col.Field)
                    kvp.Value.SortDirection = null;
        }

        // Toggle
        if (!state.SortDirection.HasValue)
            state.SortDirection = SortDirection.Ascending;
        else if (state.SortDirection == SortDirection.Ascending)
            state.SortDirection = SortDirection.Descending;
        else
            state.SortDirection = null;

        _pageState.CurrentPage = 1;

        if (EventsRef?.Sorted.HasDelegate == true)
            await EventsRef.Sorted.InvokeAsync(new SortEventArgs { Field = col.Field, Direction = state.SortDirection ?? SortDirection.Ascending });
    }

    // ── Filtering ────────────────────────────────────────────────────────

    private async Task ApplyFilter(string field, string? value)
    {
        var state = GetColumnState(field);

        if (EventsRef?.Filtering.HasDelegate == true)
        {
            var args = new FilterEventArgs { Field = field, Value = value };
            await EventsRef.Filtering.InvokeAsync(args);
            if (args.Cancel) return;
        }

        state.FilterValue = string.IsNullOrWhiteSpace(value) ? null : value;
        _pageState.CurrentPage = 1;

        if (EventsRef?.Filtered.HasDelegate == true)
            await EventsRef.Filtered.InvokeAsync(new FilterEventArgs { Field = field, Value = value });
    }

    private void ToggleFilterPopup(string field)
    {
        _filterPopupField = _filterPopupField == field ? null : field;
    }

    private void CloseFilterPopup() => _filterPopupField = null;

    private void ToggleCheckboxFilter(string field, string value)
    {
        var state = GetColumnState(field);
        if (!state.CheckedFilterValues.Remove(value))
            state.CheckedFilterValues.Add(value);
        _pageState.CurrentPage = 1;
    }

    private void ClearFilter(string field)
    {
        var state = GetColumnState(field);
        state.FilterValue = null;
        state.CheckedFilterValues.Clear();
        _pageState.CurrentPage = 1;
    }

    private List<string> GetDistinctValues(string field)
    {
        return (DataSource ?? Enumerable.Empty<TValue>())
            .Select(item => GetPropertyValue(item, field)?.ToString() ?? "")
            .Distinct()
            .OrderBy(v => v)
            .ToList();
    }

    // ── Search ───────────────────────────────────────────────────────────

    private async Task ApplySearchDebounced()
    {
        _searchCts?.Cancel();
        _searchCts = new CancellationTokenSource();
        var token = _searchCts.Token;

        try
        {
            await Task.Delay(300, token);
            _pageState.CurrentPage = 1;
            StateHasChanged();
        }
        catch (TaskCanceledException) { }
    }

    // ── Paging ───────────────────────────────────────────────────────────

    private async Task GoToPage(int page)
    {
        if (page < 1 || page > _pageState.TotalPages || page == _pageState.CurrentPage)
            return;

        if (EventsRef?.PageChanging.HasDelegate == true)
        {
            var args = new PageChangeEventArgs { PreviousPage = _pageState.CurrentPage, CurrentPage = page };
            await EventsRef.PageChanging.InvokeAsync(args);
            if (args.Cancel) return;
        }

        var prev = _pageState.CurrentPage;
        _pageState.CurrentPage = page;

        if (EventsRef?.PageChanged.HasDelegate == true)
            await EventsRef.PageChanged.InvokeAsync(new PageChangeEventArgs { PreviousPage = prev, CurrentPage = page });
    }

    private void HandlePageSizeChange(ChangeEventArgs e)
    {
        if (int.TryParse(e.Value?.ToString(), out var size))
        {
            _pageState.PageSize = size;
            _pageState.CurrentPage = 1;
        }
    }

    private IEnumerable<int> GetPageNumbers()
    {
        var total = _pageState.TotalPages;
        var current = _pageState.CurrentPage;
        var count = PageSettingsRef?.PageCount ?? 5;

        var start = Math.Max(1, current - count / 2);
        var end = Math.Min(total, start + count - 1);
        start = Math.Max(1, end - count + 1);

        return Enumerable.Range(start, end - start + 1);
    }

    // ── Selection ────────────────────────────────────────────────────────

    private async Task HandleRowClick(TValue item, int rowIndex, MouseEventArgs? mouseArgs = null)
    {
        if (EventsRef?.OnRecordClick.HasDelegate == true)
            await EventsRef.OnRecordClick.InvokeAsync(new CellClickEventArgs<TValue> { Data = item, RowIndex = rowIndex });

        if (!AllowSelection) return;
        if (SelectionSettingsRef?.Mode == SelectionMode.Cell && SelectionSettingsRef?.CheckboxOnly != true)
            return;

        await SelectRow(item, rowIndex, mouseArgs);
    }

    private async Task HandleRowDblClick(TValue item, int rowIndex)
    {
        if (EventsRef?.OnRecordDoubleClick.HasDelegate == true)
            await EventsRef.OnRecordDoubleClick.InvokeAsync(new CellClickEventArgs<TValue> { Data = item, RowIndex = rowIndex });

        if (EditSettingsRef?.AllowEditing == true && EditSettingsRef.AllowEditOnDblClick)
            await StartEdit(item, rowIndex);
    }

    private async Task SelectRow(TValue item, int rowIndex, MouseEventArgs? mouseArgs = null)
    {
        var selType = SelectionSettingsRef?.Type ?? SelectionType.Single;
        var isCtrl = mouseArgs?.CtrlKey == true || mouseArgs?.MetaKey == true;
        var isShift = mouseArgs?.ShiftKey == true;

        if (EventsRef?.RowSelecting.HasDelegate == true)
        {
            var args = new RowSelectEventArgs<TValue> { Data = item, RowIndex = rowIndex };
            await EventsRef.RowSelecting.InvokeAsync(args);
            if (args.Cancel) return;
        }

        if (selType == SelectionType.Single && !isCtrl && !isShift)
        {
            var wasSelected = _selectedItems.Contains(item);
            _selectedItems.Clear();
            if (!wasSelected || SelectionSettingsRef?.EnableToggle != true)
                _selectedItems.Add(item);
        }
        else if (selType == SelectionType.Multiple || isCtrl || isShift)
        {
            if (isShift && _lastSelectedRowIndex.HasValue)
            {
                // Range selection: select all rows between last selected and current
                var allData = PagedData.ToList();
                var resolvedCurrent = ResolveRowIndex(item, rowIndex);
                var start = Math.Min(_lastSelectedRowIndex.Value, resolvedCurrent);
                var end = Math.Max(_lastSelectedRowIndex.Value, resolvedCurrent);

                if (!isCtrl) _selectedItems.Clear();

                var sourceData = (DataSource as IList<TValue>) ?? (DataSource?.ToList() ?? new List<TValue>());
                for (var i = start; i <= end && i < sourceData.Count; i++)
                {
                    _selectedItems.Add(sourceData[i]);
                }
            }
            else if (isCtrl)
            {
                // Toggle single item
                if (!_selectedItems.Remove(item))
                    _selectedItems.Add(item);
            }
            else
            {
                if (!_selectedItems.Remove(item))
                    _selectedItems.Add(item);
            }
        }

        _lastSelectedRowIndex = ResolveRowIndex(item, rowIndex);

        if (EventsRef?.RowSelected.HasDelegate == true)
            await EventsRef.RowSelected.InvokeAsync(new RowSelectEventArgs<TValue> { Data = item, RowIndex = rowIndex });
    }

    private async Task HandleCellClick(TValue item, int rowIndex, int cellIndex, MouseEventArgs args)
    {
        if (!AllowSelection || SelectionSettingsRef?.Mode != SelectionMode.Cell)
            return;

        var resolvedRowIndex = ResolveRowIndex(item, rowIndex);
        if (EventsRef?.CellSelecting.HasDelegate == true)
        {
            var selectingArgs = new CellSelectingEventArgs<TValue>
            {
                Data = item,
                RowIndex = resolvedRowIndex,
                CellIndex = cellIndex
            };
            await EventsRef.CellSelecting.InvokeAsync(selectingArgs);
            if (selectingArgs.Cancel)
                return;
        }

        var isCtrl = args.CtrlKey || args.MetaKey;
        var isShift = args.ShiftKey;

        if (!isCtrl && !isShift)
        {
            _selectedCells.Clear();
        }

        if (isShift && _lastSelectedCell.HasValue && _lastSelectedCell.Value.RowIndex == resolvedRowIndex)
        {
            var start = Math.Min(_lastSelectedCell.Value.CellIndex, cellIndex);
            var end = Math.Max(_lastSelectedCell.Value.CellIndex, cellIndex);
            for (var i = start; i <= end; i++)
            {
                var cell = (resolvedRowIndex, i);
                if (!_selectedCells.Contains(cell))
                    _selectedCells.Add(cell);
            }
        }
        else
        {
            var cell = (resolvedRowIndex, cellIndex);
            if (isCtrl && _selectedCells.Contains(cell))
                _selectedCells.Remove(cell);
            else if (!_selectedCells.Contains(cell))
                _selectedCells.Add(cell);
        }

        _lastSelectedCell = (resolvedRowIndex, cellIndex);

        if (EventsRef?.CellSelected.HasDelegate == true)
        {
            var value = GetPropertyValue(item, Columns.Count > cellIndex ? Columns[cellIndex].Field : "");
            await EventsRef.CellSelected.InvokeAsync(new CellSelectEventArgs<TValue>
            {
                Data = item,
                RowIndex = resolvedRowIndex,
                CellIndex = cellIndex,
                CurrentValue = value,
                IsCtrlPressed = isCtrl,
                IsShiftPressed = isShift
            });
        }
    }

    private int ResolveRowIndex(TValue item, int fallbackIndex)
    {
        if (DataSource is IList<TValue> list)
        {
            var idx = list.IndexOf(item);
            if (idx >= 0)
                return idx;
        }

        return fallbackIndex;
    }

    private async Task ToggleRowSelection(TValue item, int rowIndex)
    {
        await SelectRow(item, rowIndex);
    }

    private void ToggleSelectAll(ChangeEventArgs e)
    {
        if (AllRowsSelected)
            _selectedItems.Clear();
        else
            foreach (var item in PagedData)
                _selectedItems.Add(item);
    }

    // ── Editing ──────────────────────────────────────────────────────────

    private async Task StartEdit(TValue item, int rowIndex)
    {
        if (EventsRef?.OnBeginEdit.HasDelegate == true)
        {
            var args = new RowEditEventArgs<TValue> { Data = item, RowIndex = rowIndex };
            await EventsRef.OnBeginEdit.InvokeAsync(args);
            if (args.Cancel) return;
        }

        _editItem = CloneItem(item);
        _editingRowIndex = rowIndex;
        _isEditing = true;
    }

    private void StartAdd()
    {
        _editItem = Activator.CreateInstance<TValue>();
        _editingRowIndex = -1;
        _isEditing = true;
    }

    private async Task SaveEdit()
    {
        if (_editItem == null) return;

        if (_editingRowIndex >= 0)
        {
            // Update existing
            if (EventsRef?.RowUpdating.HasDelegate == true)
            {
                var args = new RowEditEventArgs<TValue> { Data = _editItem, RowIndex = _editingRowIndex };
                await EventsRef.RowUpdating.InvokeAsync(args);
                if (args.Cancel) return;
            }

            var list = DataSource as IList<TValue>;
            if (list != null)
            {
                var pagedList = SortedData.ToList();
                var actualIdx = AllowPaging
                    ? (_pageState.CurrentPage - 1) * _pageState.PageSize + _editingRowIndex
                    : _editingRowIndex;

                if (actualIdx >= 0 && actualIdx < pagedList.Count)
                {
                    var original = pagedList[actualIdx];
                    var origIdx = list.IndexOf(original);
                    if (origIdx >= 0)
                        CopyProperties(_editItem, list[origIdx]!);
                }
            }

            if (EventsRef?.RowUpdated.HasDelegate == true)
                await EventsRef.RowUpdated.InvokeAsync(new RowEditEventArgs<TValue> { Data = _editItem, RowIndex = _editingRowIndex });
        }
        else
        {
            // Add new
            if (EventsRef?.RowCreating.HasDelegate == true)
            {
                var args = new RowEditEventArgs<TValue> { Data = _editItem };
                await EventsRef.RowCreating.InvokeAsync(args);
                if (args.Cancel) return;
            }

            var list = DataSource as IList<TValue>;
            if (list != null)
            {
                if (EditSettingsRef?.NewRowPosition == NewRowPosition.Top)
                    list.Insert(0, _editItem);
                else
                    list.Add(_editItem);
            }

            if (EventsRef?.RowCreated.HasDelegate == true)
                await EventsRef.RowCreated.InvokeAsync(new RowEditEventArgs<TValue> { Data = _editItem });
        }

        _isEditing = false;
        _editItem = default;
        _editingRowIndex = -1;
    }

    private void CancelEdit()
    {
        _isEditing = false;
        _editItem = default;
        _editingRowIndex = -1;
    }

    private async Task DeleteRow(TValue item, int rowIndex)
    {
        if (EditSettingsRef?.ShowConfirmDialog == true)
        {
            // In a real implementation, show a confirmation dialog.
        }

        if (EventsRef?.RowDeleting.HasDelegate == true)
        {
            var args = new RowEditEventArgs<TValue> { Data = item, RowIndex = rowIndex };
            await EventsRef.RowDeleting.InvokeAsync(args);
            if (args.Cancel) return;
        }

        var list = DataSource as IList<TValue>;
        list?.Remove(item);
        _selectedItems.Remove(item);

        if (EventsRef?.RowDeleted.HasDelegate == true)
            await EventsRef.RowDeleted.InvokeAsync(new RowEditEventArgs<TValue> { Data = item, RowIndex = rowIndex });
    }

    private async Task HandleCommand(string type, TValue item, int rowIndex)
    {
        switch (type.ToLower())
        {
            case "edit":
                await StartEdit(item, rowIndex);
                break;
            case "delete":
                await DeleteRow(item, rowIndex);
                break;
            case "save":
                await SaveEdit();
                break;
            case "cancel":
                CancelEdit();
                break;
        }
    }

    // ── Toolbar ──────────────────────────────────────────────────────────

    private async Task OnToolbarClick(string item)
    {
        switch (item.ToLower())
        {
            case "add":
                StartAdd();
                break;
            case "edit":
                if (_selectedItems.Count == 1)
                {
                    var sel = _selectedItems.First();
                    var idx = PagedData.ToList().IndexOf(sel);
                    await StartEdit(sel, idx);
                }
                break;
            case "delete":
                if (_selectedItems.Count > 0)
                {
                    foreach (var sel in _selectedItems.ToList())
                        await DeleteRow(sel, 0);
                }
                break;
        }

        if (OnToolbarItemClick.HasDelegate)
            await OnToolbarItemClick.InvokeAsync(item);
    }

    // ── Keyboard Navigation ──────────────────────────────────────────────

    private void HandleKeyDown(KeyboardEventArgs e)
    {
        if (e.Key == "Escape" && _isEditing)
            CancelEdit();
    }

    // ══════════════════════════════════════════════════════════════════════
    // ── DRAG & DROP GROUPING ─────────────────────────────────────────────
    // ══════════════════════════════════════════════════════════════════════

    private void StartColumnDrag(string colField)
    {
        _draggingColumnField = colField;
        _draggingGroupChipField = null;
    }

    private void EndColumnDrag()
    {
        _draggingColumnField = null;
        _dragOverGroupArea = false;
    }

    private void StartGroupChipDrag(string groupField)
    {
        _draggingGroupChipField = groupField;
        _draggingColumnField = null;
    }

    private void HandleGroupAreaDragOver(DragEventArgs e)
    {
        _dragOverGroupArea = true;
    }

    private void HandleGroupAreaDragLeave(DragEventArgs e)
    {
        _dragOverGroupArea = false;
    }

    private async Task HandleGroupAreaDrop(DragEventArgs e)
    {
        _dragOverGroupArea = false;

        if (!string.IsNullOrEmpty(_draggingColumnField))
        {
            await AddGroup(_draggingColumnField);
            _draggingColumnField = null;
        }
        else if (!string.IsNullOrEmpty(_draggingGroupChipField))
        {
            // Re-ordering groups (simple: move to end)
            var existing = _groupDescriptors.FirstOrDefault(g => g.Field == _draggingGroupChipField);
            if (existing != null)
            {
                _groupDescriptors.Remove(existing);
                _groupDescriptors.Add(existing);
            }
            _draggingGroupChipField = null;
        }
    }

    private async Task AddGroup(string colField)
    {
        if (_groupDescriptors.Any(g => g.Field == colField))
            return;

        var col = VisibleColumns.FirstOrDefault(c => c.Field == colField);
        if (col == null || !col.AllowGrouping) return;

        if (EventsRef?.Grouping.HasDelegate == true)
        {
            var args = new GroupEventArgs { Field = colField };
            await EventsRef.Grouping.InvokeAsync(args);
            if (args.Cancel) return;
        }

        _groupDescriptors.Add(new GroupDescriptor
        {
            Field = colField,
            HeaderText = col.DisplayHeader
        });
        _pageState.CurrentPage = 1;

        if (EventsRef?.Grouped.HasDelegate == true)
            await EventsRef.Grouped.InvokeAsync(new GroupEventArgs { Field = colField });
    }

    private async Task RemoveGroup(string colField)
    {
        if (EventsRef?.Ungrouping.HasDelegate == true)
        {
            var args = new GroupEventArgs { Field = colField };
            await EventsRef.Ungrouping.InvokeAsync(args);
            if (args.Cancel) return;
        }

        _groupDescriptors.RemoveAll(g => g.Field == colField);
        _pageState.CurrentPage = 1;

        if (EventsRef?.Ungrouped.HasDelegate == true)
            await EventsRef.Ungrouped.InvokeAsync(new GroupEventArgs { Field = colField });
    }

    private void ToggleGroupCollapse(GroupResult<TValue> group)
    {
        group.IsCollapsed = !group.IsCollapsed;
    }

    // ── Render Grouped Rows ──────────────────────────────────────────────

    private RenderFragment RenderGroupedRows(IEnumerable<GroupResult<TValue>> groups, int level) => builder =>
    {
        foreach (var group in groups)
        {
            // Group header row
            builder.OpenElement(0, "tr");
            builder.AddAttribute(1, "class", "hf-group-header-row");
            builder.AddAttribute(2, "onclick", EventCallback.Factory.Create(this, () => ToggleGroupCollapse(group)));

            // Indent cells for nesting
            for (int i = 0; i < level; i++)
            {
                builder.OpenElement(10, "td");
                builder.AddAttribute(11, "class", "hf-cell hf-group-indent");
                builder.AddAttribute(12, "style", "width:32px;");
                builder.CloseElement();
            }

            // Expand/collapse + group header
            var totalSpan = TotalColumnCount - level;
            builder.OpenElement(20, "td");
            builder.AddAttribute(21, "colspan", totalSpan);
            builder.AddAttribute(22, "class", "hf-cell hf-group-header-cell");

            builder.OpenElement(30, "span");
            builder.AddAttribute(31, "class", $"hf-group-expand-icon {(group.IsCollapsed ? "collapsed" : "expanded")}");
            builder.AddContent(32, group.IsCollapsed ? "▶" : "▼");
            builder.CloseElement();

            builder.OpenElement(40, "span");
            builder.AddAttribute(41, "class", "hf-group-header-text");
            builder.AddContent(42, $"{group.HeaderText}: {group.Key}");
            builder.CloseElement();

            builder.OpenElement(50, "span");
            builder.AddAttribute(51, "class", "hf-group-count");
            builder.AddContent(52, $"({group.Count} items)");
            builder.CloseElement();

            // Inline caption aggregates (e.g. "Sum: $1,234.56" next to group header)
            if (AggregateRows is { Count: > 0 } && group.Aggregates.Count > 0)
            {
                var captionRows = AggregateRows.Where(r => r.ShowInGroupCaption).ToList();
                foreach (var aggRow in captionRows)
                {
                    foreach (var aggCol in aggRow.Columns)
                    {
                        var key = $"{aggCol.Field}_{aggCol.Type}";
                        if (group.Aggregates.TryGetValue(key, out var val) && val != null)
                        {
                            builder.OpenElement(55, "span");
                            builder.AddAttribute(56, "class", "hf-group-caption-agg");
                            builder.AddContent(57, FormatAggregateValue(aggCol, val, aggCol.GroupCaptionTemplate));
                            builder.CloseElement();
                        }
                    }
                }
            }

            builder.CloseElement(); // td
            builder.CloseElement(); // tr

            if (!group.IsCollapsed)
            {
                // If this group has sub-groups, recurse
                if (group.SubGroups.Any())
                {
                    builder.AddContent(60, RenderGroupedRows(group.SubGroups, level + 1));
                }
                else
                {
                    // Render actual data rows
                    var rowIdx = 0;
                    foreach (var item in group.Items)
                    {
                        var currentIdx = rowIdx;
                        var resolvedRowIdx = ResolveRowIndex(item, currentIdx);
                        var isSelected = _selectedItems.Contains(item);

                        builder.OpenElement(70, "tr");
                        builder.AddAttribute(71, "class",
                            $"hf-row {(rowIdx % 2 == 1 && EnableAltRow ? "hf-alt-row" : "")} {(isSelected ? "hf-selected" : "")} {(EnableHover ? "hf-hover" : "")}");
                        builder.AddAttribute(72, "onclick", EventCallback.Factory.Create<MouseEventArgs>(this, e => HandleRowClick(item, currentIdx, e)));
                        if (RowHeight > 0)
                            builder.AddAttribute(73, "style", $"height:{RowHeight}px;");

                        // Group indent cells
                        for (int i = 0; i < _groupDescriptors.Count; i++)
                        {
                            builder.OpenElement(80, "td");
                            builder.AddAttribute(81, "class", "hf-cell hf-group-indent");
                            builder.AddAttribute(82, "style", "width:32px;");
                            builder.CloseElement();
                        }

                        // Checkbox column
                        if (ShowCheckboxColumn)
                        {
                            builder.OpenElement(90, "td");
                            builder.AddAttribute(91, "class", "hf-cell hf-checkbox-cell");
                            builder.AddAttribute(92, "style", "width:50px;");
                            builder.OpenElement(93, "input");
                            builder.AddAttribute(94, "type", "checkbox");
                            builder.AddAttribute(95, "checked", isSelected);
                            builder.CloseElement();
                            builder.CloseElement();
                        }

                        // Data cells
                        var colIdx = 0;
                        foreach (var col in VisibleColumns)
                        {
                            var capturedColIdx = colIdx;
                            var isCellSelected = _selectedCells.Contains((resolvedRowIdx, capturedColIdx));
                            builder.OpenElement(100, "td");
                            builder.AddAttribute(101, "class", isCellSelected ? "hf-cell hf-cell-selected" : "hf-cell");
                            builder.AddAttribute(102, "style", col.GetCellStyle());
                            builder.AddAttribute(103, "onclick", EventCallback.Factory.Create<MouseEventArgs>(this, args => HandleCellClick(item, resolvedRowIdx, capturedColIdx, args)));

                            if (col.Commands is { Count: > 0 })
                            {
                                foreach (var cmd in col.Commands)
                                {
                                    var cmdType = cmd.Type;
                                    var capturedItem = item;
                                    var capturedIdx = currentIdx;
                                    builder.OpenElement(110, "button");
                                    builder.AddAttribute(111, "class", $"hf-cmd-btn hf-cmd-{cmdType.ToLower()}");
                                    builder.AddAttribute(112, "onclick", EventCallback.Factory.Create(this, () => HandleCommand(cmdType, capturedItem, capturedIdx)));
                                    builder.AddContent(113, cmdType);
                                    builder.CloseElement();
                                }
                            }
                            else if (col.EffectiveTemplate != null)
                            {
                                builder.AddContent(120, col.EffectiveTemplate((object)item!));
                            }
                            else if (col.Type == ColumnType.CheckBox)
                            {
                                builder.OpenElement(130, "input");
                                builder.AddAttribute(131, "type", "checkbox");
                                builder.AddAttribute(132, "disabled", true);
                                builder.AddAttribute(133, "checked", GetBoolValue(item, col.Field));
                                builder.CloseElement();
                            }
                            else
                            {
                                builder.AddContent(140, GetCellDisplayValue(item, col));
                            }

                            builder.CloseElement(); // td
                            colIdx++;
                        }

                        builder.CloseElement(); // tr
                        rowIdx++;
                    }
                }

                // ── Group Footer (aggregate totals for this group) ──
                if (AggregateRows is { Count: > 0 } && group.Aggregates.Count > 0)
                {
                    var footerRows = AggregateRows.Where(r => r.ShowInGroupFooter).ToList();
                    foreach (var aggRow in footerRows)
                    {
                        builder.OpenElement(200, "tr");
                        builder.AddAttribute(201, "class", "hf-group-footer-row");

                        // Indent cells
                        for (int i = 0; i < _groupDescriptors.Count; i++)
                        {
                            builder.OpenElement(210, "td");
                            builder.AddAttribute(211, "class", "hf-cell hf-group-indent");
                            builder.AddAttribute(212, "style", "width:32px;");
                            builder.CloseElement();
                        }

                        // Render aggregate cells aligned with columns
                        foreach (var col in VisibleColumns)
                        {
                            var aggCol = aggRow.Columns.FirstOrDefault(a => a.Field == col.Field);
                            builder.OpenElement(220, "td");
                            builder.AddAttribute(221, "class", "hf-cell hf-aggregate-cell");
                            builder.AddAttribute(222, "style", col.GetCellStyle());

                            if (aggCol != null)
                            {
                                var key = $"{aggCol.Field}_{aggCol.Type}";
                                group.Aggregates.TryGetValue(key, out var val);
                                builder.OpenElement(230, "span");
                                builder.AddAttribute(231, "class", "hf-aggregate-value");
                                builder.AddContent(232, FormatAggregateValue(aggCol, val, aggCol.GroupFooterTemplate));
                                builder.CloseElement();
                            }

                            builder.CloseElement(); // td
                        }

                        builder.CloseElement(); // tr
                    }
                }
            }
        }
    };

    // ── Render Data Cells (flat mode) ────────────────────────────────────

    private RenderFragment RenderDataCells(TValue item, int rowIndex, Func<HfGridColumn, HfGridColumn> transform) => builder =>
    {
        var resolvedRowIndex = ResolveRowIndex(item, rowIndex);
        var colIdx = 0;
        foreach (var col in VisibleColumns)
        {
            builder.OpenElement(0, "td");
            var isCellSelected = _selectedCells.Contains((resolvedRowIndex, colIdx));
            builder.AddAttribute(1, "class", isCellSelected ? "hf-cell hf-cell-selected" : "hf-cell");
            builder.AddAttribute(2, "style", col.GetCellStyle());
            if (col.ClipMode == ClipMode.EllipsisWithTooltip)
                builder.AddAttribute(3, "title", GetCellDisplayValue(item, col));
            builder.AddAttribute(4, "onclick", EventCallback.Factory.Create<MouseEventArgs>(this, args => HandleCellClick(item, resolvedRowIndex, colIdx, args)));

            if (col.Commands is { Count: > 0 })
            {
                foreach (var cmd in col.Commands)
                {
                    var cmdType = cmd.Type;
                    var capturedItem = item;
                    builder.OpenElement(10, "button");
                    builder.AddAttribute(11, "class", $"hf-cmd-btn hf-cmd-{cmdType.ToLower()}");
                    builder.AddAttribute(12, "onclick", EventCallback.Factory.Create(this, () => HandleCommand(cmdType, capturedItem, 0)));
                    builder.AddContent(13, cmdType);
                    builder.CloseElement();
                }
            }
            else if (col.EffectiveTemplate != null)
            {
                builder.AddContent(20, col.EffectiveTemplate((object)item!));
            }
            else if (col.Type == ColumnType.CheckBox)
            {
                builder.OpenElement(30, "input");
                builder.AddAttribute(31, "type", "checkbox");
                builder.AddAttribute(32, "disabled", true);
                builder.AddAttribute(33, "checked", GetBoolValue(item, col.Field));
                builder.CloseElement();
            }
            else
            {
                builder.AddContent(40, GetCellDisplayValue(item, col));
            }

            builder.CloseElement(); // td
            colIdx++;
        }
    };

    // ══════════════════════════════════════════════════════════════════════
    // ── COLUMN RESIZE ────────────────────────────────────────────────────
    // ══════════════════════════════════════════════════════════════════════

    private void StartResize(HfGridColumn col, MouseEventArgs e)
    {
        _resizingCol = col;
        _resizeStartX = e.ClientX;

        // Parse starting width
        if (col.RuntimeWidth.HasValue)
            _resizeStartWidth = col.RuntimeWidth.Value;
        else if (!string.IsNullOrEmpty(col.Width))
        {
            var w = col.Width.Replace("px", "").Replace("%", "").Trim();
            if (double.TryParse(w, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                _resizeStartWidth = parsed;
            else
                _resizeStartWidth = 150;
        }
        else
            _resizeStartWidth = 150;
    }

    private async Task HandleResizeMove(MouseEventArgs e)
    {
        if (_resizingCol == null) return;

        var delta = e.ClientX - _resizeStartX;
        var newWidth = Math.Max(40, _resizeStartWidth + delta);

        if (EventsRef?.ColumnResizing.HasDelegate == true)
        {
            var args = new ResizeEventArgs
            {
                Field = _resizingCol.Field,
                OldWidth = _resizingCol.RuntimeWidth ?? _resizeStartWidth,
                NewWidth = newWidth
            };
            await EventsRef.ColumnResizing.InvokeAsync(args);
            if (args.Cancel) return;
        }

        _resizingCol.RuntimeWidth = newWidth;
    }

    private async Task EndResize(MouseEventArgs e)
    {
        if (_resizingCol != null && EventsRef?.ColumnResized.HasDelegate == true)
        {
            await EventsRef.ColumnResized.InvokeAsync(new ResizeEventArgs
            {
                Field = _resizingCol.Field,
                OldWidth = _resizeStartWidth,
                NewWidth = _resizingCol.RuntimeWidth ?? _resizeStartWidth
            });
        }
        _resizingCol = null;
    }

    // ── Render Helpers ───────────────────────────────────────────────────

    private RenderFragment RenderEditRow() => builder =>
    {
        builder.OpenElement(0, "tr");
        builder.AddAttribute(1, "class", "hf-row hf-edit-row");

        if (ShowCheckboxColumn)
        {
            builder.OpenElement(2, "td");
            builder.AddAttribute(3, "class", "hf-cell");
            builder.CloseElement();
        }

        // Group indent cells for edit row
        if (AllowGrouping && _groupDescriptors.Count > 0)
        {
            for (int i = 0; i < _groupDescriptors.Count; i++)
            {
                builder.OpenElement(4, "td");
                builder.AddAttribute(5, "class", "hf-cell hf-group-indent");
                builder.CloseElement();
            }
        }

        foreach (var col in VisibleColumns)
        {
            builder.OpenElement(10, "td");
            builder.AddAttribute(11, "class", "hf-cell hf-edit-cell");

            if (col.Commands is { Count: > 0 })
            {
                builder.OpenElement(20, "button");
                builder.AddAttribute(21, "class", "hf-cmd-btn hf-cmd-save");
                builder.AddAttribute(22, "onclick", EventCallback.Factory.Create(this, SaveEdit));
                builder.AddContent(23, "Save");
                builder.CloseElement();

                builder.OpenElement(30, "button");
                builder.AddAttribute(31, "class", "hf-cmd-btn hf-cmd-cancel");
                builder.AddAttribute(32, "onclick", EventCallback.Factory.Create(this, CancelEdit));
                builder.AddContent(33, "Cancel");
                builder.CloseElement();
            }
            else if (!string.IsNullOrEmpty(col.Field) && col.AllowEditing && !col.IsPrimaryKey)
            {
                if (col.EditTemplate != null && _editItem != null)
                {
                    builder.AddContent(40, col.EditTemplate((object)_editItem));
                }
                else
                {
                    var inputType = col.Type == ColumnType.Number ? "number"
                        : col.Type == ColumnType.Date ? "date" : "text";

                    builder.OpenElement(50, "input");
                    builder.AddAttribute(51, "type", inputType);
                    builder.AddAttribute(52, "class", "hf-edit-input");
                    builder.AddAttribute(53, "value", GetPropertyValue(_editItem, col.Field));
                    builder.AddAttribute(54, "onchange", EventCallback.Factory.Create<ChangeEventArgs>(this,
                        e => SetPropertyValue(_editItem, col.Field, e.Value?.ToString())));
                    builder.CloseElement();
                }
            }
            else
            {
                builder.AddContent(60, GetPropertyValue(_editItem, col.Field)?.ToString() ?? "");
            }

            builder.CloseElement(); // td
        }

        builder.CloseElement(); // tr
    };

    // ── Reflection Helpers ───────────────────────────────────────────────

    private ColumnState GetColumnState(string field)
    {
        if (!_columnStates.TryGetValue(field, out var state))
        {
            state = new ColumnState { Field = field };
            _columnStates[field] = state;
        }
        return state;
    }

    private static object? GetPropertyValue(object? item, string field)
    {
        if (item == null || string.IsNullOrEmpty(field)) return null;
        var prop = item.GetType().GetProperty(field, BindingFlags.Public | BindingFlags.Instance);
        return prop?.GetValue(item);
    }

    private static void SetPropertyValue(object? item, string field, string? value)
    {
        if (item == null || string.IsNullOrEmpty(field)) return;
        var prop = item.GetType().GetProperty(field, BindingFlags.Public | BindingFlags.Instance);
        if (prop == null || !prop.CanWrite) return;

        try
        {
            object? converted;
            var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

            if (string.IsNullOrEmpty(value))
                converted = targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
            else if (targetType == typeof(DateTime))
                converted = DateTime.Parse(value, CultureInfo.InvariantCulture);
            else
                converted = Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);

            prop.SetValue(item, converted);
        }
        catch { /* ignore conversion errors */ }
    }

    private static bool GetBoolValue(object? item, string field)
    {
        var val = GetPropertyValue(item, field);
        return val is true;
    }

    private string GetCellDisplayValue(object? item, HfGridColumn col)
    {
        var val = GetPropertyValue(item, col.Field);
        if (val == null) return "";

        if (!string.IsNullOrEmpty(col.Format))
        {
            if (val is IFormattable formattable)
                return formattable.ToString(col.Format, CultureInfo.CurrentCulture);
        }

        return val.ToString() ?? "";
    }

    private static TValue CloneItem(TValue source)
    {
        var clone = Activator.CreateInstance<TValue>();
        CopyProperties(source!, clone!);
        return clone;
    }

    private static void CopyProperties(object source, object target)
    {
        foreach (var prop in source.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (prop.CanRead && prop.CanWrite)
                prop.SetValue(target, prop.GetValue(source));
        }
    }

    // ── Public API Methods (SyncFusion equivalent) ───────────────────────
    public int TotalPages => _pageState.TotalPages;
    public int CurrentPage => _pageState.CurrentPage;

    public Task GoToPageAsync(int page) => GoToPage(page);

    public IEnumerable<TValue> GetSelectedRecords() => _selectedItems.ToList();

    public Task<List<TValue>> GetSelectedRecordsAsync() =>
        Task.FromResult(_selectedItems.ToList());

    public Task<List<(int RowIndex, int CellIndex)>> GetSelectedRowCellIndexesAsync() =>
        Task.FromResult(_selectedCells.ToList());

    public void ClearSelection() => _selectedItems.Clear();

    public Task ClearSelectionAsync()
    {
        _selectedItems.Clear();
        _selectedCells.Clear();
        return InvokeAsync(StateHasChanged);
    }

    public void SelectRow(int rowIndex)
    {
        var list = PagedData.ToList();
        if (rowIndex >= 0 && rowIndex < list.Count)
            _selectedItems.Add(list[rowIndex]);
    }

    public Task SelectCellAsync((int RowIndex, int CellIndex) cell, bool isCtrlPressed = false)
    {
        if (!isCtrlPressed)
        {
            _selectedCells.Clear();
        }

        if (!_selectedCells.Contains(cell))
        {
            _selectedCells.Add(cell);
        }

        return InvokeAsync(StateHasChanged);
    }

    public Task SelectCellsAsync(IEnumerable<(int RowIndex, int CellIndex)> cells)
    {
        _selectedCells.Clear();
        _selectedCells.AddRange(cells);
        return InvokeAsync(StateHasChanged);
    }

    public async Task SortByColumnAsync(string field, SortDirection direction)
    {
        foreach (var kvp in _columnStates) kvp.Value.SortDirection = null;
        GetColumnState(field).SortDirection = direction;
        _pageState.CurrentPage = 1;
        await InvokeAsync(StateHasChanged);
    }

    public async Task FilterByColumnAsync(string field, string value)
    {
        GetColumnState(field).FilterValue = value;
        _pageState.CurrentPage = 1;
        await InvokeAsync(StateHasChanged);
    }

    public async Task ClearFilteringAsync()
    {
        foreach (var kvp in _columnStates)
        {
            kvp.Value.FilterValue = null;
            kvp.Value.CheckedFilterValues.Clear();
        }
        _pageState.CurrentPage = 1;
        await InvokeAsync(StateHasChanged);
    }

    public async Task AddRecordAsync(TValue record)
    {
        var list = DataSource as IList<TValue>;
        list?.Add(record);
        await InvokeAsync(StateHasChanged);
    }

    public async Task DeleteRecordAsync(TValue record)
    {
        var list = DataSource as IList<TValue>;
        list?.Remove(record);
        _selectedItems.Remove(record);
        await InvokeAsync(StateHasChanged);
    }

    public async Task GroupByColumnAsync(string field)
    {
        await AddGroup(field);
        await InvokeAsync(StateHasChanged);
    }

    public async Task UngroupColumnAsync(string field)
    {
        await RemoveGroup(field);
        await InvokeAsync(StateHasChanged);
    }

    public async Task ClearGroupingAsync()
    {
        _groupDescriptors.Clear();
        _pageState.CurrentPage = 1;
        await InvokeAsync(StateHasChanged);
    }

    public async Task RefreshAsync() => await InvokeAsync(StateHasChanged);

    public Task Refresh() => RefreshAsync();

    public Task AutoFitColumnsAsync()
    {
        return Task.CompletedTask;
    }

    public Task ExportToExcelAsync()
    {
        return Task.CompletedTask;
    }

    public Task EndEditAsync()
    {
        _isEditing = false;
        _editingRowIndex = -1;
        _editItem = default;
        return InvokeAsync(StateHasChanged);
    }

    public Task ExpandAllGroupAsync()
    {
        _expandAllGroups = true;
        _allGroupsCollapsed = false;
        return InvokeAsync(StateHasChanged);
    }

    public Task ScrollIntoViewAsync(int columnIndex, int rowIndex)
    {
        return Task.CompletedTask;
    }

    public Task<object?> GetCellValueByIndexAsync(int rowIndex, int columnIndex)
    {
        var data = DataSource?.ToList() ?? [];
        if (rowIndex < 0 || rowIndex >= data.Count)
        {
            return Task.FromResult<object?>(null);
        }

        var columns = Columns;
        if (columnIndex < 0 || columnIndex >= columns.Count)
        {
            return Task.FromResult<object?>(null);
        }

        var field = columns[columnIndex].Field;
        var item = data[rowIndex];
        return Task.FromResult(GetPropertyValue(item, field));
    }

    public Task UpdateCellAsync(int rowIndex, string field, object? value)
    {
        var data = DataSource as IList<TValue> ?? DataSource?.ToList();
        if (data == null || rowIndex < 0 || rowIndex >= data.Count)
        {
            return Task.CompletedTask;
        }

        var item = data[rowIndex];
        if (item == null)
        {
            return Task.CompletedTask;
        }

        var prop = item.GetType().GetProperty(field, BindingFlags.Public | BindingFlags.Instance);
        if (prop == null || !prop.CanWrite)
        {
            return Task.CompletedTask;
        }

        var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
        object? convertedValue = value;
        if (value != null && targetType != value.GetType())
        {
            convertedValue = Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
        }
        prop.SetValue(item, convertedValue);

        return InvokeAsync(StateHasChanged);
    }
}
