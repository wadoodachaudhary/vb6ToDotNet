using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace HomeFront.Components.Hf;

public partial class HfGrid<TValue> : IHfGridOwner
{
    void IHfGridOwner.RegisterColumnsContainer(HfGridColumnsBase container)
    {
        var changed = !ReferenceEquals(_columnsContainer, container);
        _columnsContainer = container;
        _autoWidthPending = true;
        if (changed)
            _ = InvokeAsync(StateHasChanged);
    }

    void IHfGridOwner.NotifyColumnsChanged()
    {
        _autoWidthPending = true;
        _ = InvokeAsync(StateHasChanged);
    }

    // ── Injectables ─────────────────────────────────────────────────────
    [Inject] private IJSRuntime JsRuntime { get; set; } = default!;

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

    /// <summary>When true, columns currently used for grouping are hidden from the data grid.</summary>
    [Parameter] public bool HideGroupedColumns { get; set; } = true;

    /// <summary>Show item count on group header rows.</summary>
    [Parameter] public bool ShowGroupCount { get; set; }

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
    private (int RowIndex, int CellIndex)? _activeCell;
    private bool _expandAllGroups;
    private bool _allGroupsCollapsed;
    private readonly HashSet<string> _collapsedGroupPaths = new(StringComparer.OrdinalIgnoreCase);
    private (int RowIndex, int CellIndex)? _lastSelectedCell;
    private int? _lastSelectedRowIndex;

    // Editing
    private bool _isEditing;
    private int _editingRowIndex = -1;
    private TValue? _editItem;

    // Batch editing (cell-level inline editing)
    private TValue? _batchEditItem;
    private int _batchEditRowIndex = -1;
    private string? _batchEditField;
    private string? _batchEditValue;

    // Filtering popup
    private string? _filterPopupField;

    // Type-ahead buffer (multi-select numeric input)
    private string _typeAheadBuffer = "";

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
    private bool _autoWidthPending = true;

    // ── Computed Properties ──────────────────────────────────────────────

    private bool ShowCheckboxColumn =>
        SelectionSettingsRef?.CheckboxOnly == true ||
        VisibleColumns.Any(c => c.Type == ColumnType.CheckBox && string.IsNullOrEmpty(c.Field));

    private FilterType ResolvedFilterType =>
        FilterSettingsRef?.Type ?? FilterType.FilterBar;

    private int[] ResolvedPageSizes =>
        PageSettingsRef?.PageSizes ?? [5, 10, 20, 50];

    private IEnumerable<HfGridColumn> VisibleColumns
    {
        get
        {
            var cols = _columnsContainer?.Columns.Where(c => c.Visible) ?? Enumerable.Empty<HfGridColumn>();
            if (HideGroupedColumns && _groupDescriptors.Count > 0)
            {
                var groupedFields = _groupDescriptors.Select(g => g.Field).ToHashSet(StringComparer.OrdinalIgnoreCase);
                cols = cols.Where(c => !groupedFields.Contains(c.Field));
            }
            return cols;
        }
    }

    public IReadOnlyList<HfGridColumn> Columns =>
        _columnsContainer?.Columns ?? Array.Empty<HfGridColumn>();

    private bool AllRowsSelected =>
        PagedData.Any() && PagedData.All(item => _selectedItems.Contains(item));

    private int GroupedPlaceholderCount =>
        (AllowGrouping && HideGroupedColumns) ? _groupDescriptors.Count : 0;

    private int TotalColumnCount =>
        VisibleColumns.Count()
        + (ShowCheckboxColumn ? 1 : 0)
        + GroupedPlaceholderCount;

    private IReadOnlyList<HfGridColumn> GroupedLayoutColumns
    {
        get
        {
            if (!(AllowGrouping && HideGroupedColumns) || _groupDescriptors.Count == 0)
                return Array.Empty<HfGridColumn>();

            var result = new List<HfGridColumn>(_groupDescriptors.Count);
            foreach (var gd in _groupDescriptors)
            {
                var col = Columns.FirstOrDefault(c => string.Equals(c.Field, gd.Field, StringComparison.OrdinalIgnoreCase));
                if (col != null)
                    result.Add(col);
            }
            return result;
        }
    }

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
            return BuildGroups(data, 0, "");
        }
    }

    private IEnumerable<GroupResult<TValue>> BuildGroups(IEnumerable<TValue> data, int level, string parentPath)
    {
        if (level >= _groupDescriptors.Count)
            return Enumerable.Empty<GroupResult<TValue>>();

        var gd = _groupDescriptors[level];
        var groups = data
            .GroupBy(item => GetPropertyValue(item, gd.Field)?.ToString() ?? "(empty)")
            .Select(g =>
            {
                var allItems = g.ToList();
                var groupPath = string.IsNullOrEmpty(parentPath)
                    ? $"{gd.Field}:{g.Key}"
                    : $"{parentPath}/{gd.Field}:{g.Key}";
                var group = new GroupResult<TValue>
                {
                    Field = gd.Field,
                    HeaderText = gd.HeaderText,
                    Key = g.Key,
                    GroupPath = groupPath,
                    Count = allItems.Count,
                    Items = level == _groupDescriptors.Count - 1 ? allItems : Enumerable.Empty<TValue>(),
                    SubGroups = level < _groupDescriptors.Count - 1
                        ? BuildGroups(allItems, level + 1, groupPath)
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
            ApplyGroupCollapseState(groups, collapsed: false);
        }

        if (_allGroupsCollapsed)
        {
            ApplyGroupCollapseState(groups, collapsed: true);
        }
        else if (!_expandAllGroups)
        {
            foreach (var group in groups)
                group.IsCollapsed = _collapsedGroupPaths.Contains(group.GroupPath);
        }

        // Reset bulk flags after applying so individual toggles work
        if (level == 0)
        {
            _expandAllGroups = false;
            _allGroupsCollapsed = false;
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
        var text = template.Replace("{value}", formatted);
        text = text.Replace("Grand Total:", "", StringComparison.OrdinalIgnoreCase)
                   .Replace("Sub Total:", "", StringComparison.OrdinalIgnoreCase)
                   .Replace("Total:", "", StringComparison.OrdinalIgnoreCase)
                   .Trim();
        return text;
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
            _collapsedGroupPaths.Clear();
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

        if (_autoWidthPending)
        {
            _autoWidthPending = false;
            if (EnsureAutoColumnWidths())
                StateHasChanged();
        }
    }

    protected override void OnParametersSet()
    {
        _pageState.PageSize = PageSettingsRef?.PageSize ?? _pageState.PageSize;
        _autoWidthPending = true;
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
            if (!isCtrl && !isShift && selType == SelectionType.Multiple)
            {
                // When multiple rows are already selected, a plain click
                // keeps the selection intact so batch-edit double-click
                // doesn't lose the multi-selection.
                if (_selectedItems.Count <= 1)
                {
                    _selectedItems.Clear();
                    _selectedItems.Add(item);
                }
                else
                {
                    // Multi-selection exists — just ensure clicked row is included
                    _selectedItems.Add(item);
                }
            }
            else if (isShift && _lastSelectedRowIndex.HasValue)
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
        if (!AllowSelection)
            return;

        var resolvedRowIndex = ResolveRowIndex(item, rowIndex);
        SetActiveEditableCell(resolvedRowIndex, cellIndex);

        // When grouped, route selection through the cell click to avoid missed row clicks.
        if (AllowGrouping && _groupDescriptors.Count > 0 &&
            SelectionSettingsRef?.Mode != SelectionMode.Cell &&
            SelectionSettingsRef?.CheckboxOnly != true)
        {
            if (EventsRef?.OnRecordClick.HasDelegate == true)
                await EventsRef.OnRecordClick.InvokeAsync(new CellClickEventArgs<TValue> { Data = item, RowIndex = rowIndex });

            await SelectRow(item, rowIndex, args);
            return;
        }

        if (SelectionSettingsRef?.Mode != SelectionMode.Cell)
            return;

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

    private void SetActiveEditableCell(int rowIndex, int cellIndex)
    {
        var col = VisibleColumns.ElementAtOrDefault(cellIndex);
        if (col != null && col.AllowEditing)
            _activeCell = (rowIndex, cellIndex);
        else
            _activeCell = null;
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

    // ── Batch Cell Editing ─────────────────────────────────────────────

    private async Task StartBatchEdit(TValue item, int rowIndex, HfGridColumn col)
    {
        if (!col.AllowEditing || string.IsNullOrEmpty(col.Field) || col.IsPrimaryKey) return;
        if (EditSettingsRef?.AllowEditing != true || EditSettingsRef.Mode != EditMode.Batch) return;

        // Save previous edit if any
        await CommitBatchEdit();

        if (EventsRef?.OnCellEdit.HasDelegate == true)
        {
            var args = new CellEditArgs<TValue> { Data = item, ColumnName = col.Field };
            await EventsRef.OnCellEdit.InvokeAsync(args);
        }

        _batchEditItem = item;
        _batchEditRowIndex = ResolveRowIndex(item, rowIndex);
        _batchEditField = col.Field;
        _batchEditValue = GetPropertyValue(item, col.Field)?.ToString() ?? "";
    }

    private async Task CommitBatchEdit()
    {
        if (_batchEditItem == null || string.IsNullOrEmpty(_batchEditField)) return;

        var oldValue = GetPropertyValue(_batchEditItem, _batchEditField);
        SetPropertyValue(_batchEditItem, _batchEditField, _batchEditValue);

        if (EventsRef?.OnCellSave.HasDelegate == true)
        {
            await EventsRef.OnCellSave.InvokeAsync(new CellSaveArgs<TValue>
            {
                Data = _batchEditItem,
                ColumnName = _batchEditField,
                Value = _batchEditValue
            });
        }

        _batchEditItem = default;
        _batchEditRowIndex = -1;
        _batchEditField = null;
        _batchEditValue = null;
    }

    private void UpdateBatchEditValue(ChangeEventArgs e)
    {
        _batchEditValue = e.Value?.ToString() ?? "";
    }

    private async Task HandleBatchEditKeyDown(KeyboardEventArgs e)
    {
        if (e.Key == "Enter" || e.Key == "Tab")
        {
            await CommitBatchEdit();
        }
        else if (e.Key == "Escape")
        {
            // Cancel edit without saving
            _batchEditItem = default;
            _batchEditRowIndex = -1;
            _batchEditField = null;
            _batchEditValue = null;
        }
    }

    private bool IsBatchEditing(TValue item, string field)
    {
        if (_batchEditItem == null || _batchEditField != field) return false;
        return EqualityComparer<TValue>.Default.Equals(_batchEditItem, item);
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
            case "csv":
                await ExportToCsvAsync();
                break;
            case "excel":
                await ExportToExcelAsync();
                break;
            case "pdf":
                await ExportToPdfAsync();
                break;
        }

        if (OnToolbarItemClick.HasDelegate)
            await OnToolbarItemClick.InvokeAsync(item);
    }

    // ── Keyboard Navigation ──────────────────────────────────────────────

    private async Task HandleKeyDown(KeyboardEventArgs e)
    {
        if (e.Key == "Escape" && _isEditing)
        {
            CancelEdit();
            return;
        }

        // Type-ahead: when multiple rows are selected and no batch edit is active,
        // let user type digits and press Enter to apply the value.
        if (_selectedItems.Count > 1 && _batchEditItem == null)
        {
            if (e.Key.Length == 1 && (char.IsDigit(e.Key[0]) || e.Key[0] == '.'))
            {
                // Prevent multiple decimal points
                if (e.Key == "." && _typeAheadBuffer.Contains('.'))
                    return;
                _typeAheadBuffer += e.Key;
                return;
            }

            if (e.Key == "Backspace" && _typeAheadBuffer.Length > 0)
            {
                _typeAheadBuffer = _typeAheadBuffer[..^1];
                return;
            }

            if (e.Key == "Escape" && _typeAheadBuffer.Length > 0)
            {
                _typeAheadBuffer = "";
                return;
            }

            if (e.Key == "Enter" && _typeAheadBuffer.Length > 0)
            {
                if (EventsRef?.OnTypeAheadCommit.HasDelegate == true)
                {
                    await EventsRef.OnTypeAheadCommit.InvokeAsync(new TypeAheadCommitArgs<TValue>
                    {
                        SelectedItems = _selectedItems.ToList(),
                        Value = _typeAheadBuffer
                    });
                }
                _typeAheadBuffer = "";
                return;
            }
        }
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
        if (group.IsCollapsed)
            _collapsedGroupPaths.Add(group.GroupPath);
        else
            _collapsedGroupPaths.Remove(group.GroupPath);
    }

    private void ApplyGroupCollapseState(IEnumerable<GroupResult<TValue>> groups, bool collapsed)
    {
        foreach (var group in groups)
        {
            group.IsCollapsed = collapsed;
            if (collapsed)
                _collapsedGroupPaths.Add(group.GroupPath);
            else
                _collapsedGroupPaths.Remove(group.GroupPath);

            if (group.SubGroups.Any())
                ApplyGroupCollapseState(group.SubGroups, collapsed);
        }
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
            builder.AddContent(42, $"{group.Key}");
            builder.CloseElement();

            builder.OpenElement(50, "span");
            builder.AddAttribute(51, "class", "hf-group-header-right");

            if (ShowGroupCount)
            {
                builder.OpenElement(52, "span");
                builder.AddAttribute(53, "class", "hf-group-count");
                builder.AddContent(54, $"({group.Count} items)");
                builder.CloseElement();
            }

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

            builder.CloseElement(); // right container

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

                        // Grouped column placeholders (align data columns when grouped columns are hidden)
                        if (AllowGrouping && HideGroupedColumns && GroupedLayoutColumns.Count > 0)
                        {
                            foreach (var gcol in GroupedLayoutColumns)
                            {
                                builder.OpenElement(80, "td");
                                builder.AddAttribute(81, "class", "hf-cell hf-group-indent");
                                builder.AddAttribute(82, "style", GetGroupedPlaceholderStyle(gcol));
                                builder.CloseElement();
                            }
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
                            var capturedCol = col;
                            var capturedItemForEdit = item;
                            var isBatchEditing = IsBatchEditing(item, col.Field);
                            var isCellSelected = _selectedCells.Contains((resolvedRowIdx, capturedColIdx));
                            var isActiveCell = _activeCell.HasValue
                                && _activeCell.Value.RowIndex == resolvedRowIdx
                                && _activeCell.Value.CellIndex == capturedColIdx
                                && capturedCol.AllowEditing;
                            var editableClass = capturedCol.AllowEditing ? " hf-cell-editable" : string.Empty;
                            builder.OpenElement(100, "td");
                            builder.AddAttribute(101, "class",
                                (isCellSelected ? "hf-cell hf-cell-selected"
                                : isBatchEditing ? "hf-cell hf-batch-editing"
                                : isActiveCell ? "hf-cell hf-cell-active-editable"
                                : "hf-cell") + editableClass);
                            builder.AddAttribute(102, "style", col.GetCellStyle());

                            // Batch edit on double-click
                            if (EditSettingsRef?.Mode == EditMode.Batch && col.AllowEditing && !col.IsPrimaryKey && !string.IsNullOrEmpty(col.Field))
                            {
                                builder.AddAttribute(104, "ondblclick", EventCallback.Factory.Create<MouseEventArgs>(this, _ => StartBatchEdit(capturedItemForEdit, resolvedRowIdx, capturedCol)));
                            }
                            builder.AddAttribute(103, "onclick", EventCallback.Factory.Create<MouseEventArgs>(this, args => HandleCellClick(item, resolvedRowIdx, capturedColIdx, args)));
                            builder.AddEventStopPropagationAttribute(105, "onclick", true);

                            if (isBatchEditing)
                            {
                                var inputType = col.Type == ColumnType.Number ? "number"
                                    : col.Type == ColumnType.Date ? "date" : "text";
                                builder.OpenElement(145, "input");
                                builder.AddAttribute(146, "type", inputType);
                                builder.AddAttribute(147, "class", "hf-batch-input");
                                builder.AddAttribute(148, "value", _batchEditValue);
                                builder.AddAttribute(149, "oninput", EventCallback.Factory.Create<ChangeEventArgs>(this, UpdateBatchEditValue));
                                builder.AddAttribute(150, "onkeydown", EventCallback.Factory.Create<KeyboardEventArgs>(this, HandleBatchEditKeyDown));
                                builder.AddAttribute(151, "onblur", EventCallback.Factory.Create(this, CommitBatchEdit));
                                builder.CloseElement();
                            }
                            else if (col.Commands is { Count: > 0 })
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
            var capturedCol = col;
            var capturedColIdx = colIdx;
            var isBatchEditing = IsBatchEditing(item, col.Field);

            builder.OpenElement(0, "td");
            var isCellSelected = _selectedCells.Contains((resolvedRowIndex, colIdx));
            var isActiveCell = _activeCell.HasValue
                && _activeCell.Value.RowIndex == resolvedRowIndex
                && _activeCell.Value.CellIndex == colIdx
                && capturedCol.AllowEditing;
            var editableClass = capturedCol.AllowEditing ? " hf-cell-editable" : string.Empty;
            builder.AddAttribute(1, "class",
                (isCellSelected ? "hf-cell hf-cell-selected"
                : isBatchEditing ? "hf-cell hf-batch-editing"
                : isActiveCell ? "hf-cell hf-cell-active-editable"
                : "hf-cell") + editableClass);
            builder.AddAttribute(2, "style", col.GetCellStyle());
            if (col.ClipMode == ClipMode.EllipsisWithTooltip)
                builder.AddAttribute(3, "title", GetCellDisplayValue(item, col));

            // For batch mode, double-click opens edit; single-click does cell selection
            if (EditSettingsRef?.Mode == EditMode.Batch && col.AllowEditing && !col.IsPrimaryKey && !string.IsNullOrEmpty(col.Field))
            {
                builder.AddAttribute(4, "ondblclick", EventCallback.Factory.Create<MouseEventArgs>(this, _ => StartBatchEdit(item, resolvedRowIndex, capturedCol)));
            }
            builder.AddAttribute(5, "onclick", EventCallback.Factory.Create<MouseEventArgs>(this, args => HandleCellClick(item, resolvedRowIndex, capturedColIdx, args)));

            if (isBatchEditing)
            {
                // Render inline edit input
                var inputType = col.Type == ColumnType.Number ? "number"
                    : col.Type == ColumnType.Date ? "date" : "text";
                builder.OpenElement(45, "input");
                builder.AddAttribute(46, "type", inputType);
                builder.AddAttribute(47, "class", "hf-batch-input");
                builder.AddAttribute(48, "value", _batchEditValue);
                builder.AddAttribute(49, "oninput", EventCallback.Factory.Create<ChangeEventArgs>(this, UpdateBatchEditValue));
                builder.AddAttribute(50, "onkeydown", EventCallback.Factory.Create<KeyboardEventArgs>(this, HandleBatchEditKeyDown));
                builder.AddAttribute(51, "onblur", EventCallback.Factory.Create(this, CommitBatchEdit));
                builder.AddElementReferenceCapture(52, _ => { }); // auto-focus handled via CSS
                builder.CloseElement();
            }
            else if (col.Commands is { Count: > 0 })
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
        if (TryGetDictionaryValue(item, field, out var dictValue))
            return dictValue;
        var prop = item.GetType().GetProperty(field, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
        return prop?.GetValue(item);
    }

    private static void SetPropertyValue(object? item, string field, string? value)
    {
        if (item == null || string.IsNullOrEmpty(field)) return;
        if (TrySetDictionaryValue(item, field, value))
            return;
        var prop = item.GetType().GetProperty(field, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
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

    private static bool TryGetDictionaryValue(object item, string field, out object? value)
    {
        if (item is IDictionary<string, object> dict)
        {
            if (dict.TryGetValue(field, out value))
                return true;
            foreach (var kvp in dict)
            {
                if (string.Equals(kvp.Key, field, StringComparison.OrdinalIgnoreCase))
                {
                    value = kvp.Value;
                    return true;
                }
            }
        }
        else if (item is IDictionary<string, object?> dictNullable)
        {
            if (dictNullable.TryGetValue(field, out value))
                return true;
            foreach (var kvp in dictNullable)
            {
                if (string.Equals(kvp.Key, field, StringComparison.OrdinalIgnoreCase))
                {
                    value = kvp.Value;
                    return true;
                }
            }
        }

        value = null;
        return false;
    }

    private static bool TrySetDictionaryValue(object item, string field, string? value)
    {
        if (item is IDictionary<string, object> dict)
        {
            dict[field] = value ?? "";
            return true;
        }
        if (item is IDictionary<string, object?> dictNullable)
        {
            dictNullable[field] = value;
            return true;
        }
        return false;
    }

    private string GetTableStyle()
    {
        var totalWidth = GetTotalColumnWidthPx();
        if (totalWidth <= 0)
            return "width:100%;";
        return $"width:{totalWidth}px;min-width:100%;";
    }

    private string GetGroupedPlaceholderStyle(HfGridColumn col)
    {
        var width = GetColumnWidthPx(col);
        if (width <= 0)
            width = 120;
        return $"width:{width}px;min-width:{width}px;max-width:{width}px;";
    }

    private double GetTotalColumnWidthPx()
    {
        var total = 0d;
        if (ShowCheckboxColumn)
            total += 50;

        if (AllowGrouping && HideGroupedColumns && GroupedLayoutColumns.Count > 0)
        {
            foreach (var gcol in GroupedLayoutColumns)
            {
                var width = GetColumnWidthPx(gcol);
                total += width > 0 ? width : 120;
            }
        }

        foreach (var col in VisibleColumns)
        {
            var width = GetColumnWidthPx(col);
            if (width > 0)
                total += width;
        }
        return total;
    }

    private static double GetColumnWidthPx(HfGridColumn col)
    {
        if (col.RuntimeWidth.HasValue)
            return col.RuntimeWidth.Value;

        var parsed = TryParseWidthPx(col.Width);
        return parsed ?? 0;
    }

    private static double? TryParseWidthPx(string? width)
    {
        if (string.IsNullOrWhiteSpace(width))
            return null;
        var trimmed = width.Trim();
        if (trimmed.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed[..^2];
        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var px))
            return px;
        return null;
    }

    private bool EnsureAutoColumnWidths()
    {
        if (_columnsContainer == null)
            return false;

        var sample = (DataSource?.Take(50).ToList()) ?? new List<TValue>();
        var changed = false;

        foreach (var col in _columnsContainer.Columns)
        {
            if (!col.Visible)
                continue;
            if (col.RuntimeWidth.HasValue)
                continue;
            if (!string.IsNullOrWhiteSpace(col.Width))
                continue;

            var width = EstimateColumnWidth(col, sample);
            col.RuntimeWidth = width;
            changed = true;
        }

        return changed;
    }

    private double EstimateColumnWidth(HfGridColumn col, IReadOnlyList<TValue> sample)
    {
        var maxLen = col.DisplayHeader?.Length ?? 0;

        foreach (var item in sample)
        {
            var val = GetPropertyValue(item, col.Field);
            if (val == null)
                continue;
            var text = Convert.ToString(val, CultureInfo.CurrentCulture) ?? "";
            if (text.Length > maxLen)
                maxLen = text.Length;
        }

        var px = (maxLen * 7.6) + 36;
        if (col.Type == ColumnType.Number)
            px += 12;
        return Math.Clamp(px, 80, 520);
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

    public int? GetCurrentRowIndex() => _lastSelectedRowIndex;

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

    /// <summary>Export grid data to CSV and trigger browser download.</summary>
    public async Task ExportToCsvAsync(string fileName = "export.csv")
    {
        var sb = new StringBuilder();
        var cols = VisibleColumns.ToList();
        // Header row
        sb.AppendLine(string.Join(",", cols.Select(c => EscapeCsvField(c.DisplayHeader))));
        // Data rows — all filtered+sorted data (not just current page)
        foreach (var item in SortedData)
        {
            var values = cols.Select(c => EscapeCsvField(GetCellDisplayValue(item, c)));
            sb.AppendLine(string.Join(",", values));
        }
        var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(sb.ToString()));
        await JsRuntime.InvokeVoidAsync("hfGridExportDownload", fileName, base64, "text/csv");
    }

    /// <summary>Export grid data to Excel-compatible HTML (.xls) and trigger browser download.</summary>
    public async Task ExportToExcelAsync(string fileName = "export.xls")
    {
        var sb = new StringBuilder();
        var cols = VisibleColumns.ToList();
        sb.Append("<table border='1'><thead><tr>");
        foreach (var c in cols)
            sb.Append($"<th>{System.Net.WebUtility.HtmlEncode(c.DisplayHeader)}</th>");
        sb.Append("</tr></thead><tbody>");
        foreach (var item in SortedData)
        {
            sb.Append("<tr>");
            foreach (var c in cols)
                sb.Append($"<td>{System.Net.WebUtility.HtmlEncode(GetCellDisplayValue(item, c))}</td>");
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table>");
        var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(sb.ToString()));
        await JsRuntime.InvokeVoidAsync("hfGridExportDownload", fileName, base64, "application/vnd.ms-excel");
    }

    /// <summary>Export grid data to a printable HTML table and open the browser print dialog (Save as PDF).</summary>
    public async Task ExportToPdfAsync(string title = "Export")
    {
        var sb = new StringBuilder();
        var cols = VisibleColumns.ToList();
        sb.Append("<html><head><title>").Append(System.Net.WebUtility.HtmlEncode(title)).Append("</title>");
        sb.Append("<style>table{border-collapse:collapse;width:100%;font-size:11px;font-family:Arial,sans-serif}th,td{border:1px solid #ccc;padding:4px 8px;text-align:left}th{background:#f0f0f0;font-weight:bold}@@media print{body{margin:0}}</style>");
        sb.Append("</head><body>");
        sb.Append("<table><thead><tr>");
        foreach (var c in cols)
            sb.Append($"<th>{System.Net.WebUtility.HtmlEncode(c.DisplayHeader)}</th>");
        sb.Append("</tr></thead><tbody>");
        foreach (var item in SortedData)
        {
            sb.Append("<tr>");
            foreach (var c in cols)
                sb.Append($"<td>{System.Net.WebUtility.HtmlEncode(GetCellDisplayValue(item, c))}</td>");
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table></body></html>");
        var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(sb.ToString()));
        await JsRuntime.InvokeVoidAsync("hfGridExportPdf", base64);
    }

    private static string EscapeCsvField(string field)
    {
        if (string.IsNullOrEmpty(field)) return "\"\"";
        if (field.Contains(',') || field.Contains('"') || field.Contains('\n'))
            return "\"" + field.Replace("\"", "\"\"") + "\"";
        return field;
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
        _collapsedGroupPaths.Clear();
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

        var prop = item.GetType().GetProperty(field, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
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
