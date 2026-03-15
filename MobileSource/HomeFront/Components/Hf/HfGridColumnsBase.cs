using Microsoft.AspNetCore.Components;
using System;

namespace HomeFront.Components.Hf;

/// <summary>
/// Base class that collects HfGridColumn children.
/// Acts as a CascadingValue so columns can register themselves.
/// Also registers itself with the parent HfGrid via cascading parameter.
/// </summary>
public class HfGridColumnsBase : ComponentBase
{
    private readonly List<HfGridColumn> _columns = new();

    [CascadingParameter] internal IHfGridOwner? ParentGrid { get; set; }

    [Parameter] public RenderFragment? ChildContent { get; set; }
    [Parameter] public string? DebugTag { get; set; }

    public IReadOnlyList<HfGridColumn> Columns => _columns;

    internal void AddColumn(HfGridColumn column)
    {
        if (!_columns.Contains(column))
            _columns.Add(column);
        ParentGrid?.NotifyColumnsChanged();
        if (false && !string.IsNullOrWhiteSpace(DebugTag))
            Console.WriteLine($"[HfGridColumnsBase:{DebugTag}] AddColumn field='{column.Field}' visible='{column.Visible}'");
    }

    protected override void OnInitialized()
    {
        ParentGrid?.RegisterColumnsContainer(this);
        if (false && !string.IsNullOrWhiteSpace(DebugTag))
            Console.WriteLine($"[HfGridColumnsBase:{DebugTag}] Initialized");
    }

    protected override void BuildRenderTree(Microsoft.AspNetCore.Components.Rendering.RenderTreeBuilder builder)
    {
        builder.OpenComponent<CascadingValue<HfGridColumnsBase>>(0);
        builder.AddComponentParameter(1, "Value", this);
        builder.AddComponentParameter(2, "ChildContent", ChildContent);
        builder.CloseComponent();
    }
}
