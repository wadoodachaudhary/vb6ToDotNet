using Microsoft.AspNetCore.Components;

namespace HomeFront.Components.Hf;

/// <summary>
/// Configures editing for HfGrid. Equivalent to SyncFusion's GridEditSettings.
/// </summary>
public class HfEditSettings : ComponentBase
{
    [Parameter] public bool AllowEditing { get; set; }
    [Parameter] public bool AllowAdding { get; set; }
    [Parameter] public bool AllowDeleting { get; set; }
    [Parameter] public EditMode Mode { get; set; } = EditMode.Inline;
    [Parameter] public bool ShowConfirmDialog { get; set; } = true;
    [Parameter] public NewRowPosition NewRowPosition { get; set; } = NewRowPosition.Top;
    [Parameter] public bool AllowEditOnDblClick { get; set; } = true;
}
