using Microsoft.AspNetCore.Components;

namespace HomeFront.Components.Hf;

/// <summary>
/// Configures selection for HfGrid. Equivalent to SyncFusion's GridSelectionSettings.
/// </summary>
public class HfSelectionSettings : ComponentBase
{
    [Parameter] public SelectionType Type { get; set; } = SelectionType.Single;
    [Parameter] public SelectionMode Mode { get; set; } = SelectionMode.Row;
    [Parameter] public bool CheckboxOnly { get; set; }
    [Parameter] public bool PersistSelection { get; set; }
    [Parameter] public bool EnableToggle { get; set; } = true;
}
