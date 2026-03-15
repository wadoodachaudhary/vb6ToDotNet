using Microsoft.AspNetCore.Components;

namespace HomeFront.Components.Hf;

/// <summary>
/// Configures filtering behavior for HfGrid. Equivalent to SyncFusion's GridFilterSettings.
/// </summary>
public class HfFilterSettings : ComponentBase
{
    [Parameter] public FilterType Type { get; set; } = FilterType.FilterBar;
    [Parameter] public bool EnableCaseSensitivity { get; set; }
    [Parameter] public int ImmediateModeDelay { get; set; } = 300;
}
