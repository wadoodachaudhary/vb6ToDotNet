using Microsoft.AspNetCore.Components;

namespace HomeFront.Components.Hf;

/// <summary>
/// Configures pagination for HfGrid. Equivalent to SyncFusion's GridPageSettings.
/// </summary>
public class HfPageSettings : ComponentBase
{
    [Parameter] public int PageSize { get; set; } = 10;
    [Parameter] public int[] PageSizes { get; set; } = [5, 10, 20, 50];
    [Parameter] public int PageCount { get; set; } = 5;
}
