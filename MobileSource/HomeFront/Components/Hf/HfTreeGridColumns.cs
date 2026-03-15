using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Rendering;

namespace HomeFront.Components.Hf;

/// <summary>
/// Container for HfTreeGridColumn definitions. Children are rendered hidden and register
/// themselves with the parent HfTreeGrid via CascadingParameter.
/// </summary>
public class HfTreeGridColumns : ComponentBase
{
    [Parameter] public RenderFragment? ChildContent { get; set; }

    protected override void BuildRenderTree(RenderTreeBuilder builder)
    {
        builder.AddContent(0, ChildContent);
    }
}
