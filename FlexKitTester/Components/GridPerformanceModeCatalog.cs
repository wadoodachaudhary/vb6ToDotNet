using Fx.ControlKit.Grid;

namespace FlexKitTester.Components;

public sealed record GridPerformanceModeChoice(GridPerformanceMode Value, string Label);

public static class GridPerformanceModeCatalog
{
    public static IReadOnlyList<GridPerformanceModeChoice> All { get; } =
    [
        new(GridPerformanceMode.ServerBaseline, "Server Baseline"),
        new(GridPerformanceMode.ServerOptimized, "Server Optimized"),
        new(GridPerformanceMode.ClientPreview, "Client Preview"),
        new(GridPerformanceMode.ClientPreviewAndServerOptimized,
            "Client Preview + Server Optimized (Default)")
    ];
}
