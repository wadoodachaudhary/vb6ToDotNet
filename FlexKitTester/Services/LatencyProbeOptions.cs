namespace FlexKitTester.Services;

/// <summary>
/// Opt-in endpoints for the bench latency panel. Empty settings mean no
/// external probe is available. Hosts and connection strings are consumed only
/// by <see cref="LatencyProbeService"/> and are never returned to the UI.
/// </summary>
public sealed class LatencyProbeOptions
{
    public const string SectionName = "LatencyProbe";

    public int SampleCount { get; set; } = 7;
    public int TimeoutMilliseconds { get; set; } = 3000;
    public string DatabaseConnectionStringName { get; set; } = "BenchDatabase";
    public List<TcpProbeTargetOptions> TcpTargets { get; set; } = [];
}

public sealed class TcpProbeTargetOptions
{
    public TcpProbeRoute Route { get; set; }
    public string Host { get; set; } = string.Empty;
    public int Port { get; set; }
}

/// <summary>
/// Fixed route labels keep configured host names and addresses out of the UI.
/// </summary>
public enum TcpProbeRoute
{
    Unspecified,
    Lan,
    Vpn
}

public readonly record struct LatencyProbeAvailability(
    bool DatabaseConfigured,
    bool LanConfigured,
    bool VpnConfigured)
{
    public bool AnyInfrastructureConfigured =>
        DatabaseConfigured || LanConfigured || VpnConfigured;
}

public sealed record LatencyProbeResult(
    string Probe,
    int SuccessCount,
    int FailureCount,
    double? P50Milliseconds,
    double? P95Milliseconds,
    double? MaximumMilliseconds,
    string Status);

public sealed record LatencyProbeReport(IReadOnlyList<LatencyProbeResult> Results);

/// <summary>Browser-side timing DTO populated by latency-probe.js.</summary>
public sealed class BrowserLatencyReport
{
    public BrowserLatencyResult Circuit { get; set; } = new();
    public BrowserLatencyResult Http { get; set; } = new();
}

public sealed class BrowserLatencyResult
{
    public string Probe { get; set; } = string.Empty;
    public int SuccessCount { get; set; }
    public int FailureCount { get; set; }
    public double? P50Milliseconds { get; set; }
    public double? P95Milliseconds { get; set; }
    public double? MaximumMilliseconds { get; set; }
    public string Status { get; set; } = string.Empty;
}
