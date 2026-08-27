using System.Diagnostics;
using System.Net.Sockets;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Options;

namespace FlexKitTester.Services;

/// <summary>
/// Runs explicitly-requested, read-only infrastructure probes from the bench
/// host. Configuration is intentionally one-way: results contain fixed labels
/// and timings, never target hosts, addresses, database names, logins, or
/// connection strings.
/// </summary>
public sealed class LatencyProbeService
{
    private const int MinimumSamples = 3;
    private const int MaximumSamples = 30;
    private const int MinimumTimeoutMilliseconds = 250;
    private const int MaximumTimeoutMilliseconds = 10_000;

    private readonly IConfiguration _configuration;
    private readonly LatencyProbeOptions _options;

    public LatencyProbeService(
        IConfiguration configuration,
        IOptions<LatencyProbeOptions> options)
    {
        _configuration = configuration;
        _options = options.Value;
    }

    public int SampleCount => Math.Clamp(
        _options.SampleCount,
        MinimumSamples,
        MaximumSamples);

    private int TimeoutMilliseconds => Math.Clamp(
        _options.TimeoutMilliseconds,
        MinimumTimeoutMilliseconds,
        MaximumTimeoutMilliseconds);

    public LatencyProbeAvailability GetAvailability()
    {
        var targets = ValidTargets().ToList();
        return new LatencyProbeAvailability(
            !string.IsNullOrWhiteSpace(GetDatabaseConnectionString()),
            targets.Any(target => target.Route == TcpProbeRoute.Lan),
            targets.Any(target => target.Route == TcpProbeRoute.Vpn));
    }

    public async Task<LatencyProbeReport> RunConfiguredProbesAsync(
        CancellationToken cancellationToken = default)
    {
        var results = new List<LatencyProbeResult>();
        var connectionString = GetDatabaseConnectionString();
        if (!string.IsNullOrWhiteSpace(connectionString))
        {
            var databaseResults = await ProbeDatabaseAsync(
                connectionString,
                cancellationToken);
            results.AddRange(databaseResults);
        }

        var routeCounts = new Dictionary<TcpProbeRoute, int>();
        foreach (var target in ValidTargets())
        {
            routeCounts.TryGetValue(target.Route, out var routeIndex);
            routeCounts[target.Route] = ++routeIndex;
            var routeTotal = ValidTargets().Count(item => item.Route == target.Route);
            var fixedLabel = target.Route == TcpProbeRoute.Lan
                ? "LAN TCP handshake"
                : "VPN TCP handshake";
            if (routeTotal > 1)
                fixedLabel += $" {routeIndex}";

            results.Add(await SampleAsync(
                fixedLabel,
                async token =>
                {
                    using var client = new TcpClient();
                    await client.ConnectAsync(target.Host, target.Port, token);
                },
                cancellationToken));
        }

        return new LatencyProbeReport(results);
    }

    private async Task<IReadOnlyList<LatencyProbeResult>> ProbeDatabaseAsync(
        string configuredConnectionString,
        CancellationToken cancellationToken)
    {
        SqlConnectionStringBuilder builder;
        try
        {
            builder = new SqlConnectionStringBuilder(configuredConnectionString)
            {
                // Every measured OpenAsync performs a real transport/TLS/login
                // path instead of merely checking a process-local pool.
                Pooling = false,
                ConnectTimeout = Math.Max(1, (int)Math.Ceiling(TimeoutMilliseconds / 1000d))
            };
        }
        catch
        {
            return
            [
                FailedResult("SQL new connection", "Configuration is invalid"),
                FailedResult("SQL SELECT 1", "Configuration is invalid")
            ];
        }

        // Do not retain or expose parsed connection-string properties.
        var probeConnectionString = builder.ConnectionString;
        var openResult = await SampleAsync(
            "SQL new connection",
            async token =>
            {
                await using var connection = new SqlConnection(probeConnectionString);
                await connection.OpenAsync(token);
            },
            cancellationToken);

        if (openResult.SuccessCount == 0)
        {
            return
            [
                openResult,
                FailedResult("SQL SELECT 1", "Skipped because SQL open failed")
            ];
        }

        LatencyProbeResult queryResult;
        try
        {
            await using var connection = new SqlConnection(probeConnectionString);
            await RunWithTimeoutAsync(
                token => connection.OpenAsync(token),
                cancellationToken);

            queryResult = await SampleAsync(
                "SQL SELECT 1",
                async token =>
                {
                    await using var command = connection.CreateCommand();
                    command.CommandText = "SELECT 1";
                    command.CommandTimeout = Math.Max(
                        1,
                        (int)Math.Ceiling(TimeoutMilliseconds / 1000d));
                    _ = await command.ExecuteScalarAsync(token);
                },
                cancellationToken);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch
        {
            queryResult = FailedResult("SQL SELECT 1", "SQL probe failed");
        }

        return [openResult, queryResult];
    }

    private async Task<LatencyProbeResult> SampleAsync(
        string label,
        Func<CancellationToken, Task> operation,
        CancellationToken cancellationToken)
    {
        // Warm the runtime/DNS path once. Warm-up is deliberately excluded from
        // statistics so first-use JIT does not masquerade as network latency.
        try
        {
            await RunWithTimeoutAsync(operation, cancellationToken);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            return FailedResult(label, "Timed out");
        }
        catch
        {
            return FailedResult(label, "Probe failed");
        }

        var timings = new List<double>(SampleCount);
        var failures = 0;
        for (var i = 0; i < SampleCount; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var started = Stopwatch.GetTimestamp();
            try
            {
                await RunWithTimeoutAsync(operation, cancellationToken);
                timings.Add(Stopwatch.GetElapsedTime(started).TotalMilliseconds);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch
            {
                failures++;
                // An unavailable endpoint should not hold the bench for every
                // configured sample timeout.
                if (timings.Count == 0 && failures >= 2)
                    break;
            }
        }

        if (timings.Count == 0)
            return FailedResult(label, "Probe failed", failures);

        timings.Sort();
        return new LatencyProbeResult(
            label,
            timings.Count,
            failures,
            Percentile(timings, 0.50),
            Percentile(timings, 0.95),
            timings[^1],
            failures == 0 ? "OK" : "Partial");
    }

    private async Task RunWithTimeoutAsync(
        Func<CancellationToken, Task> operation,
        CancellationToken cancellationToken)
    {
        using var timeoutSource = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken);
        timeoutSource.CancelAfter(TimeoutMilliseconds);
        await operation(timeoutSource.Token);
    }

    private string? GetDatabaseConnectionString()
    {
        if (string.IsNullOrWhiteSpace(_options.DatabaseConnectionStringName))
            return null;

        return _configuration.GetConnectionString(
            _options.DatabaseConnectionStringName);
    }

    private IEnumerable<TcpProbeTargetOptions> ValidTargets() =>
        (_options.TcpTargets ?? [])
        .Where(target =>
            target.Route is TcpProbeRoute.Lan or TcpProbeRoute.Vpn
            && !string.IsNullOrWhiteSpace(target.Host)
            && target.Port is > 0 and <= 65_535);

    private static double Percentile(IReadOnlyList<double> sorted, double percentile)
    {
        var index = Math.Clamp(
            (int)Math.Ceiling(percentile * sorted.Count) - 1,
            0,
            sorted.Count - 1);
        return sorted[index];
    }

    private static LatencyProbeResult FailedResult(
        string label,
        string status,
        int failures = 1) =>
        new(label, 0, Math.Max(1, failures), null, null, null, status);
}
