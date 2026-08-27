using Microsoft.AspNetCore.SignalR;

namespace FlexKitTester.Services;

/// <summary>
/// Process-wide simulated uplink latency for the bench. Every inbound
/// SignalR invocation (keystroke dispatches, clicks, JS interop returns)
/// is delayed by <see cref="UplinkMs"/> before it reaches the circuit —
/// the same shape as a slow client link's upstream half. Render batches
/// going DOWN are not delayed, so treat the value as one-way latency.
/// </summary>
public static class BenchLatency
{
    public static volatile int UplinkMs;
}

public sealed class BenchLatencyHubFilter : IHubFilter
{
    public async ValueTask<object?> InvokeMethodAsync(
        HubInvocationContext invocationContext,
        Func<HubInvocationContext, ValueTask<object?>> next)
    {
        var ms = BenchLatency.UplinkMs;
        if (ms > 0)
            await Task.Delay(ms);
        return await next(invocationContext);
    }
}
