function percentile(sorted, fraction) {
    if (!sorted.length) return null;
    const index = Math.max(0, Math.min(
        sorted.length - 1,
        Math.ceil(sorted.length * fraction) - 1));
    return sorted[index];
}

function summarize(probe, samples, failures) {
    const sorted = [...samples].sort((a, b) => a - b);
    return {
        probe,
        successCount: sorted.length,
        failureCount: failures,
        p50Milliseconds: percentile(sorted, 0.50),
        p95Milliseconds: percentile(sorted, 0.95),
        maximumMilliseconds: sorted.length ? sorted[sorted.length - 1] : null,
        status: sorted.length === 0 ? "Probe failed" : failures ? "Partial" : "OK"
    };
}

async function collect(probe, count, operation) {
    // Exclude one warm-up so module import/JIT/first connection setup does not
    // become a fake latency outlier.
    try {
        await operation(true);
    } catch {
        return summarize(probe, [], 1);
    }

    const samples = [];
    let failures = 0;
    for (let i = 0; i < count; i++) {
        const started = performance.now();
        try {
            await operation(false);
            samples.push(performance.now() - started);
        } catch {
            failures++;
            if (!samples.length && failures >= 2) break;
        }
    }
    return summarize(probe, samples, failures);
}

async function fetchNoStore(path) {
    const separator = path.includes("?") ? "&" : "?";
    const response = await fetch(
        `${path}${separator}sample=${Date.now()}-${Math.random()}`,
        {
            method: "GET",
            cache: "no-store",
            credentials: "same-origin",
            headers: { "Cache-Control": "no-cache" }
        });
    if (!response.ok) throw new Error("HTTP probe failed");
}

export async function measureBrowserLatency(dotNetRef, httpPath, requestedCount) {
    const count = Math.max(3, Math.min(30, Number(requestedCount) || 7));

    // Run sequentially so the two probes do not compete for the same browser or
    // server resources. No endpoint detail is returned to .NET.
    const circuit = await collect(
        "Blazor circuit RTT",
        count,
        () => dotNetRef.invokeMethodAsync("PingCircuitAsync"));
    const http = await collect(
        "Browser-server HTTP RTT",
        count,
        () => fetchNoStore(httpPath));

    return { circuit, http };
}
