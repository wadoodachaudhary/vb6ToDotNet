const registrations = new WeakMap();

function number(value) {
    if (value == null || value === "") return null;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
}

function ms(value) {
    const parsed = number(value);
    return parsed == null ? "—" : `${parsed.toFixed(parsed < 10 ? 2 : 1)} ms`;
}

function percentile(values, fraction) {
    const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
    if (!sorted.length) return null;
    return sorted[Math.max(0, Math.ceil(sorted.length * fraction) - 1)];
}

function cell(row, value) {
    const td = document.createElement("td");
    td.textContent = value;
    row.appendChild(td);
}

function serverNumber(sample, key) {
    return number(sample?.serverMetrics?.[key]);
}

function render(host, samples) {
    const body = host.querySelector('[data-fx-role="samples"]');
    const summary = host.querySelector('[data-fx-role="summary"]');
    const status = host.querySelector('[data-fx-role="status"]');
    if (!body || !summary || !status) return;

    body.replaceChildren();
    const recent = samples.slice(-10).reverse();
    for (let index = 0; index < recent.length; index++) {
        const sample = recent[index];
        const row = document.createElement("tr");
        cell(row, String(samples.length - index));
        cell(row, ms(sample.releaseToPaintMs));
        cell(row, ms(sample.releaseHandlerDelayMs));
        cell(row, ms(sample.queueDelayMs));
        cell(row, ms(sample.invokePromiseDurationMs));
        cell(row, ms(sample.tokenDomMutationMs));
        cell(row, ms(sample.twoRafPaintMs));
        cell(row, ms(serverNumber(sample, "fxWindowMs")));
        cell(row, ms(serverNumber(sample, "fxRenderBuildMs")));
        cell(row, ms(serverNumber(sample, "fxGroupedBuildMs")));
        cell(row, `${sample.bodyRowsPresent ?? sample.rowsPresent ?? "—"} / ${serverNumber(sample, "fxWindowCount") ?? "—"}`);
        cell(row, ms(sample.longTaskOverlapMs));
        body.appendChild(row);
    }

    summary.replaceChildren();
    const metrics = [
        ["release→paint", "releaseToPaintMs"],
        ["input handler", "releaseHandlerDelayMs"],
        ["queue", "queueDelayMs"],
        [".NET invocation", "invokePromiseDurationMs"],
        ["Razor build", null]
    ];
    for (const [label, key] of metrics) {
        const values = key
            ? samples.map(sample => number(sample[key])).filter(Number.isFinite)
            : samples.map(sample => serverNumber(sample, "fxRenderBuildMs")).filter(Number.isFinite);
        const chip = document.createElement("span");
        chip.textContent = `${label}: p50 ${ms(percentile(values, 0.50))} · p95 ${ms(percentile(values, 0.95))} · max ${ms(percentile(values, 1))}`;
        summary.appendChild(chip);
    }

    const latest = samples.at(-1);
    status.textContent = latest
        ? `${samples.length} sample${samples.length === 1 ? "" : "s"}; latest ${ms(latest.releaseToPaintMs)} (${latest.invokeOutcome || "unknown outcome"})`
        : "Waiting for a ScrollTrack release…";
}

export function attachScrollTrackLatencyPanel(host, gridSelector) {
    detachScrollTrackLatencyPanel(host);
    if (!host || !gridSelector) return;

    const samples = [];
    const onTelemetry = event => {
        const grid = event.target instanceof Element ? event.target.closest(".fx-grid") : null;
        if (!grid || !grid.matches(gridSelector)) return;
        samples.push(event.detail || {});
        if (samples.length > 100) samples.shift();
        render(host, samples);
    };
    const onClick = event => {
        if (!(event.target instanceof Element)
            || !event.target.closest('[data-fx-action="clear"]')) return;
        samples.length = 0;
        render(host, samples);
    };

    document.addEventListener("fx-grid-deferred-scroll-telemetry", onTelemetry, true);
    host.addEventListener("click", onClick);
    registrations.set(host, { onTelemetry, onClick });
    render(host, samples);
}

export function detachScrollTrackLatencyPanel(host) {
    const registration = host ? registrations.get(host) : null;
    if (!registration) return;
    document.removeEventListener("fx-grid-deferred-scroll-telemetry", registration.onTelemetry, true);
    host.removeEventListener("click", registration.onClick);
    registrations.delete(host);
}
