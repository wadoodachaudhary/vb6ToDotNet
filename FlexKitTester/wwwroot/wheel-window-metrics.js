const registrations = new WeakMap();

function numeric(value) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
}

function percentile(samples, key, fraction) {
    const values = samples
        .map(sample => numeric(sample[key]))
        .filter(Number.isFinite)
        .sort((a, b) => a - b);
    if (!values.length) return null;
    return values[Math.max(0, Math.ceil(values.length * fraction) - 1)];
}

function format(value, suffix = "") {
    const number = numeric(value);
    if (number == null) return "—";
    return `${number.toFixed(Math.abs(number) < 10 ? 2 : 1)}${suffix}`;
}

function sameNumber(left, right, tolerance = 0.01) {
    const a = numeric(left);
    const b = numeric(right);
    if (a == null || b == null) return a === b;
    return Math.abs(a - b) <= tolerance;
}

function relativeClose(left, right, fraction, absoluteFloor) {
    const a = numeric(left);
    const b = numeric(right);
    if (a == null || b == null) return a === b;
    return Math.abs(a - b) <= Math.max(
        absoluteFloor,
        Math.max(Math.abs(a), Math.abs(b)) * fraction);
}

function dominantInputMode(sample) {
    const modes = [
        ["pixel", numeric(sample.pixelModeEvents) || 0],
        ["line", numeric(sample.lineModeEvents) || 0],
        ["page", numeric(sample.pageModeEvents) || 0]
    ];
    modes.sort((left, right) => right[1] - left[1]);
    return modes[0][1] > 0 ? modes[0][0] : "none";
}

function replaySampleIsComplete(sample) {
    if (!sample.replayKind) return true;
    const expectedEvents = numeric(sample.replayExpectedEvents);
    const expectedRawPx = numeric(sample.replayExpectedRawPx);
    return expectedEvents != null
        && expectedRawPx != null
        && (numeric(sample.wheelEvents) || 0) === expectedEvents
        && Math.abs((numeric(sample.totalAbsoluteRawDeltaPx) || 0) - expectedRawPx)
            <= 0.5;
}

function gestureMatches(sample, reference) {
    const viewport = Math.max(1, numeric(reference.viewportHeightPx) || 1);
    const replayMatches = (sample.replayKind || reference.replayKind)
        ? sample.replayKind === reference.replayKind
            && sample.replayVersion === reference.replayVersion
            && replaySampleIsComplete(sample)
            && replaySampleIsComplete(reference)
        : true;
    return replayMatches
        && Math.sign(numeric(sample.firstDirection) || 0)
            === Math.sign(numeric(reference.firstDirection) || 0)
        && (numeric(sample.directionReversals) || 0)
            === (numeric(reference.directionReversals) || 0)
        && dominantInputMode(sample) === dominantInputMode(reference)
        && relativeClose(
            sample.totalAbsoluteRawDeltaPx,
            reference.totalAbsoluteRawDeltaPx,
            0.15,
            10)
        && relativeClose(sample.wheelEvents, reference.wheelEvents, 0.15, 2)
        && relativeClose(
            sample.averageInputIntervalMs,
            reference.averageInputIntervalMs,
            0.30,
            4)
        && relativeClose(
            sample.startingDirectionalRunwayPx,
            reference.startingDirectionalRunwayPx,
            0.20,
            viewport * 0.25)
        && sameNumber(
            sample.startingScrollTop,
            reference.startingScrollTop,
            Math.max(1, numeric(reference.configuredRowHeightPx) || 1));
}

function appendCell(row, value, className = "") {
    const cell = document.createElement("td");
    cell.textContent = value;
    if (className) cell.className = className;
    row.appendChild(cell);
}

function guardSamplePasses(sample) {
    return (numeric(sample.maxVisibleSpacerPx) || 0) <= 1
        && (numeric(sample.droppedIntentPx) || 0) <= 0.5
        && (numeric(sample.finalPositionErrorPx) || 0) <= 1
        && (numeric(sample.maxInFlight) || 0) <= 1;
}

function pacingSamplePasses(sample) {
    return guardSamplePasses(sample)
        && (!sample.adaptivePacingEnabled
            || (numeric(sample.adaptiveAcceptedInputRatio) || 0) >= 0.03);
}

function comparisonChip(samples, enabled) {
    const subset = samples.filter(sample => !!sample.guardEnabled === enabled);
    const chip = document.createElement("span");
    if (!subset.length) {
        chip.textContent = `${enabled ? "ON" : "OFF"}: no samples`;
        return chip;
    }
    const maxSpacer = Math.max(...subset.map(sample => numeric(sample.maxVisibleSpacerPx) || 0));
    const maxDropped = Math.max(...subset.map(sample => numeric(sample.droppedIntentPx) || 0));
    const maxError = Math.max(...subset.map(sample => numeric(sample.finalPositionErrorPx) || 0));
    const p95Hold = percentile(subset, "maxClampMs", 0.95);
    const totalRequests = subset.reduce(
        (sum, sample) => sum + (numeric(sample.invocations) || 0),
        0);
    chip.textContent = `${enabled ? "ON" : "OFF"}: ${subset.length} bursts · max spacer ${format(maxSpacer, " px")} · max error ${format(maxError, " px")} · dropped ${format(maxDropped, " px")} · p95 hold ${format(p95Hold, " ms")} · ${totalRequests} requests`;
    chip.className = enabled && subset.every(guardSamplePasses)
        ? "wheel-window-metrics-pass"
        : maxSpacer > 1 || maxDropped > 0.5 || maxError > 1
            ? "wheel-window-metrics-fail"
            : "";
    return chip;
}

function adaptiveComparisonChip(samples, enabled) {
    const candidates = samples.filter(sample =>
        !!sample.guardEnabled
        && !!sample.adaptivePacingEnabled === enabled);
    const subset = candidates.filter(sample =>
        (numeric(sample.physicalGestures) || 1) === 1
        && !sample.hiddenAtStart
        && !sample.hiddenDuring
        && (!enabled || sample.adaptiveCalibrationPhaseAtStart === "steady"));
    const chip = document.createElement("span");
    if (!subset.length) {
        const warmups = enabled
            ? candidates.filter(
                sample => sample.adaptiveCalibrationPhaseAtStart !== "steady").length
            : 0;
        chip.textContent = `PACE ${enabled ? "ON" : "OFF"}: ${warmups ? `${warmups} warm-up; ` : ""}no clean ${enabled ? "steady" : "baseline"} samples`;
        return chip;
    }
    const p95Hold = percentile(subset, "maxClampMs", 0.95);
    const maxCatchup = Math.max(...subset.map(sample => numeric(sample.maximumTokenCatchupPx) || 0));
    const totalHits = subset.reduce(
        (sum, sample) => sum + (numeric(sample.adaptiveRateHits) || 0),
        0);
    const totalLimited = subset.reduce(
        (sum, sample) => sum + (numeric(sample.adaptiveRateLimitedInputPx) || 0),
        0);
    const totalDecreases = subset.reduce(
        (sum, sample) => sum + (numeric(sample.adaptiveImmediateDecreases) || 0),
        0);
    const latest = subset.at(-1);
    chip.textContent = `PACE ${enabled ? "ON" : "OFF"}: ${subset.length} clean bursts · current ${format(latest?.adaptiveMinimumRateRowsPerSecond, " rows/s")}:${format(latest?.adaptiveMaximumRateRowsPerSecond, " rows/s")} → next ${format(latest?.adaptiveNextRateRowsPerSecond, " rows/s")} · p95 hold ${format(p95Hold, " ms")} · max catch-up ${format(maxCatchup, " px")} · ${totalHits} limits / ${totalDecreases} fast decreases / ${format(totalLimited, " px")}`;
    chip.className = subset.every(pacingSamplePasses)
        ? "wheel-window-metrics-pass"
        : "wheel-window-metrics-fail";
    return chip;
}

function render(host, samples) {
    const body = host.querySelector('[data-fx-role="samples"]');
    const comparison = host.querySelector('[data-fx-role="comparison"]');
    const status = host.querySelector('[data-fx-role="status"]');
    if (!body || !comparison || !status) return;

    body.replaceChildren();
    const recent = samples.slice(-20).reverse();
    for (let index = 0; index < recent.length; index++) {
        const sample = recent[index];
        const row = document.createElement("tr");
        const spacer = numeric(sample.maxVisibleSpacerPx) || 0;
        appendCell(row, String(samples.length - index));
        appendCell(row, sample.guardEnabled ? "ON" : "OFF",
            sample.guardEnabled && guardSamplePasses(sample)
                ? "wheel-window-metrics-pass"
                : spacer > 1
                    || (numeric(sample.droppedIntentPx) || 0) > 0.5
                    || (numeric(sample.finalPositionErrorPx) || 0) > 1
                    ? "wheel-window-metrics-fail"
                    : "");
        appendCell(row, `${sample.windowMode || "?"} / ${sample.performanceMode || "?"} / ${sample.diagnosticDelayMs || 0} ms`);
        appendCell(row,
            `${format(sample.totalAbsoluteRawDeltaPx, " px")} → ${format(sample.totalAbsoluteDeltaPx, " px")} → ${format(sample.totalAbsolutePacedDeltaPx, " px")} → ${format(sample.totalAbsoluteSlowedDeltaPx, " px")} → ${format(sample.totalAbsoluteViewportDeltaPx, " px")} (${format((numeric(sample.wheelDeltaScale) ?? 1) * 100, "%")})`);
        appendCell(row,
            `${sample.adaptivePacingEnabled ? "ON" : "OFF"} ${format(sample.adaptiveMinimumRateRowsPerSecond, " rows/s")}:${format(sample.adaptiveMaximumRateRowsPerSecond, " rows/s")} → ${format(sample.adaptiveLearnedRateRowsPerSecond, " rows/s")} → ${format(sample.adaptivePendingRateRowsPerSecond, " rows/s")} / bucket ${format(sample.adaptiveRateCreditCapacityPx == null ? null : sample.adaptiveRateCreditCapacityPx / Math.max(1, numeric(sample.configuredRowHeightPx) || 1), " rows")} / ${sample.adaptiveRateHits ?? 0} / D=${sample.adaptiveImmediateDecreases ?? 0} C=${sample.adaptiveCongestionSamples ?? 0} R=${sample.adaptiveCongestionRecoveries ?? 0} / ${format(sample.adaptiveRateLimitedInputPx, " px")} (${format((numeric(sample.adaptiveAcceptedInputRatio) ?? 1) * 100, "%")}) / ${format(sample.adaptiveLatencyP95Ms, " ms")} / ${format(sample.adaptiveAdvanceP10Px, " px")} / N=${sample.adaptiveLatencySamplesAtStart ?? 0}→${sample.adaptiveLatencySamples ?? 0} ${sample.adaptiveCalibrationPhaseAtStart || "?"}→${sample.adaptiveCalibrationPhase || "?"} G=${sample.physicalGestures ?? 1}`);
        appendCell(row,
            `${sample.slowdownEvents ?? 0} / ${format((numeric(sample.minimumBoundaryGain) ?? 1) * 100, "%")} / ${format(sample.maximumIntentLagPx, " px")} / ${format(sample.tokenCatchupPx, " px")}:${format(sample.maximumTokenCatchupPx, " px")} / ${format(sample.maximumPendingIntentPx, " px")}`);
        appendCell(row,
            `${sample.hardClampCount ?? 0} / ${sample.zeroRunwayStops ?? 0} / ${sample.staleDirectionLimits ?? 0}`);
        appendCell(row, String(sample.wheelEvents ?? "—"));
        appendCell(row, format(spacer, " px"));
        appendCell(row, `${sample.blankFrames ?? 0} / ${format(sample.blankDurationMs, " ms")}`);
        appendCell(row, String(sample.clampCount ?? 0));
        appendCell(row, format(sample.maxClampMs, " ms"));
        appendCell(row, String(sample.invocations ?? 0));
        appendCell(row, String(sample.noRenderResponses ?? 0));
        appendCell(row, String(sample.boundaryRetries ?? 0));
        appendCell(row, `${sample.domAckTimeouts ?? 0} / ${sample.livenessRecoveries ?? 0}`);
        appendCell(row, String(sample.maxInFlight ?? 0));
        appendCell(row, `${sample.boundaryInvocations ?? 0} / ${sample.normalInvocations ?? 0}`);
        appendCell(row, String(sample.coalesced ?? 0));
        appendCell(row, format(sample.finalPositionErrorPx, " px"));
        appendCell(row, format(sample.droppedIntentPx, " px"));
        body.appendChild(row);
    }

    const latest = samples.at(-1);
    const comparable = latest && !latest.hiddenAtStart && !latest.hiddenDuring
        ? samples.filter(sample =>
            (sample.windowMode || "unknown") === (latest.windowMode || "unknown")
            && (sample.performanceMode || "unknown")
                === (latest.performanceMode || "unknown")
            && (numeric(sample.diagnosticDelayMs) || 0) === (numeric(latest.diagnosticDelayMs) || 0)
            && (numeric(sample.windowOverscanRows) || 0) === (numeric(latest.windowOverscanRows) || 0)
            && (numeric(sample.deferredOverscanRows) || 0) === (numeric(latest.deferredOverscanRows) || 0)
            && (numeric(sample.wheelDeltaScale) || 1) === (numeric(latest.wheelDeltaScale) || 1)
            && (numeric(sample.policyVersion) || 1) === (numeric(latest.policyVersion) || 1)
            && sameNumber(sample.viewportHeightPx, latest.viewportHeightPx, 1)
            && sameNumber(sample.configuredRowHeightPx, latest.configuredRowHeightPx)
            && sameNumber(sample.zoomFactor, latest.zoomFactor, 0.001)
            && !sample.hiddenAtStart
            && !sample.hiddenDuring
            && gestureMatches(sample, latest))
        : [];
    comparison.replaceChildren(
        comparisonChip(comparable, false),
        comparisonChip(comparable, true),
        adaptiveComparisonChip(comparable, false),
        adaptiveComparisonChip(comparable, true));
    status.textContent = latest
        ? `${samples.length} burst${samples.length === 1 ? "" : "s"}; comparing ${comparable.length} matching ${latest.windowMode || "unknown"}/${latest.diagnosticDelayMs || 0} ms/${format(latest.viewportHeightPx, " px viewport")}/${format((numeric(latest.wheelDeltaScale) || 1) * 100, "%")}; latest guard ${latest.guardEnabled ? "ON" : "OFF"}, adaptive pacing ${latest.adaptivePacingEnabled ? "ON" : "OFF"}, max spacer ${format(latest.maxVisibleSpacerPx, " px")}`
        : "Waiting for a wheel or trackpad burst…";
}

function wait(milliseconds) {
    return new Promise(resolve => setTimeout(resolve, milliseconds));
}

async function waitUntil(predicate, generationIsCurrent, timeoutMs = 12_000) {
    const deadline = performance.now() + timeoutMs;
    while (generationIsCurrent() && performance.now() < deadline) {
        if (predicate()) return true;
        await wait(25);
    }
    return false;
}

async function replayWheelTrace(
    host,
    gridSelector,
    kind,
    runId,
    generationIsCurrent) {
    const grid = document.querySelector(gridSelector);
    const scrollEl = grid?.querySelector(".fx-grid-content");
    const status = host.querySelector('[data-fx-role="status"]');
    if (!scrollEl) {
        if (status) status.textContent = "Replay failed: grid scroll viewport not found.";
        return;
    }

    const lane = grid.querySelector(".fx-grid-deferred-vscroll");
    if (!lane) {
        if (status) status.textContent = "Replay requires ScrollTrack release-only mode.";
        return;
    }
    clearReplayTags(scrollEl);
    if (status) status.textContent = "Resetting replay to the top and waiting for idle…";
    lane.dispatchEvent(new KeyboardEvent("keydown", {
        key: "Home",
        bubbles: true,
        cancelable: true
    }));
    const resetReady = await waitUntil(
        () => lane.getAttribute("aria-busy") !== "true"
            && Math.abs(scrollEl.scrollTop) <= 0.5,
        generationIsCurrent);
    if (!resetReady) {
        if (status) status.textContent = "Replay cancelled: top/idle reset timed out.";
        return;
    }
    // Let the Home commit's wheel telemetry/queue settle before tagging the
    // controlled trace. This prevents a re-click from merging two cohorts.
    await wait(320);
    if (!generationIsCurrent()) return;

    // Both traces normalize to the same 30-row distance using the row pitch
    // published by GridControl. They differ only in cadence/event shape,
    // exposing mouse-vs-trackpad pacing without assuming a 14px bench row.
    const configuredRowHeight = Math.max(
        1,
        numeric(scrollEl.dataset.fxConfiguredRowHeight)
            || grid.querySelector("tbody tr.fx-row")?.getBoundingClientRect().height
            || 16);
    const expectedRawPx = 30 * configuredRowHeight;
    const trace = kind === "mouse"
        ? { count: 10, deltaY: 3, deltaMode: 1, intervalMs: 80 }
        : { count: 60, deltaY: expectedRawPx / 60, deltaMode: 0, intervalMs: 12 };
    if (status) status.textContent = `Replaying ${kind} trace…`;
    scrollEl.dataset.fxWheelReplayKind = kind;
    scrollEl.dataset.fxWheelReplayVersion = "2-row-pitch";
    scrollEl.dataset.fxWheelReplayRun = runId;
    scrollEl.dataset.fxWheelReplayExpectedEvents = String(trace.count);
    scrollEl.dataset.fxWheelReplayExpectedRawPx = String(expectedRawPx);
    try {
        for (let index = 0; index < trace.count; index++) {
            if (!generationIsCurrent()) return;
            scrollEl.dispatchEvent(new WheelEvent("wheel", {
                deltaY: trace.deltaY,
                deltaMode: trace.deltaMode,
                bubbles: true,
                cancelable: true
            }));
            if (index + 1 < trace.count) await wait(trace.intervalMs);
        }
    } finally {
        clearReplayTags(scrollEl, runId);
    }
}

function clearReplayTags(scrollEl, owner = null) {
    if (!scrollEl || (owner && scrollEl.dataset.fxWheelReplayRun !== owner)) return;
    delete scrollEl.dataset.fxWheelReplayKind;
    delete scrollEl.dataset.fxWheelReplayVersion;
    delete scrollEl.dataset.fxWheelReplayRun;
    delete scrollEl.dataset.fxWheelReplayExpectedEvents;
    delete scrollEl.dataset.fxWheelReplayExpectedRawPx;
}

export function attachWheelWindowMetricsPanel(host, gridSelector) {
    detachWheelWindowMetricsPanel(host);
    if (!host || !gridSelector) return;

    const samples = [];
    const cancelledReplayRuns = new Set();
    let activeReplayRun = null;
    const onTelemetry = event => {
        const grid = event.target instanceof Element
            ? event.target.closest(".fx-grid")
            : null;
        if (!grid || !grid.matches(gridSelector)) return;
        const detail = event.detail || {};
        if (detail.replayRun && cancelledReplayRuns.has(detail.replayRun)) {
            cancelledReplayRuns.delete(detail.replayRun);
            return;
        }
        if (detail.replayRun && detail.replayRun === activeReplayRun)
            activeReplayRun = null;
        samples.push(detail);
        if (samples.length > 200) samples.shift();
        render(host, samples);
    };
    let replayGeneration = 0;
    const clearActiveReplay = () => {
        replayGeneration++;
        if (activeReplayRun) {
            cancelledReplayRuns.add(activeReplayRun);
            if (cancelledReplayRuns.size > 100)
                cancelledReplayRuns.delete(cancelledReplayRuns.values().next().value);
            activeReplayRun = null;
        }
        clearReplayTags(document.querySelector(gridSelector)
            ?.querySelector(".fx-grid-content"));
    };
    const onClick = event => {
        if (!(event.target instanceof Element)) return;
        const action = event.target.closest("[data-fx-action]")?.dataset.fxAction;
        if (action === "clear") {
            clearActiveReplay();
            samples.length = 0;
            render(host, samples);
            return;
        }
        if (action === "replay-mouse" || action === "replay-trackpad") {
            clearActiveReplay();
            const generation = ++replayGeneration;
            const runId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
            activeReplayRun = runId;
            void replayWheelTrace(
                host,
                gridSelector,
                action === "replay-mouse" ? "mouse" : "trackpad",
                runId,
                () => generation === replayGeneration);
        }
    };

    document.addEventListener("fx-grid-scroll-boundary-telemetry", onTelemetry, true);
    host.addEventListener("click", onClick);
    registrations.set(host, {
        onTelemetry,
        onClick,
        cancelReplay: clearActiveReplay
    });
    render(host, samples);
}

export function detachWheelWindowMetricsPanel(host) {
    const registration = host ? registrations.get(host) : null;
    if (!registration) return;
    document.removeEventListener(
        "fx-grid-scroll-boundary-telemetry",
        registration.onTelemetry,
        true);
    host.removeEventListener("click", registration.onClick);
    registration.cancelReplay?.();
    registrations.delete(host);
}
