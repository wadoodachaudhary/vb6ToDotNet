function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function findEditor(selector, requireClientBuffer, timeoutMs) {
    const deadline = performance.now() + timeoutMs;
    while (performance.now() < deadline) {
        const input = document.querySelector(selector);
        if (input && (!requireClientBuffer || input.dataset.fxClientBufferedEditor === "1"))
            return input;
        await delay(10);
    }

    throw new Error(requireClientBuffer
        ? "Client-buffered grid editor did not become ready."
        : "Grid editor did not become ready.");
}

function replaceSelection(input, text) {
    const start = input.selectionStart ?? input.value.length;
    const end = input.selectionEnd ?? start;
    if (typeof input.setRangeText === "function") {
        input.setRangeText(text, start, end, "end");
        return;
    }

    input.value = input.value.slice(0, start) + text + input.value.slice(end);
}

function dispatchInput(input, character) {
    let event;
    try {
        event = new InputEvent("input", {
            bubbles: true,
            composed: true,
            inputType: "insertText",
            data: character
        });
    } catch {
        event = new Event("input", { bubbles: true, composed: true });
    }
    input.dispatchEvent(event);
}

export async function runGridTypingBenchmark(
    selector,
    text,
    requireClientBuffer,
    interKeyDelayMs = 0) {
    const input = await findEditor(selector, requireClientBuffer, 4000);
    input.focus({ preventScroll: true });
    input.select();

    const started = performance.now();
    let keyDownEvents = 0;
    let inputEvents = 0;

    for (const character of text) {
        input.dispatchEvent(new KeyboardEvent("keydown", {
            key: character,
            bubbles: true,
            cancelable: true,
            composed: true
        }));
        keyDownEvents++;

        replaceSelection(input, character);
        dispatchInput(input, character);
        inputEvents++;

        if (interKeyDelayMs > 0)
            await delay(interKeyDelayMs);
    }

    const dispatchCompleted = performance.now();
    input.dispatchEvent(new KeyboardEvent("keydown", {
        key: "Enter",
        code: "Enter",
        bubbles: true,
        cancelable: true,
        composed: true
    }));
    keyDownEvents++;

    return {
        browserDispatchMs: dispatchCompleted - started,
        keyDownEvents,
        inputEvents,
        clientBufferActive: input.dataset.fxClientBufferedEditor === "1"
    };
}
