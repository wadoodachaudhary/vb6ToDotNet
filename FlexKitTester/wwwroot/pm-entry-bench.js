// /pm-entry bench metrics. Everything is measured in the BROWSER so the
// numbers include the full network + server + render path the user feels.
//
//  keydowns      typing keys pressed inside a cell editor
//  serverInputs  input events that reached DOCUMENT bubble — ClientBuffered
//                editors stopPropagation() their input events, so this counts
//                ONLY the per-keystroke dispatches of ServerBacked editors.
//                ClientBuffered typing should hold this at ~0 while typing.
//  echo p95      keydown → the editor's own input event (local paint echo)
//  commit        Enter keydown → first grid DOM mutation (the commit render
//                landing), i.e. the full round trip the user waits for
//  RTT           a JS→server→JS ping every 2s (reflects the simulated uplink
//                delay plus the real link)

let state = null;

function fmt(v) { return v == null ? "—" : Math.round(v); }

export function start(rootSelector, panelSelector, dotNetRef) {
    stop();
    // The grid remounts when the transport toggles — resolve it per event so
    // the metrics survive the swap. rootSelector should be a STABLE wrapper.
    const root = document.querySelector(rootSelector);
    const panel = document.querySelector(panelSelector);
    if (!root || !panel) return false;
    const gridEl = () => root.querySelector(".fx-grid") || root;

    const s = {
        keydowns: 0, serverInputs: 0, echoes: [], commits: [],
        keyAt: 0, enterAt: 0, rtt: null,
    };
    const set = (m, v) => {
        const el = panel.querySelector(`[data-m="${m}"]`);
        if (el) el.textContent = String(v);
    };
    const upd = () => {
        set("keys", s.keydowns);
        set("srv", s.serverInputs);
        const e = [...s.echoes].sort((a, b) => a - b);
        set("echo", e.length ? fmt(e[Math.floor(e.length * 0.95)]) : "—");
        set("commit", s.commits.length ? fmt(s.commits[s.commits.length - 1]) : "—");
        const c = [...s.commits].sort((a, b) => a - b);
        set("commitAvg", c.length ? fmt(c.reduce((x, y) => x + y, 0) / c.length) : "—");
        set("rtt", s.rtt == null ? "—" : s.rtt);
    };

    // A counted keydown that produced no input (Backspace on an empty field)
    // leaves keyAt armed; pairing it with a much-later input would log a
    // seconds-long echo. Only accept plausible key->paint gaps.
    const takeEcho = () => {
        if (!s.keyAt) return;
        const d = performance.now() - s.keyAt;
        s.keyAt = 0;
        if (d < 2000) s.echoes.push(d);
    };

    const isEditor = t => {
        const g = gridEl();
        return g && t instanceof Element && g.contains(t)
            && t.matches("input.fx-batch-input, textarea.fx-batch-input, .fx-batch-input input");
    };

    const onKeyDown = e => {
        if (!isEditor(e.target)) return;
        if (e.key === "Enter" || e.key === "NumpadEnter") { s.enterAt = performance.now(); return; }
        if (e.key.length === 1 || e.key === "Backspace" || e.key === "Delete") {
            s.keydowns++; s.keyAt = performance.now(); upd();
        }
    };
    // BUBBLE phase on document: ClientBuffered stops propagation at the
    // element, so arrivals here are exactly the server-dispatched keystrokes.
    const onInput = e => {
        if (!isEditor(e.target)) return;
        s.serverInputs++;
        takeEcho();
        upd();
    };
    // Local echo for ClientBuffered comes from the same input event but it
    // never bubbles — capture phase sees both transports.
    const onInputCapture = e => {
        if (!isEditor(e.target)) return;
        takeEcho();
    };

    const mo = new MutationObserver(muts => {
        if (!s.enterAt) return;
        // Only a STRUCTURAL change counts as the commit render landing —
        // attribute flips (grid JS bumps dataset counters on every keydown)
        // fire ~0ms after Enter and would fake an instant commit.
        const landed = muts.some(m => {
            if (m.type === "attributes") return false;
            const t = m.target instanceof Element ? m.target : m.target.parentElement;
            return t && !t.closest("input, textarea");
        });
        if (!landed) return;
        s.commits.push(performance.now() - s.enterAt);
        s.enterAt = 0;
        upd();
    });
    mo.observe(root, { subtree: true, childList: true, characterData: true, attributes: true });

    document.addEventListener("keydown", onKeyDown, true);
    document.addEventListener("input", onInput, false);
    document.addEventListener("input", onInputCapture, true);

    const ping = window.setInterval(async () => {
        try {
            const t0 = performance.now();
            await dotNetRef.invokeMethodAsync("Ping");
            s.rtt = Math.round(performance.now() - t0);
            upd();
        } catch { /* circuit gone */ }
    }, 2000);

    s.cleanup = () => {
        document.removeEventListener("keydown", onKeyDown, true);
        document.removeEventListener("input", onInput, false);
        document.removeEventListener("input", onInputCapture, true);
        mo.disconnect();
        window.clearInterval(ping);
    };
    state = s;
    upd();
    return true;
}

export function reset() {
    if (!state) return;
    state.keydowns = 0; state.serverInputs = 0;
    state.echoes = []; state.commits = [];
    state.keyAt = 0; state.enterAt = 0;
}

export function stop() {
    if (state?.cleanup) state.cleanup();
    state = null;
}
