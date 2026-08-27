// TEMP diagnostic overlay for the Edit Model & Options bench: shows the last
// grid events + selection state so screen recordings carry ground truth.
(function () {
    // Runs on ALL tester pages: Blazor SPA navigation means a load-time
    // pathname gate misses pages reached via the menu.
    const panel = document.createElement("div");
    panel.style.cssText = "position:fixed;bottom:8px;right:8px;z-index:99999;background:rgba(0,0,40,.85);color:#9f9;font:10px/1.35 monospace;padding:6px 8px;border-radius:4px;pointer-events:none;max-width:340px;white-space:pre;";
    document.addEventListener("DOMContentLoaded", () => document.body.appendChild(panel));
    if (document.body) document.body.appendChild(panel);
    const lines = [];
    const t0 = performance.now();
    const log = m => {
        lines.push(String(Math.round(performance.now() - t0)).padStart(6) + " " + m);
        while (lines.length > 14) lines.shift();
        panel.textContent = lines.join("\n");
    };
    const ariOf = el => el?.closest?.("tr[data-ari]")?.getAttribute("data-ari") ?? "-";
    document.addEventListener("mousedown", e => log("mdown ari=" + ariOf(e.target) + " tgt=" + e.target.tagName), true);
    document.addEventListener("dblclick", e => log("DBLCLK ari=" + ariOf(e.target)), true);
    document.addEventListener("keydown", e => log("key " + e.key + (e.repeat ? "(r)" : "")), true);
    document.addEventListener("focusin", e => log("focus " + e.target.tagName + "." + (e.target.className || "").toString().slice(0, 14)), true);
    let last = "";
    setInterval(() => {
        const grids = document.querySelectorAll(".fx-grid");
        const g = grids[grids.length - 1];
        if (!g) return;
        const act = ariOf(g.querySelector("td.fx-cell-active"));
        const tinted = [...g.querySelectorAll("tbody tr.fx-row")].filter(tr => {
            const cs = getComputedStyle(tr).backgroundColor;
            const rowTint = cs !== "rgba(0, 0, 0, 0)" && cs !== "rgb(255, 255, 255)";
            return rowTint; // VISUAL truth only — muted rows are excluded
        }).map(tr => tr.getAttribute("data-ari")).join("+");
        const focus = document.activeElement ? document.activeElement.tagName + (document.activeElement.closest?.(".fx-grid") ? "@grid" : "@page") : "-";
        const state = "STATE act=" + act + " sel=" + (tinted || "none") + " foc=" + focus + " scr=" + Math.round(window.scrollY);
        if (state !== last) { last = state; log(state); }
    }, 150);
})();
