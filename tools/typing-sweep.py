#!/usr/bin/env python3
"""
typing-sweep.py — HomeFront client typing / editing / keyboard-navigation sweep.

Drives the running HomeFront (PB) app with Playwright and, screen by screen,
hammers every grid with the batteries that keep regressing:

  • FAST TYPING   burst-types into the first text-editable cell (15 ms/key by
                  default — faster than any human) and verifies no character is
                  eaten, then backspaces mid-token and retypes (the
                  "sticky first character" class).
  • COMMIT        Enter must commit and CLOSE the editor without re-opening it
                  (the ServerBacked Enter-bubble class), and a click on another
                  cell must commit what the DOM showed (the "old value came
                  back" class).
  • TAB / ARROWS  after a commit the active cell must move with ArrowRight /
                  ArrowLeft / ArrowDown / Tab, with focus alive in the grid
                  (never stranded on <body> — the focus-trap class).
  • DROPDOWNS     Enter-open → Enter-pick must close the panel, keep focus in
                  the grid, and the NEXT Enter/arrows must still work (the
                  stale one-shot class).
  • TEXTAREAS     standalone .fx-textarea fields get the fast-typing check.

Latency: --rtt N spawns tools/latency-proxy.py on a side port so every screen
is exercised the way a WAN user feels it. Run both ways:

    python3 tools/typing-sweep.py --base http://localhost:5266
    python3 tools/typing-sweep.py --base http://localhost:5266 --rtt 150

Nothing is ever SAVED: the sweep never clicks Save/OK, and it answers every
"data has changed" prompt with No, so batch edits die with the session.

Output: a PASS/FAIL table on stdout plus typing-sweep-report.{json,md} in the
working directory. Exit code 1 if anything FAILed.
"""
import argparse, json, os, signal, subprocess, sys, time
from playwright.sync_api import sync_playwright

DEFAULT_ROUTES = [
    "/project-managers",
    "/attribute-lists",
    "/communities",
    "/customer",
    "/dimension-categories",
    "/po-formats",
    "/job-setup/CR0001",
    "/job-correspondence",
    "/default-vendor",
    "/contacts",
]

TOKEN = "743216"          # digits survive Number columns' commit formatting
CLICKAWAY_TOKEN = "31415"

STATE = """() => {
    const g = window.__sweepGrid;
    if (!g || !g.isConnected) return {gone: true};
    const active = g.querySelector('.fx-cell-active');
    const rows = [...g.querySelectorAll('tbody tr.fx-row')];
    const ed = g.querySelector('tbody input:not([type=checkbox]), tbody textarea');
    const panel = document.querySelector('.fx-dropdown-panel, [class*=dropdown-panel]');
    const a = document.activeElement;
    return {activeField: active ? active.getAttribute('data-field') : null,
            activeRow: active ? rows.indexOf(active.closest('tr')) : -1,
            editorOpen: !!ed, editorVal: ed ? ed.value : null,
            editorBuffered: ed ? ed.dataset.fxClientBufferedEditor === '1' : null,
            panelOpen: !!panel && panel.offsetParent !== null,
            aeBody: a === document.body,
            aeInGrid: !!(a && a.closest && a.closest('.fx-grid') === g)};
}"""


class Sweep:
    def __init__(self, pg, keydelay):
        self.pg = pg
        self.kd = keydelay
        self.results = []

    def rec(self, route, subject, check, status, detail=""):
        self.results.append(dict(route=route, subject=subject, check=check,
                                 status=status, detail=str(detail)[:220]))
        print(f"  [{status:4}] {subject:28} {check:26} {detail}"[:150], flush=True)

    # ── plumbing ────────────────────────────────────────────────────────────
    def dismiss_msgbox(self):
        for label in ("No", "Cancel"):
            btn = self.pg.query_selector(f".fx-msgbox-overlay button:has-text('{label}')")
            if btn:
                btn.click(); time.sleep(1.0); return True
        return False

    def handle_chooser(self):
        # Some screens open a picker dialog on load (jobs list etc.) — pick
        # CR0001 when present, else the first row, then Open/OK.
        dlg = self.pg.query_selector(".fx-dialog")
        if not dlg or not dlg.is_visible():
            return
        row = (self.pg.query_selector(".fx-dialog tbody tr:has-text('CR0001')")
               or self.pg.query_selector(".fx-dialog tbody tr.fx-row"))
        if row:
            row.click(); time.sleep(0.5)
        for label in ("Open", "OK"):
            btn = self.pg.query_selector(f".fx-dialog button:has-text('{label}')")
            if btn:
                btn.click(); time.sleep(2.5); return
        self.pg.keyboard.press("Escape"); time.sleep(1.0)

    def state(self):
        return self.pg.evaluate(STATE)

    def digits(self, s):
        return "".join(c for c in (s or "") if c.isdigit())

    # ── per-grid batteries ──────────────────────────────────────────────────
    def grid_cells(self, gi):
        return self.pg.evaluate("""(gi) => {
            const grids = [...document.querySelectorAll('.fx-grid')]
                .filter(g => g.offsetParent !== null && g.querySelector('tbody tr.fx-row'));
            const g = grids[gi];
            if (!g) return null;
            window.__sweepGrid = g;
            const row = g.querySelector('tbody tr.fx-row');
            return [...row.querySelectorAll('td[data-field]')].map(td => ({
                field: td.getAttribute('data-field'),
                isCheckbox: !!td.querySelector('input[type=checkbox]'),
                isPassword: !!td.querySelector('.fx-cell-password-toggle'),
                hasButton: !!td.querySelector('.fx-cell-action-btn')}));
        }""", gi)

    def cell(self, row, field):
        return self.pg.query_selector_all(
            f".fx-grid tbody tr.fx-row >> nth={row}") and self.pg.evaluate_handle(
            """([r, f]) => {
                const g = window.__sweepGrid;
                const rows = [...g.querySelectorAll('tbody tr.fx-row')];
                return rows[r] ? rows[r].querySelector(`td[data-field="${f}"]`) : null;
            }""", [row, field]).as_element()

    def open_editor(self, field, row=0):
        c = self.cell(row, field)
        if c is None:
            return None
        c.dblclick(); time.sleep(1.4)
        st = self.state()
        if st.get("panelOpen"):
            return "dropdown"
        if st.get("editorOpen"):
            return "text"
        return None

    def close_editor(self):
        self.pg.keyboard.press("Escape"); time.sleep(0.8)
        self.dismiss_msgbox()

    def find_columns(self, route, gi):
        cells = self.grid_cells(gi)
        if not cells:
            return None, None
        text_col = dd_col = None
        probed = 0
        for c in cells:
            if c["isCheckbox"] or c["isPassword"] or c["hasButton"]:
                continue
            if probed >= 6 or (text_col and dd_col):
                break
            probed += 1
            kind = self.open_editor(c["field"])
            self.close_editor()
            if kind == "text" and text_col is None:
                text_col = c["field"]
            elif kind == "dropdown" and dd_col is None:
                dd_col = c["field"]
        return text_col, dd_col

    def battery_text(self, route, subject, field):
        # 1) fast typing
        kind = self.open_editor(field)
        if kind != "text":
            self.rec(route, subject, "fast-typing", "SKIP", "editor did not open")
            return
        ed = self.pg.query_selector(".fx-grid tbody input:not([type=checkbox]), .fx-grid tbody textarea")
        try:
            ed.select_text()
        except Exception:
            self.pg.keyboard.press("Meta+A")
        self.pg.keyboard.press("Backspace"); time.sleep(0.2)
        self.pg.keyboard.type(TOKEN, delay=self.kd); time.sleep(0.4)
        st = self.state()
        ok = st.get("editorVal") == TOKEN
        self.rec(route, subject, "fast-typing", "PASS" if ok else "FAIL",
                 f"typed {TOKEN} got {st.get('editorVal')!r} buffered={st.get('editorBuffered')}")

        # 2) mid-edit backspace + retype (sticky-char class)
        for _ in range(3):
            self.pg.keyboard.press("Backspace")
        time.sleep(0.3)
        self.pg.keyboard.type("89", delay=self.kd); time.sleep(0.4)
        expect = TOKEN[:-3] + "89"
        st = self.state()
        ok = st.get("editorVal") == expect
        self.rec(route, subject, "edit-backspace-retype", "PASS" if ok else "FAIL",
                 f"expected {expect} got {st.get('editorVal')!r}")

        # 3) Enter commit: closes, does not reopen, focus alive
        self.pg.keyboard.press("Enter"); time.sleep(1.8)
        st = self.state()
        ok = (not st.get("editorOpen")) and (not st.get("aeBody"))
        self.rec(route, subject, "enter-commit-closes", "PASS" if ok else "FAIL", st)

        # 4) arrows + tab move the active cell
        moves = []
        for key in ("ArrowRight", "ArrowLeft", "ArrowDown", "Tab"):
            before = self.state()
            self.pg.keyboard.press(key); time.sleep(0.9)
            after = self.state()
            moved = (before.get("activeField"), before.get("activeRow")) != \
                    (after.get("activeField"), after.get("activeRow"))
            trapped = after.get("aeBody")
            moves.append((key, moved, trapped))
        bad = [m for m in moves if not m[1] or m[2]]
        self.rec(route, subject, "arrow-tab-navigation", "PASS" if not bad else "FAIL",
                 "; ".join(f"{k} moved={m} bodyTrap={t}" for k, m, t in moves))

        # 5) click-away commit keeps the typed value
        kind = self.open_editor(field)
        if kind == "text":
            ed = self.pg.query_selector(".fx-grid tbody input:not([type=checkbox]), .fx-grid tbody textarea")
            try:
                ed.select_text()
            except Exception:
                self.pg.keyboard.press("Meta+A")
            self.pg.keyboard.press("Backspace"); time.sleep(0.2)
            self.pg.keyboard.type(CLICKAWAY_TOKEN, delay=self.kd); time.sleep(0.4)
            rows = self.pg.evaluate("() => window.__sweepGrid.querySelectorAll('tbody tr.fx-row').length")
            other = self.cell(1 if rows > 1 else 0, field)
            if other and rows > 1:
                other.click(); time.sleep(1.8)
                shown = self.pg.evaluate("""(f) => {
                    const g = window.__sweepGrid;
                    const td = g.querySelector('tbody tr.fx-row td[data-field="' + f + '"]');
                    return td ? td.textContent : ''; }""", field)
                ok = CLICKAWAY_TOKEN in self.digits(shown)
                st = self.state()
                self.rec(route, subject, "clickaway-commit", "PASS" if ok else "FAIL",
                         f"display {shown.strip()!r} trap={st.get('aeBody')}")
            else:
                self.rec(route, subject, "clickaway-commit", "SKIP", "single row")
            self.close_editor()
        else:
            self.rec(route, subject, "clickaway-commit", "SKIP", "editor did not reopen")

    def battery_dropdown(self, route, subject, field):
        c = self.cell(0, field)
        if c is None:
            self.rec(route, subject, "dropdown-pick-nav", "SKIP", "no cell"); return
        c.click(); time.sleep(1.0)
        self.pg.keyboard.press("Enter"); time.sleep(1.4)
        st = self.state()
        if not st.get("panelOpen"):
            # second Enter opens on some flows (closed editor first)
            self.pg.keyboard.press("Enter"); time.sleep(1.4)
            st = self.state()
        if not st.get("panelOpen"):
            self.rec(route, subject, "dropdown-pick-nav", "SKIP", "panel never opened"); return
        self.pg.keyboard.press("Enter"); time.sleep(1.6)   # pick highlighted
        st1 = self.state()
        self.pg.keyboard.press("ArrowRight"); time.sleep(0.9)
        st2 = self.state()
        ok = (not st1.get("panelOpen")) and (not st1.get("aeBody")) \
            and st2.get("activeField") != st1.get("activeField") and not st2.get("aeBody")
        self.rec(route, subject, "dropdown-pick-nav", "PASS" if ok else "FAIL",
                 f"after pick {st1} then arrow {st2.get('activeField')}")

    def battery_textareas(self, route):
        tas = self.pg.query_selector_all("textarea.fx-textarea, .fx-textarea")
        for i, ta in enumerate(tas[:2]):
            if not ta.is_visible():
                continue
            try:
                ta.click(); time.sleep(0.5)
                ta.select_text(); self.pg.keyboard.press("Backspace"); time.sleep(0.2)
                msg = "fast textarea typing 12345 all chars intact"
                self.pg.keyboard.type(msg, delay=self.kd); time.sleep(0.4)
                val = ta.evaluate("el => el.value")
                self.rec(route, f"textarea#{i}", "fast-typing",
                         "PASS" if val == msg else "FAIL", f"got {val!r}")
            except Exception as ex:
                self.rec(route, f"textarea#{i}", "fast-typing", "SKIP", ex)

    # ── screen driver ───────────────────────────────────────────────────────
    def sweep_route(self, base, route):
        print(f"\n=== {route} ===", flush=True)
        try:
            self.pg.goto(base + route, wait_until="domcontentloaded")
            time.sleep(3.0)
            self.dismiss_msgbox()
            self.handle_chooser()
            self.pg.wait_for_selector(".fx-grid tbody tr.fx-row, .fx-textarea", timeout=20000)
            time.sleep(1.5)
        except Exception as ex:
            self.rec(route, "page", "load", "SKIP", ex)
            return
        grid_count = self.pg.evaluate(
            "() => [...document.querySelectorAll('.fx-grid')]"
            ".filter(g => g.offsetParent !== null && g.querySelector('tbody tr.fx-row')).length")
        for gi in range(min(grid_count, 2)):
            subject = f"grid#{gi}"
            try:
                text_col, dd_col = self.find_columns(route, gi)
                if text_col:
                    self.battery_text(route, f"{subject}:{text_col}", text_col)
                else:
                    self.rec(route, subject, "text-battery", "SKIP", "no text-editable column found")
                if dd_col:
                    self.battery_dropdown(route, f"{subject}:{dd_col}", dd_col)
            except Exception as ex:
                self.rec(route, subject, "battery", "SKIP", f"error: {ex}")
        try:
            self.battery_textareas(route)
        except Exception as ex:
            self.rec(route, "textareas", "fast-typing", "SKIP", ex)


def login(pg, base):
    # A cold server can drop the first circuit mid-handshake; reload and redo
    # the whole step until the User ID screen actually appears.
    for attempt in range(4):
        try:
            pg.goto(base, wait_until="domcontentloaded")
            pg.wait_for_selector("input[type=checkbox]", timeout=90000); time.sleep(3.0)
            bx = pg.query_selector_all("input[type=checkbox]")
            for i in (1, 2, 3):
                if i < len(bx) and bx[i].is_checked():
                    bx[i].click(); time.sleep(0.5)
            pg.get_by_role("button", name="OK").click()
            pg.wait_for_selector("input[placeholder='User ID']", timeout=20000)
            break
        except Exception:
            if attempt == 3:
                raise
            time.sleep(2.0)
    pg.fill("input[placeholder='User ID']", "ADMIN")
    pg.fill("input[placeholder='Password']", "x")
    pg.get_by_role("button", name="OK").click()
    pg.wait_for_selector("input[type=radio]", timeout=60000); time.sleep(1.5)
    pg.query_selector_all("input[type=radio]")[0].click(); time.sleep(0.6)
    pg.get_by_role("button", name="OK").click(); time.sleep(4.0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default="http://localhost:5266")
    ap.add_argument("--rtt", type=int, default=0, help="simulated round-trip ms (spawns latency-proxy.py)")
    ap.add_argument("--proxy-port", type=int, default=5378)
    ap.add_argument("--keydelay", type=int, default=15, help="ms between keystrokes")
    ap.add_argument("--routes", default=",".join(DEFAULT_ROUTES))
    ap.add_argument("--headed", action="store_true")
    ap.add_argument("--report", default="typing-sweep-report")
    args = ap.parse_args()

    base = args.base
    proxy = None
    if args.rtt > 0:
        target_port = base.rsplit(":", 1)[1]
        proxy = subprocess.Popen(
            [sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), "latency-proxy.py"),
             "--listen", str(args.proxy_port), "--target", target_port, "--rtt", str(args.rtt)],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1.0)
        base = f"http://localhost:{args.proxy_port}"
        print(f"latency proxy: {args.rtt}ms RTT on :{args.proxy_port} -> :{target_port}")

    try:
        with sync_playwright() as p:
            b = p.chromium.launch(headless=not args.headed)
            pg = b.new_page(viewport={"width": 1920, "height": 835})
            pg.on("dialog", lambda d: d.accept() if d.type == "beforeunload" else d.dismiss())
            login(pg, base)
            sweep = Sweep(pg, args.keydelay)
            for route in [r.strip() for r in args.routes.split(",") if r.strip()]:
                sweep.sweep_route(base, route)
            b.close()
    finally:
        if proxy:
            proxy.send_signal(signal.SIGTERM)

    fails = [r for r in sweep.results if r["status"] == "FAIL"]
    passes = [r for r in sweep.results if r["status"] == "PASS"]
    with open(args.report + ".json", "w") as f:
        json.dump(dict(rtt=args.rtt, keydelay=args.keydelay, results=sweep.results), f, indent=1)
    with open(args.report + ".md", "w") as f:
        f.write(f"# HomeFront typing sweep — rtt={args.rtt}ms keydelay={args.keydelay}ms\n\n")
        f.write(f"**{len(passes)} PASS / {len(fails)} FAIL / "
                f"{len(sweep.results) - len(passes) - len(fails)} SKIP**\n\n")
        f.write("| Route | Subject | Check | Status | Detail |\n|---|---|---|---|---|\n")
        for r in sweep.results:
            f.write(f"| {r['route']} | {r['subject']} | {r['check']} | {r['status']} | {r['detail']} |\n")
    print(f"\nSweep done: {len(passes)} PASS / {len(fails)} FAIL "
          f"/ {len(sweep.results) - len(passes) - len(fails)} SKIP  -> {args.report}.md")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
