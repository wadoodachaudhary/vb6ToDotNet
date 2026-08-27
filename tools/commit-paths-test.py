#!/usr/bin/env python3
"""Every way a grid cell edit can end — run against the 150ms latency proxy.

Typing is handled entirely in the browser, so the ONLY thing that matters is
that each exit path still captures (or correctly discards) the text. A silent
loss here is far worse than a slow keystroke, so every path is asserted.

    python3 commit-paths-test.py http://localhost:5267
"""
import sys, time, json
from playwright.sync_api import sync_playwright

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:5267"
out = {}


def login(pg):
    pg.goto(BASE, wait_until="domcontentloaded")
    pg.wait_for_selector("input[type=checkbox]", timeout=40000); time.sleep(1.5)
    bx = pg.query_selector_all("input[type=checkbox]")
    for i in (1, 2, 3):
        if bx[i].is_checked():
            bx[i].click(); time.sleep(0.3)
    pg.get_by_role("button", name="OK").click()
    pg.wait_for_selector("input[placeholder='User ID']", timeout=40000)
    pg.fill("input[placeholder='User ID']", "ADMIN")
    pg.fill("input[placeholder='Password']", "x")
    pg.get_by_role("button", name="OK").click()
    pg.wait_for_selector("input[type=radio]", timeout=40000); time.sleep(1.0)
    pg.query_selector_all("input[type=radio]")[0].click(); time.sleep(0.4)
    pg.get_by_role("button", name="OK").click(); time.sleep(2.5)


def open_grid(pg):
    pg.goto(BASE + "/db-grid/Cost%20Codes", wait_until="domcontentloaded")
    pg.wait_for_selector(".fx-grid tbody tr.fx-row", timeout=45000); time.sleep(3.0)


# Rows are anchored by their CostCode text, never by index: committing a value
# re-sorts the grid, so indices shift under us between steps.
def cell_xy(pg, code, col="Description"):
    return pg.evaluate("""([code, col]) => {
        const ths=[...document.querySelectorAll('.fx-grid thead th')].map(t=>t.textContent.trim());
        const idx=ths.indexOf(col), key=ths.indexOf('CostCode');
        const row=[...document.querySelectorAll('.fx-grid tbody tr.fx-row')]
                    .find(r=>r.cells[key].textContent.trim()===code);
        if(!row) return null;
        row.cells[idx].scrollIntoView({block:'center', inline:'center'});
        const r=row.cells[idx].getBoundingClientRect();
        return {x:Math.round(r.x+r.width/2), y:Math.round(r.y+r.height/2), idx};
    }""", [code, col])


def cell_text(pg, code, col="Description"):
    return pg.evaluate("""([code, col]) => {
        const ths=[...document.querySelectorAll('.fx-grid thead th')].map(t=>t.textContent.trim());
        const idx=ths.indexOf(col), key=ths.indexOf('CostCode');
        const row=[...document.querySelectorAll('.fx-grid tbody tr.fx-row')]
                    .find(r=>r.cells[key].textContent.trim()===code);
        if(!row) return null;
        // Enter commits AND re-opens the editor, so the cell may hold an <input>
        // whose value — not textContent — is the committed text.
        const inp=row.cells[idx].querySelector('input');
        return inp ? inp.value : row.cells[idx].textContent.trim();
    }""", [code, col])


def edit(pg, code, text, col="Description"):
    """Open the editor on the row whose CostCode == code, clear, type with real keys."""
    original = cell_text(pg, code, col)
    for attempt in range(3):
        c = cell_xy(pg, code, col)
        pg.mouse.click(c["x"], c["y"]); time.sleep(0.5)      # activate
        c = cell_xy(pg, code, col)                            # re-measure (grid may scroll)
        pg.mouse.dblclick(c["x"], c["y"])
        try:
            pg.wait_for_selector(".fx-batch-input", timeout=7000)
            break
        except Exception:
            if attempt == 2: raise
            time.sleep(1.0)
    time.sleep(0.6)
    pg.evaluate("""() => {const i=document.querySelector('.fx-batch-input');
                          i.value=''; i.dispatchEvent(new Event('input',{bubbles:true}));}""")
    pg.keyboard.type(text, delay=15)
    return c, original


with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    pg = b.new_page(viewport={"width": 1500, "height": 1000})
    login(pg); open_grid(pg)

    codes = pg.evaluate("""() => {
        const ths=[...document.querySelectorAll('.fx-grid thead th')].map(t=>t.textContent.trim());
        const key=ths.indexOf('CostCode');
        return [...document.querySelectorAll('.fx-grid tbody tr.fx-row')]
                 .map(r=>r.cells[key].textContent.trim()).filter(t=>t.length>3).slice(0,10);}""")
    out["_anchor_codes"] = codes[:7]

    # 1. TAB commits
    edit(pg, codes[0], "TAB-AAA")
    pg.keyboard.press("Tab"); time.sleep(2.2)
    g = cell_text(pg, codes[0]); out["tab_commits"] = {"got": g, "ok": g == "TAB-AAA"}

    # 2. ENTER commits
    edit(pg, codes[1], "ENT-BBB")
    pg.keyboard.press("Enter"); time.sleep(2.2)
    g = cell_text(pg, codes[1]); out["enter_commits"] = {"got": g, "ok": g == "ENT-BBB"}

    # 3. ESCAPE discards (must keep the ORIGINAL)
    _, original = edit(pg, codes[2], "ESC-SHOULD-VANISH")
    pg.keyboard.press("Escape"); time.sleep(2.2)
    g = cell_text(pg, codes[2])
    out["escape_discards"] = {"original": original, "got": g, "ok": g == original}

    # 4. CLICK-AWAY (blur) commits — relies on the native change event
    edit(pg, codes[3], "BLUR-CCC")
    away = cell_xy(pg, codes[8] if len(codes) > 8 else codes[6], "Description")
    pg.mouse.click(away["x"], away["y"]); time.sleep(2.5)
    g = cell_text(pg, codes[3]); out["clickaway_commits"] = {"got": g, "ok": g == "BLUR-CCC"}

    # 5. TYPE THEN IMMEDIATE TAB (zero settle) — the riskiest race
    edit(pg, codes[4], "RACE-DDD")
    pg.keyboard.press("Tab"); time.sleep(2.5)
    g = cell_text(pg, codes[4]); out["type_then_instant_tab"] = {"got": g, "ok": g == "RACE-DDD"}

    # 6. ARROW-OUT commits
    edit(pg, codes[5], "ARR-EEE")
    pg.keyboard.press("ArrowDown"); time.sleep(2.5)
    g = cell_text(pg, codes[5]); out["arrowdown_commits"] = {"got": g, "ok": g == "ARR-EEE"}

    # 7. BACKSPACE-EDIT then commit
    edit(pg, codes[6], "KEEPME-XXXX")
    for _ in range(4):
        pg.keyboard.press("Backspace", delay=0)
    pg.keyboard.press("Tab"); time.sleep(2.5)
    g = cell_text(pg, codes[6]); out["edited_then_commit"] = {"got": g, "ok": g == "KEEPME-"}

    out["dirty_marker_present"] = pg.evaluate(
        "() => { const b=[...document.querySelectorAll('button')].find(x=>x.textContent.trim()==='Save'); return b ? !b.disabled : null; }")

    b.close()

print(json.dumps(out, indent=2))
failed = [k for k, v in out.items() if isinstance(v, dict) and v.get("ok") is False]
print("\nFAILED:", failed if failed else "none")
