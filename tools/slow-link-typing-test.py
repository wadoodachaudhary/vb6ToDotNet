#!/usr/bin/env python3
"""Fast-delete test against the app through the 150ms latency proxy.

Playwright emits real rawKeyDown/char/keyUp sequences, so Backspace performs a
genuine native edit (which the other automation could not do).
"""
import sys, time
from playwright.sync_api import sync_playwright

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:5267"
results = {}

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    page.goto(BASE, wait_until="domcontentloaded")

    # ── security screen: uncheck the 3 hardening toggles, keep System Secrets ──
    page.wait_for_selector("input[type=checkbox]", timeout=30000)
    time.sleep(1.5)
    boxes = page.query_selector_all("input[type=checkbox]")
    for i in (1, 2, 3):
        if boxes[i].is_checked():
            boxes[i].click()
            time.sleep(0.3)
    page.get_by_role("button", name="OK").click()

    # ── login ──
    page.wait_for_selector("input[placeholder='User ID']", timeout=30000)
    page.fill("input[placeholder='User ID']", "ADMIN")
    page.fill("input[placeholder='Password']", "x")
    page.get_by_role("button", name="OK").click()

    # ── division ──
    page.wait_for_selector("input[type=radio]", timeout=30000)
    time.sleep(1.0)
    page.query_selector_all("input[type=radio]")[0].click()
    time.sleep(0.5)
    page.get_by_role("button", name="OK").click()
    time.sleep(2.0)

    # ── Cost Codes grid ──
    page.goto(BASE + "/db-grid/Cost%20Codes", wait_until="domcontentloaded")
    page.wait_for_selector(".fx-grid tbody tr.fx-row", timeout=40000)
    time.sleep(3.0)
    results["rows"] = page.eval_on_selector_all(".fx-grid tbody tr.fx-row", "els => els.length")

    # open the CostCode editor on row 2
    cell = page.evaluate("""() => {
        const ths = [...document.querySelectorAll('.fx-grid thead th')].map(t=>t.textContent.trim());
        const idx = ths.indexOf('CostCode');
        const row = document.querySelectorAll('.fx-grid tbody tr.fx-row')[1];
        const r = row.cells[idx].getBoundingClientRect();
        return {x: Math.round(r.x + r.width/2), y: Math.round(r.y + r.height/2)};
    }""")
    page.mouse.dblclick(cell["x"], cell["y"])
    page.wait_for_selector(".fx-batch-input", timeout=15000)
    time.sleep(0.8)
    page.evaluate("""() => { const i=document.querySelector('.fx-batch-input');
                             i.value=''; i.dispatchEvent(new Event('input',{bubbles:true})); }""")

    TEXT = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"       # 36 chars
    t0 = time.monotonic()
    page.keyboard.type(TEXT, delay=25)                   # ~40 chars/sec, fast human
    type_ms = (time.monotonic() - t0) * 1000
    time.sleep(0.6)
    typed = page.eval_on_selector(".fx-batch-input", "e => e.value")
    results["typed"] = {"expected": TEXT, "got": typed, "exact": typed == TEXT,
                        "elapsed_ms": round(type_ms)}

    # ── FAST DELETE: 20 rapid Backspaces ──
    DELETES = 20
    t0 = time.monotonic()
    for _ in range(DELETES):
        page.keyboard.press("Backspace", delay=0)        # as fast as the browser accepts
    del_ms = (time.monotonic() - t0) * 1000
    time.sleep(1.0)
    after = page.eval_on_selector(".fx-batch-input", "e => e.value")
    expected_after = TEXT[:-DELETES] if typed == TEXT else None
    results["fast_delete"] = {
        "deletes_sent": DELETES,
        "expected": expected_after,
        "got": after,
        "exact": after == expected_after,
        "chars_removed": len(typed) - len(after),
        "elapsed_ms": round(del_ms),
        "ms_per_delete": round(del_ms / DELETES, 1),
    }

    # ── held-Backspace style burst: clear the rest with no delay at all ──
    t0 = time.monotonic()
    for _ in range(len(after) + 5):
        page.keyboard.press("Backspace", delay=0)
    burst_ms = (time.monotonic() - t0) * 1000
    time.sleep(1.0)
    cleared = page.eval_on_selector(".fx-batch-input", "e => e.value")
    results["burst_clear"] = {"got": repr(cleared), "empty": cleared == "",
                              "elapsed_ms": round(burst_ms)}

    page.keyboard.press("Escape")
    browser.close()

import json
print(json.dumps(results, indent=2))
