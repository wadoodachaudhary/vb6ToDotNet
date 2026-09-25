import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');
const base = process.env.BENCH_URL || 'http://127.0.0.1:5299';
const output = process.env.RESULTS_DIR || '/tmp/flexcore-pointer-selection';
const baseline = process.env.BASELINE_MODULE;
await mkdir(output, { recursive: true });
const browser = await chromium.launch({ channel: 'chrome', headless: true });
const results = [];
let checks = 0;

async function openBench(path, width = 1400) {
    const page = await browser.newPage({ viewport: { width, height: 1000 } });
    page.setDefaultTimeout(20000);
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    let delay = 0;
    await page.routeWebSocket('**/_blazor*', socket => {
        const server = socket.connectToServer();
        socket.onMessage(message => setTimeout(() => server.send(message), delay));
        server.onMessage(message => setTimeout(() => socket.send(message), delay));
    });
    if (baseline) {
        const original = await readFile(baseline, 'utf8');
        await page.route('**/grid-control.js*', route => route.fulfill({ contentType: 'text/javascript', body: original }));
    }
    await page.goto(base + path);
    await page.waitForTimeout(700);
    return { page, errors, setDelay: value => { delay = value; } };
}

async function ready(grid) {
    await grid.locator('tbody tr.fx-row').first().waitFor();
    if (!baseline) await grid.locator(':scope[data-fx-instant-row-feedback]').waitFor();
    await grid.page().waitForTimeout(250);
}

async function watch(grid) {
    await grid.evaluate(root => {
        const blue = color => {
            const [r, g, b, a = 1] = color.match(/[\d.]+/g)?.map(Number) || [];
            return a > 0 && r > 80 && b > r + 12 && g > r;
        };
        const state = window.pointerPaintCheck = { expected: null, frames: 0, failures: [], running: true };
        root.addEventListener('pointerdown', event => {
            if (event.button === 0 && !event.ctrlKey && !event.metaKey && !event.shiftKey)
                state.expected = event.target.closest('tr.fx-row')?.dataset.ari ?? null;
        }, true);
        const sample = () => {
            if (!state.running) return;
            if (state.expected !== null) {
                const painted = [...root.querySelectorAll('tbody tr.fx-row')].filter(row =>
                    [...row.cells].some(cell => {
                        const bg = getComputedStyle(cell).backgroundColor;
                        return blue(bg) || (bg === 'rgba(0, 0, 0, 0)' && blue(getComputedStyle(row).backgroundColor));
                    })).map(row => row.dataset.ari);
                state.frames++;
                if (painted.length !== 1 || painted[0] !== state.expected)
                    state.failures.push({ expected: state.expected, painted, ack: root.dataset.fxPointerPaintAck });
            }
            requestAnimationFrame(sample);
        };
        sample();
    });
}

async function clickRow(grid, row, options = {}) {
    const cell = grid.locator(`tbody tr.fx-row[data-ari="${row}"] td[data-field]`).first();
    const box = await cell.boundingBox();
    assert.ok(box, `row ${row} is mounted`);
    await grid.page().mouse.click(box.x + Math.min(20, box.width / 2), box.y + box.height / 2, options);
}

async function settled(grid, delay) {
    await grid.page().waitForTimeout(800 + delay * 8);
    if (!baseline) {
        const residue = await grid.locator('[data-fx-muted], .fx-drag-preview').count();
        assert.equal(residue, 0, 'acknowledged selection releases all preview styles');
        checks++;
    }
}

async function selected(grid) {
    return grid.locator('tbody tr.fx-row.fx-selected').evaluateAll(rows => rows.map(row => row.dataset.ari));
}

try {
    for (const mode of (process.env.MODES || 'Single,Multiple').split(',')) {
        for (const delay of (process.env.DELAYS || '0,150,400,1000').split(',').map(Number)) {
            const { page, errors, setDelay } = await openBench('/picklist-churn');
            try {
                await page.locator('select').nth(0).selectOption('50');
                await page.locator('select').nth(1).selectOption('First');
                await page.locator('.pc-check input').uncheck();
                await page.getByRole('radio', { name: mode, exact: true }).click();
                await page.locator('#pc-open-optimized').click();
                const grid = page.locator('.pc-grid-frame .fx-grid');
                await ready(grid);
                await watch(grid);
                setDelay(delay);
                // Revisit the original row while earlier requests are still in flight.
                for (const row of [3, 6, 0, 9, 3, 0, 8, 2, 8, 4, 0, 2]) {
                    await clickRow(grid, row);
                    await page.waitForTimeout(45);
                }
                await settled(grid, delay);
                const sample = await page.evaluate(() => {
                    window.pointerPaintCheck.running = false;
                    return window.pointerPaintCheck;
                });
                results.push({ bench: 'picklist', mode, delayEachWay: delay, ...sample });
                await writeFile(join(output, `picklist-${mode}-${delay}.json`), JSON.stringify(sample, null, 2));
                assert.deepEqual(sample.failures, [], `${mode} ${delay}ms: newest row is the only painted row in every frame`);
                assert.deepEqual(await selected(grid), ['2']);
                assert.ok(sample.frames > 20);
                checks += 3;

                // A held press must survive the former 1.5-second cleanup timer.
                const cell = grid.locator('tr.fx-row[data-ari="4"] td[data-field]').first();
                const box = await cell.boundingBox();
                await page.mouse.move(box.x + 15, box.y + box.height / 2);
                await page.mouse.down();
                await page.waitForTimeout(1750);
                assert.equal(await cell.evaluate(el => getComputedStyle(el).backgroundColor), 'rgb(182, 200, 221)');
                await page.mouse.up();
                await settled(grid, delay);
                assert.deepEqual(await selected(grid), ['4']);
                checks += 2;

                // Releasing outside the grid or losing the window cancels a press.
                await page.mouse.move(box.x + 15, box.y + box.height / 2);
                await page.mouse.down();
                await page.mouse.move(5, 5);
                await page.mouse.up();
                await settled(grid, delay);
                assert.equal(await grid.locator('[data-fx-muted], .fx-drag-preview').count(), 0);
                await page.mouse.move(box.x + 15, box.y + box.height / 2);
                await page.mouse.down();
                await page.evaluate(() => window.dispatchEvent(new Event('blur')));
                await page.mouse.move(5, 5);
                await page.mouse.up();
                await settled(grid, delay);
                assert.equal(await grid.locator('[data-fx-muted], .fx-drag-preview').count(), 0);
                checks += 2;

                if (mode === 'Multiple') {
                    await page.keyboard.down(process.platform === 'darwin' ? 'Meta' : 'Control');
                    await clickRow(grid, 6);
                    await page.keyboard.up(process.platform === 'darwin' ? 'Meta' : 'Control');
                    await settled(grid, delay);
                    assert.deepEqual(await selected(grid), ['4', '6']);
                    await page.keyboard.down('Shift');
                    await clickRow(grid, 8);
                    await page.keyboard.up('Shift');
                    await settled(grid, delay);
                    assert.deepEqual(await selected(grid), ['6', '7', '8']);
                    await clickRow(grid, 2);
                    await settled(grid, delay);
                    assert.deepEqual(await selected(grid), ['2']);
                    checks += 3;
                }

                await page.keyboard.press('ArrowDown');
                await page.keyboard.press('ArrowDown');
                await settled(grid, delay);
                assert.deepEqual(await selected(grid), [mode === 'Multiple' ? '4' : '6']);
                assert.deepEqual(errors, []);
                checks += 2;
                setDelay(0);
                await page.screenshot({ path: join(output, `picklist-${mode}-${delay}.png`) });
                await page.locator('#pc-close').click();
                await page.locator('#pc-open-optimized').click();
                await ready(page.locator('.pc-grid-frame .fx-grid'));
                checks++;
                console.log(`PASS picklist ${mode}, ${delay}ms each way, ${sample.frames} painted frames`);
            } finally { await page.close(); }
        }
    }

    if (!baseline && process.env.EXTRA !== '0') {
        const { page, errors, setDelay } = await openBench('/row-selection');
        try {
            await page.locator('#btn-rows-1000').click();
            if ((await page.locator('#btn-group-toggle').innerText()).includes('grouped by'))
                await page.locator('#btn-group-toggle').click();
            const grid = page.locator('.fx-grid').first();
            await ready(grid);
            await grid.scrollIntoViewIfNeeded();
            await watch(grid);
            setDelay(250);
            for (const row of [1, 5, 2, 7, 1, 0, 4, 2]) {
                await clickRow(grid, row);
                await page.waitForTimeout(60);
            }
            await settled(grid, 250);
            const sample = await page.evaluate(() => { window.pointerPaintCheck.running = false; return window.pointerPaintCheck; });
            results.push({ bench: 'row-selection', delayEachWay: 250, ...sample });
            assert.deepEqual(sample.failures, []);
            assert.deepEqual(await selected(grid), ['2']);
            checks += 2;

            const start = await grid.locator('tr.fx-row[data-ari="3"] td[data-field]').first().boundingBox();
            const end = await grid.locator('tr.fx-row[data-ari="7"] td[data-field]').first().boundingBox();
            await page.mouse.move(start.x + 15, start.y + start.height / 2);
            await page.mouse.down();
            await page.waitForTimeout(1200);
            await page.mouse.move(end.x + 15, end.y + end.height / 2, { steps: 20 });
            await page.waitForTimeout(120);
            await page.mouse.up();
            await settled(grid, 250);
            assert.deepEqual(await selected(grid), ['3', '4', '5', '6', '7']);
            checks++;

            setDelay(0);
            await page.getByRole('radio', { name: '250 ms', exact: true }).click();
            await watch(grid);
            for (const row of [2, 5, 8, 3]) {
                await clickRow(grid, row);
                await page.waitForTimeout(80);
            }
            await page.waitForTimeout(16000);
            const uplinkSample = await page.evaluate(() => { window.pointerPaintCheck.running = false; return window.pointerPaintCheck; });
            results.push({ bench: 'row-selection', simulatedUplink: 250, ...uplinkSample });
            assert.deepEqual(uplinkSample.failures, []);
            assert.deepEqual(await selected(grid), ['3']);
            assert.deepEqual(errors, []);
            checks += 3;
            await page.getByRole('radio', { name: 'Off', exact: true }).click();
            await page.screenshot({ path: join(output, 'row-selection.png') });
            console.log('PASS row-selection: delayed clicks, drag selection and built-in uplink delay');
        } finally { await page.close(); }

        const veto = await openBench('/fdbgrid-bench');
        try {
            const grid = veto.page.locator('.fdb-grid-host .fx-grid');
            await ready(grid);
            await grid.locator('tr.fx-row[data-ari="0"] td[data-field="Description"]').dblclick();
            const editor = grid.locator('td.fx-batch-editing input').first();
            await editor.fill('Dirty row stays selected');
            await editor.press('Enter');
            await veto.page.locator('#fdb-save:enabled').waitFor();
            veto.setDelay(250);
            await clickRow(grid, 3);
            await settled(grid, 250);
            assert.deepEqual(await selected(grid), ['0'], 'server veto wins over the pointer preview');
            assert.equal(await grid.locator('[data-fx-muted], .fx-drag-preview').count(), 0);
            assert.deepEqual(veto.errors, []);
            checks += 3;
            await veto.page.screenshot({ path: join(output, 'veto.png') });
            console.log('PASS FDBGrid dirty-row veto under latency');
        } finally { await veto.page.close(); }
    }
} finally {
    await writeFile(join(output, 'summary.json'), JSON.stringify({ checks, results }, null, 2));
    await browser.close();
}
console.log(`PASS ${checks} assertions`);
