import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { resolve } from 'node:path';
import { homedir, tmpdir } from 'node:os';

const modulePath = process.env.PLAYWRIGHT_MODULE ?? resolve(homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs');
const { chromium } = await import(pathToFileURL(modulePath).href);
const base = process.env.BENCH_URL ?? 'http://127.0.0.1:5299';
const library = process.env.LIBRARY_ROOT ?? fileURLToPath(new URL('../../../FlexCore/', import.meta.url));
const source = process.env.BASELINE_MODULE ?? resolve(library, 'wwwroot/grid-control.js');
const output = process.env.OUTPUT_DIR ?? resolve(tmpdir(), 'keyboard-selection-colors');
await mkdir(output, { recursive: true });
const browser = await chromium.launch({ channel: 'chrome', headless: true });
const results = [];
const steps = process.env.LONG_NAV ? 24 : 8;
const cases = [
    { name: 'cell-default', row: false },
    { name: 'cell-grey', row: false, color: '#d0d0d0' },
    { name: 'cell-green', row: false, color: '#d9ead3' },
    { name: 'cell-blue', row: false, color: '#b6c8dd' },
    { name: 'cell-styled', row: false, styled: true },
    { name: 'cell-inline-styled', row: false, styled: 'inline', highlight: false },
    { name: 'row-grey', row: true, color: '#d0d0d0' },
    { name: 'row-blue', row: true, color: '#b6c8dd' },
    { name: 'cell-no-row-highlight', row: false, highlight: false },
    { name: 'no-selection', row: false, allow: false },
];
try {
    for (const delay of [0, 200, 400]) for (const test of cases) {
        if (process.env.CASE && test.name !== process.env.CASE) continue;
        const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
        const errors = [];
        page.on('pageerror', error => errors.push(error.message));
        let lag = 0;
        await page.routeWebSocket('**/_blazor*', socket => {
            const server = socket.connectToServer();
            socket.onMessage(message => setTimeout(() => server.send(message), lag));
            server.onMessage(message => setTimeout(() => socket.send(message), lag));
        });
        await page.route('**/grid-control.js*', async route => route.fulfill({
            contentType: 'text/javascript', body: await readFile(source, 'utf8'),
        }));
        await page.goto(`${base}/r2qa-grid?Count=50&RowSelection=${test.row}&HighlightRows=${test.highlight !== false}&AllowSelect=${test.allow !== false}`);
        const grid = page.locator('#qa-grid .fx-grid');
        await grid.locator('tbody tr.fx-row').first().waitFor();
        await page.waitForTimeout(1000);
        if (test.color) await page.addStyleTag({ content: `#qa-grid .fx-grid { ${test.row ? '--fx-grid-selected-row-bg' : '--fx-grid-cell-selected-row-bg'}: ${test.color}; }` });
        if (!test.row) await page.addStyleTag({ content: `
            #qa-grid td[data-field="OptionID"], #qa-grid td[data-field="Community"],
            #qa-grid td[data-field="Column4"] { color: #808080; }
            #qa-grid td[data-field="Column5"] { color: transparent; }
        ` });
        if (test.styled && test.styled !== 'inline') await page.addStyleTag({ content: `
            .fx-grid .fx-grid-table td[data-field="OptionID"], .fx-grid .fx-grid-table td[data-field="Community"] { background-color: #b0b0b0 !important; }
            .fx-grid .fx-grid-table td[data-field="Description"] { background-color: #ddddff !important; }
        ` });
        if (test.styled === 'inline') await grid.evaluate(root => {
            for (const td of root.querySelectorAll('td[data-field="OptionID"], td[data-field="Community"], td[data-field="Description"]'))
                td.style.setProperty('background-color', td.dataset.field === 'Description' ? '#ddddff' : '#b0b0b0', 'important');
        });
        const cell = grid.locator('tr.fx-row[data-ari="0"] td[data-field="OptionID"]');
        await cell.click();
        await page.waitForTimeout(2200);
        await grid.evaluate(root => root.focus());
        lag = delay;
        await grid.evaluate((root, test) => {
            const initial = root.querySelector('td.fx-cell-active');
            if (!initial) throw new Error('No active cell');
            const state = window.keyboardColors = {
                row: Number(initial.closest('tr').dataset.ari), cell: initial.cellIndex,
                started: false, running: true, frames: 0, failures: [],
            };
            const style = getComputedStyle(root);
            const colorValue = style.getPropertyValue(test.row ? '--fx-grid-selected-row-bg' : '--fx-grid-cell-selected-row-bg').trim();
            const probe = document.createElement('span');
            probe.style.backgroundColor = colorValue;
            root.append(probe);
            const expectedColor = getComputedStyle(probe).backgroundColor;
            probe.remove();
            state.expectedColor = expectedColor;
            // Observe trusted keys before the grid's capture handler stops them.
            const track = event => {
                if (!root.contains(event.target) || event.shiftKey || event.ctrlKey || event.metaKey || event.altKey) return;
                const key = event.key;
                if (key === 'ArrowRight') state.cell++;
                else if (key === 'ArrowLeft') state.cell--;
                else if (key === 'ArrowDown') state.row++;
                else if (key === 'ArrowUp') state.row--;
                else return;
                state.started = true;
            };
            window.addEventListener('keydown', track, true);
            const sample = () => {
                if (!state.running) { window.removeEventListener('keydown', track, true); return; }
                if (state.started) {
                    state.frames++;
                    const failures = [];
                    const cues = [];
                    for (const row of root.querySelectorAll('tbody tr.fx-row')) {
                        const index = Number(row.dataset.ari);
                        for (const td of row.querySelectorAll('td.fx-cell')) {
                            const css = getComputedStyle(td);
                            if (!test.row) {
                                const expectedText = ['OptionID', 'Community', 'Column4'].includes(td.dataset.field)
                                    ? 'rgb(128, 128, 128)' : td.dataset.field === 'Column5' ? 'rgba(0, 0, 0, 0)' : null;
                                if (expectedText && css.color !== expectedText)
                                    failures.push({ row: index, field: td.dataset.field, text: css.color, expectedText });
                            }
                            if (css.boxShadow !== 'none') cues.push([index, td.cellIndex]);
                            const actual = css.backgroundColor;
                            const nativeColor = test.styled
                                ? ['OptionID', 'Community'].includes(td.dataset.field) ? 'rgb(176, 176, 176)'
                                    : td.dataset.field === 'Description' ? 'rgb(221, 221, 255)' : null
                                : null;
                            const paintRow = test.allow !== false && test.highlight !== false;
                            if (nativeColor && (!paintRow || index !== state.row) && actual !== nativeColor)
                                    failures.push({ row: index, field: td.dataset.field, actual, expectedColor: nativeColor });
                            if (paintRow) {
                                if (index === state.row && actual !== expectedColor)
                                    failures.push({ row: index, field: td.dataset.field, actual, expectedColor });
                                if (index !== state.row && actual === expectedColor)
                                    failures.push({ row: index, field: td.dataset.field, stale: actual });
                            }
                        }
                    }
                    if (cues.length !== 1 || cues[0][0] !== state.row || cues[0][1] !== state.cell)
                        failures.push({ cues, expected: [state.row, state.cell] });
                    if (failures.length && state.failures.length < 20) state.failures.push({ frame: state.frames, failures: failures.slice(0, 5) });
                }
                requestAnimationFrame(sample);
            };
            requestAnimationFrame(sample);
        }, test);
        for (const key of [
            ...Array(steps).fill('ArrowRight'), ...Array(steps / 2).fill('ArrowLeft'),
            'ArrowDown', 'ArrowDown', 'ArrowRight', 'ArrowUp', 'ArrowLeft',
        ]) {
            await page.keyboard.press(key);
            await page.waitForTimeout(75);
        }
        await page.waitForTimeout(1000 + delay * 4);
        const report = await page.evaluate(() => { window.keyboardColors.running = false; return window.keyboardColors; });
        const active = await grid.locator('td.fx-cell-active').evaluate(td => [Number(td.closest('tr').dataset.ari), td.cellIndex]);
        results.push({ name: test.name, delay, ...report, active, errors });
        await page.screenshot({ path: resolve(output, `${test.name}-${delay}.png`) });
        await page.close();
        await writeFile(resolve(output, 'results.json'), JSON.stringify(results, null, 2));
        console.log(test.name, `${delay * 2}ms RTT`, report.frames, 'frames', report.failures.length, 'failures');
        if (!process.env.BASELINE_MODULE) {
            assert.ok(report.frames > 10);
            assert.deepEqual(report.failures, []);
            assert.deepEqual(active, [report.row, report.cell]);
            assert.deepEqual(errors, []);
        }
    }
} finally { await browser.close(); }
