import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { homedir, tmpdir } from 'node:os';

const { chromium } = await import(pathToFileURL(process.env.PLAYWRIGHT_MODULE ?? resolve(homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs')).href);
const base = process.env.BENCH_URL ?? 'http://127.0.0.1:5299';
const library = process.env.LIBRARY_ROOT ?? fileURLToPath(new URL('../../../FlexCore/', import.meta.url));
const output = process.env.OUTPUT_DIR ?? resolve(tmpdir(), 'cell-row-handoff');
await mkdir(output, { recursive: true });
const browser = await chromium.launch({ channel: 'chrome', headless: true });
const results = [];
try {
    for (const delay of [0, 200, 400]) {
        const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
        const errors = [];
        page.on('pageerror', error => errors.push(error.message));
        let lag = 0;
        await page.routeWebSocket('**/_blazor*', socket => {
            const server = socket.connectToServer();
            socket.onMessage(message => setTimeout(() => server.send(message), lag));
            server.onMessage(message => setTimeout(() => socket.send(message), lag));
        });
        await page.route('**/grid-control.js*', async route => route.fulfill({ contentType: 'text/javascript', body: await readFile(process.env.BASELINE_MODULE ?? resolve(library, 'wwwroot/grid-control.js'), 'utf8') }));
        await page.goto(`${base}/r2qa-grid?Count=50&AllowSelect=true&HighlightRows=true`);
        const grid = page.locator('#qa-grid .fx-grid');
        const cell = row => grid.locator(`tr.fx-row[data-ari="${row}"] td[data-field="${process.env.CLICK_FIELD || 'OptionID'}"]`);
        await cell(0).waitFor();
        await page.waitForTimeout(1000);
        // Distinguish a selected row from the default near-white hover shade.
        if (!process.env.WORKSHEET)
            await page.addStyleTag({ content: '#qa-grid .fx-grid { --fx-grid-cell-selected-row-bg: #d0d0d0; }' });
        if (process.env.WORKSHEET) {
            await page.locator('#qa-grid').evaluate(host => { host.classList.add('worksheet-paint'); host.dataset.paintHost = ''; });
            // Match the worksheet's scoped selector specificity, not an ID rule
            // that accidentally outranks the committed selected-row styling.
            await page.addStyleTag({ content: `
                .worksheet-paint[data-paint-host] .fx-grid-table td[data-field="Community"] { background-color: #d5ecbc !important; }
                .worksheet-paint[data-paint-host] .fx-grid-table td[data-field="Description"] { background-color: #ddddff !important; }
            ` });
            assert.equal(await grid.locator('tr[data-ari="2"] td[data-field="Community"]').evaluate(td => getComputedStyle(td).backgroundColor), 'rgb(213, 236, 188)', 'fixture green column is painted before selection');
            assert.equal(await grid.locator('tr[data-ari="2"] td[data-field="Description"]').evaluate(td => getComputedStyle(td).backgroundColor), 'rgb(221, 221, 255)', 'fixture purple column is painted before selection');
        }
        await cell(0).click();
        await page.waitForTimeout(900);
        lag = delay;
        await grid.evaluate((root, worksheet) => {
            const state = window.handoffFrames = { running: true, started: false, expected: 0, frames: 0, badFrames: 0, failures: [] };
            const probe = document.createElement('span');
            probe.style.backgroundColor = getComputedStyle(root).getPropertyValue('--fx-grid-cell-selected-row-bg');
            root.append(probe);
            const selected = getComputedStyle(probe).backgroundColor;
            probe.remove();
            if (worksheet && selected !== 'rgb(232, 232, 232)') throw new Error(`Expected light-grey row selection, got ${selected}`);
            window.addEventListener('pointerdown', event => {
                const row = event.target.closest?.('tr.fx-row');
                if (event.button !== 0 || event.ctrlKey || event.metaKey || event.shiftKey || !root.contains(row)) return;
                state.expected = Number(row.dataset.ari);
                state.started = true;
            }, true);
            window.addEventListener('keydown', event => {
                if (!root.contains(event.target) || event.ctrlKey || event.metaKey || event.shiftKey) return;
                if (event.key === 'ArrowDown') state.expected++;
                if (event.key === 'ArrowUp') state.expected--;
            }, true);
            const sample = () => {
                if (!state.running) return;
                if (state.started) {
                    state.frames++;
                    const highlighted = [];
                    const backgrounds = [];
                    for (const row of root.querySelectorAll('tr.fx-row')) {
                        if (worksheet) for (const td of row.querySelectorAll('td[data-field="Community"], td[data-field="Description"]')) {
                            const expected = Number(row.dataset.ari) === state.expected ? selected
                                : td.dataset.field === 'Community' ? 'rgb(213, 236, 188)' : 'rgb(221, 221, 255)';
                            const actual = getComputedStyle(td).backgroundColor;
                            if (actual !== expected) backgrounds.push({ row: Number(row.dataset.ari), field: td.dataset.field, actual, expected });
                        }
                        const matches = [...row.querySelectorAll('td[data-field]')].filter(td => {
                            const bg = getComputedStyle(td).backgroundColor;
                            return bg === selected || (bg === 'rgba(0, 0, 0, 0)' && getComputedStyle(row).backgroundColor === selected);
                        });
                        if (matches.length > 1) highlighted.push(Number(row.dataset.ari));
                    }
                    if (highlighted.length !== 1 || highlighted[0] !== state.expected || backgrounds.length) {
                        state.badFrames++;
                        if (state.failures.length < 30) state.failures.push({
                            frame: state.frames, expected: state.expected, highlighted, backgrounds: backgrounds.slice(0, 6),
                            rows: highlighted.map(index => {
                                const row = root.querySelector(`tr.fx-row[data-ari="${index}"]`);
                                return { index, classes: row.className, style: row.style.cssText,
                                    cells: [...row.cells].slice(0, 3).map(td => ({ classes: td.className, style: td.style.cssText })) };
                            })
                        });
                    }
                }
                requestAnimationFrame(sample);
            };
            requestAnimationFrame(sample);
        }, !!process.env.WORKSHEET);
        // Complete a click before its circuit response, then click elsewhere.
        for (const row of [1, 2, 3, 4]) {
            await cell(row).click();
            await page.waitForTimeout(100);
        }
        await page.waitForTimeout(1500 + delay * 4);
        for (const key of ['ArrowDown', 'ArrowDown', 'ArrowUp']) {
            await page.keyboard.press(key);
            await page.waitForTimeout(180);
        }
        await page.waitForTimeout(1000 + delay * 4);
        // Exercise pointer takeover while keyboard synchronization is pending.
        await page.keyboard.press('ArrowDown');
        await page.waitForTimeout(60);
        await cell(1).click();
        await page.waitForTimeout(1500 + delay * 4);
        const report = await page.evaluate(() => { window.handoffFrames.running = false; return window.handoffFrames; });
        const active = await grid.locator('td.fx-cell-active').evaluate(td => Number(td.closest('tr').dataset.ari));
        results.push({ delay, ...report, active, errors });
        await page.screenshot({ path: resolve(output, `handoff-${delay}.png`) });
        await page.close();
        await writeFile(resolve(output, 'results.json'), JSON.stringify(results, null, 2));
        console.log(`${delay * 2}ms RTT`, report.frames, 'frames', report.badFrames, 'failures');
        if (!process.env.BASELINE_MODULE) {
            assert.equal(report.badFrames, 0, JSON.stringify(report.failures.slice(0, 5)));
            assert.equal(active, report.expected);
            assert.deepEqual(errors, []);
        }
    }
} finally { await browser.close(); }
