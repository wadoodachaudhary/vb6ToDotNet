import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { mkdir } from 'node:fs/promises';

const require = createRequire(import.meta.url);
const { chromium, expect } = require(process.env.PLAYWRIGHT_MODULE || '/Users/wadood/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/test');
const base = process.env.BENCH_URL || 'http://127.0.0.1:5301';
assert(['127.0.0.1', 'localhost'].includes(new URL(base).hostname));
const output = process.env.RESULTS_DIR || '/tmp/choose-columns-popup';
await mkdir(output, { recursive: true });
const browser = await chromium.launch({ channel: 'chrome', headless: true });
let checks = 0;

try {
    for (const [width, height, delay] of [[1400, 900, 0], [1400, 900, 150], [390, 740, 0], [1024, 600, 0]]) {
        const page = await browser.newPage({ viewport: { width, height } });
        page.setDefaultTimeout(20000);
        const errors = [];
        page.on('pageerror', error => errors.push(error.message));
        let latency = 0;
        await page.routeWebSocket('**/_blazor*', socket => {
            const server = socket.connectToServer();
            socket.onMessage(message => setTimeout(() => server.send(message), latency));
            server.onMessage(message => setTimeout(() => socket.send(message), latency));
        });
        await page.goto(base + '/choose-columns-popup');
        const kit = await page.locator('#popup-kit').innerText();
        const grid = page.locator('#popup-grid-host .fx-grid');
        await expect(grid).toHaveAttribute('data-fx-instant-row-feedback', 'true');
        const popup = page.locator('.fx-choose-columns-dialog');
        const list = popup.locator('.fx-choose-columns-list');
        const header = popup.locator('.fx-dialog-header');
        const count = page.locator('#popup-context-count');
        const row = label => list.getByRole('option').filter({ has: page.getByText(label, { exact: true }) });
        const button = label => popup.getByRole('button', { name: label, exact: true });

        async function open() {
            await expect(grid).toHaveAttribute('data-fx-instant-row-feedback', 'true');
            await grid.locator('thead th[data-field]').first().click({ button: 'right' });
            await page.getByRole('button', { name: 'Insert a column...', exact: true }).click();
            await expect(popup).toBeVisible();
            await expect(list).toBeFocused();
            checks += 2;
        }
        async function closed(commits) {
            await expect(popup).toHaveCount(0);
            await expect(page.locator('#popup-commit-count')).toHaveText(String(commits));
            checks += 2;
        }
        async function contextIsolated(expected) {
            for (const target of [header, row('Code'), popup.locator('.fx-choose-columns-instructions'), button('Cancel')]) {
                await target.click({ button: 'right' });
                await page.waitForTimeout(250 + 4 * delay);
                await expect(popup).toBeVisible();
                await expect(count).toHaveText(String(expected));
                await expect(page.locator('.fx-grid-header-context-menu')).toHaveCount(0);
                checks += 3;
            }
            await page.mouse.click(width - 3, 3, { button: 'right' });
            await page.waitForTimeout(250 + 4 * delay);
            await expect(popup).toBeVisible();
            await expect(count).toHaveText(String(expected));
            checks += 2;
        }

        await grid.locator('tbody tr.fx-row').first().click({ button: 'right' });
        await expect(count).toHaveText('1');
        checks++;
        latency = delay;
        await open();
        await contextIsolated(1);

        const title = header.locator('span.fx-dialog-title');
        await expect(title).toHaveText('Choose Columns');
        await expect(title).toHaveCSS('color', 'rgb(0, 0, 0)');
        await expect(title).toHaveCSS('font-weight', '700');
        assert.equal(await title.evaluate(element => element.isContentEditable || element.matches(':disabled, [aria-disabled="true"]')), false);
        checks += 4;

        const initial = await popup.boundingBox();
        assert(initial.x >= 0 && initial.y >= 0 && initial.x + initial.width <= width && initial.y + initial.height <= height);
        await expect(header).toHaveCSS('cursor', 'move');
        await expect(header).toHaveCSS('user-select', 'none');
        await page.evaluate(() => window.getSelection().removeAllRanges());
        const bar = await header.boundingBox();
        const dx = width < 500 ? 10 : 70;
        const dy = 35;
        await page.mouse.move(bar.x + 45, bar.y + bar.height / 2);
        await page.mouse.down();
        await page.mouse.move(bar.x + 45 + dx, bar.y + bar.height / 2 + dy, { steps: 5 });
        await page.mouse.up();
        await expect.poll(async () => Math.round((await popup.boundingBox()).x - initial.x)).toBe(dx);
        await expect.poll(async () => Math.round((await popup.boundingBox()).y - initial.y)).toBe(dy);
        assert.equal(await page.evaluate(() => window.getSelection().toString()), '');
        await expect(count).toHaveText('1');
        await expect(popup).toBeVisible();
        checks += 8;
        await page.screenshot({
            path: `${output}/${kit}-${width}-${delay}.png`,
            style: 'body > div[style*="z-index: 99999"] { visibility: hidden !important; }'
        });

        // Cancel discards column edits even after dragging.
        await row('Description').click();
        await button('Move Up').click();
        await expect(list.getByRole('option').first()).toHaveText('Description');
        await row('Vendor').getByRole('checkbox').check();
        await button('Cancel').click();
        await closed(0);
        await open();
        await expect(list.getByRole('option').first()).toHaveText('Code');
        await expect(row('Vendor').getByRole('checkbox')).not.toBeChecked();
        await expect.poll(async () => Math.round((await popup.boundingBox()).x)).toBe(Math.round(initial.x));
        checks += 4;

        // OK preserves the existing order/visibility workflow.
        await row('Description').click();
        await button('Move Up').click();
        await row('Vendor').getByRole('checkbox').check();
        await button('OK').click();
        await closed(1);
        await expect(grid.locator('thead th[data-field]').first()).toHaveAttribute('data-field', 'Description');
        await expect(grid.locator('thead th[data-field="Vendor"]')).toBeVisible();
        checks += 2;

        for (const close of ['Escape', 'Close', 'Overlay']) {
            await open();
            await expect(list.getByRole('option').first()).toHaveText('Description');
            if (close === 'Escape') {
                await button('Cancel').focus();
                await page.keyboard.press('Escape');
            } else if (close === 'Close') await button('Close').click();
            else await page.mouse.click(width - 3, 3);
            await closed(1);
            checks++;
        }

        await open();
        const nextRow = list.getByRole('option').nth(1);
        const wasChecked = await nextRow.getByRole('checkbox').isChecked();
        await list.press('ArrowDown');
        await expect(nextRow).toHaveClass(/\bselected\b/);
        await page.keyboard.press('Space');
        await expect(nextRow.getByRole('checkbox')).toBeChecked({ checked: !wasChecked });
        await page.keyboard.press('Escape');
        await closed(1);
        await grid.locator('tbody tr.fx-row').first().click({ button: 'right' });
        await expect(count).toHaveText('2');
        checks += 3;

        // The same popup must isolate events when the grid is already in a dialog.
        await page.getByRole('button', { name: 'Open modal grid', exact: true }).click();
        await expect(page.getByRole('dialog', { name: 'Grid host', exact: true })).toBeVisible();
        await open();
        await contextIsolated(2);
        await button('Cancel').click();
        await closed(1);
        await expect(page.getByRole('dialog', { name: 'Grid host', exact: true })).toBeVisible();
        await open();
        await button('Cancel').focus();
        await page.keyboard.press('Escape');
        await closed(1);
        await expect(page.getByRole('dialog', { name: 'Grid host', exact: true })).toBeVisible();
        await page.getByRole('dialog', { name: 'Grid host', exact: true }).getByRole('button', { name: 'Close', exact: true }).click();
        await expect(page.getByRole('dialog')).toHaveCount(0);
        assert.deepEqual(errors, []);
        checks += 3;
        console.log(`${kit}: ${width}x${height}, ${delay}ms each way passed`);
        await page.close();
    }
    console.log(`PASS ${checks} Choose Columns browser checks`);
} finally {
    await browser.close();
}
