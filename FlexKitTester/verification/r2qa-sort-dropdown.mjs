import assert from 'node:assert/strict';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { resolve } from 'node:path';
import { homedir, tmpdir } from 'node:os';
const playwrightPath = process.env.PLAYWRIGHT_MODULE ?? resolve(homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs');
const { chromium } = await import(pathToFileURL(playwrightPath).href);
const benchUrl = process.env.BENCH_URL ?? 'http://127.0.0.1:5299';
const flexCoreRoot = fileURLToPath(new URL('../../../FlexCore/', import.meta.url));
const outputRoot = process.env.OUTPUT_DIR ?? tmpdir();
const browser=await chromium.launch({channel:'chrome',headless:true});
try{
 const page=await browser.newPage({viewport:{width:1400,height:900}});const errors=[];page.on('pageerror',e=>errors.push(e.message));
 let lag=0;await page.routeWebSocket('**/_blazor*',socket=>{const server=socket.connectToServer();socket.onMessage(m=>setTimeout(()=>server.send(m),lag));server.onMessage(m=>setTimeout(()=>socket.send(m),lag));});
 await page.goto(`${benchUrl}/r2qa-grid?Count=500`);const grid=page.locator('#qa-grid .fx-grid');await grid.locator('[data-fx-drag-starts]').waitFor();await page.waitForTimeout(600);
 await page.getByRole('button',{name:'Custom Sort',exact:true}).click();
 await page.getByRole('button',{name:'Sort order level 1',exact:true}).click();await page.getByRole('option',{name:'Z to A',exact:true}).click();
 await page.getByRole('button',{name:'Add Level',exact:true}).click();await page.getByRole('button',{name:'Apply',exact:true}).click();await page.waitForTimeout(600);
 const order=await grid.locator('tr.fx-row td[data-field="Description"]').allTextContents();
 assert.deepEqual(order.slice(0,5).map(s=>s.trim()),['Description 00499','Description 00498','Description 00497','Description 00496','Description 00495']);
 assert.equal(await grid.locator('.fx-sort-icon').count(),2);
 await grid.locator('th[data-field="Community"]').click();await page.waitForTimeout(600);assert.equal(await grid.locator('.fx-sort-icon').count(),1);
 console.log('PASS: OptionID descending + Community ascending tie-breaks; ordinary header click replaces both levels');
 lag=250;
 for(const option of ['SF','LSUM','EA']){
  await page.locator('#qa-controls .fx-dropdown').click();await page.getByRole('option',{name:option,exact:true}).click();await page.waitForTimeout(1400);
  assert.equal(await page.locator('#qa-controls .fx-dropdown-text').innerText(),option);assert.equal(await page.locator('.fx-dropdown-panel').count(),0);
 }
 await page.locator('#qa-controls .fx-dropdown').click();await page.getByRole('option',{name:'EA',exact:true}).waitFor();await page.keyboard.press('Escape');await page.waitForTimeout(1400);assert.equal(await page.locator('.fx-dropdown-panel').count(),0);
 for(const width of [1400,900,390]){await page.setViewportSize({width,height:900});await page.addStyleTag({content:'#qa-grid input[type=checkbox]{width:15px;height:15px}'});const box=grid.locator('input[type=checkbox]').first();const shape=await box.evaluate(e=>{const s=getComputedStyle(e,'::after');return{left:s.left,top:s.top,transform:s.transform};});assert.notEqual(shape.transform,'none');await page.screenshot({path:resolve(outputRoot, `r2qa-${width}.png`)});}
 console.log('PASS: early dropdown selection/Escape at 500ms RTT; checkbox layouts at 1400/900/390');assert.deepEqual(errors,[]);
}finally{await browser.close();}
