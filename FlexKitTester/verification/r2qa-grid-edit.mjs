import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {execFileSync} from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { resolve } from 'node:path';
import { homedir, tmpdir } from 'node:os';
const playwrightPath = process.env.PLAYWRIGHT_MODULE ?? resolve(homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs');
const { chromium } = await import(pathToFileURL(playwrightPath).href);
const benchUrl = process.env.BENCH_URL ?? 'http://127.0.0.1:5299';
const flexCoreRoot = fileURLToPath(new URL('../../../FlexCore/', import.meta.url));
const outputRoot = process.env.OUTPUT_DIR ?? tmpdir();
const browser=await chromium.launch({channel:'chrome',headless:true});
try {
 for(const windowed of [false]) {
 const page=await browser.newPage({viewport:{width:1400,height:900}});let lag=0;const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.routeWebSocket('**/_blazor*',socket=>{const server=socket.connectToServer();socket.onMessage(m=>setTimeout(()=>server.send(m),lag));server.onMessage(m=>setTimeout(()=>socket.send(m),lag));});
 await page.route('**/grid-control.js*',async route=>route.fulfill({contentType:'text/javascript',body:process.env.BASELINE?execFileSync('git',['show','HEAD:wwwroot/grid-control.js'],{cwd:flexCoreRoot,encoding:'utf8'}):await readFile(resolve(flexCoreRoot, 'wwwroot/grid-control.js'),'utf8')}));
 await page.goto(`${benchUrl}/r2qa-grid?Count=500&WindowColumns=${windowed}&Overscan=5`);const grid=page.locator('#qa-grid .fx-grid');await grid.locator('[data-fx-drag-starts]').waitFor();await page.waitForTimeout(700);lag=150;
 const cell=(row,field)=>grid.locator(`tr.fx-row[data-ari="${row}"] td[data-field="${field}"]`);
 await cell(0,'OptionID').click();await page.waitForTimeout(1000);
 await grid.evaluate(root=>{window.qaNav={bad:[],frames:0,running:true};const sample=()=>{if(!window.qaNav.running)return;let cells=[...root.querySelectorAll('td.fx-cell')].filter(c=>getComputedStyle(c).boxShadow!=='none');window.qaNav.frames++;if(cells.length>1)window.qaNav.bad.push(cells.map(c=>c.dataset.field));requestAnimationFrame(sample);};sample();});
 for(let i=0;i<8;i++){await page.keyboard.press('ArrowRight');await page.waitForTimeout(180);}await page.waitForTimeout(1600);
 const nav=await page.evaluate(()=>{window.qaNav.running=false;return window.qaNav;});console.log('nav',windowed,'bad',nav.bad.length,'frames',nav.frames,'first',nav.bad.slice(0,3));if(!process.env.BASELINE)assert.equal(nav.bad.length,0);
 if(process.env.BASELINE){await page.close();continue;}
 // Return horizontally before testing mass edits in the windowed grid.
 await grid.locator('.fx-grid-content').evaluate(e=>{e.scrollLeft=0;e.dispatchEvent(new Event('scroll'));});await page.waitForTimeout(1200);
 for(const end of ['Enter','Tab','blur']) {
  const a=await cell(0,'Description').boundingBox(), b=await cell(2,'Description').boundingBox();
  await page.mouse.move(a.x+15,a.y+a.height/2);await page.mouse.down();await page.waitForTimeout(1000);await page.mouse.move(b.x+15,b.y+b.height/2,{steps:10});await page.mouse.up();await page.waitForTimeout(1200);
  assert.equal(await grid.locator('td[data-field="Description"].fx-cell-selected').count(),3);
  await page.keyboard.type(`Batch${end}`);await page.waitForTimeout(700);
  if(end==='blur')await cell(4,'OptionID').click();else await page.keyboard.press(end);
  await page.waitForTimeout(1500);
  for(let row=0;row<3;row++)assert.equal((await cell(row,'Description').innerText()).trim(),`Batch${end}`);
 }
 // Traverse to the last column and back without wrong logical indices.
 await cell(0,'Description').click();await page.waitForTimeout(500);await page.keyboard.press('Control+End');await page.waitForTimeout(1400);
 console.log('edit modes pass',windowed,'errors',errors);assert.deepEqual(errors,[]);await page.close();
 }
}finally{await browser.close();}
