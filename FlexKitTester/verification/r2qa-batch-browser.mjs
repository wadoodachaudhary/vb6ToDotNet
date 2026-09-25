import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { resolve } from 'node:path';
import { homedir, tmpdir } from 'node:os';
const playwrightPath = process.env.PLAYWRIGHT_MODULE ?? resolve(homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs');
const { chromium } = await import(pathToFileURL(playwrightPath).href);
const benchUrl = process.env.BENCH_URL ?? 'http://127.0.0.1:5299';
const flexCoreRoot = fileURLToPath(new URL('../../../FlexCore/', import.meta.url));
const outputRoot = process.env.OUTPUT_DIR ?? tmpdir();
const browser = await chromium.launch({channel:'chrome',headless:true});
const results=[];
async function open(windowed, delay=0, overscan=20) {
 const page=await browser.newPage({viewport:{width:1400,height:900}});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 let lag=0;
 await page.routeWebSocket('**/_blazor*',socket=>{const server=socket.connectToServer();socket.onMessage(m=>setTimeout(()=>server.send(m),lag));server.onMessage(m=>setTimeout(()=>socket.send(m),lag));});
 for(const file of ['grid-control.js','dropdown-list-control.js']) await page.route(`**/${file}*`,async route=>route.fulfill({contentType:'text/javascript',body:await readFile(resolve(flexCoreRoot, `wwwroot/${file}`),'utf8')}));
 await page.goto(`${benchUrl}/r2qa-grid?Count=5000&WindowColumns=${windowed}&Overscan=${overscan}`);
 const grid=page.locator('#qa-grid .fx-grid');await grid.locator('tbody tr.fx-row').first().waitFor();await grid.locator('[data-fx-drag-starts]').waitFor();await page.waitForTimeout(1000);lag=delay;
 return {page,grid,errors};
}
try {
 for(const overscan of [20,5]) for(const delay of [0,150]) {
  const windowed=false;
  const {page,grid,errors}=await open(windowed,delay,overscan);
  const cells=await grid.locator('td.fx-cell').count();
  const thumb=grid.locator('.fx-grid-deferred-vscroll-thumb'),lane=grid.locator('.fx-grid-deferred-vscroll');
  const t=await thumb.boundingBox(), l=await lane.boundingBox();
  console.log('scroll',windowed,delay,cells);
  await page.mouse.move(t.x+t.width/2,t.y+t.height/2);await page.mouse.down();await page.mouse.move(t.x+t.width/2,l.y+l.height+100,{steps:5});
  const start=Date.now();await page.mouse.up();await grid.locator('tr.fx-row[data-ari="4999"]').waitFor({timeout:15000,state:'attached'});
  const elapsed=Date.now()-start;
  results.push({kind:'end-scroll',overscan,delay,cells,afterCells:await grid.locator('td.fx-cell').count(),elapsed});
  assert.deepEqual(errors,[]);await page.close();
 }
 const {page,grid,errors}=await open(false,200);
 const cell=(row,field)=>grid.locator(`tr.fx-row[data-ari="${row}"] td[data-field="${field}"]`);
 // Read-only drag selection remains a selection, not an edit.
 let a=await cell(0,'OptionID').boundingBox(), b=await cell(3,'OptionID').boundingBox();
 await page.mouse.move(a.x+20,a.y+a.height/2);await page.mouse.down();await page.waitForTimeout(1200);await page.mouse.move(b.x+20,b.y+b.height/2,{steps:10});await page.mouse.up();await page.waitForTimeout(1600);
 assert.equal(await grid.locator('td[data-field="OptionID"].fx-cell-selected').count(),4);
 assert.equal(await grid.locator('.fx-batch-input').count(),0);
 results.push({kind:'readonly-drag',selected:4});
 // Arrow navigation must never leave two cell borders in one row.
 await cell(0,'OptionID').click();await page.waitForTimeout(1500);
 await grid.evaluate(root=>{
  window.qaFrames={bad:[],frames:0,running:true};
  const sample=()=>{if(!window.qaFrames.running)return;const cues=[...root.querySelectorAll('td.fx-cell')].filter(c=>getComputedStyle(c).boxShadow!=='none');window.qaFrames.frames++;if(cues.length>1)window.qaFrames.bad.push(cues.map(c=>c.dataset.field));requestAnimationFrame(sample);};sample();
 });
 for(let i=0;i<8;i++){await page.keyboard.press('ArrowRight');await page.waitForTimeout(45);}
 await page.waitForTimeout(1800);
 const navigation=await page.evaluate(()=>{window.qaFrames.running=false;return window.qaFrames;});
 assert.equal(navigation.bad.length,0,JSON.stringify(navigation.bad.slice(0,4)));results.push({kind:'keyboard',frames:navigation.frames});
 // An arriving dropdown panel must be visible in its first animation frame.
 await page.evaluate(()=>{window.qaDropdown=[];const host=document.querySelector('#qa-controls .fx-dropdown-host');new MutationObserver(()=>{const panel=host.querySelector('.fx-dropdown-panel');if(panel)requestAnimationFrame(()=>window.qaDropdown.push({opacity:getComputedStyle(panel).opacity,height:panel.getBoundingClientRect().height}));}).observe(host,{childList:true});});
 await page.locator('#qa-controls .fx-dropdown').click();await page.locator('.fx-dropdown-option').first().waitFor();await page.waitForTimeout(1800);
 const dropdown=await page.evaluate(()=>window.qaDropdown);assert.equal(dropdown[0].opacity,'1');assert.ok(dropdown[0].height>30);results.push({kind:'dropdown',dropdown});
 await page.getByRole('option',{name:'SF',exact:true}).click();await page.waitForTimeout(1200);assert.equal(await page.locator('#qa-controls .fx-dropdown-text').innerText(),'SF');
 await page.locator('#qa-controls .fx-dropdown').click();await page.waitForTimeout(1200);await page.keyboard.press('Escape');await page.waitForTimeout(1200);assert.equal(await page.locator('.fx-dropdown-panel').count(),0);
 assert.deepEqual(errors,[]);await page.screenshot({path:resolve(outputRoot, 'r2qa-batch-browser.png')});await page.close();
 await writeFile(resolve(outputRoot, 'r2qa-batch-browser-results.json'),JSON.stringify(results,null,2));console.log(JSON.stringify(results,null,2));
} finally {await browser.close();}
