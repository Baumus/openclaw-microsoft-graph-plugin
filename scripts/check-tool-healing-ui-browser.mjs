// Synthetic Chromium lifecycle. No live Gateway/config/provider access.
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { build } from "esbuild";
import { chromium } from "playwright-core";
const output=path.resolve(process.env.UI_EVIDENCE_DIR ?? "/tmp/msgraph-tool-healing-ui-evidence");
await mkdir(output,{recursive:true});
const bundle=await build({entryPoints:["src/control-ui.ts"],bundle:true,write:false,format:"iife",globalName:"GraphFixture",platform:"browser",loader:{".css":"text"},plugins:[{name:"synthetic-host",setup(b){b.onResolve({filter:/^openclaw\/plugin-sdk\/control-ui$/},()=>({path:"host",namespace:"fixture"}));b.onLoad({filter:/.*/,namespace:"fixture"},()=>({contents:"export const defineControlUiPlugin = plugin => plugin;",loader:"js"}));}}]});
const policy={version:2,rules:{default:"deny"},services:{onedrive:{allowed_roots:[]},calendar:{agents:{}},mail:{agents:{main:{operations:["read"]}}},todo:{agents:{}}}};
const canonical=v=>JSON.stringify(normalize(v));
function normalize(v){return Array.isArray(v)?v.map(normalize):v&&typeof v==="object"?Object.fromEntries(Object.keys(v).sort().map(k=>[k,normalize(v[k])])):v;}
const policyHash=createHash("sha256").update(canonical(policy)).digest("hex");
const browser=await chromium.launch({executablePath:process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH,headless:true});
try {
  const page=await browser.newPage({viewport:{width:1150,height:850}}); await page.setContent('<div id="fixture"></div>'); await page.addScriptTag({content:bundle.outputFiles[0].text});
  await page.evaluate(({policy,policyHash})=>{
    let pageDefinition; let listener; let revision="r1"; let intent=null; let applied=false;
    let finishClaim; const waitClaim=new Promise(resolve=>{finishClaim=resolve;});
    let finishPatch; const waitPatch=new Promise(resolve=>{finishPatch=resolve;});
    const calls=[]; const entries={main:{tools:{alsoAllow:["read","outlook_mail_write"],deny:[]}}};
    const expected={main:{key:"alsoAllow",value:["read","outlook_mail_read"]}};
    const host={locale:"de",connection:{connected:true,canAdmin:true},agents:{rows:[{id:"main",name:"Synthetic agent"}]},subscribe:fn=>{listener=fn;return()=>{};},request:async(method,params)=>{
      calls.push({method,params});
      if(method==="config.get")return {hash:revision,configRevisionHash:revision,appliedConfigHash:applied?revision:"older",parsed:{plugins:{$include:"synthetic-plugins.json"}},config:{agents:{entries:structuredClone(entries)},plugins:{entries:{"microsoft-graph":{enabled:true,config:{credentialVaultKey:"synthetic",policy}}}}}};
      if(method==="microsoft-graph.configuration.applicationStatus")return {activePolicyHash:policyHash,toolReconciliation:"control-ui-cas-v1"};
      if(method==="microsoft-graph.configuration.validate")return {valid:true,requiredScopes:[],policyHash};
      if(method==="microsoft-graph.credentials.status")return {ok:true,value:{policyVersion:2,credential:{result:"valid"}}};
      if(method==="microsoft-graph.updateStatus")return {updateAvailable:false};
      if(method==="microsoft-graph.configuration.activation"){
        if(params.action==="status")return intent;
        if(params.action==="reconcile"){await waitClaim;intent={id:"synthetic",targetHash:policyHash,state:"issued",expected};return {record:structuredClone(intent),baseHash:revision,expected,patch:{main:{tools:{alsoAllow:expected.main.value}}}};}
        if(params.action==="complete"){intent=null;return {complete:true};}
      }
      if(method==="config.patch"){
        if(params.baseHash!==revision)throw Error("CAS conflict");
        const patch=JSON.parse(params.raw); if(patch.plugins)throw Error("must not rewrite included policy");
        entries.main.tools.alsoAllow=patch.agents.entries.main.tools.alsoAllow; revision="r2";
        await waitPatch; applied=true; return {};
      }
      throw Error(`Unexpected RPC ${method}`);
    },ui:{registerPage:p=>{pageDefinition=p;return()=>{};},registerNavigation:()=>()=>{}}};
    GraphFixture.default.activate(host); pageDefinition.mount(document.querySelector("#fixture"),{host,signal:new AbortController().signal});
    window.fixture={calls,releaseClaim:()=>finishClaim(),releasePatch:()=>finishPatch(),deny:()=>{entries.main.tools.alsoAllow=["read"];entries.main.tools.deny=["outlook_mail_*"];revision="r3";listener();}};
  },{policy,policyHash});
  await page.getByText("Synthetic agent (main): fehlende Tools: outlook_mail_read",{exact:true}).waitFor();
  assert.ok(await page.getByText("Synthetic agent (main): überzählige Tools: outlook_mail_write",{exact:true}).count());
  await page.screenshot({path:path.join(output,"drift-detected.png"),fullPage:true});
  await page.evaluate(()=>window.fixture.releaseClaim());
  await page.waitForFunction(()=>window.fixture.calls.some(c=>c.method==="config.patch"));
  await page.locator(".mg-save-status").waitFor();
  assert.equal(await page.getByText("Gespeichert und aktiv",{exact:true}).count(),0);
  await page.evaluate(()=>window.fixture.releasePatch());
  await page.getByText("Gespeichert und aktiv",{exact:true}).waitFor();
  await page.screenshot({path:path.join(output,"repair-complete.png"),fullPage:true});
  await page.evaluate(()=>window.fixture.deny()); await page.getByRole("button",{name:"Status prüfen",exact:true}).click();
  await page.getByText("Synthetic agent (main): durch Deny-Regeln gesperrt: outlook_mail_read",{exact:true}).waitFor();
  await page.screenshot({path:path.join(output,"explicit-deny-preserved.png"),fullPage:true});
  const calls=await page.evaluate(()=>window.fixture.calls);
  assert.equal(calls.filter(c=>c.method==="config.patch").length,1); assert.equal(calls.filter(c=>c.method==="plugins.reload").length,0);
  const result={result:"PASS",scenarios:["agent/tool-specific warning","root-only include-preserving CAS","unconfirmed tools not reported active","repair completion","explicit deny preserved"],writes:1,reloads:0};
  await writeFile(path.join(output,"result.json"),JSON.stringify(result,null,2)+"\n"); console.log(JSON.stringify(result));
} finally {await browser.close();}
