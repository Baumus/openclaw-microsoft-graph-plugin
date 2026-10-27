import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
const pluginRoot = fileURLToPath(new URL('../', import.meta.url));
const require = createRequire(import.meta.url);
const sdk = process.env.OPENCLAW_CONFIG_MUTATION_SDK || require.resolve('openclaw/plugin-sdk/config-mutation');
const dir = await fs.mkdtemp(path.join(os.tmpdir(),'graph-include-native-'));
process.env.OPENCLAW_STATE_DIR = dir;
process.env.OPENCLAW_CONFIG_PATH = path.join(dir,'openclaw.json');
process.env.OPENCLAW_TEST_FAST = '1';
const {replaceConfigFile,readConfigFileSnapshotForWrite} = await import(pathToFileURL(sdk).href);

await fs.writeFile(process.env.OPENCLAW_CONFIG_PATH,JSON.stringify({agents:{entries:{synthetic:{tools:{alsoAllow:["outlook_mail_read"]}}}},plugins:{$include:"plugins.json"}}));
await fs.writeFile(path.join(dir,'plugins.json'),JSON.stringify({load:{paths:[pluginRoot]},entries:{"microsoft-graph":{enabled:false,config:{policy:{version:2,rules:{default:"deny"},services:{onedrive:{allowed_roots:[]},calendar:{agents:{}},mail:{agents:{}},todo:{agents:{}}}}}}}}));
async function change(both) {
 const {snapshot,writeOptions}=await readConfigFileSnapshotForWrite();
 const next=structuredClone(snapshot.sourceConfig);
 if(both!=="tools") next.plugins.entries["microsoft-graph"].config.policy.services.calendar.agents.synthetic={operations:["read"],resources:["me"]};
 if(both!=="policy") next.agents.entries.synthetic.tools.alsoAllow.push("outlook_calendar_read");
 return replaceConfigFile({sourceConfig:next,baseHash:snapshot.hash,writeOptions:{...writeOptions,skipOutputLogs:true},afterWrite:{mode:"none",reason:"synthetic ownership test"}});
}
let rejected=false;
try {await change("both")} catch(e){rejected=String(e).includes("$include");console.log("combined",e.code,String(e.message).slice(0,180));}
if(!rejected) throw Error("combined must reject");
await change("policy");await change("tools");
const root=JSON.parse(await fs.readFile(process.env.OPENCLAW_CONFIG_PATH));
const include=JSON.parse(await fs.readFile(path.join(dir,'plugins.json')));
if(root.plugins.$include!=="plugins.json" || include.entries["microsoft-graph"].config.policy.services.calendar.agents.synthetic.operations[0]!=="read" || !root.agents.entries.synthetic.tools.alsoAllow.includes("outlook_calendar_read"))throw Error("bad result");
console.log("split writes PASS: include retained; exact calendar tools and policy persisted");
await fs.rm(dir,{recursive:true,force:true});
