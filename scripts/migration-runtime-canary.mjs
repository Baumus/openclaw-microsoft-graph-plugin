import assert from 'node:assert/strict';
import { randomBytes, createHash, createCipheriv } from 'node:crypto';
import { mkdir, writeFile, readFile, readdir, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { parseArgs } from 'node:util';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { join, resolve, dirname } from 'node:path';
import { createRequire } from 'node:module';
// Synthetic Gateway/providers only. The entry's own dependency tree selects the
// actual OpenClaw release; --entry can target an extracted npm archive.
const { values } = parseArgs({ options: { entry: { type: 'string' }, receipt: { type: 'string' }, auth: { type: 'string', default: 'token' }, owner: { type: 'string', default: 'independent-person' } } });
assert.ok(['token','password'].includes(values.auth));
const entryUrl = values.entry ? pathToFileURL(resolve(values.entry)) : new URL('../dist/index.js', import.meta.url);
const { default: entry } = await import(entryUrl.href);
const require = createRequire(entryUrl);
const cliEntry = require.resolve('openclaw/cli-entry');
const { version: runtimeVersion } = JSON.parse(await readFile(join(dirname(cliEntry), 'package.json'), 'utf8'));
assert.ok(['2026.10.1','2026.10.2'].includes(runtimeVersion));
const { runCommandWithTimeout } = await import(pathToFileURL(require.resolve('openclaw/plugin-sdk/process-runtime')).href);
for (const key of ['OPENCLAW_GATEWAY_PORT','OPENCLAW_GATEWAY_TOKEN','OPENCLAW_GATEWAY_PASSWORD','OPENCLAW_GATEWAY_URL','OPENCLAW_PROFILE']) delete process.env[key];
const root = await mkdtemp(join(tmpdir(), 'msgraph-migration-canary-'));
const owner = values.owner;
const manifestPath = join(dirname(fileURLToPath(entryUrl)), '..', 'openclaw.plugin.json');
const { WebSocketServer }=createRequire(cliEntry)('ws');
const roster=ids=>({profiles:ids.map(id=>({id,displayName:id,avatarMime:null,mergedInto:id==='synthetic-merged-alias'?owner:null,createdAt:1,updatedAt:1,emails:[],githubIdentity:null,hasAvatar:false,roleSource:'default'}))});
const policy={version:2,rules:{default:'deny'},services:{onedrive:{allowed_roots:[]},calendar:{agents:{}},mail:{agents:{main:{operations:['read'],resources:['me']}}},todo:{agents:{}}}};
const results={entrySelection:values.entry?'explicit':'source-built',startupEntry:true,noPersonalRpc:true,realPublicCli:true,providerFetches:0,commandCalls:[],checks:[],runtimeVersion, authMode:values.auth, ownerSynthetic:true};
const server=new WebSocketServer({host:'127.0.0.1',port:0});
await new Promise(r=>server.once('listening',r));
let ids=['gateway-owner',owner,'synthetic-merged-alias'];let methods;let realCommands=true;
const config={gateway:{mode:'local',port:server.address().port,bind:'loopback',auth: values.auth==='password'?{password:'synthetic-packed-host-auth'}:{token:'synthetic-packed-host-auth'}}};
server.on('connection',socket=>{
 socket.send(JSON.stringify({type:'event',event:'connect.challenge',payload:{nonce:'synthetic-challenge',ts:Date.now()}}));
 socket.on('message',bytes=>{
  const req=JSON.parse(bytes.toString());
  if(req.method==='connect'){
   assert.equal(req.params.auth?.[values.auth],'synthetic-packed-host-auth');
   socket.send(JSON.stringify({type:'res',id:req.id,ok:true,payload:{type:'hello-ok',protocol:req.params.maxProtocol,server:{version:runtimeVersion,connId:'synthetic'},features:{methods:['users.list','microsoft-graph.directory-host'],events:[]},snapshot:{},policy:{maxPayload:1048576,maxBufferedBytes:1048576,tickIntervalMs:30000}}}));
  }else{
   assert.ok(['users.list','microsoft-graph.directory-host'].includes(req.method));
   let payload;if(req.method==='users.list')payload=roster(ids);else methods.get(req.method).handler({params:{},respond:(_ok,value)=>payload=value});
   socket.send(JSON.stringify({type:'res',id:req.id,ok:true,payload}));
  }
 });
});
async function syntheticVault(state,key){
 const directory=join(state,'plugin-data','microsoft-graph','credentials');await mkdir(directory,{recursive:true,mode:0o700});
 const keyBytes=Buffer.from(key,'base64url');const keyId=createHash('sha256').update(keyBytes).digest().subarray(0,16).toString('base64url');
 const nonce=randomBytes(12);const format='openclaw-microsoft-graph-credential-vault';const cipher=createCipheriv('aes-256-gcm',keyBytes,nonce,{authTagLength:16});
 cipher.setAAD(Buffer.from(`${format}\0${1}\0${1}\0A256GCM\0${keyId}`));
 const ciphertext=Buffer.concat([cipher.update(JSON.stringify({clientId:'synthetic-client',refreshToken:'synthetic-canary-refresh-never-real',tenant:'common',scopes:['Mail.Read','offline_access']}),'utf8'),cipher.final()]);
 const value=JSON.stringify({format,version:1,generation:1,keyId,algorithm:'A256GCM',nonce:nonce.toString('base64url'),ciphertext:ciphertext.toString('base64url'),tag:cipher.getAuthTag().toString('base64url')})+'\n';
 const path=join(directory,'credential.vault.json');await writeFile(path,value,{mode:0o600});return {path,value};
}
function scheduler(){
 const controller=new AbortController(),pending=new Map(),running=new Set();
 const s={version:1,signal:controller.signal,now:Date.now,scope:()=>s,beginClose(){controller.abort();pending.clear();},schedule(job){assert.equal(controller.signal.aborted,false);pending.set(job.id,job.run);return {cancel:()=>pending.delete(job.id),stop:async()=>{pending.delete(job.id);await Promise.all(running);}};},async stop(){s.beginClose();await Promise.all(running);},pending,async flush(){while(pending.size&&!controller.signal.aborted){const[id,callback]=[...pending][0];pending.delete(id);const work=Promise.resolve().then(callback);running.add(work);try{await work;}finally{running.delete(work);}}}};return s;
}
async function host(name){
 const state=join(root,name+'-'+Date.now());await mkdir(state,{mode:0o700});const key=randomBytes(32).toString('base64url');const vault=await syntheticVault(state,key);
 methods=new Map();const tools=new Map();const sessions=new Map();const logs=[];const hooks=new Map(),services=[];let schedule=scheduler();const hostAbort=new AbortController();
 const api={registrationMode:'full',pluginConfig:{enabled:true,credentialVaultKey:key,policy},runtime:{version:runtimeVersion,config:{current:()=>config},state:{resolveStateDir:()=>state},agent:{session:{getSessionEntry:({sessionKey})=>sessions.get(sessionKey)}},system:{runCommandWithTimeout:async(argv,opts)=>{
  assert.ok(!JSON.stringify(argv).includes((config.gateway.auth.token??config.gateway.auth.password)));assert.ok(dirname(opts.baseEnv.OPENCLAW_STATE_DIR)===tmpdir());assert.ok(opts.baseEnv.OPENCLAW_STATE_DIR.includes('msgraph-users-read-'));assert.notEqual(opts.baseEnv.OPENCLAW_STATE_DIR,state);
  results.commandCalls.push({kind:argv.includes('users')?'users.list':'proof',real:realCommands});
  if(realCommands)return runCommandWithTimeout(argv,opts);
  let value;if(argv.includes('users'))value=roster(ids);else methods.get('microsoft-graph.directory-host').handler({params:{},respond:(_ok,v)=>value=v});
  return {stdout:JSON.stringify(value),stderr:'',code:0,killed:false,signal:null,termination:'exit'};
 }}},registerTool:(factory,options)=>tools.set(options.name??options.names[0],factory),registerGatewayMethod:(name,handler,options)=>methods.set(name,{handler,options}),registerService:service=>services.push(service),on:(name,handler)=>hooks.set(name,handler),logger:{info:v=>logs.push(v),warn:v=>logs.push(v),error:v=>logs.push(v)}};
 entry.register(api);
 const start=async()=>{await services.at(-1).start({scheduler:schedule});assert.equal(schedule.pending.size,0,'initial start cannot schedule before publication');await hooks.get('gateway_start')({port:config.gateway.port},{config,abortSignal:hostAbort.signal});};
 const restart=async()=>{await services.at(-1).stop();entry.register(api);schedule=scheduler();await start();await schedule.flush();};
 const restartRetained=async()=>{const definition=services.at(-1);await definition.stop();schedule=scheduler();config.gateway.controlUi={allowedOrigins:['https://synthetic.invalid']};await definition.start({scheduler:schedule});assert.equal(schedule.pending.size,1,'retained same definition reacquires scheduled work without gateway_start');};
 const tool=agentId=>{const sessionKey='agent:'+agentId+':packed';const sessionId='packed-'+agentId;sessions.set(sessionKey,{sessionId,lifecycleRevision:'1'});return tools.get('outlook_mail_read').create({agentId,sessionKey,sessionId,assertInvocationCurrent(){}});};
 return {state,vault,start,restart,restartRetained,flush:()=>schedule.flush(),stop:()=>services.at(-1).stop(),tool,tools,logs};
}
try{
 const h=await host('synthetic-host');
 assert.equal(methods.get('microsoft-graph.accounts.automatic-migration').options.scope,'operator.admin');
 for(const [name,m]of methods)if(name.startsWith('microsoft-graph.accounts.'))assert.equal(m.options.profileAccess,'required');
 const manifest=JSON.parse(await readFile(manifestPath,'utf8'));
 assert.deepEqual([...h.tools.keys()].sort(),manifest.contracts.tools.toSorted());results.checks.push('all 17 declared tools registered; personal methods retain genuine profile/admin gates');
 globalThis.fetch=()=>{throw new Error('startup_must_not_fetch_provider');};
 const accountDirectory=join(h.state,'plugin-data/microsoft-graph');
 await h.start();assert.equal((await readdir(accountDirectory)).includes('accounts'),false);
 await h.restartRetained();assert.equal(results.commandCalls.length,0,'cancelled queued generation did not read directory');
 await h.flush();results.checks.push('installed-shaped queued service stop/start of same registration resumes without replaying gateway_start; unrelated UI config does not revoke authority');results.startupLogs=[...h.logs];assert.equal(results.providerFetches,0);assert.ok(results.commandCalls.length>0,'startup must actually read exact host roster');const registry=JSON.parse(await readFile(join(accountDirectory,'accounts/registry.json'),'utf8'));assert.equal(registry.automatic.profileId,owner);assert.equal(registry.automatic.invalidated,false);assert.equal(Object.values(registry.accounts)[0].connected,true);
 assert.equal(await readFile(h.vault.path,'utf8'),h.vault.value);results.checks.push('post-ready registered service/hook startup uses real exported CLI/public runner with no account RPC/UI/profile/admin/provider; encrypted legacy backup unchanged');
 globalThis.fetch=async(url,options)=>{results.providerFetches++;const value=String(url);if(value.startsWith('https://login.microsoftonline.com/')){assert.ok(String(options.body).includes('synthetic-canary-refresh-never-real'));return new Response(JSON.stringify({access_token:'synthetic-canary-token-never-real',expires_in:3600}));}if(value.startsWith('https://graph.microsoft.com/')){assert.equal(new Headers(options.headers).get('authorization'),'Bearer synthetic-canary-token-never-real');return new Response(JSON.stringify({value:[{id:'packed-synthetic-mail',subject:'packaged path'}]}));}throw new Error('unexpected_provider_url');};
 const result=await h.tool('main').execute('packed-allowed',{action:'list_messages'},AbortSignal.timeout(90000));assert.equal(result.details.ok,true);assert.ok(JSON.stringify(result).includes('packed-synthetic-mail'));assert.ok(!JSON.stringify(result).includes('synthetic-canary-token-never-real'));assert.ok(!JSON.stringify(result).includes('synthetic-canary-refresh-never-real'));results.checks.push('real public CLI authorized unbound packaged mail tool to synthetic OAuth/Graph; no model-facing credentials');
 // Remaining package negatives use synthetic argv output: avoid redundant slow CLI startups.
 realCommands=false;const dispatched=results.providerFetches;
 const denied=await h.tool('not-authorized').execute('packed-denied',{action:'list_messages'});assert.equal(denied.details.error,'access_denied');assert.equal(results.providerFetches,dispatched);results.checks.push('unapproved agent denies before OAuth/Graph');
 ids.push('second-independent-person');const changed=await h.tool('main').execute('packed-multi',{action:'list_messages'});assert.equal(changed.details.error,'graph_single_user_topology_changed');assert.equal(results.providerFetches,dispatched);results.checks.push('second canonical human permanently disables automatic tool path before provider');
 ids=['gateway-owner',owner,'synthetic-merged-alias'];await h.restart();const sticky=await h.tool('main').execute('packed-sticky',{action:'list_messages'});assert.equal(sticky.details.error,'graph_single_user_topology_changed');results.checks.push('singleton return plus fresh service generation/restart cannot revive automatic grant');
 ids=['gateway-owner',owner,'second-independent-person'];const second=await host('synthetic-multiuser');await second.start();await second.flush();
 const multi=await second.tool('main').execute('packed-multi-fresh',{action:'list_messages'});assert.equal(multi.details.ok,false);results.checks.push('fresh multi-user startup does not choose a legacy owner');
 assert.equal(await readFile(second.vault.path,'utf8'),second.vault.value);assert.equal(results.providerFetches,dispatched);assert.deepEqual(second.logs,[]);await second.stop();await h.stop();
 results.pass=true;if(values.receipt)await writeFile(resolve(values.receipt),JSON.stringify(results,null,2)+'\n');console.log(JSON.stringify({pass:true,checks:results.checks,providerFetches:results.providerFetches,commands:results.commandCalls.length}));
}catch(error){results.pass=false;results.error={message:error.message,stack:error.stack};if(values.receipt)await writeFile(resolve(values.receipt),JSON.stringify(results,null,2)+'\n');throw error;}finally{for(const client of server.clients)client.terminate();await new Promise(r=>server.close(r));await rm(root,{recursive:true,force:true});}
