import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { randomBytes, createHmac } from "node:crypto";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import entry from "./index.js";
import { AccountStore, blankAccountPolicy } from "./accounts.js";
import { createVaultCredential, readVaultCredential, vaultCredentialPath, vaultQuarantinePath } from "./credential-vault.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
const roots: string[] = [];
afterEach(async () => { vi.unstubAllGlobals(); vi.unstubAllEnvs(); await Promise.all(roots.splice(0).map(root => rm(root, { recursive: true, force: true }))); });
function scheduler() {
  const abort = new AbortController(); const pending = new Map<string, () => unknown>(); const running = new Set<Promise<unknown>>();
  const s = { version: 1 as const, signal: abort.signal, now: Date.now, scope: () => s, beginClose() { abort.abort(); pending.clear(); },
    schedule(p: { id: string; run: () => unknown }) { if (abort.signal.aborted) throw new Error("closed"); pending.set(p.id, p.run); return { cancel: () => pending.delete(p.id), stop: async () => { pending.delete(p.id); await Promise.all(running); } }; },
    async stop() { s.beginClose(); await Promise.all(running); },
    async flushOne() { const first = [...pending][0]; if (!first || abort.signal.aborted) return; const [id, callback] = first; pending.delete(id); const work = Promise.resolve().then(callback); running.add(work); try { await work; } finally { running.delete(work); } },
    async flush() { while (pending.size && !abort.signal.aborted) { const [id, callback] = [...pending][0]; pending.delete(id); const work = Promise.resolve().then(callback); running.add(work); try { await work; } finally { running.delete(work); } } }, pending };
  return s;
}
const row = (id: string, mergedInto: string | null = null) => ({ id, mergedInto, displayName: id, avatarMime: null, createdAt: 1, updatedAt: 1, emails: [], githubIdentity: null, hasAvatar: false, roleSource: "default" });
async function host(options: { mode?: string; enabled?: boolean; legacy?: boolean; users?: string[]; runtimeVersion?: string; strictCeiling?: boolean } = {}) {
  const root = await mkdtemp(join(tmpdir(), "graph-startup-test-")); roots.push(root); const key = randomBytes(32).toString("base64url");
  const policy = graphPolicyFixture(); const config = { gateway: { mode: "local", port: 18789, bind: "loopback", auth: { mode: "token", token: "synthetic-only" } } };
  const methods = new Map<string, any>(), hooks = new Map<string, any>(), tools = new Map<string, any>(), services: any[] = [];
  let directory = () => JSON.stringify({ profiles: (options.users ?? ["gateway-owner", "A"]).map(id => row(id)) });
  const abort = new AbortController(), schedule = scheduler();
  const pluginConfig = { enabled: options.enabled ?? true, credentialVaultKey: key, policy, ...(options.strictCeiling ? { accountPolicyCeilings: { A: blankAccountPolicy() } } : {}) };
  const api: any = { registrationMode: options.mode ?? "full", pluginConfig,
    runtime: { version: options.runtimeVersion ?? "2026.10.1", config: { current: () => config }, state: { resolveStateDir: () => root },
      agent: { session: { getSessionEntry: () => ({ sessionId: "s", lifecycleRevision: "1" }) } },
      system: { runCommandWithTimeout: vi.fn(async (argv: string[], opts: any) => {
        let stdout: string;
        if (argv.includes("users")) stdout = directory();
        else { let value; methods.get("microsoft-graph.directory-host")({ params: {}, respond: (_ok: boolean, v: unknown) => value = v }); stdout = JSON.stringify(value); }
        return { stdout, stderr: "", code: 0, signal: null, killed: false, termination: "exit" };
      }) } }, registerTool: (factory: any, opts: any) => tools.set(opts.name ?? opts.names[0], factory), registerGatewayMethod: (name: string, fn: any) => methods.set(name, fn),
    registerService: (service: any) => services.push(service), on: (name: string, fn: any) => hooks.set(name, fn), logger: { warn: vi.fn(), info: vi.fn(), error: vi.fn() } };
  const store = new AccountStore(root, {}, policy);
  if (options.legacy !== false) await createVaultCredential(root, { clientId: "synthetic-client", tenant: "common", refreshToken: "synthetic-never-real", scopes: ["Mail.Read", "offline_access"] }, key);
  entry.register(api);
  const service = services[0];
  const start = async () => { await service?.start({ scheduler: schedule }); await hooks.get("gateway_start")?.({ port: 18789 }, { config, abortSignal: abort.signal }); };
  return { api, config, pluginConfig, root, key, store, schedule, service, services, hooks, tools, abort, start, directory: (fn: () => string) => directory = fn };
}
describe("registered trusted host singleton startup", () => {
  it("migrates on post-ready full-mode startup with no personal RPC, UI, admin or provider access; preserves backup and ordinary tool rights", async () => {
    const h = await host(); const original = await readFile(vaultCredentialPath(h.root)); const fetch = vi.fn(); vi.stubGlobal("fetch", fetch);
    h.directory(() => JSON.stringify({ profiles: [row("gateway-owner"), row("alias", "A"), row("A")] }));
    await h.start(); expect(await h.store.readAccount("A")).toBeUndefined(); await h.schedule.flush();
    const a = (await h.store.readAccount("A"))!; expect(a).toMatchObject({ connected: true, shared: false, unattended: false }); expect(a.policy).toEqual(h.pluginConfig.policy);
    expect(await h.store.automaticStatus()).toMatchObject({ profileId: "A", invalidated: false }); expect(await readFile(vaultCredentialPath(h.root))).toEqual(original); expect(fetch).not.toHaveBeenCalled();
    vi.stubGlobal("fetch", vi.fn(async (url: string) => new Response(JSON.stringify(url.includes("login.microsoftonline") ? { access_token: "synthetic-token", expires_in: 3600 } : { value: [{ id: "synthetic-mail" }] }))));
    const tool = h.tools.get("outlook_mail_read").create({ agentId: "main", sessionKey: "agent:main:startup", sessionId: "s", assertInvocationCurrent() {} });
    const result = await tool.execute("startup-read", { action: "list_messages" }, AbortSignal.timeout(90000));
    expect(result.details.ok).toBe(true); expect(JSON.stringify(result)).toContain("synthetic-mail");
    expect(JSON.stringify(result)).not.toContain("synthetic-never-real"); expect(JSON.stringify(result)).not.toContain("synthetic-token");
    expect(h.api.runtime.system.runCommandWithTimeout.mock.calls.some(([, opts]: any) => opts.signal instanceof AbortSignal)).toBe(true);
    expect(h.api.logger.warn).not.toHaveBeenCalled();
  });
  it.each(["2026.10.1", "2026.10.2"])("automatically assigns independent sole canonical owner on %s with aliases and arbitrary state root", async runtimeVersion => {
    const h = await host({ runtimeVersion, users: ["gateway-owner", "independent-owner"] });
    h.directory(() => JSON.stringify({ profiles: [row("independent-owner"), row("other-alias", "independent-owner"), row("gateway-owner")] }));
    const original = await readFile(vaultCredentialPath(h.root)); const fetch = vi.fn(); vi.stubGlobal("fetch", fetch);
    await h.start(); await h.schedule.flush();
    expect(await h.store.automaticStatus()).toMatchObject({ profileId: "independent-owner", invalidated: false });
    expect(await h.store.readAccount("independent-owner")).toMatchObject({ connected: true, shared: false, unattended: false });
    expect(await h.store.readAccount("A")).toBeUndefined(); expect(await readFile(vaultCredentialPath(h.root))).toEqual(original); expect(fetch).not.toHaveBeenCalled();
    await h.service.stop();
  });
  it.each(["2026.10.1", "2026.10.2"].flatMap(runtimeVersion => ["multiuser", "malformed", "foreign", "cancelled"].map(reason => ({ runtimeVersion, reason }))))("portable startup rejects $runtimeVersion $reason", async ({ runtimeVersion, reason }) => {
    const h = await host({ runtimeVersion, users: ["gateway-owner", "independent-owner"] });
    if (reason === "multiuser") h.directory(() => JSON.stringify({ profiles: [row("independent-owner"), row("another-person")] }));
    if (reason === "malformed") h.directory(() => '{"profiles":[{"id":"independent-owner"}]}');
    if (reason === "foreign") h.api.runtime.system.runCommandWithTimeout.mockResolvedValue({ stdout: JSON.stringify({ hostInstanceId: "another-server" }), stderr: "", code: 0, signal: null, killed: false, termination: "exit" });
    await h.start(); if (reason === "cancelled") h.abort.abort(); await h.schedule.flush();
    expect(await h.store.automaticStatus()).toBeUndefined(); expect(await h.store.readAccount("independent-owner")).toBeUndefined(); await h.service.stop();
  });
  it.each(["OPENCLAW_PROFILE", "OPENCLAW_GATEWAY_PASSWORD"])("startup closes authority when %s changes", async key => {
    const h = await host(); await h.start(); vi.stubEnv(key, "changed-synthetic-transport"); await h.schedule.flush(); expect(await h.store.automaticStatus()).toBeUndefined(); await h.service.stop();
  });
  it.each(["cli-metadata", "discovery", "tool-discovery", "setup"])("does not register migration in %s mode", async mode => { const h = await host({ mode }); expect(h.service).toBeUndefined(); expect(h.hooks.has("gateway_start")).toBe(false); });
  it.each([{ enabled: false }, { users: ["gateway-owner"] }, { users: ["gateway-owner", "A", "B"] }, { legacy: false }, { runtimeVersion: "unsupported" }])("does not provision an account in ineligible startup %j", async options => {
    const h = await host(options); await h.start(); await h.schedule.flush(); expect(await h.store.readAccount("A")).toBeUndefined(); expect(await h.store.automaticStatus()).toBeUndefined();
  });
  it("retries unknown directory only three times and never interprets malformed output as singleton", async () => { const h = await host(); h.directory(() => "{}"); await h.start(); await h.schedule.flush(); expect(h.api.runtime.system.runCommandWithTimeout.mock.calls.filter(([a]: any) => a.includes("users"))).toHaveLength(3); expect(await h.store.readAccount("A")).toBeUndefined(); });
  it("retired same-config registration cancels already queued callbacks", async () => { const h = await host(); await h.start(); entry.register(h.api); await h.schedule.flush(); expect(h.schedule.signal.aborted).toBe(true); expect(await h.store.readAccount("A")).toBeUndefined(); });
  it("initial service start never schedules before the published host-ready hook", async () => {
    const h = await host(); await h.service.start({ scheduler: h.schedule });
    expect(h.schedule.pending.size).toBe(0); await h.schedule.flush();
    expect(h.api.runtime.system.runCommandWithTimeout).not.toHaveBeenCalled();
    expect(await h.store.readAccount("A")).toBeUndefined();
  });
  it.each(["queued", "unknown-retry"])("retained same-definition %s stop/start resumes without replaying host-ready and permits ordinary tool use", async state => {
    const h = await host(); const fetch = vi.fn(); vi.stubGlobal("fetch", fetch);
    if (state === "unknown-retry") h.directory(() => "{}");
    await h.start(); if (state === "unknown-retry") await h.schedule.flushOne();
    expect(h.schedule.pending.size).toBe(1); const retiredCallback = [...h.schedule.pending.values()][0];
    await h.service.stop(); const restarted = scheduler();
    // An unrelated UI change must not invalidate the established host transport.
    Object.assign(h.config.gateway, { controlUi: { allowedOrigins: ["https://synthetic.invalid"] } });
    h.directory(() => JSON.stringify({ profiles: [row("gateway-owner"), row("A")] }));
    await h.service.start({ scheduler: restarted }); expect(restarted.pending.size).toBe(1);
    await retiredCallback(); expect(await h.store.readAccount("A")).toBeUndefined();
    await restarted.flush(); expect(fetch).not.toHaveBeenCalled();
    expect(await h.store.automaticStatus()).toMatchObject({ profileId: "A", invalidated: false });
    expect(await h.store.readAccount("A")).toMatchObject({ connected: true, shared: false, unattended: false });
    vi.stubGlobal("fetch", vi.fn(async (url: string) => new Response(JSON.stringify(url.includes("login.microsoftonline") ? { access_token: "synthetic-token", expires_in: 3600 } : { value: [{ id: "retained-restart-mail" }] }))));
    const tool = h.tools.get("outlook_mail_read").create({ agentId: "main", sessionKey: "agent:main:restart", sessionId: "s", assertInvocationCurrent() {} });
    const result = await tool.execute("retained-read", { action: "list_messages" }, AbortSignal.timeout(90000));
    expect(result.details.ok).toBe(true); expect(JSON.stringify(result)).toContain("retained-restart-mail");
    expect(JSON.stringify(result)).not.toContain("synthetic-token");
    await h.service.stop();
  });
  it.each(["aborted-while-stopped", "aborted-after-restart", "gateway-stop", "auth-change", "replacement", "unknown", "multiuser", "foreign"])("retained same-definition restart denies $0", async reason => {
    const h = await host(); await h.start(); await h.service.stop(); const restarted = scheduler();
    if (reason === "aborted-while-stopped") h.abort.abort();
    if (reason === "gateway-stop") h.hooks.get("gateway_stop")();
    if (reason === "auth-change") h.config.gateway.auth.token = "changed-synthetic";
    if (reason === "replacement") entry.register(h.api);
    if (reason === "unknown") h.directory(() => "{}");
    if (reason === "multiuser") h.directory(() => JSON.stringify({ profiles: [row("gateway-owner"), row("A"), row("B")] }));
    if (reason === "foreign") h.api.runtime.system.runCommandWithTimeout.mockResolvedValue({ stdout: JSON.stringify({ hostInstanceId: "foreign" }), stderr: "", code: 0, signal: null, killed: false, termination: "exit" });
    await h.service.start({ scheduler: restarted });
    if (["aborted-while-stopped", "gateway-stop", "auth-change", "replacement"].includes(reason)) expect(restarted.pending.size).toBe(0);
    if (reason === "aborted-after-restart") { expect(restarted.pending.size).toBe(1); h.abort.abort(); expect(restarted.signal.aborted).toBe(true); }
    await restarted.flush(); expect(await h.store.readAccount("A")).toBeUndefined(); expect(await h.store.automaticStatus()).toBeUndefined();
    await h.service.stop();
  });
  it("service stop joins running directory work and denies publication", async () => {
    const h = await host(); let release!: () => void, entered!: () => void; const enteredPromise = new Promise<void>(r => entered = r); const blocked = new Promise<void>(r => release = r);
    const runner = h.api.runtime.system.runCommandWithTimeout.getMockImplementation(); h.api.runtime.system.runCommandWithTimeout.mockImplementation(async (...args: any[]) => { if (args[0].includes("users")) { entered(); await blocked; } return runner(...args); });
    await h.start(); const work = h.schedule.flush(); await enteredPromise; let settled = false; const stopped = h.service.stop().then(() => settled = true); await Promise.resolve(); expect(settled).toBe(false); release(); await Promise.all([work, stopped]); expect(await h.store.readAccount("A")).toBeUndefined();
  });
  it.each([2, 3, 4].flatMap(checkpoint => ["drain", "config"].map(reason => ({ reason, checkpoint }))))("denies a positive migration at vault/publication checkpoint $checkpoint on $reason", async ({ reason, checkpoint }) => {
    const h = await host(); const original = await readFile(vaultCredentialPath(h.root)); let reads = 0; h.directory(() => { if (++reads === checkpoint) { if (reason === "drain") h.abort.abort(); else h.pluginConfig.enabled = false; } return JSON.stringify({ profiles: [row("gateway-owner"), row("A")] }); });
    await h.start(); await h.schedule.flush(); expect(await h.store.automaticStatus()).toBeUndefined(); expect((await h.store.readAccount("A"))?.connected).not.toBe(true); expect(await readFile(vaultCredentialPath(h.root))).toEqual(original);
  });
  it.each(["config", "drain"])("records topology mismatch despite %s cancellation, and restart cannot resurrect the tombstone", async reason => {
    const h = await host(); let reads = 0; h.directory(() => { if (++reads === 2) { if (reason === "drain") h.abort.abort(); else h.pluginConfig.enabled = false; return JSON.stringify({ profiles: [row("gateway-owner"), row("A"), row("B")] }); } return JSON.stringify({ profiles: [row("gateway-owner"), row("A")] }); });
    await h.start(); await h.schedule.flush(); expect(await h.store.automaticStatus()).toMatchObject({ invalidated: true }); expect((await h.store.readAccount("A"))?.connected).toBe(false);
    h.pluginConfig.enabled = true; h.directory(() => JSON.stringify({ profiles: [row("gateway-owner"), row("A")] }));
    const restarted = scheduler(); entry.register(h.api);
    // Registration created a fresh owned generation; no previous host signal is reused.
    // The API's services collector retains the newly registered lifetime.
    await h.services.at(-1).start({ scheduler: restarted });
    await h.hooks.get("gateway_start")({ port: 18789 }, { config: h.config, abortSignal: new AbortController().signal });
    await restarted.flush();
    expect(await h.store.automaticStatus()).toMatchObject({ invalidated: true }); expect((await h.store.readAccount("A"))?.connected).toBe(false);
  });
  it.each(["in_flight", "quarantined", "malformed"])("startup retains source %s refresh markers without laundering credentials", async state => {
    const h = await host(); const original = await readFile(vaultCredentialPath(h.root));
    const record = await readVaultCredential(h.root, h.key);
    const payload = { format: "openclaw-microsoft-graph-refresh-quarantine", version: 2, generation: record.envelope.generation, keyId: record.envelope.keyId, vaultDigest: record.digest, state, createdAt: "2026-10-01T00:00:00.000Z", updatedAt: "2026-10-01T00:00:00.000Z" };
    const marker = state === "malformed" ? "not-a-marker" : JSON.stringify({ ...payload, mac: createHmac("sha256", Buffer.from(h.key, "base64url")).update(JSON.stringify(payload)).digest("base64url") }) + "\n";
    await writeFile(vaultQuarantinePath(h.root), marker, { mode: 0o600 });
    const fetch = vi.fn(); vi.stubGlobal("fetch", fetch); await h.start(); await h.schedule.flush();
    expect(await h.store.automaticStatus()).toBeUndefined(); expect((await h.store.readAccount("A"))?.connected).not.toBe(true);
    expect(await readFile(vaultCredentialPath(h.root))).toEqual(original); expect(await readFile(vaultQuarantinePath(h.root), "utf8")).toBe(marker); expect(fetch).not.toHaveBeenCalled();
  });
  it("startup intersects a stricter profile ceiling and a clean restart never overwrites the migrated vault", async () => {
    const h = await host({ strictCeiling: true }); await h.start(); await h.schedule.flush();
    const a = (await h.store.readAccount("A"))!; expect(a.policy).toEqual(blankAccountPolicy()); expect(a.shared).toBe(false); expect(a.unattended).toBe(false);
    const original = await readFile(vaultCredentialPath(h.store.accountPath("A", a.generation)));
    await h.service.stop(); const restarted = scheduler(); entry.register(h.api); await h.services.at(-1).start({ scheduler: restarted });
    await h.hooks.get("gateway_start")({ port: 18789 }, { config: h.config, abortSignal: new AbortController().signal }); await restarted.flush();
    expect(await h.store.readAccount("A")).toEqual(a); expect(await readFile(vaultCredentialPath(h.store.accountPath("A", a.generation)))).toEqual(original);
    expect(await h.store.automaticStatus()).toMatchObject({ invalidated: false });
  });
  it.each(["foreign", "truncated"])("startup denies a %s directory result without guessing an account", async state => {
    const h = await host(); const runner = h.api.runtime.system.runCommandWithTimeout.getMockImplementation();
    h.api.runtime.system.runCommandWithTimeout.mockImplementation(async (...args: any[]) => { const output = await runner(...args); return state === "foreign" ? { ...output, stdout: JSON.stringify({ hostInstanceId: "foreign" }) } : { ...output, stdoutTruncatedBytes: 1 }; });
    await h.start(); await h.schedule.flush(); expect(await h.store.readAccount("A")).toBeUndefined(); expect(await h.store.automaticStatus()).toBeUndefined();
  });
  it("does not promote an explicitly provisioned/configured account", async () => { const h = await host(); await h.store.provision("A", () => {}); await h.start(); await h.schedule.flush(); expect(await h.store.automaticStatus()).toBeUndefined(); expect((await h.store.readAccount("A"))?.connected).toBe(false); });
});
