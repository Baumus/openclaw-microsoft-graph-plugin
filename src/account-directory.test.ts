import { describe, expect, it, vi, afterEach } from "vitest";
import { createAccountDirectory, parseDirectorySnapshot } from "./account-directory.js";
const row = (id: string, mergedInto: string | null = null) => ({ id, mergedInto, displayName: id, avatarMime: null, createdAt: 1, updatedAt: 1, emails: [] as string[], githubIdentity: null, hasAvatar: false, roleSource: "default" });
const raw = (rows = [row("gateway-owner"), row("A")]) => JSON.stringify({ profiles: rows });
afterEach(() => { vi.unstubAllEnvs(); vi.unstubAllGlobals(); });
describe("supported full-host directory oracle", () => {
  it("excludes only exact shared owner; handles canonical merged aliases and complete roster", () => {
    expect(parseDirectorySnapshot(raw()).soleProfileId).toBe("A");
    expect(parseDirectorySnapshot(raw([row("A"), row("alias", "A")])).soleProfileId).toBe("A");
    expect(parseDirectorySnapshot(raw([row("gateway-owner"), row("A"), row("B")])).soleProfileId).toBeUndefined();
    expect(parseDirectorySnapshot(raw([row("Owner"), row("A")])).soleProfileId).toBeUndefined();
    expect(parseDirectorySnapshot(raw([row("gateway-owner")])).soleProfileId).toBeUndefined();
  });
  it.each(["invalid", '{}', '{"profiles":[],"cursor":"next"}', raw([row("A"), row("A")]), raw([row("alias", "missing")]), raw([row("a", "b"), row("b", "a")]), JSON.stringify({ profiles: [{ id: "A", mergedInto: null }] }), JSON.stringify({ profiles: [{ ...row("A"), futureField: true }] })])("denies unknown, partial, malformed or inconsistent roster %s", value => {
    expect(() => parseDirectorySnapshot(value)).toThrow("graph_single_user_directory_unavailable");
  });
  it("fingerprints identities and mutations without storing names/emails", () => {
    const a = parseDirectorySnapshot(raw()); expect(a.digest).toMatch(/^[a-f0-9]{64}$/);
    expect(parseDirectorySnapshot(raw([row("A"), row("gateway-owner")])).digest).toBe(a.digest);
    expect(parseDirectorySnapshot(raw([row("gateway-owner"), { ...row("A"), emails: ["synthetic@example.invalid"] }])).digest).not.toBe(a.digest);
    expect(JSON.stringify(a)).not.toContain("example");
  });
  function api(outputs?: unknown[]) {
    const config = { gateway: { mode: "local", port: 18789, bind: "loopback", auth: { mode: "token", token: "synthetic-host-auth" } } };
    const replies = outputs ?? [{ hostInstanceId: "same-host" }, JSON.parse(raw()), { hostInstanceId: "same-host" }];
    const request = vi.fn(async () => replies.shift());
    const stop = vi.fn(); const stopAndWait = vi.fn(async () => {});
    let options: any; const Client = vi.fn(function(this: any, opts: any) { options = opts; Object.assign(this, { start: () => opts.onHelloOk({}), request, stop, stopAndWait }); });
    const host = { runtime: { version: "2026.10.1", config: { current: () => config } } };
    const read = (signal?: AbortSignal) => createAccountDirectory(host as any, "same-host", Client as any).read(signal);
    return { config, request, Client, host, read, stop, stopAndWait, options: () => options };
  }
  it("uses one fresh public read-only CLI client and three bounded RPCs, with no persistent state", async () => {
    vi.stubEnv("OPENCLAW_GATEWAY_URL", "wss://foreign.invalid"); vi.stubEnv("OPENCLAW_GATEWAY_TOKEN", undefined); vi.stubEnv("OPENCLAW_GATEWAY_PORT", undefined);
    const r = api(); expect((await r.read()).soleProfileId).toBe("A");
    expect(r.Client).toHaveBeenCalledTimes(1); expect(r.request).toHaveBeenCalledTimes(3);
    expect(r.request.mock.calls.map((args: any) => args[0])).toEqual(["microsoft-graph.directory-host", "users.list", "microsoft-graph.directory-host"]);
    expect(r.options()).toMatchObject({ url: "ws://127.0.0.1:18789", token: "synthetic-host-auth", clientName: "cli", mode: "cli", role: "operator", scopes: ["operator.read"], deviceIdentity: null, sharedStateMode: "read-only", env: {} });
    expect(r.stopAndWait).toHaveBeenCalledWith({ timeoutMs: 1000 });
    expect(r.request.mock.calls.every((args: any) => args[2].timeoutMs === 5000)).toBe(true);
  });
  it.each(["2026.10.1", "2026.10.2"])("supports %s independent IDs, aliases, defaults and local listeners", async version => {
    vi.stubEnv("OPENCLAW_GATEWAY_TOKEN", undefined); vi.stubEnv("OPENCLAW_GATEWAY_PASSWORD", undefined); vi.stubEnv("OPENCLAW_GATEWAY_PORT", undefined); vi.stubEnv("OPENCLAW_PROFILE", undefined);
    for (const shape of [
      { bind: "loopback", port: 24567, auth: { token: "synthetic-config-first" } },
      { bind: "lan", auth: { mode: "password", password: "synthetic-password" } },
      { bind: "auto", auth: { password: "synthetic-inferred-password" } },
      { auth: { mode: "token", token: "synthetic-resolved-runtime-value" } },
    ]) {
      const r = api([{ hostInstanceId: "same-host" }, { profiles: [row("gateway-owner"), row("person-independent"), row("merged-alias", "person-independent")] }, { hostInstanceId: "same-host" }]);
      r.host.runtime.version = version; r.config.gateway = shape as any;
      expect((await r.read()).soleProfileId).toBe("person-independent");
      expect(r.options().url).toBe(`ws://127.0.0.1:${shape.port ?? 18789}`);
      expect(r.options().password ?? r.options().token).toBe(shape.auth.password ?? shape.auth.token);
      expect(Object.keys(r.options()).filter(k => ["token", "password"].includes(k))).toHaveLength(1);
    }
  });
  it("uses public config-first auth and named-profile/env port without inheriting host paths", async () => {
    vi.stubEnv("OPENCLAW_GATEWAY_TOKEN", "synthetic-environment-token"); vi.stubEnv("OPENCLAW_GATEWAY_PASSWORD", undefined); vi.stubEnv("OPENCLAW_GATEWAY_PORT", "127.0.0.1:29123"); vi.stubEnv("OPENCLAW_PROFILE", "different-installation");
    vi.stubEnv("OPENCLAW_STATE_DIR", "C:\\Users\\Another\\OpenClaw"); vi.stubEnv("OPENCLAW_CONFIG_PATH", "/srv/another-install/config.json");
    const r = api(); await r.read(); expect(r.options().url).toBe("ws://127.0.0.1:29123"); expect(r.options().token).toBe("synthetic-host-auth"); expect(r.options().env).toEqual({});
    vi.stubEnv("OPENCLAW_GATEWAY_PORT", undefined); const named = api(); delete (named.config.gateway as any).port;
    await named.read(); const { resolveGatewayPort } = await import("openclaw/plugin-sdk/core"); expect(named.options().url).toBe(`ws://127.0.0.1:${resolveGatewayPort(named.config as any, process.env)}`); expect(named.options().url).not.toBe("ws://127.0.0.1:18789");
  });
  it.each(["token", "password"])("does not substitute env for unresolved active %s SecretRef", async mode => {
    vi.stubEnv("OPENCLAW_GATEWAY_TOKEN", "synthetic-env-token"); vi.stubEnv("OPENCLAW_GATEWAY_PASSWORD", "synthetic-env-password");
    const r = api(); r.config.gateway.auth = { mode, [mode]: { source: "env", provider: "default", id: "SYNTHETIC_UNRESOLVED" } } as any;
    await expect(r.read()).rejects.toThrow("graph_single_user_directory_unavailable"); expect(r.Client).not.toHaveBeenCalled();
  });
  it.each(["2026.10.0", "2026.10.3", "2026.10.2-beta.1", "unknown"])("rejects unverified runtime %s", async version => {
    const r = api(); r.host.runtime.version = version; await expect(r.read()).rejects.toThrow("graph_single_user_directory_unavailable"); expect(r.Client).not.toHaveBeenCalled();
  });
  it.each(["OPENCLAW_PROFILE", "OPENCLAW_GATEWAY_PASSWORD", "OPENCLAW_GATEWAY_PORT", "OPENCLAW_GATEWAY_TOKEN"])("pins changed transport environment %s", async key => {
    const r = api(); r.request.mockImplementation(async () => { vi.stubEnv(key, "changed-synthetic"); return { hostInstanceId: "same-host" }; });
    await expect(r.read()).rejects.toThrow("graph_single_user_directory_unavailable"); expect(r.stopAndWait).toHaveBeenCalled();
  });
  it("cancels before/while connecting and joins cleanup", async () => {
    const abort = new AbortController(); abort.abort(); const r = api(); await expect(r.read(abort.signal)).rejects.toThrow("graph_single_user_directory_unavailable"); expect(r.Client).not.toHaveBeenCalled();
    const waiting = api(); const pending = new AbortController(); waiting.Client.mockImplementation(function(this: any, options: any) { Object.assign(this, { start: () => pending.abort(), stopAndWait: waiting.stopAndWait }); });
    await expect(waiting.read(pending.signal)).rejects.toThrow("graph_single_user_directory_unavailable"); expect(waiting.stopAndWait).toHaveBeenCalled();
  });
  it("completes bounded safety after-proof if drain follows a fully observed roster", async () => {
    const r = api(), abort = new AbortController(); let calls = 0;
    r.request.mockImplementation(async () => { if (++calls === 2) { abort.abort(); return JSON.parse(raw([row("gateway-owner"), row("A"), row("B")])); } return { hostInstanceId: "same-host" }; });
    expect((await r.read(abort.signal)).soleProfileId).toBeUndefined(); expect(calls).toBe(3);
    expect((r.request.mock.calls as any)[0][2].signal).toBe(abort.signal); expect((r.request.mock.calls as any)[2][2].signal).toBeUndefined();
  });
  it("rejects foreign instance, oversize, timeout, disconnect/reconnect and config changes without leaking", async () => {
    const foreign = api([{ hostInstanceId: "foreign" }]); await expect(foreign.read()).rejects.toThrow("graph_single_user_directory_unavailable"); expect(foreign.request).toHaveBeenCalledTimes(1);
    for (const value of [undefined, { hostInstanceId: "same-host", extra: true }, { hostInstanceId: "x".repeat(1024 * 1024) }]) { const r = api([value]); await expect(r.read()).rejects.toThrow("graph_single_user_directory_unavailable"); }
    const timeout = api(); timeout.request.mockRejectedValue(new Error("sensitive-sdk-error")); await expect(timeout.read()).rejects.toThrow(/^graph_single_user_directory_unavailable$/);
    for (const mutate of [(r: any) => r.config.gateway.port++, (r: any) => r.options().onClose(), (r: any) => r.options().onHelloOk({})]) { const r = api(); r.request.mockImplementation(async () => { mutate(r); return { hostInstanceId: "same-host" }; }); await expect(r.read()).rejects.toThrow("graph_single_user_directory_unavailable"); }
    for (const change of [{ mode: "remote" }, { tls: { enabled: true } }, { auth: { mode: "none" } }, { auth: { mode: "token", token: { source: "env" } } }]) { const r = api(); Object.assign(r.config.gateway, change); await expect(r.read()).rejects.toThrow("graph_single_user_directory_unavailable"); expect(r.Client).not.toHaveBeenCalled(); }
  });
});

it("exercises exported SDK client against an isolated synthetic Gateway", async () => {
  const { createRequire } = await import("node:module");
  const { mkdtemp, writeFile, readdir, rm } = await import("node:fs/promises");
  const { join } = await import("node:path"); const { tmpdir } = await import("node:os");
  const { fileURLToPath } = await import("node:url");
  const require = createRequire(import.meta.resolve("openclaw/cli-entry"));
  const { WebSocketServer } = require("ws");
  const runnerSpecifier = "openclaw/plugin-sdk/process-runtime";
  const { runCommandWithTimeout } = await import(runnerSpecifier);
  const root = await mkdtemp(join(tmpdir(), "graph-directory-cli-test-"));
  const server = new WebSocketServer({ host: "127.0.0.1", port: 0 });
  await new Promise<void>(resolve => server.once("listening", resolve));
  const port = server.address().port;
  const config = { gateway: { mode: "local", port, bind: "loopback", auth: { mode: "token", token: "synthetic-only-shared-auth" } } };
  const configFile = join(root, "openclaw.json"); await writeFile(configFile, JSON.stringify(config), { mode: 0o600 });
  let authenticated = 0; const called: string[] = [];
  let getHostProof = () => ({ hostInstanceId: "synthetic-instance" });
  server.on("connection", (socket: any) => {
    socket.send(JSON.stringify({ type: "event", event: "connect.challenge", payload: { nonce: "synthetic-challenge", ts: Date.now() } }));
    socket.on("message", (bytes: Buffer) => {
      const request = JSON.parse(bytes.toString());
      if (request.method === "connect") {
        expect(request.params.client.id).toBe("cli"); expect(request.params.client.mode).toBe("cli"); expect(request.params.role).toBe("operator"); expect(request.params.scopes).toEqual(["operator.read"]); expect(request.params.device).toBeUndefined();
        if (request.params.auth?.token !== "synthetic-only-shared-auth") { socket.close(); return; }
        authenticated++;
        socket.send(JSON.stringify({ type: "res", id: request.id, ok: true, payload: { type: "hello-ok", protocol: request.params.maxProtocol, server: { version: "2026.10.1", connId: "synthetic" }, features: { methods: ["users.list", "microsoft-graph.directory-host"], events: [] }, snapshot: {}, policy: { maxPayload: 1048576, maxBufferedBytes: 1048576, tickIntervalMs: 30000 } } }));
      } else {
        called.push(request.method);
        socket.send(JSON.stringify({ type: "res", id: request.id, ok: true, payload: request.method === "users.list" ? JSON.parse(raw()) : getHostProof() }));
      }
    });
  });
  try {
    vi.stubEnv("OPENCLAW_CONFIG_PATH", configFile); vi.stubEnv("OPENCLAW_STATE_DIR", root);
    vi.stubEnv("OPENCLAW_GATEWAY_PORT", undefined); vi.stubEnv("OPENCLAW_GATEWAY_TOKEN", undefined);
    const api = { runtime: { version: "2026.10.1", config: { current: () => config }, system: { runCommandWithTimeout } } };
    const snapshot = await createAccountDirectory(api as any, "synthetic-instance").read();
    expect(snapshot.soleProfileId).toBe("A"); expect(authenticated).toBe(1);
    expect(called).toEqual(["microsoft-graph.directory-host", "users.list", "microsoft-graph.directory-host"]);
    expect(await readdir(root)).toEqual(["openclaw.json"]); // No device/token/config side effects.
    expect(fileURLToPath(import.meta.resolve("openclaw/cli-entry"))).toMatch(/openclaw\.mjs$/);

    // Prove default tool timing with genuine SDK directory transport and only
    // synthetic OAuth/Graph provider responses. No production state/provider.
    const { default: entry } = await import("./index.js");
    const { createVaultCredential } = await import("./credential-vault.js");
    const { graphPolicyFixture } = await import("./fixtures/graph-access-policy.js");
    const { randomBytes } = await import("node:crypto");
    const key = randomBytes(32).toString("base64url");
    const methods = new Map<string, any>(); const tools = new Map<string, any>();
    const session = { sessionId: "synthetic-session", lifecycleRevision: "1" };
    const host = {
      ...api, pluginConfig: { enabled: true, credentialVaultKey: key, policy: graphPolicyFixture() },
      runtime: { ...api.runtime, state: { resolveStateDir: () => root }, agent: { session: { getSessionEntry: () => session } } },
      registerTool: (factory: any, options: any) => tools.set(options.name ?? options.names[0], factory),
      registerGatewayMethod: (name: string, handler: any) => methods.set(name, handler), on: () => {},
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    };
    entry.register(host as any);
    getHostProof = () => { let value: any; methods.get("microsoft-graph.directory-host")({ params: {}, respond: (_ok: boolean, result: unknown) => value = result }); return value; };
    await createVaultCredential(root, { clientId: "synthetic-client", tenant: "common", refreshToken: "synthetic-never-production", scopes: ["Mail.Read", "offline_access"] }, key);
    let migration: any;
    await methods.get("microsoft-graph.accounts.automatic-migration")({ params: {}, client: { authenticatedUserProfile: { profileId: "A" } }, hasCurrentClientAuthority: () => true, respond: (ok: boolean, value: unknown, error: unknown) => migration = { ok, value, error } });
    expect(migration.ok).toBe(true);
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.startsWith("https://login.microsoftonline.com/")) return new Response(JSON.stringify({ access_token: "synthetic-default-tool", expires_in: 3600 }));
      if (url.startsWith("https://graph.microsoft.com/")) return new Response(JSON.stringify({ value: [{ id: "synthetic-mail", subject: "synthetic-only" }] }));
      throw new Error("unexpected_provider");
    }));
    const context = { agentId: "main", sessionKey: "agent:main:synthetic", sessionId: session.sessionId, assertInvocationCurrent() {} };
    const tool = tools.get("outlook_mail_read").create(context);
    const result = await tool.execute("synthetic-default", { action: "list_messages" }, AbortSignal.timeout(90000));
    expect(result.details.ok).toBe(true); expect(JSON.stringify(result)).not.toContain("synthetic-never-production");
    expect(host.logger.warn).not.toHaveBeenCalled();
  } finally {
    for (const client of server.clients) client.terminate();
    await new Promise<void>(resolve => server.close(() => resolve()));
    await rm(root, { recursive: true, force: true });
  }
}, 180000);
