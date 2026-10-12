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
    const run = vi.fn(async () => ({ stdout: JSON.stringify(replies.shift()), stderr: "", code: 0, killed: false, signal: null, termination: "exit" }));
    return { config, run, host: { runtime: { version: "2026.10.1", config: { current: () => config }, system: { runCommandWithTimeout: run } } } };
  }
  it("pins public CLI export, this exact instance and loopback auth via env only", async () => {
    vi.stubEnv("OPENCLAW_GATEWAY_URL", "wss://foreign.invalid"); vi.stubEnv("OPENCLAW_GATEWAY_TOKEN", undefined); vi.stubEnv("OPENCLAW_GATEWAY_PORT", undefined);
    const r = api(); const result = await createAccountDirectory(r.host as any, "same-host").read(); expect(result.soleProfileId).toBe("A");
    expect(r.run).toHaveBeenCalledTimes(3);
    for (const [argv, options] of r.run.mock.calls as unknown as [string[], any][]) {
      expect(argv[0]).toBe(process.execPath); expect(argv[1]).toMatch(/openclaw\.mjs$/);
      expect(JSON.stringify(argv)).not.toContain("synthetic-host-auth");
      expect(options.baseEnv.OPENCLAW_GATEWAY_URL).toBe("ws://127.0.0.1:18789");
      expect(options.baseEnv.OPENCLAW_GATEWAY_TOKEN).toBe("synthetic-host-auth");
      expect(options.baseEnv.NODE_OPTIONS).toBeUndefined();
      expect(options.baseEnv.OPENCLAW_STATE_DIR).toMatch(/msgraph-users-read-/);
      expect(options.baseEnv.OPENCLAW_CONFIG_PATH).toContain(options.baseEnv.OPENCLAW_STATE_DIR);
      expect(options.baseEnv.OPENCLAW_CONFIG_PATH.endsWith("absent-config.json")).toBe(true);
      expect(options.requireProcessTreeExtinction).toBe(true);
    }
  });
  it("rejects foreign instance, truncation, timeout, runtime/schema and configuration changes", async () => {
    vi.stubEnv("OPENCLAW_GATEWAY_TOKEN", undefined); vi.stubEnv("OPENCLAW_GATEWAY_PORT", undefined);
    const foreign = api([{ hostInstanceId: "foreign" }]); await expect(createAccountDirectory(foreign.host as any, "same-host").read()).rejects.toThrow("graph_single_user_directory_unavailable"); expect(foreign.run).toHaveBeenCalledTimes(1);
    const truncated = api(); truncated.run.mockImplementation(async () => ({ stdout: raw(), stdoutTruncatedBytes: 1, stderr: "sensitive-secret-must-not-leak", code: 0, killed: false, signal: null, termination: "exit" }));
    await expect(createAccountDirectory(truncated.host as any, "same-host").read()).rejects.toThrow("graph_single_user_directory_unavailable");
    const changed = api(); changed.run.mockImplementation(async () => { changed.config.gateway.port++; return { stdout: JSON.stringify({ hostInstanceId: "same-host" }), stderr: "", code: 0, killed: false, signal: null, termination: "exit" }; });
    await expect(createAccountDirectory(changed.host as any, "same-host").read()).rejects.toThrow("graph_single_user_directory_unavailable");
    for (const change of [{ mode: "remote" }, { bind: "lan" }, { tls: { enabled: true } }, { auth: { mode: "none" } }, { auth: { mode: "token", token: { source: "env" } } }]) {
      const r = api(); Object.assign(r.config.gateway, change); await expect(createAccountDirectory(r.host as any, "same-host").read()).rejects.toThrow("graph_single_user_directory_unavailable"); expect(r.run).not.toHaveBeenCalled();
    }
  });
});

it("exercises exported CLI and public process runner against an isolated synthetic Gateway", async () => {
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
    expect(snapshot.soleProfileId).toBe("A"); expect(authenticated).toBe(3);
    expect(called).toEqual(["microsoft-graph.directory-host", "users.list", "microsoft-graph.directory-host"]);
    expect(await readdir(root)).toEqual(["openclaw.json"]); // No device/token/config side effects.
    expect(fileURLToPath(import.meta.resolve("openclaw/cli-entry"))).toMatch(/openclaw\.mjs$/);

    // Prove default tool timing with genuine CLI directory transport and only
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
