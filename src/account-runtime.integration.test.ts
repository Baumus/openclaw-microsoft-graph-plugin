import { writeFileSync as requireSyntheticWrite } from "node:fs";
import { mkdtemp, rm, writeFile, readFile, utimes } from "node:fs/promises";
import { randomBytes, createHmac } from "node:crypto";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import entry, { NativeApprovalSnapshotStore, WarningApprovalTrustStore } from "./index.js";
import { accountScope } from "./account-scope.js";
import { AccountStore } from "./accounts.js";
import { acquireVaultLock, createVaultCredential, readVaultCredential, vaultCredentialPath, vaultQuarantinePath } from "./credential-vault.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
import { authorizeOperation, authorizeRoot, type GraphPolicy } from "./policy.js";
const dirs: string[] = [];
afterEach(async () => { vi.unstubAllGlobals(); for (const dir of dirs.splice(0)) await rm(dir, { recursive: true, force: true }); });
async function runtime(worker = false, accountPolicyCeilings: Record<string, GraphPolicy> = {}, legacyPolicy: GraphPolicy = graphPolicyFixture(), directory?: () => string) {
  const root = await mkdtemp(join(tmpdir(), "graph-host-test-")); dirs.push(root); const key = randomBytes(32).toString("base64url");
  const policy = structuredClone(legacyPolicy);
  if (worker) for (const service of ["mail", "todo", "calendar"] as const) policy.services[service].agents["graph-a"] = { operations: ["read"] };
  const methods = new Map<string, { handler: (ctx: any) => Promise<void>; options: any }>(); const tools = new Map<string, any>(); const hooks: Record<string, any> = {};
  const entries = new Map<string, any>();
  const hostConfig = { gateway: { mode: "local", port: 18789, bind: "loopback", auth: { mode: "token", token: "synthetic-gateway-auth" } } };
  const api = {
    pluginConfig: { enabled: true, credentialVaultKey: key, policy, accountWorkers: worker ? { "graph-a": "A" } : {}, accountPolicyCeilings },
    runtime: { ...(directory ? { version: "2026.10.1", config: { current: () => hostConfig }, system: { runCommandWithTimeout: vi.fn(async (argv: string[]) => {
      let stdout: string;
      if (argv.includes("users")) stdout = directory();
      else { let proof: unknown; await methods.get("microsoft-graph.directory-host")!.handler({ params: {}, respond: (_ok: boolean, value: unknown) => { proof = value; } }); stdout = JSON.stringify(proof); }
      return { stdout, stderr: "", code: 0, signal: null, killed: false, termination: "exit" };
    }) } } : {}), state: { resolveStateDir: () => root }, agent: { session: { getSessionEntry: ({ sessionKey }: any) => entries.get(sessionKey) } } },
    registerTool: (factory: any, options: any) => tools.set(options.name ?? options.names[0], factory),
    registerGatewayMethod: (name: string, handler: any, options: any) => methods.set(name, { handler, options }),
    on: (name: string, handler: any) => hooks[name] = handler,
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  };
  entry.register(api as any);
  const rpc = async (actor: string | undefined, name: string, params: Record<string, unknown> = {}, current = () => true, authority?: any, signal?: AbortSignal) => {
    let reply: any;
    const ctx = { params, signal, client: actor ? { authenticatedUserProfile: { profileId: actor } } : {}, hasCurrentClientAuthority: current,
      sessionAccessAuthority: authority, withSessionTurnAuthority: async (_target: unknown, consume: any) => consume(entries.get(params.sessionKey as string)), respond: (ok: boolean, value: unknown, error: any) => { reply = { ok, value, error }; } };
    await methods.get(`microsoft-graph.accounts.${name}`)!.handler(ctx);
    if (!reply.ok) throw new Error(reply.error.message); return reply.value;
  };
  const store = new AccountStore(root, worker ? { "graph-a": "A" } : {}, policy, accountPolicyCeilings);
  const provision = async (id: string, shared = true, unattended = true) => {
    const a = await rpc(id, "provision"); await rpc(id, "configure", { expectedRevision: a.revision, policy, shared, unattended });
    await createVaultCredential(store.accountPath(id, a.generation), { clientId: "synthetic-client", tenant: "common", refreshToken: `refresh-${id}`, scopes: ["Mail.Read", "Mail.ReadWrite", "Mail.Send", "Tasks.Read", "Calendars.Read", "Files.Read", "offline_access"] }, key);
    await store.markConnected(id, a.generation, () => {});
  };
  const context = (id: string, agentId = "main", visibility = "shared") => {
    const sessionKey = `agent:${agentId}:${id}`; const e = { sessionId: id, lifecycleRevision: "1", visibility, createdActor: { type: "human", source: "profile", id: "A" } }; entries.set(sessionKey, e);
    let active = true;
    const ctx = { agentId, sessionKey, sessionId: id, assertInvocationCurrent: () => { if (!active) throw new Error("host_invocation_revoked"); } };
    return { ctx, idle: () => active = false, resume: () => active = true, revoke: () => active = false, authority: { target: { ...ctx, lifecycleRevision: "1" }, assertCurrent: () => {} } };
  };
  const bind = async (actor: string, profileId: string, c: ReturnType<typeof context>, expectedRevision: number | null = null) => rpc(actor, "bind", { agentId: c.ctx.agentId, sessionKey: c.ctx.sessionKey, sessionId: c.ctx.sessionId, profileId, expectedRevision }, () => true, c.authority);
  return { root, key, policy, methods, tools, hooks, rpc, store, provision, context, bind, entries };
}
describe("unmocked public V2 Graph host boundary", () => {
  it("requires authenticated profile on all personal methods, never registers legacy global credential APIs", async () => {
    const r = await runtime();
    for (const [name, method] of r.methods) if (name.startsWith("microsoft-graph.accounts.")) expect(method.options.profileAccess).toBe("required");
    expect(r.methods.has("microsoft-graph.credentials.status")).toBe(false);
    await expect(r.rpc(undefined, "status")).rejects.toThrow("graph_authenticated_profile_required");
    await expect(r.rpc("A", "provision", {}, () => false)).rejects.toThrow("graph_authenticated_profile_required");
    const a = await r.rpc("A", "provision"); expect(a.profileId).toBe("A"); expect((await r.rpc("B", "status")).account).toBeNull();
    await expect(r.rpc("B", "status", { profileId: "A" })).rejects.toThrow("invalid_rpc_parameters");
  });
  it("uses the bound account for parallel real registered me-tools, not request text/sender", async () => {
    const r = await runtime(); await r.provision("A"); await r.provision("B");
    const A = r.context("A-chat"), B = r.context("B-chat"); await r.bind("A", "A", A); await r.bind("B", "B", B);
    const calls: string[] = [];
    vi.stubGlobal("fetch", vi.fn(async (url: string, init: RequestInit) => {
      if (url.includes("login.microsoftonline.com")) { const id = String((init.body as URLSearchParams).get("refresh_token")).slice(-1); return new Response(JSON.stringify({ access_token: `token-${id}`, expires_in: 3600 }), { status: 200 }); }
      const auth = (init.headers as Record<string, string>).authorization; calls.push(auth);
      return new Response(JSON.stringify({ value: [{ id: auth, subject: auth }] }), { status: 200 });
    }));
    const factory = r.tools.get("outlook_mail_read"); expect(factory.contextVersion).toBe(2);
    const tools = [factory.create(A.ctx), factory.create(B.ctx)];
    const result = await Promise.all(tools.map((tool, i) => tool.execute(`call-${i}`, { action: "list_messages" })));
    expect(result.every(x => x.details.ok)).toBe(true); expect(new Set(calls)).toEqual(new Set(["Bearer token-A", "Bearer token-B"]));
    A.idle();
    await r.bind("B", "A", A, 1); // Explicit shared participant still selects A; no automatic B credential.
    A.resume();
    await tools[0].execute("steering", { action: "list_messages" }); expect(calls.at(-1)).toBe("Bearer token-A");
    expect(JSON.stringify(result)).not.toContain("refresh-");
  });
  it("rejects missing binding/profile/worker conflict before credential exchange or Graph", async () => {
    const r = await runtime(true); await r.provision("A"); const c = r.context("unbound"); const fetch = vi.fn(); vi.stubGlobal("fetch", fetch);
    const t = r.tools.get("outlook_mail_read").create(c.ctx);
    expect((await t.execute("x", { action: "list_messages" })).details.error).toBe("graph_execution_account_required"); expect(fetch).not.toHaveBeenCalled();
    const worker = r.context("job", "graph-a"); await expect(r.bind("A", "A", worker)).rejects.toThrow("graph_execution_account_conflict");
    c.revoke(); expect((await t.execute("revoked", { action: "list_messages" })).details.ok).toBe(false); expect(fetch).not.toHaveBeenCalled();
  });
  it("revalidates host/profile after an awaited boundary before remote mutation", async () => {
    const r = await runtime(); await r.provision("A"); const c = r.context("race"); await r.bind("A", "A", c);
    let requests = 0;
    vi.stubGlobal("fetch", vi.fn(async (url: string) => { requests++; if (url.includes("login.microsoftonline.com")) { c.revoke(); return new Response(JSON.stringify({ access_token: "never-use", expires_in: 3600 }), { status: 200 }); } throw new Error("remote Graph must not be reached"); }));
    const tool = r.tools.get("outlook_mail_read").create(c.ctx);
    expect((await tool.execute("race-call", { action: "list_messages" })).details.ok).toBe(false); expect(requests).toBe(1);
  });
  it("checks exact target/incarnation and refuses personal binding in an unproven shared context", async () => {
    const r = await runtime(); await r.provision("A", false, false); const c = r.context("shared");
    await expect(r.bind("A", "A", c)).rejects.toThrow("access_denied");
    const draft = r.context("draft", "main", "draft"); await expect(r.bind("A", "A", draft)).rejects.toThrow("access_denied");
    const account = (await r.rpc("A", "status")).account;
    await r.rpc("A", "configure", { expectedRevision: account.revision, policy: r.policy, shared: true, unattended: false });
    await expect(r.bind("A", "A", draft)).resolves.toMatchObject({ profileId: "A" });
    const bad = { ...draft, authority: { target: { ...draft.authority.target, sessionId: "other" }, assertCurrent() {} } };
    await expect(r.bind("A", "A", bad, 1)).rejects.toThrow("graph_account_conflict");
    r.entries.get(draft.ctx.sessionKey).lifecycleRevision = "2";
    const fetch = vi.fn(); vi.stubGlobal("fetch", fetch); const tool = r.tools.get("outlook_mail_read").create(draft.ctx);
    expect((await tool.execute("privacy", { action: "list_messages" })).details.ok).toBe(false); expect(fetch).not.toHaveBeenCalled();
  });
  it("binds approval and allow-always scopes to account identity, including callbacks outside ALS", async () => {
    const r = await runtime(); await r.provision("A"); await r.provision("B"); const c = r.context("approval"); await r.bind("A", "A", c);
    const tool = r.tools.get("outlook_mail_write").create(c.ctx); const params = { action: "mark_read", messageId: "id-1", isRead: true };
    const approval = await r.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "approved-A", params }, c.ctx);
    expect(approval.requireApproval).toBeDefined();
    await expect(r.bind("B", "B", c, 1)).rejects.toThrow("graph_account_busy");
    await approval.requireApproval.onResolution("allow-once");
    const a = (await r.store.readAccount("A"))!; await r.rpc("A", "disconnect", { expectedRevision: a.revision });
    await expect(approval.requireApproval.onResolution("allow-once")).rejects.toThrow("graph_account_context_changed");
    vi.stubGlobal("fetch", vi.fn()); expect((await tool.execute("approved-A", params)).details.ok).toBe(false); expect(fetch).not.toHaveBeenCalled();
    const snapshots = new NativeApprovalSnapshotStore(); const trust = new WarningApprovalTrustStore();
    const scope = { identity: "A:1", stateDir: "A", policy: {}, assertCurrent() {}, async assertFresh() {} };
    accountScope.run(scope, () => snapshots.record("x", { agentId: "main", sessionId: "s", toolName: "t", params: "{}" }));
    expect(accountScope.run({ ...scope, identity: "B:1" }, () => snapshots.consume("x", "main", "s", "t", {}))).toBe(false);
    trust.grant({ accountIdentity: "A:1", agentId: "main", toolName: "t", action: "a" }); expect(trust.has({ accountIdentity: "B:1", agentId: "main", toolName: "t", action: "a" })).toBe(false);
  });
  it("keeps device sign-in codes/status per profile and denies foreign session handles", async () => {
    const r = await runtime();
    for (const id of ["A", "B"]) { const a = await r.rpc(id, "provision"); await r.rpc(id, "configure", { expectedRevision: a.revision, policy: r.policy, shared: false, unattended: false }); }
    let n = 0;
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({ device_code: `device-${++n}`, user_code: `CODE-${n}ABC`, verification_uri: "https://microsoft.com/devicelogin", expires_in: 600, interval: 5 }), { status: 200 })));
    const params = { clientId: "12345678-1234-1234-1234-123456789012", tenant: "common" };
    const A = await r.rpc("A", "device-start", params); const B = await r.rpc("B", "device-start", params);
    expect(A.sessionId).not.toBe(B.sessionId); expect(A.userCode).not.toBe(B.userCode);
    await expect(r.rpc("B", "device-status", { sessionId: A.sessionId })).rejects.toThrow("device_session_unavailable");
    await expect(r.rpc("B", "device-cancel", { sessionId: A.sessionId })).rejects.toThrow("device_session_unavailable");
    await r.rpc("A", "device-cancel", { sessionId: A.sessionId }); await r.rpc("B", "device-cancel", { sessionId: B.sessionId });
  });
  it("does not release fetched account data after consent is revoked during response", async () => {
    const r = await runtime(); await r.provision("A"); const c = r.context("response-race"); await r.bind("A", "A", c);
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.includes("login.microsoftonline.com")) return new Response(JSON.stringify({ access_token: "response-token", expires_in: 3600 }), { status: 200 });
      const a = (await r.rpc("A", "status")).account;
      await r.rpc("A", "configure", { expectedRevision: a.revision, policy: r.policy, shared: false, unattended: false });
      return new Response(JSON.stringify({ value: [{ id: "private-data", subject: "must-not-release-after-revocation" }] }), { status: 200 });
    }));
    const tool = r.tools.get("outlook_mail_read").create(c.ctx); const result = await tool.execute("response-race", { action: "list_messages" });
    expect(result.details.error).toBe("graph_account_context_changed"); expect(JSON.stringify(result)).not.toContain("must-not-release-after-revocation");
  });
  it("migrates legacy once to the explicitly authenticated owner and preserves backup without sharing", async () => {
    const r = await runtime(); await createVaultCredential(r.root, { clientId: "legacy", refreshToken: "synthetic-legacy", tenant: "common", scopes: ["Mail.Read", "offline_access"] }, r.key);
    const receipt = await r.rpc("A", "migrate-legacy"); expect(receipt).toMatchObject({ migrated: true, profileId: "A", retainedLegacyBackup: true });
    const a = (await r.rpc("A", "status")).account; expect(a.shared).toBe(false); expect(a.unattended).toBe(false);
    await expect(r.rpc("A", "migrate-legacy")).resolves.toEqual(receipt);
    await expect(r.rpc("B", "migrate-legacy")).rejects.toThrow("graph_account_conflict");
    await r.rpc("A", "disconnect", { expectedRevision: a.revision }); await expect(r.rpc("A", "migrate-legacy")).rejects.toThrow("graph_account_conflict");
  });
  it("migrates under the stricter profile ceiling and resolves only its authorized operations", async () => {
    const personal = graphPolicyFixture();
    personal.services.mail.agents.main.operations = ["read"];
    personal.services.calendar.agents.main.resources = ["synthetic-calendar"];
    personal.services.calendar.agents["secondary-agent"].resources = ["disjoint-resource"];
    personal.services.onedrive.allowed_roots[0].permissions.write = false;
    const r = await runtime(false, { A: personal });
    await createVaultCredential(r.root, { clientId: "legacy", refreshToken: "synthetic-legacy", tenant: "common", scopes: ["Mail.Read", "offline_access"] }, r.key);
    await r.rpc("A", "migrate-legacy");
    const migrated = (await r.rpc("A", "status")).account;
    const policy = migrated.policy as GraphPolicy;
    expect(() => authorizeOperation(policy, "main", "mail", "read", "me")).not.toThrow();
    expect(() => authorizeOperation(policy, "main", "mail", "send", "me")).toThrow("access_denied");
    expect(() => authorizeOperation(policy, "main", "calendar", "read", "synthetic-calendar")).not.toThrow();
    expect(() => authorizeOperation(policy, "main", "calendar", "read", "me")).toThrow("access_denied");
    expect(policy.services.calendar.agents["secondary-agent"]).toBeUndefined();
    expect(() => authorizeRoot(policy, "fixture-reader", "synthetic_documents", "read")).not.toThrow();
    expect(() => authorizeRoot(policy, "main", "synthetic_documents", "write")).toThrow("access_denied");
    await r.rpc("A", "configure", { expectedRevision: migrated.revision, policy: migrated.policy, shared: true, unattended: false });
    const c = r.context("migrated"); await r.bind("A", "A", c);
    const resolved = await r.store.resolve({ ...c.ctx, lifecycleRevision: "1", assertCurrent() {} });
    expect(() => authorizeOperation(resolved.scope.policy as typeof personal, "main", "mail", "read", "me")).not.toThrow();
    expect(() => authorizeOperation(resolved.scope.policy as typeof personal, "main", "mail", "send", "me")).toThrow("access_denied");
    expect(() => authorizeOperation(resolved.scope.policy as GraphPolicy, "main", "calendar", "read", "synthetic-calendar")).not.toThrow();
    expect(() => authorizeRoot(resolved.scope.policy as GraphPolicy, "main", "synthetic_documents", "write")).toThrow("access_denied");
  });
  it("does not widen legacy grants when the profile ceiling is broader", async () => {
    const legacy = graphPolicyFixture();
    legacy.services.mail.agents.main.operations = ["read"];
    legacy.services.calendar.agents.main.operations = ["read"];
    legacy.services.calendar.agents.main.resources = ["me"];
    legacy.services.onedrive.allowed_roots[0].permissions.write = false;
    const legacyOther = structuredClone(legacy.services.onedrive.allowed_roots[0]);
    legacyOther.label = "legacy_other"; legacyOther.path = "/Synthetic/Other"; legacyOther.item_id = "other-root";
    legacy.services.onedrive.allowed_roots.push(legacyOther);
    legacy.rules.warningApprovalsByService = { mail: true, onedrive: false };
    const broader = graphPolicyFixture();
    broader.services.calendar.agents.main.resources = ["me", "synthetic-calendar"];
    broader.services.onedrive.allowed_roots[0].agents["fixture-reader"].permissions.write = true;
    const mismatched = structuredClone(legacyOther); mismatched.drive_id = "different-drive";
    broader.services.onedrive.allowed_roots.push(mismatched);
    broader.rules.warningApprovalsByService = { mail: false, onedrive: true };
    const r = await runtime(false, { A: broader }, legacy);
    await createVaultCredential(r.root, { clientId: "legacy", refreshToken: "synthetic-legacy", tenant: "common", scopes: ["Mail.Read", "Mail.Send", "offline_access"] }, r.key);
    await r.rpc("A", "migrate-legacy");
    const migrated = (await r.rpc("A", "status")).account;
    const policy = migrated.policy as GraphPolicy;
    expect(() => authorizeOperation(policy, "main", "mail", "read", "me")).not.toThrow();
    expect(() => authorizeOperation(policy, "main", "mail", "send", "me")).toThrow("access_denied");
    expect(() => authorizeOperation(policy, "main", "calendar", "read", "me")).not.toThrow();
    expect(() => authorizeOperation(policy, "main", "calendar", "read", "synthetic-calendar")).toThrow("access_denied");
    expect(() => authorizeRoot(policy, "fixture-reader", "synthetic_documents", "read")).not.toThrow();
    expect(() => authorizeRoot(policy, "fixture-reader", "synthetic_documents", "write")).toThrow("access_denied");
    expect(() => authorizeRoot(policy, "fixture-reader", "legacy_other", "read")).toThrow("access_denied");
    expect(policy.rules.warningApprovalsByService).toEqual({ mail: true, onedrive: true });
    await r.rpc("A", "configure", { expectedRevision: migrated.revision, policy, shared: true, unattended: false });
    const c = r.context("broader-ceiling"); await r.bind("A", "A", c);
    const resolved = await r.store.resolve({ ...c.ctx, lifecycleRevision: "1", assertCurrent() {} });
    expect(() => authorizeOperation(resolved.scope.policy as GraphPolicy, "main", "mail", "read", "me")).not.toThrow();
    expect(() => authorizeOperation(resolved.scope.policy as GraphPolicy, "main", "mail", "send", "me")).toThrow("access_denied");
  });
});

const hostRoster = (ids = ["gateway-owner", "A"]) => JSON.stringify({ profiles: ids.map(id => ({ id, displayName: id, avatarMime: null, mergedInto: null, createdAt: 1, updatedAt: 1, emails: [], githubIdentity: null, hasAvatar: false, roleSource: "default" })) });
describe("automatic sole-user legacy continuity", () => {
  async function marker(r: Awaited<ReturnType<typeof runtime>>, path: string, state: "in_flight" | "quarantined") {
    const record = await readVaultCredential(path, r.key);
    const payload = { format: "openclaw-microsoft-graph-refresh-quarantine", version: 2, generation: record.envelope.generation, keyId: record.envelope.keyId, vaultDigest: record.digest, state, createdAt: "2026-10-01T00:00:00.000Z", updatedAt: "2026-10-01T00:00:00.000Z" };
    const raw = JSON.stringify({ ...payload, mac: createHmac("sha256", Buffer.from(r.key, "base64url")).update(JSON.stringify(payload)).digest("base64url") }) + "\n";
    await writeFile(vaultQuarantinePath(path), raw, { mode: 0o600 }); return raw;
  }
  async function deniedPublication(r: Awaited<ReturnType<typeof runtime>>) {
    expect((await r.store.readAccount("A"))?.connected ?? false).toBe(false);
    expect(await r.store.automaticStatus()).toBeUndefined();
    const c = r.context("migration-denied"); const fetch = vi.fn(); vi.stubGlobal("fetch", fetch);
    expect((await r.tools.get("outlook_mail_read").create(c.ctx).execute("denied", { action: "list_messages" })).details.ok).toBe(false);
    expect(fetch).not.toHaveBeenCalled();
  }
  it.each(["in_flight", "quarantined", "malformed"])("both migration paths deny source %s without laundering or changing backup", async state => {
    const r = await runtime(false, {}, graphPolicyFixture(), () => hostRoster()); await legacy(r);
    const source = await readFile(vaultCredentialPath(r.root));
    const raw = state === "malformed" ? "{invalid-marker" : await marker(r, r.root, state as "in_flight" | "quarantined");
    if (state === "malformed") await writeFile(vaultQuarantinePath(r.root), raw, { mode: 0o600 });
    for (const method of ["automatic-migration", "migrate-legacy"]) await expect(r.rpc("A", method)).rejects.toThrow(/credential_(reauthorization_required|vault_unavailable)/);
    expect(await readFile(vaultCredentialPath(r.root))).toEqual(source);
    expect(await readFile(vaultQuarantinePath(r.root), "utf8")).toBe(raw);
    await deniedPublication(r);
  });
  it.each(["source", "destination"])("denies held and stale %s vault locks without publishing", async side => {
    const r = await runtime(false, {}, graphPolicyFixture(), () => hostRoster()); await legacy(r); const a = await r.store.provision("A", () => {});
    const path = side === "source" ? r.root : r.store.accountPath("A", a.generation);
    const source = await readFile(vaultCredentialPath(r.root)); const lock = await acquireVaultLock(path, "synthetic-holder");
    try { await expect(r.rpc("A", "automatic-migration")).rejects.toThrow("credential_vault_locked"); } finally { await lock.release(); }
    const lockPath = vaultCredentialPath(path) + ".lock";
    await writeFile(lockPath, JSON.stringify({ pid: 999999, operation: "stale", startedAt: "2000-01-01T00:00:00.000Z" }), { mode: 0o600 });
    const old = new Date("2000-01-01T00:00:00.000Z"); await utimes(lockPath, old, old);
    await expect(r.rpc("A", "migrate-legacy")).rejects.toThrow("credential_vault_locked");
    expect(await readFile(vaultCredentialPath(r.root))).toEqual(source); await deniedPublication(r);
  });
  it.each(["source", "destination"])("denies compromised %s lock during copied-vault admission", async side => {
    let reads = 0; let r: Awaited<ReturnType<typeof runtime>>;
    r = await runtime(false, {}, graphPolicyFixture(), () => {
      if (++reads === 4) {
        const path = side === "source" ? r.root : r.store.accountPath("A", 1);
        // Synchronous corruption at the snapshot boundary: both locks are held.
        requireSyntheticWrite(vaultCredentialPath(path) + ".lock", "compromised");
      }
      return hostRoster();
    });
    await legacy(r); const source = await readFile(vaultCredentialPath(r.root));
    await expect(r.rpc("A", "automatic-migration")).rejects.toThrow("credential_vault_locked");
    expect(await readFile(vaultCredentialPath(r.root))).toEqual(source); await deniedPublication(r);
  });
  it.each(["in_flight", "quarantined", "malformed", "different-token", "orphan-marker"])("interrupted destination %s is never admitted by retry", async state => {
    const r = await runtime(false, {}, graphPolicyFixture(), () => hostRoster()); await legacy(r); const a = await r.store.provision("A", () => {});
    const path = r.store.accountPath("A", a.generation); const source = await readVaultCredential(r.root, r.key);
    if (state !== "orphan-marker") await createVaultCredential(path, { ...source.credential, ...(state === "different-token" ? { refreshToken: "different-synthetic-token" } : {}) }, r.key);
    else { const lock = await acquireVaultLock(path, "synthetic-orphan-marker"); await lock.release(); }
    if (state === "malformed" || state === "orphan-marker") await writeFile(vaultQuarantinePath(path), "not-a-marker", { mode: 0o600 });
    else if (state !== "different-token") await marker(r, path, state as "in_flight" | "quarantined");
    const destination = state === "orphan-marker" ? undefined : await readFile(vaultCredentialPath(path));
    for (const method of ["automatic-migration", "migrate-legacy"]) await expect(r.rpc("A", method)).rejects.toThrow(/credential_|graph_account_conflict/);
    if (destination) expect(await readFile(vaultCredentialPath(path))).toEqual(destination);
    else await expect(readFile(vaultCredentialPath(path))).rejects.toMatchObject({ code: "ENOENT" });
    expect(await readFile(vaultCredentialPath(r.root))).toEqual(source.raw); await deniedPublication(r);
  });
  it("caller cancellation before copy cannot publish a credential or positive grant", async () => {
    let reads = 0; const controller = new AbortController();
    const r = await runtime(false, {}, graphPolicyFixture(), () => {
      if (++reads === 3) controller.abort(); return hostRoster();
    });
    await legacy(r); const source = await readFile(vaultCredentialPath(r.root));
    await expect(r.rpc("A", "automatic-migration", {}, () => true, undefined, controller.signal)).rejects.toThrow();
    await expect(readFile(vaultCredentialPath(r.store.accountPath("A", 1)))).rejects.toMatchObject({ code: "ENOENT" });
    expect(await readFile(vaultCredentialPath(r.root))).toEqual(source); await deniedPublication(r);
  });
  it.each(["authority", "caller"])("persists a monotone tombstone despite %s cancellation during observed-transition persistence", async cancellation => {
    let reads = 0; let inject = true; let active = true; const controller = new AbortController();
    const r = await runtime(false, {}, graphPolicyFixture(), () => {
      if (inject && ++reads === 4) {
        setImmediate(() => { if (cancellation === "authority") active = false; else controller.abort(); });
        return hostRoster(["gateway-owner", "A", "B"]);
      }
      return hostRoster();
    });
    await legacy(r); const source = await readFile(vaultCredentialPath(r.root)); const fetch = vi.fn(); vi.stubGlobal("fetch", fetch);
    await expect(r.rpc("A", "automatic-migration", {}, () => active, undefined, controller.signal)).rejects.toThrow("graph_single_user_topology_changed");
    expect((await r.store.readAccount("A"))!.connected).toBe(false);
    const tombstone = await r.store.automaticStatus(); expect(tombstone).toMatchObject({ profileId: "A", generation: 1, invalidated: true });
    inject = false; active = true;
    await expect(r.rpc("A", "automatic-migration")).rejects.toThrow("graph_single_user_topology_changed");
    // A fresh store instance reads the durable captured binding, not process state.
    const reloaded = new AccountStore(r.root); expect(await reloaded.automaticStatus()).toEqual(tombstone);
    await expect(reloaded.checkAutomatic()).rejects.toThrow("graph_single_user_topology_changed");
    await expect(r.rpc("A", "migrate-legacy")).resolves.toMatchObject({ migrated: true });
    expect(await reloaded.automaticStatus()).toEqual(tombstone); expect((await reloaded.readAccount("A"))!.connected).toBe(true);
    expect(await readFile(vaultCredentialPath(r.root))).toEqual(source); expect(fetch).not.toHaveBeenCalled();
  });

  async function legacy(r: Awaited<ReturnType<typeof runtime>>) { await createVaultCredential(r.root, { clientId: "synthetic-legacy", tenant: "common", refreshToken: "synthetic-sole-user", scopes: ["Mail.Read", "offline_access"] }, r.key); }
  it("automatically migrates only authenticated sole admin; old/new/jobs/delegation need no bindings", async () => {
    const r = await runtime(false, {}, graphPolicyFixture(), () => hostRoster()); await legacy(r);
    const before = await import("./credential-vault.js").then(m => m.readVaultCredential(r.root, r.key));
    expect(r.methods.get("microsoft-graph.accounts.automatic-migration")!.options.scope).toBe("operator.admin");
    await expect(r.rpc(undefined, "automatic-migration")).rejects.toThrow("graph_authenticated_profile_required");
    expect(await r.rpc("B", "automatic-migration")).toMatchObject({ automatic: false });
    expect((await r.rpc("A", "status")).account).toBeNull();
    expect(await r.rpc("A", "automatic-migration")).toMatchObject({ automatic: true, retainedLegacyBackup: true });
    expect(await r.rpc("A", "automatic-migration")).toMatchObject({ automatic: true });
    const a = (await r.rpc("A", "status")).account; expect(a.policy).toEqual(r.policy); expect(a.shared).toBe(false); expect(a.unattended).toBe(false);
    const after = await import("./credential-vault.js").then(m => m.readVaultCredential(r.root, r.key)); expect(after.raw).toEqual(before.raw);
    vi.stubGlobal("fetch", vi.fn(async (url: string) => url.includes("login.microsoftonline.com") ? new Response(JSON.stringify({ access_token: "synthetic-sole-token", expires_in: 3600 })) : new Response(JSON.stringify({ value: [{ id: "synthetic-result" }] }))));
    for (const id of ["old-session", "continued-session", "new-session", "cron-job", "ordinary-delegation"]) {
      const c = r.context(id); const tool = r.tools.get("outlook_mail_read").create(c.ctx);
      expect((await tool.execute(id, { action: "list_messages" })).details.ok).toBe(true);
      expect((await r.store.status({ ...c.ctx, assertCurrent() {} }, "A")).binding).toBeNull();
    }
    const denied = r.context("new-unapproved-agent", "unknown-agent"); const result = await r.tools.get("outlook_mail_read").create(denied.ctx).execute("denied", { action: "list_messages" }); expect(result.details.error).toBe("access_denied");
  });
  it("durably invalidates automatic access on second human/identity change and never resurrects", async () => {
    let ids = ["gateway-owner", "A"];
    const r = await runtime(false, {}, graphPolicyFixture(), () => hostRoster(ids)); await legacy(r); await r.rpc("A", "automatic-migration");
    const c = r.context("old"); const tool = r.tools.get("outlook_mail_read").create(c.ctx); const fetch = vi.fn(); vi.stubGlobal("fetch", fetch);
    ids = ["gateway-owner", "A", "B"];
    expect((await tool.execute("second", { action: "list_messages" })).details.error).toBe("graph_single_user_topology_changed"); expect(fetch).not.toHaveBeenCalled();
    ids = ["gateway-owner", "A"];
    expect((await tool.execute("returned", { action: "list_messages" })).details.error).toBe("graph_single_user_topology_changed");
    await expect(r.rpc("A", "automatic-migration")).rejects.toThrow("graph_single_user_topology_changed");
    expect((await r.store.automaticStatus())!.invalidated).toBe(true);
  });
  it("rechecks topology after OAuth/response waits and withholds output after detected transition", async () => {
    let ids = ["gateway-owner", "A"];
    const r = await runtime(false, {}, graphPolicyFixture(), () => hostRoster(ids)); await legacy(r); await r.rpc("A", "automatic-migration");
    const c = r.context("wait"); let requests = 0;
    vi.stubGlobal("fetch", vi.fn(async () => { requests++; ids.push("B"); return new Response(JSON.stringify({ access_token: "must-not-be-dispatched", expires_in: 3600 })); }));
    const result = await r.tools.get("outlook_mail_read").create(c.ctx).execute("after-oauth", { action: "list_messages" });
    expect(result.details.ok).toBe(false); expect(requests).toBe(1); expect(JSON.stringify(result)).not.toContain("must-not-be-dispatched");
  });
  it("withholds fetched output and invalidates pending approvals after a detected topology transition", async () => {
    let ids = ["gateway-owner", "A"];
    const r = await runtime(false, {}, graphPolicyFixture(), () => hostRoster(ids)); await legacy(r); await r.rpc("A", "automatic-migration");
    const c = r.context("approval-snapshot"); r.tools.get("outlook_mail_write").create(c.ctx);
    const approval = await r.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "auto-approval", params: { action: "mark_read", messageId: "synthetic-id", isRead: true } }, c.ctx);
    expect(approval.requireApproval).toBeDefined(); ids.push("B");
    await expect(approval.requireApproval.onResolution("allow-once")).rejects.toThrow("graph_single_user_topology_changed");
    const second = await runtime(false, {}, graphPolicyFixture(), () => hostRoster(ids)); ids = ["gateway-owner", "A"]; await legacy(second); await second.rpc("A", "automatic-migration");
    const read = second.context("output");
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.includes("login.microsoftonline.com")) return new Response(JSON.stringify({ access_token: "synthetic-token", expires_in: 3600 }));
      ids.push("B"); return new Response(JSON.stringify({ value: [{ subject: "must-never-release-topology-private-output" }] }));
    }));
    const result = await second.tools.get("outlook_mail_read").create(read.ctx).execute("output", { action: "list_messages" });
    expect(result.details.ok).toBe(false); expect(JSON.stringify(result)).not.toContain("must-never-release-topology-private-output");
  });
  it("fails closed on outages, migration authority loss, and preserves explicit worker precedence", async () => {
    let unavailable = false;
    const r = await runtime(true, {}, graphPolicyFixture(), () => unavailable ? "malformed" : hostRoster()); await legacy(r); await r.rpc("A", "automatic-migration");
    const worker = r.context("fixed", "graph-a"); const fetch = vi.fn(); vi.stubGlobal("fetch", fetch);
    expect((await r.tools.get("outlook_mail_read").create(worker.ctx).execute("fixed", { action: "list_messages" })).details.error).toBe("access_denied");
    unavailable = true; const c = r.context("outage"); expect((await r.tools.get("outlook_mail_read").create(c.ctx).execute("outage", { action: "list_messages" })).details.error).toBe("graph_single_user_directory_unavailable"); expect(fetch).not.toHaveBeenCalled();
    const second = await runtime(false, {}, graphPolicyFixture(), () => hostRoster()); await legacy(second);
    await expect(second.rpc("A", "automatic-migration", {}, () => false)).rejects.toThrow("graph_authenticated_profile_required"); expect((await second.rpc("A", "status")).account).toBeNull();
  });
  it("recovers a copied-vault interrupted migration without altering encrypted backup", async () => {
    let reads = 0; let fault = true;
    const r = await runtime(false, {}, graphPolicyFixture(), () => ++reads === 4 && fault ? "malformed" : hostRoster()); await legacy(r);
    await expect(r.rpc("A", "automatic-migration")).rejects.toThrow("graph_single_user_directory_unavailable");
    expect((await r.store.readAccount("A"))!.connected).toBe(false);
    expect(await r.store.automaticStatus()).toBeUndefined();
    fault = false;
    expect(await r.rpc("A", "automatic-migration")).toMatchObject({ automatic: true, retainedLegacyBackup: true });
    expect((await r.store.readAccount("A"))!.connected).toBe(true);
  });
  it("records a permanent tombstone for a topology change detected during migration", async () => {
    let reads = 0; let changed = true;
    const r = await runtime(false, {}, graphPolicyFixture(), () => changed && ++reads === 4 ? hostRoster(["gateway-owner", "A", "B"]) : hostRoster()); await legacy(r);
    await expect(r.rpc("A", "automatic-migration")).rejects.toThrow("graph_single_user_topology_changed");
    expect((await r.store.automaticStatus())!.invalidated).toBe(true);
    changed = false;
    await expect(r.rpc("A", "automatic-migration")).rejects.toThrow("graph_single_user_topology_changed");
    // Explicit self-migration remains usable after an interrupted automatic
    // bootstrap, but cannot resurrect the revoked convenience receipt.
    expect(await r.rpc("A", "migrate-legacy")).toMatchObject({ migrated: true });
    expect((await r.store.automaticStatus())!.invalidated).toBe(true);
    expect((await r.store.readAccount("A"))!.connected).toBe(true);
  });
  it("revokes convenience on account disconnect, not explicit fallback authority", async () => {
    const r = await runtime(false, {}, graphPolicyFixture(), () => hostRoster()); await legacy(r); await r.rpc("A", "automatic-migration");
    const a = (await r.store.readAccount("A"))!;
    await r.rpc("A", "disconnect", { expectedRevision: a.revision });
    expect((await r.store.automaticStatus())!.invalidated).toBe(true);
    const c = r.context("disconnected"); vi.stubGlobal("fetch", vi.fn());
    const result = await r.tools.get("outlook_mail_write").create(c.ctx).execute("disconnected", { action: "mark_read", messageId: "synthetic-id", isRead: true });
    expect(result.details.error).toBe("graph_single_user_topology_changed"); expect(result.details.remoteEffectsMayHaveOccurred).toBe(true); expect(result.details.outcome).toBe("unknown");
    expect(fetch).not.toHaveBeenCalled();
  });
  it("honors stricter profile ceiling and explicit account-policy changes end automatic mode", async () => {
    const cap = structuredClone(graphPolicyFixture()); cap.services.mail.agents.main.operations = ["read"];
    const r = await runtime(false, { A: cap }, graphPolicyFixture(), () => hostRoster()); await legacy(r); await r.rpc("A", "automatic-migration");
    const a = (await r.rpc("A", "status")).account; expect(a.policy.services.mail.agents.main.operations).toEqual(["read"]);
    await r.rpc("A", "configure", { expectedRevision: a.revision, policy: a.policy, shared: true, unattended: true });
    expect((await r.store.automaticStatus())!.invalidated).toBe(true);
    const c = r.context("explicit"); await r.bind("A", "A", c); expect((await r.store.readAccount("A"))!.shared).toBe(true);
  });
});
