import { mkdtemp, rm } from "node:fs/promises";
import { randomBytes } from "node:crypto";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import entry, { NativeApprovalSnapshotStore, WarningApprovalTrustStore } from "./index.js";
import { accountScope } from "./account-scope.js";
import { AccountStore } from "./accounts.js";
import { createVaultCredential } from "./credential-vault.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
import { authorizeOperation, authorizeRoot, type GraphPolicy } from "./policy.js";
const dirs: string[] = [];
afterEach(async () => { vi.unstubAllGlobals(); for (const dir of dirs.splice(0)) await rm(dir, { recursive: true, force: true }); });
async function runtime(worker = false, accountPolicyCeilings: Record<string, GraphPolicy> = {}, legacyPolicy: GraphPolicy = graphPolicyFixture()) {
  const root = await mkdtemp(join(tmpdir(), "graph-host-test-")); dirs.push(root); const key = randomBytes(32).toString("base64url");
  const policy = structuredClone(legacyPolicy);
  if (worker) for (const service of ["mail", "todo", "calendar"] as const) policy.services[service].agents["graph-a"] = { operations: ["read"] };
  const methods = new Map<string, { handler: (ctx: any) => Promise<void>; options: any }>(); const tools = new Map<string, any>(); const hooks: Record<string, any> = {};
  const entries = new Map<string, any>();
  const api = {
    pluginConfig: { enabled: true, credentialVaultKey: key, policy, accountWorkers: worker ? { "graph-a": "A" } : {}, accountPolicyCeilings },
    runtime: { state: { resolveStateDir: () => root }, agent: { session: { getSessionEntry: ({ sessionKey }: any) => entries.get(sessionKey) } } },
    registerTool: (factory: any, options: any) => tools.set(options.name ?? options.names[0], factory),
    registerGatewayMethod: (name: string, handler: any, options: any) => methods.set(name, { handler, options }),
    on: (name: string, handler: any) => hooks[name] = handler,
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  };
  entry.register(api as any);
  const rpc = async (actor: string | undefined, name: string, params: Record<string, unknown> = {}, current = () => true, authority?: any) => {
    let reply: any;
    const ctx = { params, client: actor ? { authenticatedUserProfile: { profileId: actor } } : {}, hasCurrentClientAuthority: current,
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
