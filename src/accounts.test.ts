import { mkdtemp, rm, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { AccountStore, assertPolicyCeiling, blankAccountPolicy, type AccountHostContext } from "./accounts.js";
import { accountScope, assertAccountEffect } from "./account-scope.js";
import { OneDriveAgentsSessionCache } from "./onedrive-agents-instructions.js";
import { ContinuationStore } from "./continuation.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
const dirs: string[] = [];
afterEach(async () => { for (const dir of dirs.splice(0)) await rm(dir, { recursive: true, force: true }); });
const target = (id: string, agentId = "main"): AccountHostContext => ({ agentId, sessionKey: `agent:${agentId}:${id}`, sessionId: id, lifecycleRevision: "1", assertCurrent() {} });
async function setup(workers: Record<string, string> = {}) {
  const root = await mkdtemp(join(tmpdir(), "graph-accounts-test-")); dirs.push(root);
  const policy = graphPolicyFixture();
  for (const service of ["calendar", "mail", "todo"] as const) for (const agent of Object.keys(workers)) policy.services[service].agents[agent] = { operations: ["read"] };
  const store = new AccountStore(root, workers, policy);
  for (const id of ["A", "B"]) { const a = await store.provision(id, () => {}); await store.update(id, a.revision, policy, true, true, () => {}); await store.markConnected(id, a.generation, () => {}); }
  return { root, policy, store };
}
describe("protected account/session contract", () => {
  it("single/multi deployment revisions invalidate even explicitly bound scopes and callbacks", async () => {
    const { store } = await setup(); const t = target("bound"); await store.bind("A", "A", t, null);
    await store.setInstallation("multi-user", () => {}); const old = await store.resolve(t);
    await expect(store.setInstallation("single-user", () => {}, { soleProfileId: "A", digest: "a".repeat(64) })).rejects.toThrow("graph_account_conflict");
    expect((await store.installationStatus())?.revision).toBe(1);
    // A separate one-account installation can create deliberate deployment grants.
    const root = await mkdtemp(join(tmpdir(), "graph-single-mode-")); dirs.push(root);
    const one = new AccountStore(root, {}, graphPolicyFixture()); await one.setInstallation("single-user", () => {}, { soleProfileId: "A", digest: "a".repeat(64) });
    const a = await one.provision("A", () => {}); await one.update("A", a.revision, graphPolicyFixture(), false, false, () => {}); await one.markConnected("A", 1, () => {});
    const scope = (await one.resolve(target("cron"))).scope; await one.setInstallation("multi-user", () => {});
    expect(() => scope.assertCurrent()).toThrow("graph_account_context_changed"); await expect(scope.assertFresh()).rejects.toThrow("graph_account_context_changed");
    await expect(one.resolve(target("new-job"))).rejects.toThrow("graph_execution_account_required");
    await one.setInstallation("single-user", () => {}, { soleProfileId: "A", digest: "a".repeat(64) });
    expect((await one.resolve(target("unchanged-job"))).scope.identity).not.toBe(scope.identity);
    expect(() => old.scope.assertCurrent()).not.toThrow(); // Failed mode change never revokes the valid previous deployment.
  });
  it("keeps simultaneous A/B contexts, credentials and continuations separate", async () => {
    const { store } = await setup(); const A = target("A-session"), B = target("B-session");
    await store.bind("A", "A", A, null); await store.bind("B", "B", B, null);
    const [a, b] = await Promise.all([store.resolve(A), store.resolve(B)]);
    expect(a.scope.stateDir).not.toBe(b.scope.stateDir); expect(a.scope.identity).not.toBe(b.scope.identity);
    const continuation = new ContinuationStore(); const binding = { agentId: "main", service: "mail" as const, action: "list", resource: "me", criteria: "{}" };
    const handle = continuation.issue({ ...binding, accountIdentity: a.scope.identity }, "/me/messages?$skip=2", "/me/messages");
    expect(() => continuation.resolve(handle, { ...binding, accountIdentity: b.scope.identity }, "/me/messages")).toThrow("invalid_continuation");
    await Promise.all([accountScope.run(a.scope, async () => { await assertAccountEffect(); expect(accountScope.getStore()?.identity).toBe(a.scope.identity); }), accountScope.run(b.scope, async () => { await assertAccountEffect(); expect(accountScope.getStore()?.identity).toBe(b.scope.identity); })]);
  });
  it("shared steering does not change A; an explicit between-work switch increments binding", async () => {
    const { store } = await setup(); const t = target("shared"); const a = await store.bind("A", "A", t, null);
    expect((await store.status(t, "B")).binding?.profileId).toBe("A");
    const old = await store.resolve(t); const b = await store.bind("B", "B", t, a.revision);
    expect(b.profileId).toBe("B"); expect(b.revision).toBe(a.revision + 1); expect(() => old.scope.assertCurrent()).toThrow("graph_account_context_changed");
  });
  it("blocks account switches during execution and pending/accepted approval", async () => {
    const { store } = await setup(); const t = target("busy"); const b = await store.bind("A", "A", t, null); const r = await store.resolve(t);
    const release = store.hold(r.holdKey);
    await expect(store.bind("B", "B", t, b.revision)).rejects.toThrow("graph_account_busy"); release();
    store.pendingApproval("call", r.holdKey); await expect(store.bind("B", "B", t, b.revision)).rejects.toThrow("graph_account_busy");
    store.finishApproval("call"); await expect(store.bind("B", "B", t, b.revision)).resolves.toMatchObject({ profileId: "B" });
  });
  it("rejects unshared personal contexts when continuous private output authority is unproven", async () => {
    const { store } = await setup(); const a = (await store.readAccount("A"))!;
    await store.update("A", a.revision, a.policy, false, false, () => {});
    await expect(store.bind("B", "A", target("bad"), null)).rejects.toThrow("access_denied");
    await expect(store.bind("A", "A", target("shared"), null)).rejects.toThrow("access_denied");
    const personal = { ...target("draft"), privateOwner: "A" };
    await expect(store.bind("A", "A", personal, null)).rejects.toThrow("access_denied");
    await expect(store.resolve({ ...personal, privateOwner: undefined })).rejects.toThrow("graph_execution_account_required");
  });
  it("unbound new session/fork never falls back to the only connected account", async () => {
    const { store } = await setup(); await expect(store.resolve(target("fresh"))).rejects.toThrow("graph_execution_account_required");
    await store.bind("A", "A", target("old"), null);
    await expect(store.resolve({ ...target("old"), lifecycleRevision: "2" })).rejects.toThrow("graph_execution_account_required");
  });
  it("fixed worker is deterministic and cannot be selected away or used without unattended consent", async () => {
    const { root, policy, store } = await setup({ "graph-a": "A" });
    const t = target("job", "graph-a"); const r = await store.resolve(t);
    expect(r.scope.stateDir).toBe(store.accountPath("A", 1));
    await expect(store.bind("B", "B", t, null)).rejects.toThrow("graph_execution_account_conflict");
    const otherStore = new AccountStore(root, {}, policy); await otherStore.bind("B", "B", t, null);
    await expect(store.resolve(t)).rejects.toThrow("graph_execution_account_conflict");
    const a = (await store.readAccount("A"))!; await store.update("A", a.revision, a.policy, true, false, () => {});
    await expect(store.resolve(target("other-job", "graph-a"))).rejects.toThrow("access_denied");
  });
  it("disconnect/reconnect invalidates old scope/handles without resurrecting encrypted generations", async () => {
    const { store } = await setup(); const t = target("revoked"); await store.bind("A", "A", t, null); const r = await store.resolve(t);
    const a = (await store.readAccount("A"))!; await store.disconnect("A", a.revision, () => {});
    await expect(accountScope.run(r.scope, assertAccountEffect)).rejects.toThrow("graph_account_context_changed");
    const next = (await store.readAccount("A"))!; expect(next.generation).toBe(2); expect(next.shared).toBe(false); expect(next.unattended).toBe(false);
    await expect(store.resolve(t)).rejects.toThrow("graph_execution_account_required");
  });
  it("persists explicit ownership and bindings across restart, but rejects corrupt state", async () => {
    const { root, policy, store } = await setup(); const t = target("restart"); await store.bind("A", "A", t, null);
    const next = new AccountStore(root, {}, policy); expect((await next.status(t, "A")).binding?.profileId).toBe("A");
    const file = join(root, "plugin-data/microsoft-graph/accounts/registry.json"); const raw = await readFile(file, "utf8"); expect(raw).not.toContain("refreshToken");
    await writeFile(file, "corrupt", { mode: 0o600 }); await expect(next.resolve(t)).rejects.toThrow("graph_account_store_unavailable");
  });
  it("never publishes an unauthorized registry change, and rejects stale revisions", async () => {
    const { store, policy } = await setup(); const a = (await store.readAccount("A"))!;
    await expect(store.update("A", a.revision, blankAccountPolicy(), false, false, () => { throw new Error("revoked"); })).rejects.toThrow("revoked");
    expect((await store.readAccount("A"))?.shared).toBe(true);
    await expect(store.update("A", 99, policy, true, true, () => {})).rejects.toThrow("graph_account_conflict");
  });
  it("never transfers OneDrive instruction content or acknowledgements across account identities", async () => {
    const { store } = await setup(); const A = target("instructions-A"), B = target("instructions-B"); await store.bind("A", "A", A, null); await store.bind("B", "B", B, null);
    const a = await store.resolve(A), b = await store.resolve(B); const cache = new OneDriveAgentsSessionCache();
    const input = { agentId: "main", sessionId: "same-session", rootPin: "same-root", rootLabel: "root", relativeDirectory: "" };
    const first = await accountScope.run(a.scope, () => cache.discover({ ...input, load: async () => new TextEncoder().encode("A-instructions") }));
    const second = await accountScope.run(b.scope, () => cache.discover({ ...input, load: async () => new TextEncoder().encode("B-instructions") }));
    expect(first.acknowledgement).not.toBe(second.acknowledgement); expect(second.instructions?.[0].content).toBe("B-instructions");
    expect(accountScope.run(b.scope, () => cache.cachedAcknowledgement({ ...input, relativeDirectories: [""], acknowledgement: first.acknowledgement }))).toBe("invalid");
    cache.clearSession(input.sessionId);
  });
  it("enforces per-account agent/operation/resource ceiling and approval floor", () => {
    const policy = graphPolicyFixture(); const invalid = graphPolicyFixture(); invalid.services.mail.agents.main.resources = ["foreign-mailbox"];
    expect(() => assertPolicyCeiling(invalid, policy)).toThrow("access_denied");
    const weak = graphPolicyFixture(); weak.rules.warningApprovalsByService = { mail: false }; expect(() => assertPolicyCeiling(weak, policy)).toThrow("access_denied");
  });
});
