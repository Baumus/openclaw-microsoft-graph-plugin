import type { AccountDirectory, DirectorySnapshot } from "./account-directory.js";
import { createHash } from "node:crypto";
import { join } from "node:path";
import { createSecretFileAtomic } from "@openclaw/fs-safe/secret";
import { replaceFileAtomic } from "@openclaw/fs-safe/atomic";
import { readSecureFile } from "@openclaw/fs-safe/secure-file";
import { acquireFileLock } from "@openclaw/fs-safe/file-lock";
import { accountScope, type AccountScope } from "./account-scope.js";
import { validatePolicy, authorizeOperation, authorizeRoot, type GraphPolicy } from "./policy.js";

export const blankAccountPolicy = (): GraphPolicy => ({ version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: {} }, mail: { agents: {} }, todo: { agents: {} } } });
export const SHARING_WARNING = "Gemeinsam bedienbare Agents können die hier freigegebenen Microsoft-Aktionen in diesem Konto auch auf Anweisung anderer zugelassener Benutzer ausführen. Leser gemeinsamer Sessions können dort ausgegebene Microsoft-Daten sehen. OAuth-Tokens und die Verwaltung deiner Verbindung werden dadurch nicht geteilt.";
export type Account = { profileId: string; generation: number; revision: number; connected: boolean; shared: boolean; unattended: boolean; policySource?: "installation" | "personal"; policy: GraphPolicy };
export type SessionTarget = { agentId: string; sessionKey: string; sessionId: string; lifecycleRevision?: string };
export type Binding = SessionTarget & { profileId: string; generation: number; revision: number; privateOwner?: string };
export type AutomaticAccount = { profileId: string; generation: number; topologyDigest: string; invalidated: boolean };
export type InstallationMode = "auto" | "single-user" | "multi-user";
export type Installation = { mode: Exclude<InstallationMode, "auto">; revision: number; profileId?: string };
export type AccountRegistry = { installation?: Installation; automatic?: AutomaticAccount; legacyMigration?: { profileId: string; generation: number; sourceDigest: string }; version: 1; accounts: Record<string, Account>; sessions: Record<string, Binding> };
type Registry = AccountRegistry;
export type AccountHostContext = SessionTarget & { assertCurrent: () => void; privateOwner?: string };
export type WorkerBindings = Record<string, string>;
const fingerprint = (value: unknown) => createHash("sha256").update(JSON.stringify(value)).digest("hex");
const profileKey = (id: string) => fingerprint(["profile", id]);
const sessionKey = (target: SessionTarget) => fingerprint([target.agentId, target.sessionKey, target.sessionId, target.lifecycleRevision ?? null]);
function validProfile(id: unknown): asserts id is string { if (typeof id !== "string" || !id || id.length > 256 || /[\u0000-\u001f]/.test(id)) throw new Error("graph_authenticated_profile_required"); }
function validateRegistry(value: unknown): Registry {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("graph_account_store_unavailable");
  const r = value as Registry;
  const closed = (v: unknown, fields: string[]) => !!v && typeof v === "object" && !Array.isArray(v) && Object.keys(v).every(k => fields.includes(k));
  if (!closed(r, ["version", "accounts", "sessions", "legacyMigration", "automatic", "installation"])) throw new Error("graph_account_store_unavailable");
  if (r.installation) {
    const i = r.installation;
    if (!closed(i, ["mode", "revision", "profileId"]) || !["single-user", "multi-user"].includes(i.mode) || !Number.isSafeInteger(i.revision) || i.revision < 1) throw new Error("graph_account_store_unavailable");
    if (i.mode === "single-user") { validProfile(i.profileId); if (i.profileId === "gateway-owner") throw new Error("graph_account_store_unavailable"); }
    else if (i.profileId !== undefined) throw new Error("graph_account_store_unavailable");
  }
  if (r.automatic) validProfile(r.automatic.profileId);
  if (r.automatic && (!closed(r.automatic, ["profileId", "generation", "topologyDigest", "invalidated"]) || !r.automatic.profileId || !Number.isSafeInteger(r.automatic.generation) || r.automatic.generation < 1 || !/^[a-f0-9]{64}$/.test(r.automatic.topologyDigest) || typeof r.automatic.invalidated !== "boolean")) throw new Error("graph_account_store_unavailable");
  if (r.legacyMigration && (!closed(r.legacyMigration, ["profileId", "generation", "sourceDigest"]) || typeof r.legacyMigration.profileId !== "string" || !Number.isSafeInteger(r.legacyMigration.generation) || !/^[A-Za-z0-9_-]{43}$/.test(r.legacyMigration.sourceDigest))) throw new Error("graph_account_store_unavailable");
  if (r.version !== 1 || !r.accounts || !r.sessions || typeof r.accounts !== "object" || typeof r.sessions !== "object" || Array.isArray(r.accounts) || Array.isArray(r.sessions)) throw new Error("graph_account_store_unavailable");
  for (const [key, a] of Object.entries(r.accounts)) {
    if (!closed(a, ["profileId", "generation", "revision", "connected", "shared", "unattended", "policySource", "policy"])) throw new Error("graph_account_store_unavailable");
    validProfile(a.profileId);
    if (key !== profileKey(a.profileId) || !Number.isSafeInteger(a.generation) || a.generation < 1 || !Number.isSafeInteger(a.revision) || a.revision < 1 || [a.connected, a.shared, a.unattended].some(x => typeof x !== "boolean")) throw new Error("graph_account_store_unavailable");
    if (a.policySource !== undefined && !["installation", "personal"].includes(a.policySource)) throw new Error("graph_account_store_unavailable");
    validatePolicy(a.policy);
  }
  if (r.automatic && !r.automatic.invalidated) {
    const account = r.accounts[profileKey(r.automatic.profileId)];
    if (!account?.connected || account.generation !== r.automatic.generation || r.legacyMigration?.profileId !== account.profileId || r.legacyMigration.generation !== account.generation) throw new Error("graph_account_store_unavailable");
  }
  for (const [key, b] of Object.entries(r.sessions)) {
    if (!closed(b, ["agentId", "sessionKey", "sessionId", "lifecycleRevision", "profileId", "generation", "revision", "privateOwner"])) throw new Error("graph_account_store_unavailable");
    validProfile(b.profileId);
    if (key !== sessionKey(b) || !b.agentId || !b.sessionId || !b.sessionKey || !Number.isSafeInteger(b.revision) || b.revision < 1 || !Number.isSafeInteger(b.generation) || b.generation < 1) throw new Error("graph_account_store_unavailable");
  }
  return r;
}
/** The administrator's existing policy is a ceiling; personal policy can only restrict it. */
export function assertPolicyCeiling(policy: GraphPolicy, ceiling: GraphPolicy): void {
  for (const service of ["calendar", "mail", "todo"] as const) for (const [agent, grant] of Object.entries(policy.services[service].agents)) {
    for (const op of grant.operations) for (const resource of grant.resources ?? ["me"]) authorizeOperation(ceiling, agent, service, op, resource);
  }
  for (const root of policy.services.onedrive.allowed_roots) for (const [agent, grant] of Object.entries(root.agents)) for (const op of ["read", "write", "delete"] as const) if (grant.permissions[op]) {
    const c = authorizeRoot(ceiling, agent, root.label, op);
    if (root.drive_id !== c.drive_id || root.item_id !== c.item_id || root.agents_instructions !== c.agents_instructions || !root.permissions[op]) throw new Error("access_denied");
  }
  for (const service of ["onedrive", "calendar", "mail", "todo"] as const) if (policy.rules.warningApprovalsByService?.[service] === false && ceiling.rules.warningApprovalsByService?.[service] !== false) throw new Error("access_denied");
}
/** Migration keeps only legacy grants that the destination profile ceiling still permits. */
export function intersectMigrationPolicy(legacy: GraphPolicy, ceiling: GraphPolicy): GraphPolicy {
  const policy = structuredClone(legacy);
  for (const service of ["calendar", "mail", "todo"] as const) {
    const grants = policy.services[service].agents;
    for (const [agent, grant] of Object.entries(grants)) {
      const cap = ceiling.services[service].agents[agent];
      if (!cap) { delete grants[agent]; continue; }
      const operations = grant.operations.filter(operation => cap.operations.includes(operation));
      const resources = (grant.resources ?? ["me"]).filter(resource => (cap.resources ?? ["me"]).includes(resource));
      if (!operations.length || !resources.length) { delete grants[agent]; continue; }
      grant.operations = operations;
      if (grant.resources !== undefined) grant.resources = resources;
    }
  }
  policy.services.onedrive.allowed_roots = policy.services.onedrive.allowed_roots.flatMap(root => {
    const cap = ceiling.services.onedrive.allowed_roots.find(candidate => candidate.label === root.label);
    if (!cap || root.path !== cap.path || root.drive_id !== cap.drive_id || root.item_id !== cap.item_id || root.include_descendants !== cap.include_descendants || root.agents_instructions !== cap.agents_instructions) return [];
    for (const operation of ["read", "write", "delete"] as const) root.permissions[operation] = root.permissions[operation] && cap.permissions[operation];
    for (const [agent, grant] of Object.entries(root.agents)) {
      const capGrant = cap.agents[agent];
      for (const operation of ["read", "write", "delete"] as const) if (grant.permissions[operation]) grant.permissions[operation] = !!(root.permissions[operation] && capGrant?.permissions[operation]);
      if (!(["read", "write", "delete"] as const).some(operation => grant.permissions[operation] === true)) delete root.agents[agent];
    }
    return Object.keys(root.agents).length ? [root] : [];
  });
  for (const service of ["onedrive", "calendar", "mail", "todo"] as const) {
    const former = legacy.rules.warningApprovalsByService?.[service];
    const cap = ceiling.rules.warningApprovalsByService?.[service];
    if (former === false && cap !== false) {
      if (cap === true) (policy.rules.warningApprovalsByService ??= {})[service] = true;
      else delete policy.rules.warningApprovalsByService?.[service];
    } else if (former === undefined && cap === true) (policy.rules.warningApprovalsByService ??= {})[service] = true;
  }
  validatePolicy(policy);
  assertPolicyCeiling(policy, legacy);
  assertPolicyCeiling(policy, ceiling);
  return policy;
}

export class AccountStore {
  private registry: Registry = { version: 1, accounts: {}, sessions: {} };
  private readonly holds = new Map<string, number>();
  private readonly pending = new Map<string, { target: string; expires: number }>();
  constructor(readonly root: string, readonly workers: WorkerBindings = {}, readonly ceiling: GraphPolicy = blankAccountPolicy(), readonly ceilings: Record<string, GraphPolicy> = {}, readonly hostBusy: (target: SessionTarget) => boolean = () => false, readonly directory?: AccountDirectory, readonly installationMode: InstallationMode = "auto") {}
  ceilingFor(profileId: string): GraphPolicy { return this.ceilings[profileId] ?? this.ceiling; }
  private get file() { return join(this.root, "plugin-data", "microsoft-graph", "accounts", "registry.json"); }
  private async load(): Promise<Registry> {
    try {
      const result = await readSecureFile({ filePath: this.file, label: "Graph account registry", trust: { trustedDirs: [join(this.root, "plugin-data", "microsoft-graph", "accounts")] }, io: { maxBytes: 2 * 1024 * 1024, timeoutMs: 5000 } });
      if (process.platform !== "win32" && (result.stat.mode & 0o777) !== 0o600) throw new Error("graph_account_store_unavailable");
      return validateRegistry(JSON.parse(result.buffer.toString("utf8")));
    } catch (error) {
      // Only an actual missing file is an empty store, never corrupt/unreadable state.
      let e: unknown = error;
      for (let i = 0; i < 5 && e && typeof e === "object"; i++, e = (e as { cause?: unknown }).cause) if (["ENOENT", "not-found"].includes(String((e as { code?: unknown }).code))) return { version: 1, accounts: {}, sessions: {} };
      throw new Error("graph_account_store_unavailable");
    }
  }
  assertInstallationCurrent(expected: Installation | undefined): void {
    if (fingerprint(this.registry.installation ?? null) !== fingerprint(expected ?? null)) throw new Error("graph_account_context_changed");
  }
  async installationStatus(): Promise<Installation | undefined> { await this.refresh(); return this.registry.installation ? structuredClone(this.registry.installation) : undefined; }
  /** Initialization-only detection; an explicit config change creates a new deployment grant.
   * Old automatic tombstones and credential generations are never reset. */
  async setInstallation(mode: Installation["mode"], current: () => void, snapshot?: DirectorySnapshot): Promise<Installation> {
    return this.mutate(r => {
      if (r.installation?.mode === mode) return r.installation;
      if (mode === "single-user") {
        const id = snapshot?.soleProfileId;
        if (!id || id === "gateway-owner") throw new Error("graph_single_user_directory_unavailable");
        const accounts = Object.values(r.accounts);
        if (accounts.some(a => a.profileId !== id) || Object.values(r.sessions).some(b => b.profileId !== id)) throw new Error("graph_account_conflict");
        // Only fresh installs may select an owner before personal OAuth exists.
        if (r.installation && !accounts.find(a => a.profileId === id)?.connected) throw new Error("graph_execution_account_required");
        r.installation = { mode, revision: (r.installation?.revision ?? 0) + 1, profileId: id };
      } else {
        r.installation = { mode, revision: (r.installation?.revision ?? 0) + 1 };
        if (r.automatic) r.automatic.invalidated = true;
      }
      return r.installation;
    }, current);
  }
  private assertPersonalOwner(profileId: string): void {
    if (this.registry.installation?.mode === "single-user" && this.registry.installation.profileId !== profileId) throw new Error("access_denied");
  }
  async effectivePolicy(account: Account): Promise<GraphPolicy> {
    await this.refresh();
    return this.registry.installation?.mode === "single-user" && this.registry.installation.profileId === account.profileId && account.policySource === "installation"
      ? intersectMigrationPolicy(this.ceiling, this.ceilingFor(account.profileId)) : structuredClone(account.policy);
  }
  async automaticStatus(): Promise<AutomaticAccount | undefined> { await this.refresh(); return this.registry.automatic ? structuredClone(this.registry.automatic) : undefined; }
  async checkAutomatic(snapshot?: DirectorySnapshot): Promise<AutomaticAccount | undefined> {
    await this.refresh();
    const a = this.registry.automatic;
    if (!a) return;
    if (a.invalidated) throw new Error("graph_single_user_topology_changed");
    const fresh = snapshot ?? await this.directory?.read();
    if (!fresh) throw new Error("graph_single_user_directory_unavailable");
    if (fresh.soleProfileId !== a.profileId || fresh.digest !== a.topologyDigest) {
      // Sticky, durable tombstone. Returning to a singleton never resurrects it.
      await this.mutate(r => { if (r.automatic) r.automatic.invalidated = true; }, () => {});
      throw new Error("graph_single_user_topology_changed");
    }
    await this.refresh();
    if (fingerprint(this.registry.automatic) !== fingerprint(a)) throw new Error("graph_account_context_changed");
    return structuredClone(a);
  }
  async refresh(): Promise<void> { this.registry = await this.load(); }
  accountPath(profileId: string, generation: number): string { validProfile(profileId); return join(this.root, "plugin-data", "microsoft-graph", "accounts", profileKey(profileId), String(generation)); }
  async readAccount(profileId: string): Promise<Account | undefined> { validProfile(profileId); await this.refresh(); const a = this.registry.accounts[profileKey(profileId)]; return a ? structuredClone(a) : undefined; }
  private busy(target: string): boolean {
    const now = Date.now(); for (const [id, p] of this.pending) if (p.expires <= now) this.pending.delete(id);
    return !!this.holds.get(target) || [...this.pending.values()].some(p => p.target === target);
  }
  private accountBusy(profileId: string): boolean {
    return Object.entries(this.registry.sessions).some(([key, b]) => b.profileId === profileId && this.busy(key)) || this.busy(`worker:${profileId}`);
  }
  async mutate<T>(action: (r: Registry) => T | Promise<T>, assertCurrent: () => void, admission: { beforePublication?: () => Promise<void>; automaticAttempt?: { profileId: string; generation: number; topologyDigest: string } } = {}): Promise<T> {
    const directory = join(this.root, "plugin-data", "microsoft-graph", "accounts");
    try { await createSecretFileAtomic({ rootDir: this.root, filePath: join(directory, ".initialized"), content: "graph-accounts-v1\n", mode: 0o600, dirMode: 0o700, durable: true }); } catch { /* Existing sentinel; secure reads/lock fail closed. */ }
    const lock = await acquireFileLock(this.file, { managerKey: "graph-accounts-v1", staleRecovery: "fail-closed", timeoutMs: 5000, staleMs: 300000, payload: () => ({ pid: process.pid, operation: "account-registry" }) });
    try {
      this.registry = await this.load();
      let next = structuredClone(this.registry);
      let value: T | undefined; let topologyError: unknown;
      try { value = await action(next); }
      catch (error) {
        if (!admission.automaticAttempt || !(error instanceof Error) || error.message !== "graph_single_user_topology_changed") throw error;
        // A detected transition only REDUCES authority. Discard all partially
        // mutated grants/connection state, serialize the tombstone under this
        // same registry lock, and do not depend on the cancelled actor's authority.
        next = structuredClone(this.registry);
        next.automatic ??= { ...admission.automaticAttempt, invalidated: true };
        next.automatic.invalidated = true;
        topologyError = error;
      }
      const publicationGuard = topologyError ? () => {} : assertCurrent;
      const verifyPublication = async () => {
        publicationGuard();
        if (!await lock.verifyStillHeld()) throw new Error("graph_account_store_unavailable");
        if (!topologyError) await admission.beforePublication?.();
        publicationGuard();
      };
      await verifyPublication();
      const content = JSON.stringify(next) + "\n";
      if (Buffer.byteLength(content) > 2 * 1024 * 1024) throw new Error("graph_account_store_capacity");
      try { await readSecureFile({ filePath: this.file, label: "Graph account registry", trust: { trustedDirs: [directory] }, io: { maxBytes: 2 * 1024 * 1024, timeoutMs: 5000 } }); }
      catch (error) {
        if (Object.keys(this.registry.accounts).length || Object.keys(this.registry.sessions).length) throw error;
        await createSecretFileAtomic({ rootDir: this.root, filePath: this.file, content, mode: 0o600, dirMode: 0o700, durable: true });
        this.registry = next; if (topologyError) throw topologyError; return structuredClone(value as T);
      }
      await replaceFileAtomic({ filePath: this.file, content, mode: 0o600, dirMode: 0o700, preserveExistingMode: false, copyFallbackOnPermissionError: false, destinationHardlinks: "reject", syncTempFile: true, syncParentDir: true, beforeRename: verifyPublication });
      this.registry = next; if (topologyError) throw topologyError; return structuredClone(value as T);
    } finally { await lock.release(); }
  }
  async provision(profileId: string, assertCurrent: () => void, requireEmptyRegistry = false): Promise<Account> {
    validProfile(profileId);
    return this.mutate(r => {
      this.assertPersonalOwner(profileId);
      if (requireEmptyRegistry && (Object.keys(r.accounts).length || Object.keys(r.sessions).length || r.legacyMigration || r.automatic)) throw new Error("graph_account_conflict");
      return r.accounts[profileKey(profileId)] ??= { profileId, generation: 1, revision: 1, connected: false, shared: false, unattended: false, policy: blankAccountPolicy() };
    }, assertCurrent);
  }
  async update(profileId: string, expectedRevision: number, policy: GraphPolicy, shared: boolean, unattended: boolean, assertCurrent: () => void): Promise<Account> {
    validProfile(profileId); validatePolicy(policy); assertPolicyCeiling(policy, this.ceilingFor(profileId));
    return this.mutate(r => {
      const a = r.accounts[profileKey(profileId)];
      if (!a || a.revision !== expectedRevision) throw new Error("graph_account_conflict");
      if (r.automatic?.profileId === profileId) r.automatic.invalidated = true;
      a.policy = structuredClone(policy); a.policySource = "personal"; a.shared = shared; a.unattended = unattended; a.revision++;
      return a;
    }, assertCurrent);
  }
  async markConnected(profileId: string, generation: number, assertCurrent: () => void): Promise<void> {
    await this.mutate(r => { const a = r.accounts[profileKey(profileId)]; if (!a || a.generation !== generation) throw new Error("graph_account_conflict"); a.connected = true; }, assertCurrent);
  }
  async disconnect(profileId: string, expectedRevision: number, assertCurrent: () => void): Promise<void> {
    await this.mutate(r => {
      const a = r.accounts[profileKey(profileId)]; if (!a || a.revision !== expectedRevision) throw new Error("graph_account_conflict");
      // Revocation may interrupt live work; generation checks stop all subsequent effects.
      if (r.automatic?.profileId === profileId) r.automatic.invalidated = true;
      a.generation++; a.revision++; a.connected = false; a.shared = false; a.unattended = false;
      for (const [key, b] of Object.entries(r.sessions)) if (b.profileId === profileId) delete r.sessions[key];
    }, assertCurrent);
  }
  async bind(actor: string, profileId: string, target: AccountHostContext, expectedRevision: number | null): Promise<Binding> {
    validProfile(actor); validProfile(profileId);
    if (this.workers[target.agentId]) throw new Error("graph_execution_account_conflict");
    return this.mutate(r => {
      const key = sessionKey(target); const old = r.sessions[key]; const a = r.accounts[profileKey(profileId)];
      if (!a?.connected || !a.shared) throw new Error("access_denied");
      if ((old?.revision ?? null) !== expectedRevision) throw new Error("graph_account_conflict");
      if (this.busy(key) || this.hostBusy(target)) throw new Error("graph_account_busy");
      const b: Binding = { agentId: target.agentId, sessionKey: target.sessionKey, sessionId: target.sessionId, lifecycleRevision: target.lifecycleRevision, profileId, generation: a.generation, revision: (old?.revision ?? 0) + 1, ...(!a.shared ? { privateOwner: profileId } : {}) };
      r.sessions[key] = b; return b;
    }, target.assertCurrent);
  }
  async status(target: AccountHostContext, actor: string): Promise<{ binding: Binding | null; available: string[]; workerProfileId?: string }> {
    await this.refresh(); target.assertCurrent();
    const b = this.registry.sessions[sessionKey(target)];
    return { ...(this.workers[target.agentId] ? { workerProfileId: this.workers[target.agentId] } : {}), binding: b ? structuredClone(b) : null, available: Object.values(this.registry.accounts).filter(a => a.connected && (a.profileId === actor || a.shared)).map(a => a.profileId) };
  }
  async resolve(target: AccountHostContext): Promise<{ scope: AccountScope; holdKey: string }> {
    await this.refresh(); target.assertCurrent();
    const key = sessionKey(target); const binding = this.registry.sessions[key]; const worker = this.workers[target.agentId];
    if (worker && binding && binding.profileId !== worker) throw new Error("graph_execution_account_conflict");
    const installation = this.registry.installation;
    if (this.installationMode !== "auto" && installation?.mode !== this.installationMode) throw new Error("graph_account_context_changed");
    // An initialization-only deployment contract, NOT a cached directory lease.
    // The host SDK exposes no canonical current-human identity on tool contexts.
    // Do not invent one from session owners, prompt text or requesterSenderId.
    const single = installation?.mode === "single-user" ? installation.profileId : undefined;
    if (single && (worker && worker !== single || binding && binding.profileId !== single)) throw new Error("graph_execution_account_conflict");
    const convenience = !!single && !worker && !binding;
    target.assertCurrent();
    const profileId = worker ?? binding?.profileId ?? (convenience ? single : undefined);
    if (!profileId) throw new Error("graph_execution_account_required");
    const a = this.registry.accounts[profileKey(profileId)];
    if (!a?.connected || (binding && binding.generation !== a.generation)) throw new Error("graph_execution_account_required");
    if (!convenience && (!a.shared || (worker && !a.unattended))) throw new Error("access_denied");
    const policy = single && a.policySource === "installation" ? intersectMigrationPolicy(this.ceiling, this.ceilingFor(profileId)) : a.policy;
    assertPolicyCeiling(policy, this.ceilingFor(profileId));
    const selection = worker ?? binding ?? installation;
    const signature = fingerprint([a, selection, installation]);
    const assertCurrent = () => {
      target.assertCurrent();
      const current = this.registry.accounts[profileKey(profileId)];
      const currentBinding = worker ?? this.registry.sessions[key] ?? (convenience ? this.registry.installation : undefined);
      if (fingerprint([current, currentBinding, this.registry.installation]) !== signature) throw new Error("graph_account_context_changed");
    };
    const holdKey = worker || convenience ? `worker:${profileId}` : key;
    return { holdKey, scope: { identity: fingerprint([profileId, a.generation, a.revision, selection, installation, policy]), stateDir: this.accountPath(profileId, a.generation), policy: structuredClone(policy), assertCurrent, assertFresh: async () => { await this.refresh(); assertCurrent(); } } };
  }
  hold(key: string): () => void { this.holds.set(key, (this.holds.get(key) ?? 0) + 1); let done = false; return () => { if (done) return; done = true; const n = (this.holds.get(key) ?? 1) - 1; if (n) this.holds.set(key, n); else this.holds.delete(key); }; }
  pendingApproval(id: string, key: string): void { this.pending.set(id, { target: key, expires: Date.now() + 180000 }); }
  finishApproval(id: string): void { this.pending.delete(id); }
  async run<T>(target: AccountHostContext, action: () => Promise<T>): Promise<T> {
    const { scope, holdKey } = await this.resolve(target); const release = this.hold(holdKey);
    try { return await accountScope.run(scope, async () => { const result = await action(); await scope.assertFresh(); scope.assertCurrent(); return result; }); } finally { release(); }
  }
}
