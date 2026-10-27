import { createHash, randomUUID } from "node:crypto";
import { join } from "node:path";
import { readSecureFile } from "@openclaw/fs-safe/secure-file";
import { createSecretFileAtomic } from "@openclaw/fs-safe/secret";
import { replaceFileAtomic } from "@openclaw/fs-safe/atomic";
import { acquireFileLock } from "@openclaw/fs-safe/file-lock";
import { readConfigFileSnapshotForWrite } from "openclaw/plugin-sdk/config-mutation";
import { canonicalPolicy } from "./config-policy-identity.js";
import { validatePolicy } from "./policy.js";
import { planChangedToolExposure, planPolicyToolReconciliation, requiredPolicyTools, planToolExposure, oneDriveToolExposureSaved } from "./control-ui-tool-exposure.js";

type Entries = Parameters<typeof planToolExposure>[1];
type Expected = ReturnType<typeof planToolExposure>["expected"];
type Snapshot = { hash: string; config: { tools?: { deny?: string[] }; agents?: { entries?: Entries }; plugins?: { entries?: Record<string, { config?: { policy?: unknown } }> } }; parsed?: unknown };
export type ActivationIntent = { version: 1; id: string; createdAt: number; targetHash: string; ownerHash: string; sourceHash: string; beforeToolsHash: string; required: Record<string, string[]>; state: "prepared" | "issued" | "complete" | "cancelled"; expected?: Expected };
const digest = (v: unknown) => createHash("sha256").update(canonicalPolicy(v)).digest("hex");
const policyHash = (v: unknown) => digest(validatePolicy(v));
const savedPolicy = (s: Snapshot) => s.config.plugins?.entries?.["microsoft-graph"]?.config?.policy;
function sourceHash(s: Snapshot) {
  // Bind authored include ownership without persisting paths or the policy.
  const plugins = (s.parsed as { plugins?: { $include?: unknown; entries?: Record<string, { config?: { policy?: unknown } }> } } | undefined)?.plugins;
  const policy = plugins?.entries?.["microsoft-graph"]?.config?.policy;
  return digest({ pluginsInclude: plugins?.$include ?? null, policyInclude: policy && typeof policy === "object" && "$include" in policy ? policy.$include : null });
}
function toolsHash(s: Snapshot, required: Record<string, string[]>) {
  return digest({ deny: s.config.tools?.deny ?? [], agents: Object.fromEntries(Object.keys(required).map(id => [id, s.config.agents?.entries?.[id] ?? null])) });
}
function parseRecord(v: unknown): ActivationIntent | null {
  if (v === null) return null;
  if (!v || typeof v !== "object" || Array.isArray(v)) throw Error("invalid_activation_record");
  const r = v as ActivationIntent;
  if (r.version !== 1 || typeof r.id !== "string" || !/^[a-f0-9-]{36}$/.test(r.id) || !Number.isSafeInteger(r.createdAt)
    || ![r.targetHash, r.ownerHash, r.sourceHash, r.beforeToolsHash].every(x => typeof x === "string" && /^[a-f0-9]{64}$/.test(x))
    || !["prepared", "issued", "complete", "cancelled"].includes(r.state) || !r.required || typeof r.required !== "object" || Array.isArray(r.required)
    || Object.entries(r.required).some(([id, names]) => !/^[a-zA-Z0-9_-]{1,128}$/.test(id) || !Array.isArray(names) || names.length > 32 || names.some(n => typeof n !== "string" || !/^(onedrive_[a-z_]+|outlook_(calendar|mail|contacts)_(read|write)|microsoft_todo_(read|write))$/.test(n)))) throw Error("invalid_activation_record");
  if (r.expected && (typeof r.expected !== "object" || Array.isArray(r.expected) || Object.entries(r.expected).some(([id, e]) => !Object.hasOwn(r.required, id) || !e || !["allow", "alsoAllow"].includes(e.key) || !Array.isArray(e.value) || e.value.some(n => typeof n !== "string")))) throw Error("invalid_activation_record");
  return r;
}
function missing(e: unknown): boolean {
  for (let i = 0; i < 6 && e && typeof e === "object"; i++, e = (e as { cause?: unknown }).cause) if (["not-found", "ENOENT"].includes(String((e as { code?: unknown }).code))) return true;
  return false;
}
/** Durable intent only, never credentials, resources or a policy draft. Native config.patch owns all config writes. */
export class ActivationJournal {
  private readonly file: string;
  private readonly directory: string;
  private readonly ownerHash: string;
  constructor(private stateDir: string, private snapshot: () => Promise<Snapshot> = async () => (await readConfigFileSnapshotForWrite()).snapshot as unknown as Snapshot, owner: string = "", private projectHash: (rawHash: string) => string = hash => hash) {
    if (!owner) throw Error("activation_identity_required");
    this.ownerHash = digest({ authority: "operator.admin", owner });
    this.directory = join(stateDir, "plugin-data", "microsoft-graph", "activation"); this.file = join(this.directory, "pending.json");
  }
  private async read(): Promise<ActivationIntent | null> {
    try {
      const value = await readSecureFile({ filePath: this.file, label: "Microsoft Graph activation intent", trust: { trustedDirs: [this.directory] }, io: { maxBytes: 64 * 1024, timeoutMs: 5000 } });
      if (process.platform !== "win32" && (value.stat.mode & 0o777) !== 0o600) throw Error("activation_permissions");
      return parseRecord(JSON.parse(value.buffer.toString("utf8")));
    } catch (e) { if (missing(e)) return null; throw e; }
  }
  private async transaction<T>(fn: (current: ActivationIntent | null, write: (record: ActivationIntent) => Promise<void>) => Promise<T>): Promise<T> {
    try { await createSecretFileAtomic({ rootDir: this.stateDir, filePath: this.file, content: "null\n", mode: 0o600, dirMode: 0o700, durable: true }); }
    catch { await this.read(); }
    const lock = await acquireFileLock(this.file, { managerKey: "microsoft-graph-activation", staleRecovery: "fail-closed", staleMs: 60000, timeoutMs: 5000, payload: () => ({ pid: process.pid, operation: "activation-intent" }) });
    try {
      return await fn(await this.read(), async record => {
        parseRecord(record);
        if (Buffer.byteLength(JSON.stringify(record)) > 64 * 1024) throw Error("activation_record_too_large");
        await replaceFileAtomic({ filePath: this.file, content: JSON.stringify(record) + "\n", mode: 0o600, dirMode: 0o700, preserveExistingMode: false, destinationHardlinks: "reject", copyFallbackOnPermissionError: false, syncTempFile: true, syncParentDir: true, beforeRename: async () => { if (!await lock.verifyStillHeld()) throw Error("activation_lock_lost"); } });
      });
    } finally { await lock.release(); }
  }
  async prepare(baseHash: string, baseline: unknown, target: unknown) {
    const before = validatePolicy(baseline); const after = validatePolicy(target);
    return this.transaction(async (current, write) => {
      if (current && current.ownerHash !== this.ownerHash && ["prepared", "issued"].includes(current.state) && Date.now() - current.createdAt <= 24 * 3600000) throw Error("activation_owner_conflict");
      const s = await this.snapshot();
      if (this.projectHash(s.hash) !== baseHash || policyHash(savedPolicy(s)) !== policyHash(before)) throw Error("activation_conflict");
      const plan = planChangedToolExposure(before, after, s.config.agents?.entries, s.config.tools?.deny);
      const required = Object.fromEntries(Object.entries(plan.expected).map(([id, e]) => [id, e.value.filter(n => !(s.config.agents?.entries?.[id]?.tools?.[e.key] ?? []).includes(n))]));
      if (!Object.keys(required).length) throw Error("activation_no_tools");
      const record: ActivationIntent = { version: 1, ownerHash: this.ownerHash, id: randomUUID(), createdAt: Date.now(), targetHash: policyHash(after), sourceHash: sourceHash(s), beforeToolsHash: toolsHash(s, required), required, state: "prepared" };
      await write(record); return record;
    });
  }
  /** Single claimed repair: only the already-saved, active policy is authority.
   * Persist the exact target before replying so fresh pages reconcile, not replay. */
  async reconcile(baseHash: string, activeHash: string | undefined) {
    return this.transaction(async (current, write) => {
      if (current && ["prepared", "issued"].includes(current.state) && Date.now() - current.createdAt <= 24 * 3600000) throw Error("activation_pending");
      const s = await this.snapshot();
      const policy = validatePolicy(savedPolicy(s)); const targetHash = policyHash(policy);
      if (this.projectHash(s.hash) !== baseHash || activeHash !== targetHash) throw Error("activation_conflict");
      const plan = planPolicyToolReconciliation(policy, s.config.agents?.entries, s.config.tools?.deny);
      if (!Object.keys(plan.patch).length) return null;
      const desired = requiredPolicyTools(policy);
      const required = Object.fromEntries(Object.keys(plan.patch).map(agent => [agent, desired[agent] ?? []]));
      const record: ActivationIntent = { version: 1, ownerHash: this.ownerHash, id: randomUUID(), createdAt: Date.now(), targetHash, sourceHash: sourceHash(s), beforeToolsHash: toolsHash(s, required), required, expected: plan.expected, state: "issued" };
      await write(record);
      return { record, patch: plan.patch, expected: plan.expected, baseHash };
    });
  }
  async status() {
    const record = await this.read();
    if (!record || record.ownerHash !== this.ownerHash || ["complete", "cancelled"].includes(record.state) || Date.now() - record.createdAt > 24 * 3600000) return null;
    const s = await this.snapshot();
    if (policyHash(savedPolicy(s)) !== record.targetHash || sourceHash(s) !== record.sourceHash) return null;
    return record;
  }
  async claim(id: string, activeHash: string | undefined) {
    return this.transaction(async (record, write) => {
      if (!record || record.ownerHash !== this.ownerHash || record.id !== id || record.state !== "prepared" || Date.now() - record.createdAt > 24 * 3600000) throw Error("activation_not_claimable");
      const s = await this.snapshot();
      if (activeHash !== record.targetHash || policyHash(savedPolicy(s)) !== record.targetHash || sourceHash(s) !== record.sourceHash || toolsHash(s, record.required) !== record.beforeToolsHash) throw Error("activation_conflict");
      const plan = planToolExposure(record.required, s.config.agents?.entries, s.config.tools?.deny);
      const baseHash = this.projectHash(s.hash);
      record.state = "issued"; record.expected = plan.expected;
      // Claim before dispatch. A lost response cannot authorize a second write;
      // fresh hosts only reconcile it or offer explicit reviewed recovery.
      await write(record); return { record, ...plan, baseHash };
    });
  }
  async cancel(id: string) {
    return this.transaction(async (record, write) => {
      if (!record || record.ownerHash !== this.ownerHash || record.id !== id) throw Error("activation_conflict");
      record.state = "cancelled"; await write(record); return { cancelled: true };
    });
  }
  async complete(id: string, activeHash: string | undefined) {
    return this.transaction(async (record, write) => {
      if (!record || record.ownerHash !== this.ownerHash || record.id !== id || record.state !== "issued" || !record.expected) throw Error("activation_conflict");
      const s = await this.snapshot();
      if (activeHash !== record.targetHash || policyHash(savedPolicy(s)) !== record.targetHash || !oneDriveToolExposureSaved(s.config.agents?.entries, record.expected)) throw Error("activation_conflict");
      record.state = "complete"; await write(record); return { complete: true };
    });
  }
}
