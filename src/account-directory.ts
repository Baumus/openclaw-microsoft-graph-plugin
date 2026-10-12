import { createHash } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import type { OpenClawPluginApi } from "openclaw/plugin-sdk/core";

/** A fresh whole-host snapshot, NOT an atomic catalog lease. */
export type DirectorySnapshot = { soleProfileId?: string; digest: string };
export type AccountDirectory = { read: (signal?: AbortSignal) => Promise<DirectorySnapshot> };
const fail = () => new Error("graph_single_user_directory_unavailable");
const fields = new Set(["id", "displayName", "avatarMime", "mergedInto", "role", "createdAt", "updatedAt", "emails", "githubIdentity", "hasAvatar", "effectiveRole", "roleSource"]);
const idValid = (v: unknown): v is string => typeof v === "string" && !!v && v.length <= 256 && !/[\u0000-\u001f\u007f]/.test(v);
export function parseDirectorySnapshot(raw: string): DirectorySnapshot {
  let value: unknown; try { value = JSON.parse(raw); } catch { throw fail(); }
  if (!value || typeof value !== "object" || Array.isArray(value) || Object.keys(value).join() !== "profiles") throw fail();
  const profiles = (value as { profiles: unknown }).profiles;
  if (!Array.isArray(profiles) || profiles.length > 10000) throw fail();
  const ids = new Set<string>();
  const rows = profiles.map(row => {
    if (!row || typeof row !== "object" || Array.isArray(row) || Object.keys(row).some(k => !fields.has(k)) || !idValid(row.id) || ids.has(row.id)) throw fail();
    ids.add(row.id);
    if (!(row.mergedInto === null || idValid(row.mergedInto)) || !Number.isSafeInteger(row.createdAt) || !Number.isSafeInteger(row.updatedAt) || row.createdAt < 0 || row.updatedAt < row.createdAt || !Array.isArray(row.emails) || row.emails.some((v: unknown) => typeof v !== "string" || v.length > 512) || typeof row.hasAvatar !== "boolean" || !Object.hasOwn(row, "githubIdentity") || !["default", "assigned", "githubLogin"].includes(row.roleSource)) throw fail();
    if (!(row.displayName === null || typeof row.displayName === "string") || !(row.avatarMime === null || typeof row.avatarMime === "string") || (row.role !== undefined && !idValid(row.role)) || (row.effectiveRole !== undefined && !idValid(row.effectiveRole))) throw fail();
    if (row.githubIdentity !== null) {
      const identity = row.githubIdentity;
      if (!identity || typeof identity !== "object" || Array.isArray(identity) || Object.keys(identity).sort().join() !== "avatarUrl,login,profileUrl" || !idValid(identity.login) || typeof identity.profileUrl !== "string" || typeof identity.avatarUrl !== "string") throw fail();
    }
    // No directory PII is persisted or returned, only the fingerprint.
    return { ...row, emails: [...row.emails].sort() };
  });
  const byId = new Map(rows.map(row => [row.id, row]));
  for (const row of rows) if (row.mergedInto !== null) {
    const seen = new Set<string>([row.id]); let current = row;
    while (current.mergedInto !== null) { if (seen.has(current.mergedInto)) throw fail(); seen.add(current.mergedInto); current = byId.get(current.mergedInto); if (!current) throw fail(); }
    if (current.id === "gateway-owner") throw fail();
  }
  const humans = rows.filter(row => row.mergedInto === null && row.id !== "gateway-owner");
  // Stable key ordering prevents CLI object ordering changes from revoking grants.
  const projection = rows.sort((a, b) => a.id.localeCompare(b.id)).map(row => Object.fromEntries(Object.keys(row).sort().map(k => [k, row[k]])));
  return { ...(humans.length === 1 ? { soleProfileId: humans[0].id } : {}), digest: createHash("sha256").update(JSON.stringify(projection)).digest("hex") };
}

/** Public CLI entry and public process runner; never private dispatch or host DB. */
export function createAccountDirectory(api: OpenClawPluginApi, hostInstanceId: string): AccountDirectory {
  return { async read(signal?: AbortSignal) {
    let scratch: string | undefined;
    try {
      signal?.throwIfAborted();
      if (api.runtime.version !== "2026.10.1") throw fail();
      const config = api.runtime.config.current();
      const gateway = config.gateway;
      const port = process.env.OPENCLAW_GATEWAY_PORT ? Number(process.env.OPENCLAW_GATEWAY_PORT) : gateway?.port;
      const token = process.env.OPENCLAW_GATEWAY_TOKEN || gateway?.auth?.token;
      // Exact local token-auth shape only. Remote/TLS/custom/anonymous/trusted-proxy
      // paths keep explicit migration instead of borrowing a foreign identity.
      if (gateway?.mode === "remote" || gateway?.tls?.enabled || ![undefined, "loopback"].includes(gateway?.bind) || gateway?.auth?.mode !== "token" || typeof token !== "string" || !token.trim() || !Number.isSafeInteger(port) || port! < 1 || port! > 65535 || !api.runtime.system?.runCommandWithTimeout) throw fail();
      const entry = fileURLToPath(import.meta.resolve("openclaw/cli-entry"));
      // The public CLI may initialize its own runtime state even for a read.
      // Give it private disposable scratch state, never the host's real store.
      scratch = await mkdtemp(join(tmpdir(), "msgraph-users-read-"));
      const env: NodeJS.ProcessEnv = {};
      for (const key of ["HOME", "USERPROFILE", "SYSTEMROOT", "WINDIR", "XDG_CONFIG_HOME"]) if (process.env[key]) env[key] = process.env[key];
      env.OPENCLAW_STATE_DIR = scratch;
      env.OPENCLAW_CONFIG_PATH = join(scratch, "absent-config.json");
      env.OPENCLAW_GATEWAY_URL = `ws://127.0.0.1:${port}`;
      env.OPENCLAW_GATEWAY_TOKEN = token;
      const before = JSON.stringify([gateway, process.env.OPENCLAW_GATEWAY_PORT, process.env.OPENCLAW_GATEWAY_TOKEN]);
      const run = (args: string[], commandSignal: AbortSignal | null | undefined = signal) => api.runtime.system.runCommandWithTimeout([process.execPath, entry, ...args], { baseEnv: env, signal: commandSignal ?? undefined, timeoutMs: 10000, maxOutputBytes: { stdout: 1024 * 1024, stderr: 4096 }, maxCombinedOutputBytes: 1024 * 1024 + 4096, terminateOnOutputLimit: true, killProcessTree: true, requireProcessTreeExtinction: true });
      const valid = (result: Awaited<ReturnType<typeof run>>) => result.code === 0 && !result.killed && !result.signal && result.termination === "exit" && !result.stdoutTruncatedBytes && !result.stderrTruncatedBytes && !result.outputLimitExceeded && result.cleanup !== "uncertain";
      const proveHost = async (safetyCompletion = false) => {
        const proof = await run(["gateway", "call", "microsoft-graph.directory-host", "--json", "--timeout", "5000"], safetyCompletion ? null : signal);
        if (!valid(proof)) throw fail();
        const value = JSON.parse(proof.stdout);
        if (value?.hostInstanceId !== hostInstanceId || Object.keys(value).join() !== "hostInstanceId") throw fail();
      };
      await proveHost(); signal?.throwIfAborted();
      const result = await run(["users", "list", "--json", "--timeout", "5000"]);
      if (before !== JSON.stringify([api.runtime.config.current().gateway, process.env.OPENCLAW_GATEWAY_PORT, process.env.OPENCLAW_GATEWAY_TOKEN]) || !valid(result)) throw fail();
      const snapshot = parseDirectorySnapshot(result.stdout);
      // Once a complete roster is observed, finish its bounded after-proof even
      // during drain. This returns safety evidence only: positive callers must
      // still check their live authority after awaiting the snapshot.
      await proveHost(true);
      if (before !== JSON.stringify([api.runtime.config.current().gateway, process.env.OPENCLAW_GATEWAY_PORT, process.env.OPENCLAW_GATEWAY_TOKEN])) throw fail();
      return snapshot;
    } catch { throw fail(); } // Never expose subprocess stderr, roster, or shared auth.
    finally { if (scratch) await rm(scratch, { recursive: true, force: true }); }
  } };
}
