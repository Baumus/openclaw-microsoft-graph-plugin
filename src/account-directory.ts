import { createHash } from "node:crypto";
import { resolveGatewayPort, type OpenClawPluginApi } from "openclaw/plugin-sdk/core";
import { GatewayClient, resolveGatewayAuth } from "openclaw/plugin-sdk/gateway-runtime";

/** A fresh whole-host snapshot, NOT an atomic catalog lease. */
export type DirectorySnapshot = { soleProfileId?: string; digest: string };
export type AccountDirectory = { read: (signal?: AbortSignal) => Promise<DirectorySnapshot> };
// These exact releases share the verified public roster/auth/SDK contracts.
// A future release is not silently assumed compatible.
const directoryRuntimeVersions = new Set(["2026.10.1", "2026.10.2"]);
const transportEnvironment = () => [process.env.OPENCLAW_GATEWAY_PORT, process.env.OPENCLAW_PROFILE, process.env.OPENCLAW_GATEWAY_TOKEN, process.env.OPENCLAW_GATEWAY_PASSWORD];
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
  // Stable key ordering prevents RPC object ordering changes from revoking grants.
  const projection = rows.sort((a, b) => a.id.localeCompare(b.id)).map(row => Object.fromEntries(Object.keys(row).sort().map(k => [k, row[k]])));
  return { ...(humans.length === 1 ? { soleProfileId: humans[0].id } : {}), digest: createHash("sha256").update(JSON.stringify(projection)).digest("hex") };
}

/** Public SDK client; never restricted runtime dispatch or private host storage.
 * A new read-only connection for each snapshot avoids cached auth/roster leases.
 */
export function createAccountDirectory(api: OpenClawPluginApi, hostInstanceId: string, Client: typeof GatewayClient = GatewayClient): AccountDirectory {
  return { async read(signal?: AbortSignal) {
    let client: GatewayClient | undefined;
    let connectTimer: ReturnType<typeof setTimeout> | undefined;
    let rejectConnect: (() => void) | undefined;
    const cancelConnect = () => rejectConnect?.();
    try {
      signal?.throwIfAborted();
      if (!directoryRuntimeVersions.has(api.runtime.version)) throw fail();
      const config = api.runtime.config.current();
      const gateway = config.gateway;
      const port = resolveGatewayPort(config as Parameters<typeof resolveGatewayPort>[0], process.env);
      // Keep the current runtime snapshot's SecretRef authority metadata intact.
      const auth = resolveGatewayAuth({ authConfig: gateway?.auth as Parameters<typeof resolveGatewayAuth>[0]["authConfig"], env: process.env, tailscaleMode: gateway?.tailscale?.mode });
      if ([gateway?.auth?.token, gateway?.auth?.password].some(value => value !== undefined && typeof value !== "string")) throw fail();
      const secret = auth.mode === "token" ? auth.token : auth.mode === "password" ? auth.password : undefined;
      if (gateway?.mode === "remote" || gateway?.tls?.enabled || ![undefined, "loopback", "lan", "auto"].includes(gateway?.bind) || typeof secret !== "string" || !secret.trim() || !Number.isSafeInteger(port) || port < 1 || port > 65535) throw fail();
      const identity = () => JSON.stringify([api.runtime.config.current().gateway, transportEnvironment()]);
      const before = identity();
      let closed = false, hellos = 0;
      const connected = new Promise<void>((resolve, reject) => {
        rejectConnect = () => reject(fail());
        // Ordinary local CLI/shared-auth admission, NOT the reserved backend
        // identity. No device identity, stored token, write, argv or subprocess.
        client = new Client({
          url: `ws://127.0.0.1:${port}`, [auth.mode]: secret,
          clientName: "cli", mode: "cli", role: "operator", scopes: ["operator.read"],
          deviceIdentity: null, sharedStateMode: "read-only", env: {},
          hostDeps: { logDebug() {}, logError() {} },
          connectChallengeTimeoutMs: 5000, requestTimeoutMs: 5000,
          onHelloOk: () => { if (++hellos !== 1 || closed) { closed = true; client?.stop(); rejectConnect?.(); } else resolve(); },
          onConnectError: () => { closed = true; rejectConnect?.(); },
          onClose: () => { closed = true; rejectConnect?.(); },
        });
        connectTimer = setTimeout(() => rejectConnect?.(), 5000);
        signal?.addEventListener("abort", cancelConnect, { once: true });
        if (signal?.aborted) rejectConnect(); else client.start();
      });
      await connected;
      clearTimeout(connectTimer); signal?.removeEventListener("abort", cancelConnect);
      const request = async (method: string, requestSignal?: AbortSignal) => {
        if (closed || hellos !== 1 || before !== identity()) throw fail();
        const value = await client!.request(method, {}, { signal: requestSignal, timeoutMs: 5000 });
        if (closed || hellos !== 1 || before !== identity()) throw fail();
        const raw = JSON.stringify(value);
        if (typeof raw !== "string" || Buffer.byteLength(raw) > 1024 * 1024) throw fail();
        return raw;
      };
      const proveHost = async (requestSignal?: AbortSignal) => {
        const value = JSON.parse(await request("microsoft-graph.directory-host", requestSignal));
        if (value?.hostInstanceId !== hostInstanceId || Object.keys(value).join() !== "hostInstanceId") throw fail();
      };
      await proveHost(signal); signal?.throwIfAborted();
      const snapshot = parseDirectorySnapshot(await request("users.list", signal));
      // A fully observed roster still receives its bounded safety after-proof
      // during drain. Positive consumers must recheck their own live authority.
      await proveHost();
      return snapshot;
    } catch { throw fail(); } // No roster, auth, transport errors or SDK logs escape.
    finally {
      clearTimeout(connectTimer); signal?.removeEventListener("abort", cancelConnect);
      if (client) { try { await client.stopAndWait({ timeoutMs: 1000 }); } catch { throw fail(); } }
    }
  } };
}
