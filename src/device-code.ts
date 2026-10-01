import { randomUUID } from "node:crypto";
import { acquireVaultLock, createVaultCredential, decodeVaultKey, inspectVaultCredential } from "./credential-vault.js";
import { requiredPolicyScopes } from "./credential-cli.js";
import { selectScope } from "./credential.js";
import { validatePolicy, type GraphPolicy } from "./policy.js";

const MAX_LIFETIME_MS = 15 * 60_000;
const MAX_RESPONSE_BYTES = 64 * 1024;
const SAFE_TENANT = /^(?:[0-9a-fA-F-]{36}|[A-Za-z0-9.-]{1,253})$/;
const SAFE_CLIENT = /^[0-9a-fA-F-]{36}$/;
const SAFE_SESSION = /^[0-9a-fA-F-]{36}$/;
const sleep = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms));

type DeviceSession = { id: string; grantedScopes?: string[]; state: "pending" | "created" | "failed"; committing?: boolean; expiresAt: number; error?: string; interval: number; deviceCode: string; userCode: string; verificationUri: string; clientId: string; tenant: string; scopes: string[] };
type StartResult = { sessionId: string; userCode: string; verificationUri: string; expiresAt: string; scopes: string[] };
type StatusResult = { state: "pending" | "created" | "failed"; error?: string; scopes?: string[] };

async function oauthPost(url: string, body: URLSearchParams, fetchFn: typeof fetch, timeoutMs = 10_000): Promise<{ status: number; value: Record<string, unknown> }> {
  const response = await fetchFn(url, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body, signal: AbortSignal.timeout(timeoutMs), redirect: "error" });
  const raw = await response.text();
  if (Buffer.byteLength(raw) > MAX_RESPONSE_BYTES) throw new Error("device_authorization_failed");
  let value: unknown;
  try { value = JSON.parse(raw); } catch { throw new Error("device_authorization_failed"); }
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("device_authorization_failed");
  return { status: response.status, value: value as Record<string, unknown> };
}
function checkedScopes(policy: GraphPolicy): string[] { return [...new Set([...requiredPolicyScopes(validatePolicy(policy)), "offline_access"])].sort(); }
function endpoint(tenant: string, suffix: string): string { return `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/${suffix}`; }

export class DeviceCodeSignIn {
  private session?: DeviceSession;
  private starting = false;
  constructor(private readonly config: { policy?: GraphPolicy; credentialVaultKey?: unknown }, private readonly stateDir: () => string, private readonly fetchFn: typeof fetch = fetch, private readonly wait: (ms: number) => Promise<void> = sleep) {}

  async start(clientId: string, tenant: string): Promise<StartResult> {
    if (!SAFE_CLIENT.test(clientId) || !SAFE_TENANT.test(tenant) || tenant.includes("..")) throw new Error("invalid_rpc_parameters");
    if (this.starting) throw new Error("device_authorization_in_progress");
    if (this.session?.state === "pending" && this.session.expiresAt > Date.now()) {
      if (this.session.clientId !== clientId || this.session.tenant !== tenant) throw new Error("device_authorization_in_progress");
      return { sessionId: this.session.id, userCode: this.session.userCode, verificationUri: this.session.verificationUri, expiresAt: new Date(this.session.expiresAt).toISOString(), scopes: this.session.scopes };
    }
    this.starting = true;
    try {
      const key = this.config.credentialVaultKey;
      if (typeof key !== "string") throw new Error("credential_vault_unavailable");
      decodeVaultKey(key);
      const policy = validatePolicy(this.config.policy);
      const scopes = checkedScopes(policy);
      if (scopes.length < 2) throw new Error("invalid_policy");
      if ((await inspectVaultCredential(this.stateDir(), key)).result !== "missing") throw new Error("credential_vault_conflict");
      const result = await oauthPost(endpoint(tenant, "devicecode"), new URLSearchParams({ client_id: clientId, scope: scopes.join(" ") }), this.fetchFn);
      const value = result.value;
      if (result.status !== 200 || typeof value.device_code !== "string" || !value.device_code || value.device_code.length > 4096
        || typeof value.user_code !== "string" || !/^[A-Za-z0-9-]{4,32}$/.test(value.user_code)
        || typeof value.verification_uri !== "string" || !["https://microsoft.com/devicelogin", "https://www.microsoft.com/devicelogin"].includes(value.verification_uri)
        || !Number.isSafeInteger(value.expires_in) || Number(value.expires_in) < 30 || Number(value.expires_in) > 3600
        || !Number.isSafeInteger(value.interval) || Number(value.interval) < 1 || Number(value.interval) > 60) throw new Error("device_authorization_failed");
      const now = Date.now();
      const session: DeviceSession = { id: randomUUID(), state: "pending", expiresAt: now + Math.min(Number(value.expires_in) * 1000, MAX_LIFETIME_MS), interval: Number(value.interval) * 1000, deviceCode: value.device_code, userCode: value.user_code, verificationUri: value.verification_uri, clientId, tenant, scopes };
      this.session = session;
      void this.poll(session, key).catch(() => { if (this.session === session) { session.state = "failed"; session.error = "device_authorization_failed"; session.deviceCode = ""; session.userCode = ""; } });
      return { sessionId: session.id, userCode: value.user_code, verificationUri: value.verification_uri, expiresAt: new Date(session.expiresAt).toISOString(), scopes };
    } finally { this.starting = false; }
  }

  cancel(sessionId: string): StatusResult {
    if (!SAFE_SESSION.test(sessionId) || !this.session || this.session.id !== sessionId) throw new Error("device_session_unavailable");
    if (this.session.committing) throw new Error("device_authorization_in_progress");
    if (this.session.state === "pending") { this.session.state = "failed"; this.session.error = "device_authorization_cancelled"; this.session.deviceCode = ""; this.session.userCode = ""; }
    return this.status(sessionId);
  }

  status(sessionId: string): StatusResult {
    if (!SAFE_SESSION.test(sessionId) || !this.session || this.session.id !== sessionId) throw new Error("device_session_unavailable");
    if (this.session.state === "pending" && Date.now() >= this.session.expiresAt) { this.session.state = "failed"; this.session.error = "device_authorization_expired"; this.session.deviceCode = ""; this.session.userCode = ""; }
    return { state: this.session.state, ...(this.session.error ? { error: this.session.error } : {}), ...(this.session.grantedScopes ? { scopes: this.session.grantedScopes } : {}) };
  }

  private async poll(session: DeviceSession, key: string): Promise<void> {
    while (this.session === session && session.state === "pending" && Date.now() < session.expiresAt) {
      await this.wait(session.interval);
      if (this.session !== session || session.state !== "pending" || Date.now() >= session.expiresAt) break;
      let response: { status: number; value: Record<string, unknown> };
      try { response = await oauthPost(endpoint(session.tenant, "token"), new URLSearchParams({ grant_type: "urn:ietf:params:oauth:grant-type:device_code", client_id: session.clientId, device_code: session.deviceCode }), this.fetchFn); }
      catch { session.state = "failed"; session.error = "device_authorization_failed"; break; }
      if (this.session !== session || session.state !== "pending" || Date.now() >= session.expiresAt) break;
      const value = response.value;
      if (response.status !== 200) {
        if (value.error === "authorization_pending") continue;
        if (value.error === "slow_down") { session.interval = Math.min(session.interval + 5000, 60_000); continue; }
        session.state = "failed";
        session.error = value.error === "authorization_declined" ? "device_authorization_declined" : "device_authorization_failed";
        break;
      }
      try {
        if (typeof value.refresh_token !== "string" || !value.refresh_token || value.refresh_token.length > 32 * 1024
          || typeof value.scope !== "string" || value.scope.length > 8192) throw new Error("device_authorization_failed");
        const granted = [...new Set(value.scope.split(/\s+/).filter(Boolean))];
        const credential = { clientId: session.clientId, tenant: session.tenant, refreshToken: value.refresh_token, scopes: granted.some((scope) => scope.toLowerCase() === "offline_access") ? granted : [...granted, "offline_access"] };
        for (const scope of session.scopes.filter((scope) => scope !== "offline_access")) selectScope(credential, [scope]);
        session.committing = true;
        const lock = await acquireVaultLock(this.stateDir(), "device-code-sign-in");
        try {
          if ((await inspectVaultCredential(this.stateDir(), key)).result !== "missing") throw new Error("credential_vault_conflict");
          if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
          await createVaultCredential(this.stateDir(), credential, key);
        } finally { await lock.release().catch(() => undefined); }
        session.grantedScopes = credential.scopes;
        session.state = "created";
      } catch (error) {
        session.state = "failed";
        session.error = error instanceof Error && ["credential_scope_missing", "credential_vault_conflict"].includes(error.message) ? error.message : "device_authorization_failed";
      }
      break;
    }
    if (session.state === "pending") { session.state = "failed"; session.error = "device_authorization_expired"; }
    session.deviceCode = ""; session.userCode = "";
  }
}
