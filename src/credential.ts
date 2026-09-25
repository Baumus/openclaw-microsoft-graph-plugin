import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { decodeVaultKey, refreshVaultCredential, type VaultBinding } from "./credential-vault.js";
import type { GraphPolicy } from "./policy.js";

const execFileAsync = promisify(execFile);
const ACCESS_TOKEN_EXPIRY_SKEW_MS = 60_000;
const ACCESS_TOKEN_CACHE_MAX_RESIDENCY_MS = 60 * 60_000;
const ACCESS_TOKEN_CACHE_MAX_ENTRIES = 128;
const ACCESS_TOKEN_MAX_LENGTH = 16 * 1024;
const ACCESS_TOKEN_PATTERN = /^[A-Za-z0-9._~+\/-]+=*$/;
const EXPIRES_IN_MAX_SECONDS = 86_400;

type CachedAccessToken = { token: string; expiresAt: number };
export type TokenExchangeResult = { accessToken: string; replacementRefreshToken?: string; expiresAt?: number };
type AccessTokenRefresh = { promise: Promise<TokenExchangeResult>; controller: AbortController; waiters: number; settled: boolean };
const accessTokenCache = new Map<string, CachedAccessToken>();
const accessTokenRefreshes = new Map<string, AccessTokenRefresh>();

function accessTokenCacheKey(credential: Credential, requestedScopes: string[], durableIdentity?: string): string {
  const refreshFingerprint = durableIdentity ?? createHash("sha256").update(credential.refreshToken).digest("base64url");
  return createHash("sha256").update(JSON.stringify({
    tenant: credential.tenant.toLowerCase(),
    clientId: credential.clientId,
    scopes: [...new Set(requestedScopes.map((scope) => scope.toLowerCase()))].sort(),
    refreshFingerprint,
  })).digest("base64url");
}

function pruneAccessTokenCache(now = Date.now()): void {
  for (const [key, cached] of accessTokenCache) if (cached.expiresAt <= now) accessTokenCache.delete(key);
}

function cachedAccessToken(key: string): CachedAccessToken | undefined {
  pruneAccessTokenCache();
  const cached = accessTokenCache.get(key);
  if (!cached) return undefined;
  accessTokenCache.delete(key);
  accessTokenCache.set(key, cached);
  return cached;
}

function cacheAccessToken(key: string, token: string, expiresAt: number): void {
  pruneAccessTokenCache();
  if (!Number.isFinite(expiresAt) || expiresAt <= Date.now()) return;
  accessTokenCache.delete(key);
  while (accessTokenCache.size >= ACCESS_TOKEN_CACHE_MAX_ENTRIES) {
    const oldest = accessTokenCache.keys().next().value;
    if (oldest === undefined) break;
    accessTokenCache.delete(oldest);
  }
  accessTokenCache.set(key, { token, expiresAt });
}

async function waitForToken(refresh: AccessTokenRefresh, signal?: AbortSignal): Promise<TokenExchangeResult> {
  signal?.throwIfAborted();
  refresh.waiters += 1;
  let abort: (() => void) | undefined;
  try {
    if (!signal) return await refresh.promise;
    return await Promise.race([
      refresh.promise,
      new Promise<never>((_resolve, reject) => {
        abort = () => reject(signal.reason);
        signal.addEventListener("abort", abort, { once: true });
      }),
    ]);
  } finally {
    if (abort) signal?.removeEventListener("abort", abort);
    refresh.waiters -= 1;
    if (refresh.waiters === 0 && !refresh.settled) refresh.controller.abort();
  }
}

/** Test-only reset for the process-local secret cache; production never serializes it. */
export function clearAccessTokenCacheForTests(): void {
  accessTokenCache.clear();
  accessTokenRefreshes.clear();
}

/** Sanitized test-only cache observability; never returns token material or cache identities. */
export function accessTokenCacheStateForTests(): { entries: number; refreshes: number } {
  pruneAccessTokenCache();
  return { entries: accessTokenCache.size, refreshes: accessTokenRefreshes.size };
}

export type Credential = { clientId: string; refreshToken: string; tenant: string; scopes: string[] };

export function parseCredential(raw: string): Credential {
  const trimmed = raw.trim();
  if (!trimmed) throw new Error("credential_unavailable");
  let value: Record<string, unknown> = {};
  try {
    const parsed = JSON.parse(trimmed) as unknown;
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) value = parsed as Record<string, unknown>;
  } catch {
    for (const line of trimmed.split(/\r?\n/)) {
      const match = line.match(/^\s*([^:=]+)\s*[:=]\s*(.*?)\s*$/);
      if (match) value[match[1]] = match[2];
    }
  }
  const clientId = value.clientId ?? value.client_id;
  const refreshToken = value.refreshToken ?? value.refresh_token;
  const tenant = value.tenant ?? "common";
  const rawScopes = value.scopes;
  if (typeof clientId !== "string" || !clientId || typeof refreshToken !== "string" || !refreshToken || rawScopes === undefined) {
    throw new Error("credential_unavailable");
  }
  const scopes = Array.isArray(rawScopes) ? rawScopes.map(String) : String(rawScopes).split(/\s+/);
  return { clientId, refreshToken, tenant: String(tenant), scopes: [...new Set(scopes.filter(Boolean).concat("offline_access"))] };
}

export async function readCredential(secretRef: string, signal?: AbortSignal): Promise<Credential> {
  if (!/^[A-Za-z0-9._/@+-]+$/.test(secretRef)) throw new Error("invalid_secret_reference");
  signal?.throwIfAborted();
  try {
    const { stdout } = await execFileAsync("pass", ["show", secretRef], { encoding: "utf8", timeout: 10_000, maxBuffer: 64 * 1024, signal });
    return parseCredential(stdout);
  } catch {
    if (signal?.aborted) throw signal.reason;
    throw new Error("credential_unavailable");
  }
}

export function assertScopes(credential: Credential, anyOf: string[]): void {
  const present = new Set(credential.scopes.map((scope) => scope.toLowerCase()));
  if (!anyOf.some((scope) => present.has(scope.toLowerCase()))) throw new Error("credential_scope_missing");
}

export function selectScope(credential: Credential, allowed: string[]): string {
  const present = new Set(credential.scopes.map((scope) => scope.toLowerCase()));
  const impliedBy: Record<string, string[]> = {
    "files.read": ["files.readwrite"],
    "calendars.read": ["calendars.readwrite"],
    "mail.read": ["mail.readwrite"],
    "tasks.read": ["tasks.readwrite"],
  };
  const selected = allowed.find((scope) => present.has(scope.toLowerCase()) || (impliedBy[scope.toLowerCase()] ?? []).some((superset) => present.has(superset)));
  if (!selected) throw new Error("credential_scope_missing");
  return selected;
}

async function boundedJson(response: Response, maximum = 64 * 1024): Promise<Record<string, unknown>> {
  const length = Number(response.headers.get("content-length"));
  if (Number.isFinite(length) && length > maximum) throw new Error("authentication_failed");
  if (!response.body) return {};
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > maximum) { await reader.cancel(); throw new Error("authentication_failed"); }
      chunks.push(value);
    }
  } finally { reader.releaseLock(); }
  try {
    const parsed = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(Buffer.concat(chunks.map((chunk) => Buffer.from(chunk)), total))) as unknown;
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string, unknown> : {};
  } catch { throw new Error("authentication_failed"); }
}

export type TokenExchangeOptions = {
  requestTimeoutMs?: number;
  /** Vault mode reads a durable-generation cache but defers admission until publication succeeds. */
  transactionalCacheIdentity?: string;
  quarantineOnTransportFailure?: boolean;
};

async function performTokenExchange(credential: Credential, requestedScopes: string[], signal: AbortSignal | undefined, fetchFn: typeof fetch, requestTimeoutMs: number, quarantineOnTransportFailure: boolean): Promise<TokenExchangeResult> {
  const url = `https://login.microsoftonline.com/${encodeURIComponent(credential.tenant)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "refresh_token",
    client_id: credential.clientId,
    refresh_token: credential.refreshToken,
    scope: [...new Set(requestedScopes)].join(" "),
  });
  signal?.throwIfAborted();
  const controller = new AbortController();
  const timeoutReason = new DOMException("OAuth token exchange timed out", "TimeoutError");
  const timer = setTimeout(() => controller.abort(timeoutReason), requestTimeoutMs);
  let abortParent: (() => void) | undefined;
  if (signal) {
    abortParent = () => controller.abort(signal.reason);
    signal.addEventListener("abort", abortParent, { once: true });
  }
  let dispatched = false;
  try {
    let responsePromise: Promise<Response>;
    try {
      responsePromise = fetchFn(url, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body, signal: controller.signal });
      dispatched = true;
    } catch (error) {
      if (signal?.aborted) throw signal.reason;
      throw error;
    }
    let response: Response;
    try { response = await responsePromise; }
    catch (error) {
      if (quarantineOnTransportFailure && dispatched) throw new Error("credential_refresh_outcome_uncertain");
      if (controller.signal.aborted) throw controller.signal.reason;
      throw error;
    }
    if (!response.ok) {
      if (quarantineOnTransportFailure && response.status !== 400 && response.status !== 401) throw new Error("credential_refresh_outcome_uncertain");
      throw new Error("authentication_failed");
    }
    try {
      const payload = await boundedJson(response);
      if (typeof payload.access_token !== "string"
        || payload.access_token.length < 1
        || payload.access_token.length > ACCESS_TOKEN_MAX_LENGTH
        || !ACCESS_TOKEN_PATTERN.test(payload.access_token)) throw new Error("authentication_failed");
      if (payload.refresh_token !== undefined && (typeof payload.refresh_token !== "string" || !payload.refresh_token || payload.refresh_token.length > 32 * 1024)) throw new Error("credential_reauthorization_required");
      if (typeof payload.expires_in !== "number" || !Number.isSafeInteger(payload.expires_in) || payload.expires_in < 1 || payload.expires_in > EXPIRES_IN_MAX_SECONDS) throw new Error("authentication_failed");
      const cacheResidencyMs = Math.min(payload.expires_in * 1_000, ACCESS_TOKEN_CACHE_MAX_RESIDENCY_MS);
      const expiresAt = Date.now() + cacheResidencyMs - ACCESS_TOKEN_EXPIRY_SKEW_MS;
      return {
        accessToken: payload.access_token,
        ...(typeof payload.refresh_token === "string" ? { replacementRefreshToken: payload.refresh_token } : {}),
        expiresAt,
      };
    } catch (error) {
      if (quarantineOnTransportFailure) throw new Error("credential_refresh_outcome_uncertain");
      throw error;
    }
  } finally {
    clearTimeout(timer);
    if (abortParent) signal?.removeEventListener("abort", abortParent);
  }
}

export async function exchangeRefreshTokenDetailed(credential: Credential, requestedScopes: string[], signal?: AbortSignal, fetchFn: typeof fetch = fetch, options: TokenExchangeOptions = {}): Promise<TokenExchangeResult> {
  if (!requestedScopes.length) throw new Error("credential_scope_missing");
  for (const scope of requestedScopes) selectScope(credential, [scope]);
  signal?.throwIfAborted();
  const requestTimeoutMs = options.requestTimeoutMs ?? 10_000;
  if (!Number.isSafeInteger(requestTimeoutMs) || requestTimeoutMs < 1 || requestTimeoutMs > 30_000) throw new Error("invalid_request_timeout");
  const cacheKey = accessTokenCacheKey(credential, requestedScopes, options.transactionalCacheIdentity);
  const cached = cachedAccessToken(cacheKey);
  if (cached) return { accessToken: cached.token, expiresAt: cached.expiresAt };
  if (options.transactionalCacheIdentity) {
    try {
      return await performTokenExchange(credential, requestedScopes, signal, fetchFn, requestTimeoutMs, options.quarantineOnTransportFailure === true);
    } catch (error) {
      if (error instanceof Error && ["authentication_failed", "credential_reauthorization_required", "credential_refresh_outcome_uncertain"].includes(error.message)) throw error;
      if (signal?.aborted) throw signal.reason;
      throw new Error("authentication_failed");
    }
  }
  const existing = accessTokenRefreshes.get(cacheKey);
  if (existing) return waitForToken(existing, signal);
  const controller = new AbortController();
  const refresh = { controller, waiters: 0, settled: false } as AccessTokenRefresh;
  refresh.promise = (async () => {
    const result = await performTokenExchange(credential, requestedScopes, controller.signal, fetchFn, requestTimeoutMs, false);
    if (result.expiresAt !== undefined) cacheAccessToken(cacheKey, result.accessToken, result.expiresAt);
    return result;
  })().finally(() => {
    refresh.settled = true;
    if (accessTokenRefreshes.get(cacheKey) === refresh) accessTokenRefreshes.delete(cacheKey);
  });
  accessTokenRefreshes.set(cacheKey, refresh);
  try {
    return await waitForToken(refresh, signal);
  } catch (error) {
    if (signal?.aborted) throw signal.reason;
    if (error instanceof DOMException && error.name === "TimeoutError") throw error;
    if (error instanceof Error && (error.message === "authentication_failed" || error.message === "credential_reauthorization_required")) throw error;
    throw new Error("authentication_failed");
  }
}

export async function exchangeRefreshToken(credential: Credential, requestedScopes: string[], signal?: AbortSignal, fetchFn: typeof fetch = fetch, options: TokenExchangeOptions = {}): Promise<string> {
  return (await exchangeRefreshTokenDetailed(credential, requestedScopes, signal, fetchFn, options)).accessToken;
}

export type CredentialBackendConfig = {
  credentialVaultKey?: unknown;
};

function vaultCacheIdentity(binding: VaultBinding): string {
  return `vault:${binding.generation}:${binding.keyId}:${binding.digest}:${binding.binding}`;
}

/** Called only after service/resource authorization has succeeded. */
export async function tokenForAuthorizedOperation(params: {
  config: CredentialBackendConfig;
  policy: GraphPolicy;
  allowedScopes: string[];
  signal?: AbortSignal;
  stateDir?: string;
  requestTimeoutMs?: number;
  fetchFn?: typeof fetch;
  /** @internal deterministic vault publication fault injection for producer tests. */
  vaultTestHooks?: { beforePublication?: () => void | Promise<void>; afterPublication?: () => void | Promise<void> };
}): Promise<string> {
  if (!params.stateDir) throw new Error("credential_vault_unavailable");
  const key = params.config.credentialVaultKey;
  decodeVaultKey(key);
  let exchangedCredential: Credential | undefined;
  let exchangedScopes: string[] | undefined;
  return refreshVaultCredential({
    stateDir: params.stateDir, key, signal: params.signal,
    exchange: async (credential, binding) => {
      const scope = selectScope(credential, params.allowedScopes);
      exchangedCredential = credential;
      exchangedScopes = [scope];
      return exchangeRefreshTokenDetailed(credential, [scope], params.signal, params.fetchFn ?? fetch, {
        requestTimeoutMs: params.requestTimeoutMs,
        transactionalCacheIdentity: vaultCacheIdentity(binding),
        quarantineOnTransportFailure: true,
      });
    },
    onDurableResult: (result, predecessor, durable) => {
      if (!exchangedCredential || !exchangedScopes || result.expiresAt === undefined) return;
      const oldKey = accessTokenCacheKey(exchangedCredential, exchangedScopes, vaultCacheIdentity(predecessor));
      const durableKey = accessTokenCacheKey(exchangedCredential, exchangedScopes, vaultCacheIdentity(durable));
      if (oldKey !== durableKey) accessTokenCache.delete(oldKey);
      cacheAccessToken(durableKey, result.accessToken, result.expiresAt);
    },
    testHooks: params.vaultTestHooks,
  });
}
