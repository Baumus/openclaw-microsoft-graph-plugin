import { createCipheriv, createDecipheriv, createHash, createHmac, randomBytes, timingSafeEqual } from "node:crypto";
import { unlink } from "node:fs/promises";
import { join } from "node:path";
import { replaceFileAtomic } from "@openclaw/fs-safe/atomic";
import { pinDirectory } from "@openclaw/fs-safe/durability";
import { acquireFileLock, type FileLockHandle } from "@openclaw/fs-safe/file-lock";
import { createSecretFileAtomic } from "@openclaw/fs-safe/secret";
import { readSecureFile } from "@openclaw/fs-safe/secure-file";

export type VaultCredential = { clientId: string; refreshToken: string; tenant: string; scopes: string[] };
export type VaultEnvelope = { format: typeof VAULT_FORMAT; version: 1; generation: number; keyId: string; algorithm: "A256GCM"; nonce: string; ciphertext: string; tag: string };
export type VaultRecord = { envelope: VaultEnvelope; credential: VaultCredential; digest: string; raw: Buffer };
export type VaultBinding = { generation: number; keyId: string; digest: string; binding: string };
export type VaultRefreshResult = { accessToken: string; replacementRefreshToken?: string; expiresAt?: number };

export const VAULT_FORMAT = "openclaw-microsoft-graph-credential-vault" as const;
export const VAULT_MAX_BYTES = 64 * 1024;
const READ_TIMEOUT_MS = 5_000;
const ENVELOPE_FIELDS = new Set(["format", "version", "generation", "keyId", "algorithm", "nonce", "ciphertext", "tag"]);
const CREDENTIAL_FIELDS = new Set(["clientId", "refreshToken", "tenant", "scopes"]);
const QUARANTINE_FORMAT = "openclaw-microsoft-graph-refresh-quarantine" as const;
const LEGACY_QUARANTINE_FIELDS = new Set(["format", "version", "generation", "keyId", "vaultDigest", "reason", "createdAt", "mac"]);
const REFRESH_STATE_FIELDS = new Set(["format", "version", "generation", "keyId", "vaultDigest", "state", "createdAt", "updatedAt", "mac"]);
const REFRESH_STATE_ALL_FIELDS = new Set([...LEGACY_QUARANTINE_FIELDS, ...REFRESH_STATE_FIELDS]);
type RefreshState = "in_flight" | "quarantined";
type LegacyQuarantineMarker = { format: typeof QUARANTINE_FORMAT; version: 1; generation: number; keyId: string; vaultDigest: string; reason: "oauth_outcome_uncertain"; createdAt: string; mac: string };
type RefreshStateMarker = { format: typeof QUARANTINE_FORMAT; version: 2; generation: number; keyId: string; vaultDigest: string; state: RefreshState; createdAt: string; updatedAt: string; mac: string };
type ParsedRefreshStateMarker = RefreshStateMarker | (LegacyQuarantineMarker & { state: "quarantined"; updatedAt: string });
type MutationHooks = {
  beforePublication?: () => void | Promise<void>;
  afterPublication?: () => void | Promise<void>;
  afterMarkerPublication?: (state: RefreshState) => void | Promise<void>;
  beforeMarkerTransition?: () => void | Promise<void>;
  afterExchange?: () => void | Promise<void>;
};

function unavailable(): Error { return new Error("credential_vault_unavailable"); }
function missing(error: unknown): boolean {
  let current: unknown = error;
  for (let depth = 0; depth < 5 && current && typeof current === "object"; depth += 1) {
    const code = (current as { code?: unknown }).code;
    if (code === "not-found" || code === "ENOENT") return true;
    current = (current as { cause?: unknown }).cause;
  }
  return false;
}
function strictBase64Url(value: unknown, bytes?: number): Buffer {
  if (typeof value !== "string" || !value || !/^[A-Za-z0-9_-]+$/.test(value)) throw unavailable();
  const decoded = Buffer.from(value, "base64url");
  if (decoded.toString("base64url") !== value || (bytes !== undefined && decoded.byteLength !== bytes)) throw unavailable();
  return decoded;
}
export function decodeVaultKey(value: unknown): Buffer { return strictBase64Url(value, 32); }
export function vaultKeyId(key: Buffer): string {
  if (key.byteLength !== 32) throw unavailable();
  return createHash("sha256").update(key).digest().subarray(0, 16).toString("base64url");
}
export function vaultCredentialDirectory(stateDir: string): string {
  if (typeof stateDir !== "string" || !stateDir) throw unavailable();
  return join(stateDir, "plugin-data", "microsoft-graph", "credentials");
}
export function vaultCredentialPath(stateDir: string): string { return join(vaultCredentialDirectory(stateDir), "credential.vault.json"); }
export function vaultQuarantinePath(stateDir: string): string { return join(vaultCredentialDirectory(stateDir), "credential.quarantine.json"); }
function sentinelPath(stateDir: string): string { return join(vaultCredentialDirectory(stateDir), ".initialized"); }
function digest(raw: Uint8Array): string { return createHash("sha256").update(raw).digest("base64url"); }

export function vaultRecordBinding(record: Pick<VaultRecord, "envelope" | "digest">): VaultBinding {
  const { generation, keyId } = record.envelope;
  const binding = createHash("sha256").update(JSON.stringify({ generation, keyId, digest: record.digest })).digest("base64url");
  return { generation, keyId, digest: record.digest, binding };
}

function scanStringEnd(raw: string, start: number): number {
  for (let index = start + 1; index < raw.length; index += 1) {
    if (raw[index] === "\\") { index += 1; continue; }
    if (raw[index] === "\"") return index;
    if (raw.charCodeAt(index) < 0x20) throw unavailable();
  }
  throw unavailable();
}
function parseStrictObject(raw: string, allowed: Set<string>): Record<string, unknown> {
  if (!raw || raw.length > VAULT_MAX_BYTES) throw unavailable();
  const seen = new Set<string>();
  let depth = 0;
  for (let index = 0; index < raw.length; index += 1) {
    const char = raw[index];
    if (char === "\"") {
      const end = scanStringEnd(raw, index);
      if (depth === 1) {
        let next = end + 1;
        while (/\s/.test(raw[next] ?? "")) next += 1;
        if (raw[next] === ":") {
          let key: unknown;
          try { key = JSON.parse(raw.slice(index, end + 1)); } catch { throw unavailable(); }
          if (typeof key !== "string" || !allowed.has(key) || seen.has(key)) throw unavailable();
          seen.add(key);
        }
      }
      index = end;
      continue;
    }
    if (char === "{") depth += 1;
    else if (char === "}") depth -= 1;
    if (depth < 0 || depth > 1) throw unavailable();
  }
  if (depth !== 0) throw unavailable();
  let parsed: unknown;
  try { parsed = JSON.parse(raw); } catch { throw unavailable(); }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) throw unavailable();
  const keys = Object.keys(parsed);
  if (keys.length !== seen.size || keys.some((key) => !allowed.has(key))) throw unavailable();
  return parsed as Record<string, unknown>;
}

export function parseVaultCredential(raw: string): VaultCredential {
  const value = parseStrictObject(raw, CREDENTIAL_FIELDS);
  if (Object.keys(value).length !== CREDENTIAL_FIELDS.size
    || typeof value.clientId !== "string" || value.clientId.length < 1 || value.clientId.length > 512
    || typeof value.refreshToken !== "string" || value.refreshToken.length < 1 || value.refreshToken.length > 32 * 1024
    || typeof value.tenant !== "string" || value.tenant.length < 1 || value.tenant.length > 512
    || !Array.isArray(value.scopes) || value.scopes.length < 1 || value.scopes.length > 128
    || value.scopes.some((scope) => typeof scope !== "string" || scope.length < 1 || scope.length > 512)) throw unavailable();
  const scopes = [...new Set(value.scopes as string[])];
  if (!scopes.some((scope) => scope.toLowerCase() === "offline_access")) scopes.push("offline_access");
  return { clientId: value.clientId, refreshToken: value.refreshToken, tenant: value.tenant, scopes };
}
export function parseVaultEnvelope(raw: string): VaultEnvelope {
  const value = parseStrictObject(raw, ENVELOPE_FIELDS);
  if (Object.keys(value).length !== ENVELOPE_FIELDS.size || value.format !== VAULT_FORMAT || value.version !== 1
    || !Number.isSafeInteger(value.generation) || Number(value.generation) < 1 || value.algorithm !== "A256GCM") throw unavailable();
  strictBase64Url(value.keyId, 16); strictBase64Url(value.nonce, 12); strictBase64Url(value.tag, 16);
  const ciphertext = strictBase64Url(value.ciphertext);
  if (ciphertext.byteLength < 1 || ciphertext.byteLength > VAULT_MAX_BYTES) throw unavailable();
  return value as VaultEnvelope;
}
function canonicalCredential(credential: VaultCredential): string {
  const parsed = parseVaultCredential(JSON.stringify(credential));
  return JSON.stringify({ clientId: parsed.clientId, refreshToken: parsed.refreshToken, tenant: parsed.tenant, scopes: parsed.scopes });
}
function aad(envelope: Pick<VaultEnvelope, "generation" | "keyId">): Buffer {
  return Buffer.from(`${VAULT_FORMAT}\0${1}\0${envelope.generation}\0A256GCM\0${envelope.keyId}`, "utf8");
}
export function encryptVaultCredential(generation: number, credential: VaultCredential, keyValue: unknown, nonce = randomBytes(12)): VaultEnvelope {
  if (!Number.isSafeInteger(generation) || generation < 1 || nonce.byteLength !== 12) throw unavailable();
  const key = decodeVaultKey(keyValue);
  const header = { generation, keyId: vaultKeyId(key) };
  const cipher = createCipheriv("aes-256-gcm", key, nonce, { authTagLength: 16 });
  cipher.setAAD(aad(header));
  const ciphertext = Buffer.concat([cipher.update(canonicalCredential(credential), "utf8"), cipher.final()]);
  return { format: VAULT_FORMAT, version: 1, generation, keyId: header.keyId, algorithm: "A256GCM", nonce: nonce.toString("base64url"), ciphertext: ciphertext.toString("base64url"), tag: cipher.getAuthTag().toString("base64url") };
}
export function decryptVaultEnvelope(envelopeInput: VaultEnvelope | string, keyValue: unknown): VaultCredential {
  try {
    const envelope = typeof envelopeInput === "string" ? parseVaultEnvelope(envelopeInput) : parseVaultEnvelope(JSON.stringify(envelopeInput));
    const key = decodeVaultKey(keyValue);
    if (envelope.keyId !== vaultKeyId(key)) throw unavailable();
    const decipher = createDecipheriv("aes-256-gcm", key, strictBase64Url(envelope.nonce, 12), { authTagLength: 16 });
    decipher.setAAD(aad(envelope)); decipher.setAuthTag(strictBase64Url(envelope.tag, 16));
    return parseVaultCredential(Buffer.concat([decipher.update(strictBase64Url(envelope.ciphertext)), decipher.final()]).toString("utf8"));
  } catch { throw unavailable(); }
}
export function serializeVaultEnvelope(envelope: VaultEnvelope): string { return `${JSON.stringify(parseVaultEnvelope(JSON.stringify(envelope)))}\n`; }

export async function ensureVaultCredentialDirectory(stateDir: string): Promise<void> {
  const directory = vaultCredentialDirectory(stateDir);
  const sentinel = sentinelPath(stateDir);
  const validSentinel = async () => {
    try {
      const existing = await readSecureFile({ filePath: sentinel, label: "Microsoft Graph credential vault directory sentinel", trust: { trustedDirs: [directory] }, io: { maxBytes: 128, timeoutMs: READ_TIMEOUT_MS } });
      if (existing.buffer.toString("utf8") !== "openclaw-microsoft-graph-credential-vault\n" || (process.platform !== "win32" && (existing.stat.mode & 0o777) !== 0o600)) throw unavailable();
      return true;
    } catch (error) { if (missing(error)) return false; throw unavailable(); }
  };
  if (!await validSentinel()) {
    try { await createSecretFileAtomic({ rootDir: stateDir, filePath: sentinel, content: "openclaw-microsoft-graph-credential-vault\n", mode: 0o600, dirMode: 0o700, durable: true }); }
    catch { if (!await validSentinel()) throw unavailable(); }
  }
  const pinned = await pinDirectory(directory, { label: "Microsoft Graph credential vault directory" });
  try { if (process.platform !== "win32" && (pinned.receipt.identity.mode & 0o777) !== 0o700) throw unavailable(); await pinned.assertCurrent(); }
  finally { await pinned.close(); }
}
async function readVaultRaw(stateDir: string): Promise<Buffer> {
  const directory = vaultCredentialDirectory(stateDir);
  const pinned = await pinDirectory(directory, { label: "Microsoft Graph credential vault directory" });
  try {
    if (process.platform !== "win32" && (pinned.receipt.identity.mode & 0o777) !== 0o700) throw unavailable();
    const result = await readSecureFile({ filePath: vaultCredentialPath(stateDir), label: "Microsoft Graph credential vault", trust: { trustedDirs: [directory] }, io: { maxBytes: VAULT_MAX_BYTES, timeoutMs: READ_TIMEOUT_MS } });
    if (process.platform !== "win32" && (result.stat.mode & 0o777) !== 0o600) throw unavailable();
    await pinned.assertCurrent(); return result.buffer;
  } finally { await pinned.close(); }
}
export async function readVaultCredential(stateDir: string, keyValue: unknown): Promise<VaultRecord> {
  try { const raw = await readVaultRaw(stateDir); const envelope = parseVaultEnvelope(raw.toString("utf8")); return { envelope, credential: decryptVaultEnvelope(envelope, keyValue), digest: digest(raw), raw }; }
  catch { throw unavailable(); }
}
export async function createVaultCredential(stateDir: string, credential: VaultCredential, keyValue: unknown): Promise<VaultRecord> {
  const content = serializeVaultEnvelope(encryptVaultCredential(1, credential, keyValue));
  try {
    await createSecretFileAtomic({ rootDir: stateDir, filePath: vaultCredentialPath(stateDir), content, mode: 0o600, dirMode: 0o700, durable: true });
    const record = await readVaultCredential(stateDir, keyValue);
    if (record.envelope.generation !== 1 || record.digest !== digest(Buffer.from(content))) throw unavailable();
    return record;
  } catch { throw new Error("credential_vault_write_failed"); }
}
export async function acquireVaultLock(stateDir: string, operation: string): Promise<FileLockHandle> {
  try {
    await ensureVaultCredentialDirectory(stateDir);
    return await acquireFileLock(vaultCredentialPath(stateDir), { managerKey: "microsoft-graph-credential-vault", staleMs: 5 * 60_000, timeoutMs: 1_000, retry: { retries: 20, factor: 1.2, minTimeout: 20, maxTimeout: 250, randomize: true }, staleRecovery: "fail-closed", payload: () => ({ pid: process.pid, operation, startedAt: new Date().toISOString() }), compromiseCheckIntervalMs: 250 });
  } catch { throw new Error("credential_vault_locked"); }
}

function markerMac(marker: Omit<LegacyQuarantineMarker, "mac"> | Omit<RefreshStateMarker, "mac">, keyValue: unknown): string { return createHmac("sha256", decodeVaultKey(keyValue)).update(JSON.stringify(marker)).digest("base64url"); }
function parseMarker(raw: string, keyValue: unknown): ParsedRefreshStateMarker {
  const value = parseStrictObject(raw, REFRESH_STATE_ALL_FIELDS);
  if (value.format !== QUARANTINE_FORMAT || typeof value.mac !== "string") throw unavailable();
  if (value.version === 1) {
    if (Object.keys(value).length !== LEGACY_QUARANTINE_FIELDS.size || Object.keys(value).some((field) => !LEGACY_QUARANTINE_FIELDS.has(field))
      || !Number.isSafeInteger(value.generation) || Number(value.generation) < 1 || typeof value.keyId !== "string" || typeof value.vaultDigest !== "string"
      || value.reason !== "oauth_outcome_uncertain" || typeof value.createdAt !== "string" || !Number.isFinite(Date.parse(value.createdAt))) throw unavailable();
    const marker = value as unknown as LegacyQuarantineMarker; const { mac, ...payload } = marker;
    if (!timingSafeEqual(strictBase64Url(mac, 32), strictBase64Url(markerMac(payload, keyValue), 32))) throw unavailable();
    return { ...marker, state: "quarantined", updatedAt: marker.createdAt };
  }
  if (value.version !== 2 || Object.keys(value).length !== REFRESH_STATE_FIELDS.size || Object.keys(value).some((field) => !REFRESH_STATE_FIELDS.has(field))
    || !Number.isSafeInteger(value.generation) || Number(value.generation) < 1 || typeof value.keyId !== "string" || typeof value.vaultDigest !== "string"
    || (value.state !== "in_flight" && value.state !== "quarantined")
    || typeof value.createdAt !== "string" || !Number.isFinite(Date.parse(value.createdAt))
    || typeof value.updatedAt !== "string" || !Number.isFinite(Date.parse(value.updatedAt))) throw unavailable();
  const marker = value as unknown as RefreshStateMarker; const { mac, ...payload } = marker;
  if (!timingSafeEqual(strictBase64Url(mac, 32), strictBase64Url(markerMac(payload, keyValue), 32))) throw unavailable();
  return marker;
}
async function readMarker(stateDir: string, keyValue: unknown): Promise<ParsedRefreshStateMarker | undefined> {
  const directory = vaultCredentialDirectory(stateDir); const pinned = await pinDirectory(directory, { label: "Microsoft Graph credential vault directory" });
  try {
    let result;
    try { result = await readSecureFile({ filePath: vaultQuarantinePath(stateDir), label: "Microsoft Graph refresh quarantine", trust: { trustedDirs: [directory] }, io: { maxBytes: 8 * 1024, timeoutMs: READ_TIMEOUT_MS } }); }
    catch (error) { if (missing(error)) return undefined; throw error; }
    if (process.platform !== "win32" && (result.stat.mode & 0o777) !== 0o600) throw unavailable();
    await pinned.assertCurrent(); return parseMarker(result.buffer.toString("utf8"), keyValue);
  } finally { await pinned.close(); }
}
function markerPayload(record: VaultRecord, state: RefreshState, createdAt = new Date().toISOString()): Omit<RefreshStateMarker, "mac"> {
  return { format: QUARANTINE_FORMAT, version: 2, generation: record.envelope.generation, keyId: record.envelope.keyId, vaultDigest: record.digest, state, createdAt, updatedAt: new Date().toISOString() };
}
async function publishInitialMarker(stateDir: string, record: VaultRecord, keyValue: unknown, lock: FileLockHandle, hooks?: MutationHooks): Promise<void> {
  const payload = markerPayload(record, "in_flight");
  const marker = { ...payload, mac: markerMac(payload, keyValue) };
  try {
    if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
    await createSecretFileAtomic({ rootDir: stateDir, filePath: vaultQuarantinePath(stateDir), content: `${JSON.stringify(marker)}\n`, mode: 0o600, dirMode: 0o700, durable: true });
    await hooks?.afterMarkerPublication?.("in_flight");
    if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
    const published = await readMarker(stateDir, keyValue);
    if (!published || published.state !== "in_flight" || published.vaultDigest !== record.digest) throw unavailable();
    if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
  } catch (error) { if (error instanceof Error && error.message === "credential_vault_locked") throw error; throw new Error("credential_vault_write_failed"); }
}
async function transitionMarkerToQuarantined(stateDir: string, record: VaultRecord, keyValue: unknown, lock: FileLockHandle, hooks?: MutationHooks): Promise<void> {
  const existing = await readMarker(stateDir, keyValue);
  if (!existing) throw new Error("credential_vault_write_failed");
  const payload = markerPayload(record, "quarantined", existing.createdAt);
  const marker = { ...payload, mac: markerMac(payload, keyValue) };
  try {
    await replaceFileAtomic({
      filePath: vaultQuarantinePath(stateDir), content: `${JSON.stringify(marker)}\n`, mode: 0o600, dirMode: 0o700,
      preserveExistingMode: false, copyFallbackOnPermissionError: false, destinationHardlinks: "reject", syncTempFile: true, syncParentDir: true,
      beforeRename: async () => { if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked"); await hooks?.beforeMarkerTransition?.(); },
    });
    const published = await readMarker(stateDir, keyValue);
    if (!published || published.state !== "quarantined" || published.vaultDigest !== record.digest) throw unavailable();
  } catch (error) { if (error instanceof Error && error.message === "credential_vault_locked") throw error; throw new Error("credential_vault_write_failed"); }
}
async function clearMarker(stateDir: string, keyValue: unknown, lock: FileLockHandle): Promise<void> {
  if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
  const pinned = await pinDirectory(vaultCredentialDirectory(stateDir), { label: "Microsoft Graph credential vault directory" });
  try { await pinned.assertCurrent(); await unlink(vaultQuarantinePath(stateDir)); await pinned.sync(); await pinned.assertCurrent(); }
  catch { throw new Error("credential_vault_write_failed"); }
  finally { await pinned.close(); }
  if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
  if (await readMarker(stateDir, keyValue)) throw new Error("credential_vault_write_failed");
}
async function replaceRecord(stateDir: string, current: VaultRecord, credential: VaultCredential, keyValue: unknown, lock: FileLockHandle, hooks?: MutationHooks): Promise<VaultRecord> {
  if (current.envelope.generation >= Number.MAX_SAFE_INTEGER) throw new Error("credential_vault_write_failed");
  const next = encryptVaultCredential(current.envelope.generation + 1, credential, keyValue); const content = serializeVaultEnvelope(next); const expected = digest(Buffer.from(content));
  try {
    await replaceFileAtomic({ filePath: vaultCredentialPath(stateDir), content, mode: 0o600, dirMode: 0o700, preserveExistingMode: false, copyFallbackOnPermissionError: false, destinationHardlinks: "reject", syncTempFile: true, syncParentDir: true, beforeRename: async () => { if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked"); if (digest(await readVaultRaw(stateDir)) !== current.digest) throw new Error("credential_vault_conflict"); await hooks?.beforePublication?.(); } });
  } catch (error) { if (error instanceof Error && ["credential_vault_locked", "credential_vault_conflict"].includes(error.message)) throw error; throw new Error("credential_vault_write_failed"); }
  try { await hooks?.afterPublication?.(); const published = await readVaultCredential(stateDir, keyValue); if (published.digest !== expected || published.envelope.generation !== next.generation || published.envelope.keyId !== next.keyId) throw unavailable(); return published; }
  catch { throw new Error("credential_vault_write_failed"); }
}

export async function refreshVaultCredential(params: { stateDir: string; key: unknown; signal?: AbortSignal; exchange: (credential: VaultCredential, binding: VaultBinding) => Promise<VaultRefreshResult>; onDurableResult?: (result: VaultRefreshResult, predecessor: VaultBinding, durable: VaultBinding) => void; testHooks?: MutationHooks }): Promise<string> {
  params.signal?.throwIfAborted(); const lock = await acquireVaultLock(params.stateDir, "refresh");
  try {
    const current = await readVaultCredential(params.stateDir, params.key);
    if (await readMarker(params.stateDir, params.key)) throw new Error("credential_reauthorization_required");
    await publishInitialMarker(params.stateDir, current, params.key, lock, params.testHooks);
    params.signal?.throwIfAborted();
    if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
    let result: VaultRefreshResult;
    try {
      result = await params.exchange(current.credential, vaultRecordBinding(current));
      await params.testHooks?.afterExchange?.();
    } catch (error) {
      if (error instanceof Error && error.message === "authentication_failed") {
        await clearMarker(params.stateDir, params.key, lock);
        throw error;
      }
      if (error instanceof Error && error.message === "credential_refresh_outcome_uncertain") {
        try { await transitionMarkerToQuarantined(params.stateDir, current, params.key, lock, params.testHooks); }
        catch (transitionError) { if (transitionError instanceof Error && transitionError.message === "credential_vault_locked") throw transitionError; }
      }
      throw error;
    }
    if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
    if (typeof result.accessToken !== "string" || !result.accessToken || (result.replacementRefreshToken !== undefined && (!result.replacementRefreshToken || result.replacementRefreshToken.length > 32 * 1024))) {
      try { await transitionMarkerToQuarantined(params.stateDir, current, params.key, lock, params.testHooks); }
      catch (transitionError) { if (transitionError instanceof Error && transitionError.message === "credential_vault_locked") throw transitionError; }
      throw new Error("credential_refresh_outcome_uncertain");
    }
    let durable = current;
    if (result.replacementRefreshToken !== undefined && result.replacementRefreshToken !== current.credential.refreshToken) {
      try { durable = await replaceRecord(params.stateDir, current, { ...current.credential, refreshToken: result.replacementRefreshToken }, params.key, lock, params.testHooks); }
      catch {
        const observed = await readVaultCredential(params.stateDir, params.key).catch(() => current);
        try { await transitionMarkerToQuarantined(params.stateDir, observed, params.key, lock, params.testHooks); } catch { /* The authenticated in-flight marker remains the fail-closed fallback. */ }
        throw new Error("credential_refresh_outcome_uncertain");
      }
    }
    if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
    const verified = await readVaultCredential(params.stateDir, params.key);
    if (verified.digest !== durable.digest || verified.envelope.generation !== durable.envelope.generation || verified.envelope.keyId !== durable.envelope.keyId) throw new Error("credential_refresh_outcome_uncertain");
    await clearMarker(params.stateDir, params.key, lock);
    params.onDurableResult?.(result, vaultRecordBinding(current), vaultRecordBinding(durable)); return result.accessToken;
  } finally { await lock.release().catch(() => undefined); }
}

export async function clearVaultQuarantine(stateDir: string, keyValue: unknown, expectedBinding: string, apply: boolean): Promise<VaultRecord> {
  if (!apply) {
    const current = await readVaultCredential(stateDir, keyValue); const binding = vaultRecordBinding(current);
    if (!expectedBinding || expectedBinding !== binding.binding || !await readMarker(stateDir, keyValue)) throw new Error("credential_vault_conflict");
    return current;
  }
  const lock = await acquireVaultLock(stateDir, "recover-refresh");
  try {
    const current = await readVaultCredential(stateDir, keyValue); const binding = vaultRecordBinding(current);
    if (!expectedBinding || expectedBinding !== binding.binding || !await readMarker(stateDir, keyValue)) throw new Error("credential_vault_conflict");
    if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
    const pinned = await pinDirectory(vaultCredentialDirectory(stateDir), { label: "Microsoft Graph credential vault directory" });
    try { await pinned.assertCurrent(); await unlink(vaultQuarantinePath(stateDir)); await pinned.sync(); await pinned.assertCurrent(); } finally { await pinned.close(); }
    if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
    if (await readMarker(stateDir, keyValue)) throw new Error("credential_vault_write_failed"); return current;
  } finally { await lock.release().catch(() => undefined); }
}
export async function inspectVaultCredential(stateDir: string, keyValue?: unknown): Promise<{ result: "missing" | "valid" | "quarantined" | "unavailable"; generation?: number; keyId?: string; digest?: string; binding?: string }> {
  try {
    const raw = await readVaultRaw(stateDir); const envelope = parseVaultEnvelope(raw.toString("utf8")); const digestValue = digest(raw); let quarantined = false;
    if (keyValue !== undefined) { decryptVaultEnvelope(envelope, keyValue); quarantined = Boolean(await readMarker(stateDir, keyValue)); }
    return { result: quarantined ? "quarantined" : "valid", generation: envelope.generation, keyId: envelope.keyId, digest: digestValue, binding: vaultRecordBinding({ envelope, digest: digestValue }).binding };
  } catch (error) { return missing(error) ? { result: "missing" } : { result: "unavailable" }; }
}
