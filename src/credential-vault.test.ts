import { chmod, link, lstat, mkdir, mkdtemp, readFile, symlink, utimes, writeFile } from "node:fs/promises";
import { randomBytes } from "node:crypto";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { acquireVaultLock, clearVaultQuarantine, createVaultCredential, decodeVaultKey, decryptVaultEnvelope, encryptVaultCredential, inspectVaultCredential, parseVaultCredential, parseVaultEnvelope, readVaultCredential, refreshVaultCredential, serializeVaultEnvelope, vaultCredentialPath, vaultQuarantinePath, vaultRecordBinding } from "./credential-vault.js";

const temporary: string[] = [];
const credential = { clientId: "synthetic-client", refreshToken: "synthetic-refresh", tenant: "synthetic-tenant", scopes: ["Files.Read", "Files.ReadWrite", "offline_access"] };
const key = () => randomBytes(32).toString("base64url");
async function stateDir() { const path = await mkdtemp(join(tmpdir(), "microsoft-graph-vault-test-")); temporary.push(path); return path; }
afterEach(async () => { const { rm } = await import("node:fs/promises"); await Promise.all(temporary.splice(0).map((path) => rm(path, { recursive: true, force: true }))); });

describe("single credential vault", () => {
  it("accepts only canonical 32-byte keys and authenticates the entire envelope", () => {
    const active = key(); expect(decodeVaultKey(active)).toHaveLength(32);
    expect(() => decodeVaultKey(`${active}=`)).toThrow("credential_vault_unavailable");
    const first = encryptVaultCredential(1, credential, active); const second = encryptVaultCredential(1, credential, active);
    expect(first.nonce).not.toBe(second.nonce); expect(decryptVaultEnvelope(first, active)).toEqual(credential);
    expect(() => decryptVaultEnvelope(first, key())).toThrow("credential_vault_unavailable");
    expect(() => decryptVaultEnvelope({ ...first, generation: 2 }, active)).toThrow("credential_vault_unavailable");
  });

  it("strictly rejects malformed, empty, duplicate, oversized, and overflowing inputs", () => {
    const serialized = serializeVaultEnvelope(encryptVaultCredential(1, credential, key())).trim();
    for (const raw of ["", "{", serialized.replace(/}$/, ',"extra":true}'), serialized.replace('"version":1', '"version":1,"version":1'), serialized.replace('"generation":1', `"generation":${Number.MAX_SAFE_INTEGER + 1}`), "x".repeat(64 * 1024 + 1)]) expect(() => parseVaultEnvelope(raw)).toThrow("credential_vault_unavailable");
    const payload = JSON.stringify(credential);
    expect(() => parseVaultCredential(payload.replace(/}$/, ',"unknown":true}'))).toThrow("credential_vault_unavailable");
    expect(() => parseVaultCredential(payload.replace('"tenant":"synthetic-tenant"', '"tenant":"synthetic-tenant","tenant":"duplicate"'))).toThrow("credential_vault_unavailable");
  });

  it("uses one private record and refuses unsafe or oversized files", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    expect((await lstat(vaultCredentialPath(state))).mode & 0o777).toBe(0o600);
    const directory = join(state, "plugin-data", "microsoft-graph", "credentials"); expect((await lstat(directory)).mode & 0o777).toBe(0o700);
    await chmod(directory, 0o755); await expect(readVaultCredential(state, active)).rejects.toThrow("credential_vault_unavailable"); await chmod(directory, 0o700);
    const other = await stateDir(); const otherDirectory = join(other, "plugin-data", "microsoft-graph", "credentials"); await mkdir(otherDirectory, { recursive: true, mode: 0o700 });
    const target = join(otherDirectory, "target"); await writeFile(target, await readFile(vaultCredentialPath(state)), { mode: 0o600 }); await symlink(target, vaultCredentialPath(other));
    await expect(readVaultCredential(other, active)).rejects.toThrow("credential_vault_unavailable");
    const hard = await stateDir(); const hardDirectory = join(hard, "plugin-data", "microsoft-graph", "credentials"); await mkdir(hardDirectory, { recursive: true, mode: 0o700 }); await link(vaultCredentialPath(state), vaultCredentialPath(hard));
    await expect(readVaultCredential(hard, active)).rejects.toThrow("credential_vault_unavailable");
    const oversized = await stateDir(); const oversizedDirectory = join(oversized, "plugin-data", "microsoft-graph", "credentials"); await mkdir(oversizedDirectory, { recursive: true, mode: 0o700 }); await writeFile(vaultCredentialPath(oversized), Buffer.alloc(64 * 1024 + 1), { mode: 0o600 });
    await expect(readVaultCredential(oversized, active)).rejects.toThrow("credential_vault_unavailable");
  });

  it("serializes all scopes through one lock and persists rotation before return", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    let entered!: () => void; const started = new Promise<void>((resolve) => { entered = resolve; }); let release!: () => void; const gate = new Promise<void>((resolve) => { release = resolve; }); const order: string[] = [];
    const first = refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { order.push("first"); entered(); await gate; return { accessToken: "one", replacementRefreshToken: "rotated" }; } });
    await started;
    const second = refreshVaultCredential({ stateDir: state, key: active, exchange: async (current) => { order.push(current.refreshToken); return { accessToken: "two" }; } });
    await new Promise((resolve) => setTimeout(resolve, 30)); expect(order).toEqual(["first"]); release(); await expect(first).resolves.toBe("one"); await expect(second).resolves.toBe("two");
    expect(order).toEqual(["first", "rotated"]); expect((await readVaultCredential(state, active)).credential.refreshToken).toBe("rotated");
  });

  it("quarantines uncertain refresh outcomes durably and blocks a second exchange", async () => {
    const state = await stateDir(); const active = key(); const current = await createVaultCredential(state, credential, active); let exchanges = 0;
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { exchanges += 1; throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow("credential_refresh_outcome_uncertain");
    expect((await lstat(vaultQuarantinePath(state))).mode & 0o777).toBe(0o600);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { exchanges += 1; return { accessToken: "no" }; } })).rejects.toThrow("credential_reauthorization_required"); expect(exchanges).toBe(1);
    const binding = vaultRecordBinding(current).binding; await clearVaultQuarantine(state, active, binding, true); await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => ({ accessToken: "recovered" }) })).resolves.toBe("recovered");
  });

  it("does not dispatch when durable in-flight marker verification fails and blocks the next call", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active); let exchanges = 0;
    await expect(refreshVaultCredential({
      stateDir: state, key: active,
      exchange: async () => { exchanges += 1; return { accessToken: "never" }; },
      testHooks: { afterMarkerPublication: () => { throw new Error("synthetic_marker_readback_failure"); } },
    })).rejects.toThrow("credential_vault_write_failed");
    expect(exchanges).toBe(0);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { exchanges += 1; return { accessToken: "never" }; } })).rejects.toThrow("credential_reauthorization_required");
    expect(exchanges).toBe(0);
  });

  it("retains the in-flight marker when the lock is compromised after exchange", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active); let exchanges = 0;
    await expect(refreshVaultCredential({
      stateDir: state, key: active,
      exchange: async () => { exchanges += 1; return { accessToken: "uncertain" }; },
      testHooks: { afterExchange: async () => { await writeFile(`${vaultCredentialPath(state)}.lock`, "compromised", { mode: 0o600 }); } },
    })).rejects.toThrow("credential_vault_locked");
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { exchanges += 1; return { accessToken: "never" }; } })).rejects.toThrow();
    expect(exchanges).toBe(1);
    expect(await inspectVaultCredential(state, active)).toMatchObject({ result: "quarantined" });
  });

  it("retains the in-flight marker when post-dispatch quarantine publication fails", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active); let exchanges = 0;
    await expect(refreshVaultCredential({
      stateDir: state, key: active,
      exchange: async () => { exchanges += 1; throw new Error("credential_refresh_outcome_uncertain"); },
      testHooks: { beforeMarkerTransition: () => { throw new Error("synthetic_quarantine_publication_failure"); } },
    })).rejects.toThrow("credential_refresh_outcome_uncertain");
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { exchanges += 1; return { accessToken: "never" }; } })).rejects.toThrow("credential_reauthorization_required");
    expect(exchanges).toBe(1);
    expect(await inspectVaultCredential(state, active)).toMatchObject({ result: "quarantined" });
  });

  it("fails closed for live and stale locks", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active); const lock = await acquireVaultLock(state, "holder");
    try { await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => ({ accessToken: "no" }) })).rejects.toThrow("credential_vault_locked"); } finally { await lock.release(); }
    const stale = await stateDir(); await createVaultCredential(stale, credential, active); const lockPath = `${vaultCredentialPath(stale)}.lock`; await writeFile(lockPath, JSON.stringify({ pid: 999999, operation: "stale", startedAt: "2000-01-01T00:00:00.000Z" }), { mode: 0o600 }); const old = new Date("2000-01-01T00:00:00.000Z"); await utimes(lockPath, old, old);
    await expect(refreshVaultCredential({ stateDir: stale, key: active, exchange: async () => ({ accessToken: "no" }) })).rejects.toThrow("credential_vault_locked");
  }, 10_000);

  it("reports exactly one sanitized credential status", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    expect(await inspectVaultCredential(state, active)).toMatchObject({ result: "valid", generation: 1, binding: expect.any(String) });
  });
});
