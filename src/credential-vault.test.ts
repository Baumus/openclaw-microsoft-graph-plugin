import { fork } from "node:child_process";
import { build } from "esbuild";
import { chmod, link, lstat, mkdir, mkdtemp, readFile, symlink, utimes, writeFile } from "node:fs/promises";
import { createHmac, randomBytes } from "node:crypto";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { acquireVaultLock, clearVaultQuarantine, createVaultCredential, decodeVaultKey, decryptVaultEnvelope, encryptVaultCredential, inspectVaultCredential, parseVaultCredential, parseVaultEnvelope, readVaultCredential, refreshVaultCredential, serializeVaultEnvelope, vaultCredentialPath, vaultQuarantinePath, vaultRecordBinding } from "./credential-vault.js";

const temporary: string[] = [];
const credential = { clientId: "synthetic-client", refreshToken: "synthetic-refresh", tenant: "synthetic-tenant", scopes: ["Files.Read", "Files.ReadWrite", "offline_access"] };
const key = () => randomBytes(32).toString("base64url");
async function stateDir() { const path = await mkdtemp(join(tmpdir(), "microsoft-graph-vault-test-")); temporary.push(path); return path; }
afterEach(async () => { vi.restoreAllMocks(); const { rm } = await import("node:fs/promises"); await Promise.all(temporary.splice(0).map((path) => rm(path, { recursive: true, force: true }))); });

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
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { exchanges += 1; return { accessToken: "no" }; } })).rejects.toThrow("credential_refresh_retry_later"); expect(exchanges).toBe(1);
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
  it("durably consumes one recovery attempt across new calls and exhausts it on failure", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    const exchange = vi.fn(async () => { throw new Error("credential_refresh_outcome_uncertain"); });
    const params = { stateDir: state, key: active, exchange };
    await expect(refreshVaultCredential(params)).rejects.toThrow("credential_refresh_outcome_uncertain");
    await expect(refreshVaultCredential(params)).rejects.toThrow("credential_refresh_retry_later");
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    await expect(refreshVaultCredential({ ...params })).rejects.toThrow("credential_refresh_outcome_uncertain");
    expect(JSON.parse(await readFile(vaultQuarantinePath(state), "utf8"))).toMatchObject({ version: 3, reason: "exhausted", recoveryUsed: 1 });
    await expect(refreshVaultCredential({ ...params })).rejects.toThrow("credential_reauthorization_required");
    expect(exchange).toHaveBeenCalledTimes(2);
  });

  it("a definitive transient response during recovery cannot reset its durable budget", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_temporarily_unavailable"); } })).rejects.toThrow("credential_refresh_temporarily_unavailable");
    const exchange = vi.fn();
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).rejects.toThrow("credential_reauthorization_required");
    expect(exchange).not.toHaveBeenCalled();
  });

  it.each([1, 2])("legacy v%i markers cannot prove OAuth-only uncertainty and remain operator-bound", async (version) => {
    const state = await stateDir(); const active = key(); const current = await createVaultCredential(state, credential, active);
    const base = { format: "openclaw-microsoft-graph-refresh-quarantine", version, generation: 1, keyId: current.envelope.keyId, vaultDigest: current.digest };
    const payload = version === 1 ? { ...base, reason: "oauth_outcome_uncertain", createdAt: "2000-01-01T00:00:00Z" } : { ...base, state: "quarantined", createdAt: "2000-01-01T00:00:00Z", updatedAt: "2000-01-01T00:00:00Z" };
    const mac = createHmac("sha256", Buffer.from(active, "base64url")).update(JSON.stringify(payload)).digest("base64url");
    await writeFile(vaultQuarantinePath(state), JSON.stringify({ ...payload, mac }), { mode: 0o600 });
    const exchange = vi.fn(); await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).rejects.toThrow("credential_reauthorization_required"); expect(exchange).not.toHaveBeenCalled();
  });

  it.each(["generation", "keyId", "vaultDigest", "mac"])("rejects changed %s before recovery dispatch", async (field) => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    const { mac: _mac, ...payload } = JSON.parse(await readFile(vaultQuarantinePath(state), "utf8"));
    if (field !== "mac") payload[field] = field === "generation" ? 2 : "changed";
    const mac = createHmac("sha256", Buffer.from(active, "base64url")).update(JSON.stringify(payload)).digest("base64url");
    await writeFile(vaultQuarantinePath(state), JSON.stringify({ ...payload, mac: field === "mac" ? "a".repeat(43) : mac }), { mode: 0o600 });
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    const exchange = vi.fn(); await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).rejects.toThrow(); expect(exchange).not.toHaveBeenCalled();
  });

  it("recovery handoff serializes concurrent callers, observes rotation, and cleans up after success", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    let entered!: () => void; const started = new Promise<void>((resolve) => { entered = resolve; });
    let release!: () => void; const gate = new Promise<void>((resolve) => { release = resolve; });
    const observed: string[] = [];
    const first = refreshVaultCredential({ stateDir: state, key: active, exchange: async (current) => { observed.push(current.refreshToken); entered(); await gate; return { accessToken: "recovered", replacementRefreshToken: "rotated-after-recovery" }; } });
    await started;
    expect(JSON.parse(await readFile(vaultQuarantinePath(state), "utf8"))).toMatchObject({ state: "in_flight", recoveryUsed: 1 });
    const second = refreshVaultCredential({ stateDir: state, key: active, exchange: async (current) => { observed.push(current.refreshToken); return { accessToken: "later" }; } });
    release(); await expect(first).resolves.toBe("recovered"); await expect(second).resolves.toBe("later");
    expect(observed).toEqual([credential.refreshToken, "rotated-after-recovery"]);
    expect(await inspectVaultCredential(state, active)).toMatchObject({ result: "valid", generation: 2 });
  });

  it("a crash after recovery claim or failed claim publication cannot re-enable recovery", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    const controller = new AbortController(); const exchange = vi.fn();
    await expect(refreshVaultCredential({ stateDir: state, key: active, signal: controller.signal, exchange, testHooks: { beforeMarkerTransition: () => { controller.abort(); } } })).rejects.toMatchObject({ name: "AbortError" });
    expect(exchange).not.toHaveBeenCalled();
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).rejects.toThrow("credential_reauthorization_required");
    expect(JSON.parse(await readFile(vaultQuarantinePath(state), "utf8"))).toMatchObject({ state: "in_flight", recoveryUsed: 1 });
  });

  it("does not automatically recover uncertain vault publication, even after cooldown", async () => {
    for (const afterPublication of [false, true]) {
      const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
      const fault = () => { throw new Error("synthetic-publication-fault"); };
      const exchange = vi.fn(async () => ({ accessToken: "must-not-return", replacementRefreshToken: "new-secret" }));
      await expect(refreshVaultCredential({ stateDir: state, key: active, exchange, testHooks: afterPublication ? { afterPublication: fault } : { beforePublication: fault } })).rejects.toThrow("credential_refresh_outcome_uncertain");
      vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
      await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).rejects.toThrow("credential_reauthorization_required");
      expect(exchange).toHaveBeenCalledTimes(1);
      expect((await readVaultCredential(state, active)).credential.refreshToken).toBe(afterPublication ? "new-secret" : credential.refreshToken);
    }
  });

  it("serializes recovery with an independent process and hands off only the durable generation", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    const now = Date.now() + 60_001; vi.spyOn(Date, "now").mockReturnValue(now);
    const compiled = await build({ stdin: { contents: `
      import { refreshVaultCredential } from "./src/credential-vault.ts";
      process.once("message", async ({ stateDir, key, now }) => {
        Date.now = () => now;
        try {
          const result = await refreshVaultCredential({ stateDir, key, exchange: async () => {
            const gate = new Promise(resolve => process.once("message", resolve));
            process.send({ ready: true }); await gate;
            return { accessToken: "child-access", replacementRefreshToken: "child-rotated" };
          } });
          process.send({ result });
        } catch { process.send({ failed: true }); }
        process.disconnect();
      });`, resolveDir: process.cwd() }, bundle: true, packages: "external", platform: "node", format: "esm", write: false });
    // Keep the disposable bundle beneath this package so its owned dependencies resolve.
    const directory = await mkdtemp(join(process.cwd(), "node_modules", ".vault-process-test-")); temporary.push(directory);
    const entry = join(directory, "child.mjs"); await writeFile(entry, compiled.outputFiles[0]!.contents);
    const child = fork(entry, [], { stdio: ["ignore", "ignore", "pipe", "ipc"] });
    const exited = new Promise<number | null>((resolve) => child.once("exit", resolve));
    const messages: unknown[] = [];
    try {
      await new Promise<void>((resolve, reject) => {
        child.on("message", (message) => { messages.push(message); if ((message as { ready?: boolean }).ready) resolve(); if ((message as { failed?: boolean }).failed) reject(new Error("synthetic child failed")); });
        child.once("error", reject); child.send({ stateDir: state, key: active, now });
      });
      let observed: string | undefined;
      const next = refreshVaultCredential({ stateDir: state, key: active, exchange: async (current) => { observed = current.refreshToken; return { accessToken: "parent-access" }; } });
      child.send({ release: true });
      await expect(next).resolves.toBe("parent-access");
      expect(await exited).toBe(0);
      expect(messages).toContainEqual({ result: "child-access" });
      expect(observed).toBe("child-rotated");
      expect(await inspectVaultCredential(state, active)).toMatchObject({ result: "valid", generation: 2 });
    } finally { if (child.exitCode === null) { child.kill(); await exited; } }
  }, 10_000);

  it("retains a consumed claim when recovery publication readback fails", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    const exchange = vi.fn();
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange, testHooks: { afterMarkerPublication: () => { throw new Error("synthetic-readback-failure"); } } })).rejects.toThrow("credential_vault_write_failed");
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).rejects.toThrow("credential_reauthorization_required");
    expect(exchange).not.toHaveBeenCalled();
  });

  it("checks the cooldown boundary without a timer or busy retry", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    const marker = JSON.parse(await readFile(vaultQuarantinePath(state), "utf8"));
    const clock = vi.spyOn(Date, "now").mockReturnValue(Date.parse(marker.updatedAt) + 59_999);
    const exchange = vi.fn(async () => ({ accessToken: "after-boundary" }));
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).rejects.toThrow("credential_refresh_retry_later");
    expect(exchange).not.toHaveBeenCalled(); clock.mockReturnValue(Date.parse(marker.updatedAt) + 60_000);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).resolves.toBe("after-boundary");
    expect(exchange).toHaveBeenCalledTimes(1);
  });

  it("retains fail-closed state when a known OAuth rejection observes a changed vault", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => {
      await writeFile(vaultCredentialPath(state), serializeVaultEnvelope(encryptVaultCredential(2, { ...credential, refreshToken: "changed" }, active)), { mode: 0o600 });
      throw new Error("credential_reauthorization_required");
    } })).rejects.toThrow("credential_vault_conflict");
    expect(await inspectVaultCredential(state, active)).toMatchObject({ result: "quarantined", generation: 2 });
  });

  it("does not dispatch after the vault changes during initial marker publication", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    const exchange = vi.fn();
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange, testHooks: { afterMarkerPublication: async () => {
      await writeFile(vaultCredentialPath(state), serializeVaultEnvelope(encryptVaultCredential(2, credential, active)), { mode: 0o600 });
    } } })).rejects.toThrow("credential_vault_conflict");
    expect(exchange).not.toHaveBeenCalled();
  });

  it("does not consume a recovery budget when another lock owner prevents acquisition", async () => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    const lock = await acquireVaultLock(state, "other-process"); const exchange = vi.fn(async () => ({ accessToken: "after-lock" }));
    try {
      await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).rejects.toThrow("credential_vault_locked");
      expect(JSON.parse(await readFile(vaultQuarantinePath(state), "utf8")).recoveryUsed).toBe(0);
    } finally { await lock.release(); }
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange })).resolves.toBe("after-lock");
    expect(exchange).toHaveBeenCalledTimes(1);
  }, 10_000);

  it.each(["credential_reauthorization_required", "authentication_failed"])("cannot reset the consumed recovery budget after %s", async (failure) => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow("credential_refresh_outcome_uncertain");
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    const recovery = vi.fn(async () => { throw new Error(failure); });
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: recovery })).rejects.toThrow(failure);
    expect(recovery).toHaveBeenCalledTimes(1);
    expect(JSON.parse(await readFile(vaultQuarantinePath(state), "utf8"))).toMatchObject({ state: "quarantined", reason: "exhausted", recoveryUsed: 1 });
    // Recreate the invocation with independent callbacks; disk state, not a
    // process-local counter, must block any further exchange.
    const independent = vi.fn(async () => ({ accessToken: "must-not-dispatch" }));
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: independent })).rejects.toThrow("credential_reauthorization_required");
    expect(independent).not.toHaveBeenCalled();
    expect(await inspectVaultCredential(state, active)).toMatchObject({ result: "quarantined", generation: 1 });
  });

  it.each(["credential_reauthorization_required", "authentication_failed"])("retains the consumed claim when %s recovery observes a changed vault", async (failure) => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => {
      await writeFile(vaultCredentialPath(state), serializeVaultEnvelope(encryptVaultCredential(2, { ...credential, refreshToken: "changed-during-recovery" }, active)), { mode: 0o600 });
      throw new Error(failure);
    } })).rejects.toThrow("credential_vault_write_failed");
    expect(JSON.parse(await readFile(vaultQuarantinePath(state), "utf8"))).toMatchObject({ state: "in_flight", recoveryUsed: 1, generation: 1 });
    const independent = vi.fn();
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: independent })).rejects.toThrow("credential_reauthorization_required");
    expect(independent).not.toHaveBeenCalled();
    expect((await readVaultCredential(state, active)).credential.refreshToken).toBe("changed-during-recovery");
  });

  it.each(["credential_reauthorization_required", "authentication_failed"])("retains the consumed claim when the lock is compromised during %s recovery", async (failure) => {
    const state = await stateDir(); const active = key(); await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow();
    vi.spyOn(Date, "now").mockReturnValue(Date.now() + 60_001);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => {
      await writeFile(`${vaultCredentialPath(state)}.lock`, "synthetic-compromised-lock", { mode: 0o600 });
      throw new Error(failure);
    } })).rejects.toThrow("credential_vault_locked");
    expect(JSON.parse(await readFile(vaultQuarantinePath(state), "utf8"))).toMatchObject({ state: "in_flight", recoveryUsed: 1 });
    const independent = vi.fn();
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: independent })).rejects.toThrow("credential_vault_locked");
    expect(independent).not.toHaveBeenCalled();
  }, 10_000);

});
