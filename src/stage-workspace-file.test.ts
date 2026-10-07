import { createHash, randomUUID } from "node:crypto";
import * as fsPromises from "node:fs/promises";
import { chmod, link, mkdtemp, mkdir, readFile, readlink, readdir, rename, rm, symlink, truncate, utimes, writeFile } from "node:fs/promises";
import { hostname, tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { reconcileWorkspaceStaging, stageWorkspaceFile, validateWorkspaceRelativeFilePath, workspaceFileContentType, workspaceStagingProcessIdentity, WorkspaceStagingStore, WORKSPACE_STAGING_QUOTA_BYTES } from "./stage-workspace-file.js";

vi.mock("node:fs/promises", async (importOriginal) => {
  const actual = await importOriginal<typeof import("node:fs/promises")>();
  return { ...actual, lstat: vi.fn(actual.lstat) };
});

const directories: string[] = [];
afterEach(async () => {
  await Promise.all(directories.splice(0).map((directory) => rm(directory, { recursive: true, force: true })));
});

async function workspace() {
  const path = await mkdtemp(join(tmpdir(), "mg-stage-file-"));
  directories.push(path);
  await mkdir(join(path, "reports"), { recursive: true });
  return path;
}

function saver(directory: string, fileName = "synthetic-id.pdf") {
  const saved: Buffer[] = [];
  const save = vi.fn(async (source: AsyncIterable<Uint8Array>, _mime: string, kind: string) => {
    expect(kind).toMatch(/^inbound\/baumus-msgraph-workspace-staging\/[0-9a-f-]{36}$/);
    const path = join(directory, "media", kind, fileName);
    await mkdir(join(directory, "media", kind), { recursive: true });
    for await (const chunk of source) saved.push(Buffer.from(chunk));
    await writeFile(path, Buffer.concat(saved));
    return { id: fileName, path, size: Buffer.concat(saved).byteLength, contentType: "application/pdf" };
  });
  return { save: save as never, saved, calls: save };
}

async function owner(runId: string, pid = process.pid, processStart?: string) {
  const identity = await workspaceStagingProcessIdentity();
  return JSON.stringify({ runId, pid, host: hostname(), ...identity,
    processStart: processStart ?? identity.processStart });
}

describe("one-call workspace source staging", () => {
  it("obtains a verifiable macOS boot and process identity for safe stale-owner recovery", async () => {
    if (process.platform !== "darwin") return;
    const identity = await workspaceStagingProcessIdentity();
    expect(identity.pidNamespace).toBe("platform:darwin");
    expect(identity.bootId).toMatch(/^darwin:\d+$/);
    expect(identity.processStart).toMatch(/^[0-9a-f]{64}$/);
  });

  it("streams a regular nested file with a stable private URI and fingerprint", async () => {
    const directory = await workspace();
    const bytes = Buffer.from("%PDF-1.7\nsynthetic");
    await writeFile(join(directory, "reports", "onepager.pdf"), bytes);
    const { save, saved, calls } = saver(directory);
    const result = await stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", save);
    expect(result).toMatchObject({
      sourceSha256: createHash("sha256").update(bytes).digest("hex"),
      sourceByteSize: bytes.byteLength,
    });
    expect(result.sourceMediaUri).toMatch(/^media:\/\/inbound\/baumus-msgraph-workspace-staging\/[0-9a-f-]{36}\/synthetic-id\.pdf$/);
    expect(Buffer.concat(saved)).toEqual(bytes);
    expect(calls).toHaveBeenCalledTimes(1);
    await result.lease.cleanup();
    expect(await readdir(join(directory, "media", "inbound", "baumus-msgraph-workspace-staging", result.sourceMediaUri.split("/")[4]))).toEqual([]);
    expect(workspaceFileContentType("reports/onepager.PDF")).toBe("application/pdf");
  });

  it("stages through a trusted state root whose parent has a filesystem alias", async () => {
    const directory = await workspace();
    const actual = join(directory, "actual");
    const alias = join(directory, "alias");
    const stateDir = join(alias, "state");
    await mkdir(join(actual, "state"), { recursive: true });
    await symlink(actual, alias);
    await writeFile(join(directory, "reports", "onepager.pdf"), "synthetic");
    const store = new WorkspaceStagingStore();
    const { save } = saver(stateDir);
    const staged = await stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", save, undefined, store, stateDir);
    expect(staged.sourceByteSize).toBe(9);
    await staged.lease.cleanup();
  });

  it("rejects absolute, traversal, and ambiguous paths before opening", () => {
    for (const value of ["/tmp/file.pdf", "../secret.pdf", "reports/../secret.pdf", "reports//file.pdf", "reports\\file.pdf", ""]) {
      expect(() => validateWorkspaceRelativeFilePath(value)).toThrow("invalid_workspace_relative_path");
    }
  });

  it("fails closed for missing context, missing files, symlinks, and hardlinks", async () => {
    const directory = await workspace();
    const { save, calls } = saver(directory);
    await expect(stageWorkspaceFile(undefined, "reports/a.pdf", "application/pdf", save)).rejects.toThrow("workspace_context_unavailable");
    await expect(stageWorkspaceFile(directory, "reports/missing.pdf", "application/pdf", save)).rejects.toThrow("workspace_file_unavailable");
    await writeFile(join(directory, "reports", "real.pdf"), "synthetic");
    await symlink("real.pdf", join(directory, "reports", "link.pdf"));
    await expect(stageWorkspaceFile(directory, "reports/link.pdf", "application/pdf", save)).rejects.toThrow("workspace_file_unavailable");
    await link(join(directory, "reports", "real.pdf"), join(directory, "reports", "hard.pdf"));
    await expect(stageWorkspaceFile(directory, "reports/hard.pdf", "application/pdf", save)).rejects.toThrow("workspace_file_unavailable");
    expect(calls).not.toHaveBeenCalled();
  });

  it("does not save when a local preflight has been cancelled", async () => {
    const directory = await workspace();
    await writeFile(join(directory, "reports", "cancel.pdf"), "synthetic");
    const controller = new AbortController();
    controller.abort();
    const { save, calls } = saver(directory);
    await expect(stageWorkspaceFile(directory, "reports/cancel.pdf", "application/pdf", save, controller.signal)).rejects.toMatchObject({ name: "AbortError" });
    expect(calls).not.toHaveBeenCalled();
  });

  it("does not return a receipt if the source changes during copy", async () => {
    const directory = await workspace();
    const path = join(directory, "reports", "mutable.pdf");
    await writeFile(path, "first");
    const save = (async (source: AsyncIterable<Uint8Array>) => {
      for await (const _chunk of source) await writeFile(path, "other");
      const savedPath = join(directory, "synthetic-id.pdf");
      await writeFile(savedPath, "first");
      return { id: "synthetic-id.pdf", path: savedPath, size: 5, contentType: "application/pdf" };
    }) as never;
    await expect(stageWorkspaceFile(directory, "reports/mutable.pdf", "application/pdf", save)).rejects.toThrow("workspace_file_changed");
    await expect(readdir(directory)).resolves.not.toContain("synthetic-id.pdf");
  });

  it("reserves aggregate bytes atomically and releases idempotently", () => {
    const store = new WorkspaceStagingStore();
    const first = store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2);
    const second = store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2);
    expect(() => store.reserve(1)).toThrow("workspace_staging_quota_exceeded");
    first();
    first();
    const third = store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2);
    expect(() => store.reserve(64 * 1024 * 1024 + 1)).toThrow("workspace_staging_quota_exceeded");
    second();
    third();
    expect(() => store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2)).not.toThrow();
  });

  it("limits empty-file reservations by count as well as bytes", () => {
    const store = new WorkspaceStagingStore();
    const releases = Array.from({ length: 64 }, () => store.reserve(0));
    expect(() => store.reserve(0)).toThrow("workspace_staging_quota_exceeded");
    releases.forEach((release) => release());
    expect(() => store.reserve(0)).not.toThrow();
  });

  it("rejects oversized sparse files before saving", async () => {
    const directory = await workspace();
    await writeFile(join(directory, "reports", "large.pdf"), "");
    await truncate(join(directory, "reports", "large.pdf"), 64 * 1024 * 1024 + 1);
    const { save, calls } = saver(directory);
    await expect(stageWorkspaceFile(directory, "reports/large.pdf", "application/pdf", save)).rejects.toThrow("workspace_staging_quota_exceeded");
    expect(calls).not.toHaveBeenCalled();
  });

  it("holds concurrent reservations through streaming and frees them after cleanup", async () => {
    const directory = await workspace();
    const store = new WorkspaceStagingStore();
    for (const name of ["a.pdf", "b.pdf", "c.pdf"]) {
      const path = join(directory, "reports", name);
      await writeFile(path, "");
      await truncate(path, name === "c.pdf" ? 1 : 64 * 1024 * 1024);
    }
    let resume!: () => void;
    const gate = new Promise<void>((resolve) => { resume = resolve; });
    let started = 0;
    let notify!: () => void;
    const bothStarted = new Promise<void>((resolve) => { notify = resolve; });
    const save = (async (source: AsyncIterable<Uint8Array>, _mime: string, kind: string, _limit: number, name: string) => {
      if (++started === 2) notify();
      await gate;
      let size = 0;
      for await (const chunk of source) size += chunk.byteLength;
      const path = join(directory, "media", kind, name);
      await mkdir(join(directory, "media", kind), { recursive: true });
      await writeFile(path, "");
      await truncate(path, size);
      return { id: name, path, size };
    }) as never;
    const first = stageWorkspaceFile(directory, "reports/a.pdf", "application/pdf", save, undefined, store);
    const second = stageWorkspaceFile(directory, "reports/b.pdf", "application/pdf", save, undefined, store);
    await bothStarted;
    await expect(stageWorkspaceFile(directory, "reports/c.pdf", "application/pdf", save, undefined, store)).rejects.toThrow("workspace_staging_quota_exceeded");
    resume();
    const results = await Promise.all([first, second]);
    await results[0].lease.cleanup();
    const third = await stageWorkspaceFile(directory, "reports/c.pdf", "application/pdf", save, undefined, store);
    await Promise.all([results[1].lease.cleanup(), third.lease.cleanup()]);
  });

  it("keeps a replaced foreign artifact instead of deleting it", async () => {
    const directory = await workspace();
    await writeFile(join(directory, "reports", "onepager.pdf"), "synthetic");
    const { save } = saver(directory);
    const staged = await stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", save);
    const path = join(directory, "media", staged.sourceMediaUri.slice("media://".length));
    await rename(path, `${path}.owned`);
    await writeFile(path, "not owned");
    await staged.lease.cleanup();
    expect(await readdir(join(directory, "media", "inbound", "baumus-msgraph-workspace-staging", staged.sourceMediaUri.split("/")[4]))).toContain("synthetic-id.pdf");
  });

  it("reclaims a published copy when cancellation occurs immediately after saving", async () => {
    const directory = await workspace();
    await writeFile(join(directory, "reports", "onepager.pdf"), "synthetic");
    const controller = new AbortController();
    const { save } = saver(directory);
    const abortingSave = (async (source: AsyncIterable<Uint8Array>, mime: string, kind: string) => {
      const saved = await (save as (stream: AsyncIterable<Uint8Array>, contentType: string, subdir: string) => Promise<unknown>)(source, mime, kind);
      controller.abort();
      return saved;
    }) as typeof save;
    await expect(stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", abortingSave, controller.signal))
      .rejects.toMatchObject({ name: "AbortError" });
    const [runId] = await readdir(join(directory, "media", "inbound", "baumus-msgraph-workspace-staging"));
    expect(await readdir(join(directory, "media", "inbound", "baumus-msgraph-workspace-staging", runId))).toEqual([]);
  });

  it("cleans a saved copy even when its receipt fails validation, without deleting foreign media", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    await writeFile(join(directory, "reports", "onepager.pdf"), "synthetic");
    const { save } = saver(stateDir);
    const store = new WorkspaceStagingStore();
    const badReceipt = (async (source: AsyncIterable<Uint8Array>, mime: string, kind: string) => {
      const saved = await (save as (stream: AsyncIterable<Uint8Array>, type: string, subdir: string) => Promise<{ id: string; path: string; size: number }>)(source, mime, kind);
      return { ...saved, id: "wrong-id.pdf" };
    }) as typeof save;
    await expect(stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", badReceipt, undefined, store, stateDir))
      .rejects.toThrow("workspace_file_unavailable");
    const run = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", store.runId);
    expect(await readdir(run)).toEqual([".workspace-staging-owner"]);
    const outside = join(directory, "foreign.pdf");
    await writeFile(outside, "keep");
    const foreignSave = (async () => ({ id: "foreign.pdf", path: outside, size: 4, contentType: "application/pdf" })) as typeof save;
    await expect(stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", foreignSave, undefined, store, stateDir))
      .rejects.toThrow("workspace_file_unavailable");
    expect(await readFile(outside, "utf8")).toBe("keep");
    expect(await readdir(run)).toEqual([".workspace-staging-owner"]);
  });

  it("queues an unsafe saved symlink instead of deleting foreign data or releasing its reservation", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    await writeFile(join(directory, "reports", "onepager.pdf"), "synthetic");
    const outside = join(directory, "outside.pdf");
    await writeFile(outside, "keep");
    const store = new WorkspaceStagingStore();
    const save = (async (_source: AsyncIterable<Uint8Array>, _type: string, kind: string) => {
      const path = join(stateDir, "media", kind, "claimed.pdf");
      await symlink(outside, path);
      return { id: "claimed.pdf", path, size: 4, contentType: "application/pdf" };
    }) as never;
    await expect(stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", save, undefined, store, stateDir))
      .rejects.toThrow("workspace_file_unavailable");
    expect(await readFile(outside, "utf8")).toBe("keep");
    const releaseHalf = store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2);
    expect(() => store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2)).toThrow("workspace_staging_quota_exceeded");
    releaseHalf();
    const claimed = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", store.runId, "claimed.pdf");
    expect(await readlink(claimed)).toBe(outside);
    await rm(claimed);
    await store.retryFailedCleanups();
    const release = store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2);
    expect(() => store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2)).not.toThrow();
    release();
  });

  it("retains quota and queues a saved copy when its post-save lstat cannot run", async () => {
    if (process.getuid?.() === 0) return;
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    await writeFile(join(directory, "reports", "onepager.pdf"), "synthetic");
    const store = new WorkspaceStagingStore();
    const { save } = saver(stateDir);
    const run = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", store.runId);
    const unreadableSave = (async (source: AsyncIterable<Uint8Array>, mime: string, kind: string) => {
      const saved = await (save as (stream: AsyncIterable<Uint8Array>, type: string, subdir: string) => Promise<unknown>)(source, mime, kind);
      await chmod(run, 0);
      return saved;
    }) as typeof save;
    try {
      await expect(stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", unreadableSave, undefined, store, stateDir))
        .rejects.toThrow("workspace_file_unavailable");
      const releaseHalf = store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2);
      expect(() => store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2)).toThrow("workspace_staging_quota_exceeded");
      releaseHalf();
    } finally { await chmod(run, 0o700); }
    expect(await readdir(run)).toContain("synthetic-id.pdf");
    await rm(join(run, "synthetic-id.pdf"));
    await store.retryFailedCleanups();
    const release = store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2);
    expect(() => store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2)).not.toThrow();
    release();
  });

  it("recovers a dead quota owner and a PID reused in the same boot", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    const store = new WorkspaceStagingStore();
    await store.ownRun(stateDir);
    const lock = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", ".workspace-staging-quota-lock");
    for (const [pid, start] of [[99999999, "0"], [process.pid, "0"]] as const) {
      await mkdir(lock);
      await writeFile(join(lock, ".workspace-staging-lock-owner"), JSON.stringify({ ...JSON.parse(await owner(randomUUID(), pid, start)), token: randomUUID() }));
      const reservation = await store.reserveShared(1, stateDir);
      await reservation.release();
      await expect(readdir(lock)).rejects.toMatchObject({ code: "ENOENT" });
    }
    await mkdir(lock);
    await writeFile(join(lock, ".workspace-staging-lock-owner"), JSON.stringify({ ...JSON.parse(await owner(randomUUID())), bootId: randomUUID(), token: randomUUID() }));
    const rebooted = await store.reserveShared(1, stateDir);
    await rebooted.release();
    await expect(readdir(lock)).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("retries when another recovery replaces the quota lock before identity verification", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    const store = new WorkspaceStagingStore();
    await store.ownRun(stateDir);
    const lock = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", ".workspace-staging-quota-lock");
    const marker = JSON.stringify({ ...JSON.parse(await owner(randomUUID(), 99999999, "0")), token: randomUUID() });
    await mkdir(lock);
    await writeFile(join(lock, ".workspace-staging-lock-owner"), marker);
    const { lstat: originalLstat } = await vi.importActual<typeof import("node:fs/promises")>("node:fs/promises");
    let observations = 0;
    const lstatMock = vi.mocked(fsPromises.lstat);
    lstatMock.mockImplementation(async (path) => {
      if (path === lock && ++observations === 2) {
        const previous = `${lock}-previous`;
        await rename(lock, previous);
        await mkdir(lock);
        await writeFile(join(lock, ".workspace-staging-lock-owner"), marker);
        await rm(previous, { recursive: true });
      }
      return originalLstat(path);
    });
    try {
      const reservation = await store.reserveShared(1, stateDir);
      expect(observations).toBeGreaterThan(2);
      await reservation.release();
      await expect(readdir(lock)).rejects.toMatchObject({ code: "ENOENT" });
    } finally { lstatMock.mockImplementation(originalLstat); }
  });

  it.each(["symlink", "foreign file"] as const)("does not accept a %s replacing the quota lock", async (kind) => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    const store = new WorkspaceStagingStore();
    await store.ownRun(stateDir);
    const lock = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", ".workspace-staging-quota-lock");
    await mkdir(lock);
    const foreign = join(directory, "foreign");
    await mkdir(foreign);
    await writeFile(join(foreign, "keep"), "keep");
    const { lstat: originalLstat } = await vi.importActual<typeof import("node:fs/promises")>("node:fs/promises");
    const lstatMock = vi.mocked(fsPromises.lstat);
    let observations = 0;
    lstatMock.mockImplementation(async (path) => {
      if (path === lock && ++observations === 2) {
        const previous = `${lock}-previous`;
        await rename(lock, previous);
        if (kind === "symlink") await symlink(foreign, lock);
        else {
          await mkdir(lock);
          await writeFile(join(lock, "foreign"), "keep");
        }
        await rm(previous, { recursive: true });
      }
      return originalLstat(path);
    });
    const timer = vi.spyOn(globalThis, "setTimeout").mockImplementation(((callback: () => void) => {
      queueMicrotask(callback);
      return 0 as unknown as NodeJS.Timeout;
    }) as typeof setTimeout);
    try {
      await expect(store.reserveShared(1, stateDir)).rejects.toThrow(kind === "symlink" ? "workspace_staging_namespace_invalid" : "workspace_staging_quota_exceeded");
      expect(observations).toBeGreaterThan(1);
      expect(await readFile(join(foreign, "keep"), "utf8")).toBe("keep");
      if (kind === "symlink") expect(await readlink(lock)).toBe(foreign);
      else expect(await readdir(lock)).toEqual(["foreign"]);
    } finally {
      timer.mockRestore();
      lstatMock.mockImplementation(originalLstat);
    }
  });

  it("fails closed for a live, foreign, or symlinked quota lock", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    const store = new WorkspaceStagingStore();
    await store.ownRun(stateDir);
    const lock = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", ".workspace-staging-quota-lock");
    const foreign = join(directory, "foreign");
    await mkdir(foreign);
    const timer = vi.spyOn(globalThis, "setTimeout").mockImplementation(((callback: () => void) => {
      queueMicrotask(callback);
      return 0 as unknown as NodeJS.Timeout;
    }) as typeof setTimeout);
    try {
      for (const host of [hostname(), "other-host"]) {
        await mkdir(lock);
        await writeFile(join(lock, ".workspace-staging-lock-owner"), JSON.stringify({ ...JSON.parse(await owner(randomUUID())), host, token: randomUUID() }));
        await expect(store.reserveShared(1, stateDir)).rejects.toThrow("workspace_staging_quota_exceeded");
        expect(await readdir(lock)).toEqual([".workspace-staging-lock-owner"]);
        await rm(lock, { recursive: true });
      }
      await symlink(foreign, lock);
      await expect(store.reserveShared(1, stateDir)).rejects.toThrow("workspace_staging_namespace_invalid");
      await rm(lock);
      await mkdir(lock);
      await symlink(join(foreign, "owner"), join(lock, ".workspace-staging-lock-owner"));
      await expect(store.reserveShared(1, stateDir)).rejects.toThrow("workspace_staging_quota_exceeded");
      expect(await readdir(foreign)).toEqual([]);
    } finally { timer.mockRestore(); }
  }, 30_000);

  it("reclaims only an old, empty unmarked quota lock", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    const store = new WorkspaceStagingStore();
    await store.ownRun(stateDir);
    const lock = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", ".workspace-staging-quota-lock");
    await mkdir(lock);
    const timer = vi.spyOn(globalThis, "setTimeout").mockImplementation(((callback: () => void) => {
      queueMicrotask(callback);
      return 0 as unknown as NodeJS.Timeout;
    }) as typeof setTimeout);
    try {
      await expect(store.reserveShared(1, stateDir)).rejects.toThrow("workspace_staging_quota_exceeded");
      expect(await readdir(lock)).toEqual([]);
      const clock = vi.spyOn(Date, "now").mockReturnValue(Date.now() + 9 * 24 * 60 * 60_000);
      try {
        const reservation = await store.reserveShared(1, stateDir);
        await reservation.release();
        await expect(readdir(lock)).rejects.toMatchObject({ code: "ENOENT" });
      } finally { clock.mockRestore(); }
    } finally { timer.mockRestore(); }
  });

  it("reconciles crashed runs without touching active, recent, or unrelated inbound files", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    const inbound = join(stateDir, "media", "inbound");
    const namespace = join(inbound, "baumus-msgraph-workspace-staging");
    const staleId = randomUUID();
    const recentId = randomUUID();
    const activeId = randomUUID();
    const foreignId = randomUUID();
    const fileId = randomUUID();
    const tempId = `.${randomUUID()}.123.${randomUUID()}.tmp`;
    for (const runId of [staleId, recentId, activeId, foreignId]) {
      await mkdir(join(namespace, runId), { recursive: true });
      await writeFile(join(namespace, runId, `${fileId}.pdf`), "synthetic");
      if (runId === staleId) {
        await mkdir(join(namespace, runId, tempId));
        await writeFile(join(namespace, runId, tempId, "partial"), "synthetic");
      } else await writeFile(join(namespace, runId, tempId), "synthetic");
      if (runId !== foreignId) await writeFile(join(namespace, runId, ".workspace-staging-owner"), await owner(runId, runId === staleId ? 99999999 : process.pid));
    }
    await writeFile(join(inbound, "unrelated.pdf"), "keep");
    const stale = new Date(Date.now() - 9 * 24 * 60 * 60_000);
    await utimes(join(namespace, staleId), stale, stale);
    await utimes(join(namespace, activeId), stale, stale);
    await utimes(join(namespace, foreignId), stale, stale);
    await reconcileWorkspaceStaging(stateDir, activeId);
    expect(await readdir(namespace)).toEqual(expect.arrayContaining([recentId, activeId, foreignId]));
    expect(await readdir(namespace)).not.toContain(staleId);
    expect(await readdir(inbound)).toContain("unrelated.pdf");
    expect(await readdir(join(namespace, activeId))).toContain(`${fileId}.pdf`);
    expect(await readdir(join(namespace, foreignId))).toContain(`${fileId}.pdf`);
  });

  it("reconciles a real prior run on startup, leaving foreign files inside the marked run", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    const source = join(directory, "reports", "onepager.pdf");
    await mkdir(stateDir);
    await writeFile(source, "synthetic");
    const { save } = saver(stateDir, `${randomUUID()}.pdf`);
    const previous = new WorkspaceStagingStore();
    const staged = await stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", save, undefined, previous, stateDir);
    const run = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", previous.runId);
    await writeFile(join(run, "foreign-note.txt"), "keep");
    await writeFile(join(run, ".workspace-staging-owner"), await owner(previous.runId, 99999999));
    const restarted = new WorkspaceStagingStore();
    restarted.startReconciliation(stateDir);
    await vi.waitFor(async () => expect(await readdir(run)).toEqual([".workspace-staging-owner", "foreign-note.txt"]));
    expect(await readdir(run)).toContain("foreign-note.txt");
    expect(staged.sourceMediaUri).toContain(previous.runId);
  });

  it("rejects symlinked ancestors rather than deleting outside the state directory", async () => {
    const directory = await workspace();
    const actual = join(directory, "actual");
    const stateDir = join(directory, "state");
    const namespace = join(actual, "media", "inbound", "baumus-msgraph-workspace-staging");
    const runId = randomUUID();
    await mkdir(join(namespace, runId), { recursive: true });
    await writeFile(join(namespace, runId, ".workspace-staging-owner"), await owner(runId, 99999999));
    const fileName = `${randomUUID()}.pdf`;
    await writeFile(join(namespace, runId, fileName), "keep");
    await symlink(actual, stateDir);
    await expect(reconcileWorkspaceStaging(stateDir)).rejects.toThrow("workspace_staging_namespace_invalid");
    await rm(stateDir);
    for (const ancestor of ["media", "inbound"]) {
      const staging = join(directory, `state-${ancestor}`);
      await mkdir(join(staging, ...(ancestor === "inbound" ? ["media"] : [])), { recursive: true });
      await symlink(join(actual, "media", ...(ancestor === "inbound" ? ["inbound"] : [])), join(staging, "media", ...(ancestor === "inbound" ? ["inbound"] : [])));
      await expect(reconcileWorkspaceStaging(staging)).rejects.toThrow("workspace_staging_namespace_invalid");
    }
    expect(await readdir(join(namespace, runId))).toContain(fileName);
  });

  it("reclaims a reused PID promptly, but preserves a live owner with unavailable identity", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    const namespace = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging");
    const reused = randomUUID();
    const rebooted = randomUUID();
    const legacy = randomUUID();
    for (const runId of [reused, rebooted, legacy]) {
      await mkdir(join(namespace, runId), { recursive: true });
      await writeFile(join(namespace, runId, `${randomUUID()}.pdf`), "copy");
    }
    await writeFile(join(namespace, reused, ".workspace-staging-owner"), await owner(reused, process.pid, "0"));
    await writeFile(join(namespace, rebooted, ".workspace-staging-owner"), JSON.stringify({ ...JSON.parse(await owner(rebooted)), bootId: randomUUID() }));
    await writeFile(join(namespace, legacy, ".workspace-staging-owner"), JSON.stringify({ runId: legacy, pid: process.pid, host: hostname(), pidNamespace: (await workspaceStagingProcessIdentity()).pidNamespace }));
    const old = new Date(Date.now() - 9 * 24 * 60 * 60_000);
    await utimes(join(namespace, legacy), old, old);
    await reconcileWorkspaceStaging(stateDir);
    expect(await readdir(namespace)).not.toContain(reused);
    expect(await readdir(namespace)).not.toContain(rebooted);
    expect(await readdir(namespace)).toContain(legacy);
  });

  it("enforces shared bytes and slots across independent stores, including orphan copies", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    const first = new WorkspaceStagingStore();
    const second = new WorkspaceStagingStore();
    await first.ownRun(stateDir);
    await second.ownRun(stateDir);
    const held = await first.reserveShared(64 * 1024 * 1024, stateDir);
    const other = await second.reserveShared(64 * 1024 * 1024, stateDir);
    await expect(second.reserveShared(1, stateDir)).rejects.toThrow("workspace_staging_quota_exceeded");
    await held.release();
    const orphan = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", first.runId, `${randomUUID()}.pdf`);
    await writeFile(orphan, "");
    await truncate(orphan, 64 * 1024 * 1024);
    await expect(first.reserveShared(1, stateDir)).rejects.toThrow("workspace_staging_quota_exceeded");
    await other.release();
    const slots = [];
    for (let slot = 0; slot < 63; slot++) slots.push(await second.reserveShared(0, stateDir));
    await expect(second.reserveShared(0, stateDir)).rejects.toThrow("workspace_staging_quota_exceeded");
    await Promise.all(slots.map((slot) => slot.release()));
  }, 20_000);

  it("retries after a blocked orphan cleanup without deleting a foreign symlink", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    const runId = randomUUID();
    const run = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", runId);
    const temp = `.${randomUUID()}.123.${randomUUID()}.tmp`;
    await mkdir(join(run, temp), { recursive: true });
    await writeFile(join(run, ".workspace-staging-owner"), await owner(runId, 99999999));
    await symlink(join(directory, "reports"), join(run, temp, "foreign"));
    const stale = new Date(Date.now() - 9 * 24 * 60 * 60_000);
    await utimes(run, stale, stale);
    await expect(reconcileWorkspaceStaging(stateDir)).rejects.toThrow("workspace_staging_reconciliation_failed");
    expect(await readdir(join(run, temp))).toContain("foreign");
    await rm(join(run, temp, "foreign"));
    await reconcileWorkspaceStaging(stateDir);
    expect(await readdir(join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging"))).not.toContain(runId);
  });

  it("records ownership before saving so a crash before streaming remains discoverable", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    await mkdir(stateDir);
    await writeFile(join(directory, "reports", "onepager.pdf"), "synthetic");
    const store = new WorkspaceStagingStore();
    const interruptedSave = vi.fn(async () => { throw new Error("interrupted"); }) as never;
    await expect(stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", interruptedSave, undefined, store, stateDir))
      .rejects.toThrow("workspace_file_unavailable");
    const run = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", store.runId);
    expect(await readdir(run)).toEqual([".workspace-staging-owner"]);
    await writeFile(join(run, ".workspace-staging-owner"), await owner(store.runId, 99999999));
    await reconcileWorkspaceStaging(stateDir);
    expect(await readdir(join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging"))).not.toContain(store.runId);
  });

  it("does not claim a preexisting foreign directory with the same run ID", async () => {
    const directory = await workspace();
    const stateDir = join(directory, "state");
    const store = new WorkspaceStagingStore();
    const run = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging", store.runId);
    await mkdir(run, { recursive: true });
    await writeFile(join(run, "foreign.pdf"), "keep");
    await writeFile(join(directory, "reports", "onepager.pdf"), "synthetic");
    const { save, calls } = saver(stateDir);
    await expect(stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", save, undefined, store, stateDir))
      .rejects.toThrow("workspace_file_unavailable");
    expect(calls).not.toHaveBeenCalled();
    expect(await readdir(run)).toEqual(["foreign.pdf"]);
  });

  it("retries a failed terminal cleanup without permanently holding its quota", async () => {
    const store = new WorkspaceStagingStore();
    const releaseFirst = store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2);
    const releaseSecond = store.reserve(WORKSPACE_STAGING_QUOTA_BYTES / 2);
    const lease = { cleanup: vi.fn().mockRejectedValueOnce(new Error("temporarily locked")).mockImplementation(async () => { releaseFirst(); releaseSecond(); }) };
    store.bind("failed-call", "onedrive_upload", lease, undefined, () => undefined);
    await expect(store.cleanup("failed-call")).rejects.toThrow("temporarily locked");
    expect(() => store.reserve(1)).toThrow("workspace_staging_quota_exceeded");
    await store.retryFailedCleanups();
    expect(lease.cleanup).toHaveBeenCalledTimes(2);
    store.reserve(1)();
  });
});
