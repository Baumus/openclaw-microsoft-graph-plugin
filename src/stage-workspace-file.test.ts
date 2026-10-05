import { createHash } from "node:crypto";
import { link, mkdtemp, mkdir, readdir, rename, rm, symlink, truncate, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { stageWorkspaceFile, validateWorkspaceRelativeFilePath, workspaceFileContentType, WorkspaceStagingStore, WORKSPACE_STAGING_QUOTA_BYTES } from "./stage-workspace-file.js";

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

function saver(directory: string) {
  const saved: Buffer[] = [];
  const save = vi.fn(async (source: AsyncIterable<Uint8Array>, _mime: string, kind: string) => {
    expect(kind).toBe("inbound");
    for await (const chunk of source) saved.push(Buffer.from(chunk));
    const path = join(directory, "media", "inbound", "synthetic-id.pdf");
    await mkdir(join(directory, "media", "inbound"), { recursive: true });
    await writeFile(path, Buffer.concat(saved));
    return { id: "synthetic-id.pdf", path, size: Buffer.concat(saved).byteLength, contentType: "application/pdf" };
  });
  return { save: save as never, saved, calls: save };
}

describe("one-call workspace source staging", () => {
  it("streams a regular nested file with a stable private URI and fingerprint", async () => {
    const directory = await workspace();
    const bytes = Buffer.from("%PDF-1.7\nsynthetic");
    await writeFile(join(directory, "reports", "onepager.pdf"), bytes);
    const { save, saved, calls } = saver(directory);
    const result = await stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", save);
    expect(result).toMatchObject({
      sourceMediaUri: "media://inbound/synthetic-id.pdf",
      sourceSha256: createHash("sha256").update(bytes).digest("hex"),
      sourceByteSize: bytes.byteLength,
    });
    expect(Buffer.concat(saved)).toEqual(bytes);
    expect(calls).toHaveBeenCalledTimes(1);
    await result.lease.cleanup();
    expect(await readdir(join(directory, "media", "inbound"))).toEqual([]);
    expect(workspaceFileContentType("reports/onepager.PDF")).toBe("application/pdf");
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
    const save = (async (source: AsyncIterable<Uint8Array>, _mime: string, _kind: string, _limit: number, name: string) => {
      if (++started === 2) notify();
      await gate;
      let size = 0;
      for await (const chunk of source) size += chunk.byteLength;
      const path = join(directory, name);
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
    const path = join(directory, "media", "inbound", "synthetic-id.pdf");
    await rename(path, `${path}.owned`);
    await writeFile(path, "not owned");
    await staged.lease.cleanup();
    expect(await readdir(join(directory, "media", "inbound"))).toContain("synthetic-id.pdf");
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
    expect(await readdir(join(directory, "media", "inbound"))).toEqual([]);
  });
});
