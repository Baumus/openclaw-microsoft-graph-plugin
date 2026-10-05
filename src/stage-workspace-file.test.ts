import { createHash } from "node:crypto";
import { link, mkdtemp, mkdir, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { stageWorkspaceFile, validateWorkspaceRelativeFilePath, workspaceFileContentType } from "./stage-workspace-file.js";

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

function saver() {
  const saved: Buffer[] = [];
  const save = vi.fn(async (source: AsyncIterable<Uint8Array>, _mime: string, kind: string) => {
    expect(kind).toBe("inbound");
    for await (const chunk of source) saved.push(Buffer.from(chunk));
    return { id: "synthetic-id.pdf", size: Buffer.concat(saved).byteLength, contentType: "application/pdf" };
  });
  return { save: save as never, saved, calls: save };
}

describe("one-call workspace source staging", () => {
  it("streams a regular nested file with a stable private URI and fingerprint", async () => {
    const directory = await workspace();
    const bytes = Buffer.from("%PDF-1.7\nsynthetic");
    await writeFile(join(directory, "reports", "onepager.pdf"), bytes);
    const { save, saved, calls } = saver();
    const result = await stageWorkspaceFile(directory, "reports/onepager.pdf", "application/pdf", save);
    expect(result).toEqual({
      sourceMediaUri: "media://inbound/synthetic-id.pdf",
      sourceSha256: createHash("sha256").update(bytes).digest("hex"),
      sourceByteSize: bytes.byteLength,
    });
    expect(Buffer.concat(saved)).toEqual(bytes);
    expect(calls).toHaveBeenCalledTimes(1);
    expect(workspaceFileContentType("reports/onepager.PDF")).toBe("application/pdf");
  });

  it("rejects absolute, traversal, and ambiguous paths before opening", () => {
    for (const value of ["/tmp/file.pdf", "../secret.pdf", "reports/../secret.pdf", "reports//file.pdf", "reports\\file.pdf", ""]) {
      expect(() => validateWorkspaceRelativeFilePath(value)).toThrow("invalid_workspace_relative_path");
    }
  });

  it("fails closed for missing context, missing files, symlinks, and hardlinks", async () => {
    const directory = await workspace();
    const { save, calls } = saver();
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
    const { save, calls } = saver();
    await expect(stageWorkspaceFile(directory, "reports/cancel.pdf", "application/pdf", save, controller.signal)).rejects.toMatchObject({ name: "AbortError" });
    expect(calls).not.toHaveBeenCalled();
  });

  it("does not return a receipt if the source changes during copy", async () => {
    const directory = await workspace();
    const path = join(directory, "reports", "mutable.pdf");
    await writeFile(path, "first");
    const save = (async (source: AsyncIterable<Uint8Array>) => {
      for await (const _chunk of source) await writeFile(path, "other");
      return { id: "synthetic-id.pdf", size: 5, contentType: "application/pdf" };
    }) as never;
    await expect(stageWorkspaceFile(directory, "reports/mutable.pdf", "application/pdf", save)).rejects.toThrow("workspace_file_changed");
  });
});
