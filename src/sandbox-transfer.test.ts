import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { expect, test } from "vitest";
import { copyToSandbox, openSandboxSnapshot } from "./sandbox-transfer.js";

test("round trips binary data and rejects traversal and links", async () => {
  const cwd = mkdtempSync(join(tmpdir(), "graph-transfer-test-"));
  const sandbox = { async run(p: { argv: readonly string[]; input?: string; maxOutputBytes?: number }) {
    const out = spawnSync(p.argv[0]!, p.argv.slice(1), { cwd, input: p.input, encoding: "utf8", maxBuffer: p.maxOutputBytes });
    return { code: out.status ?? -1, stdout: out.stdout ?? "", stderr: out.stderr ?? "" };
  } };
  try {
    const bytes = Buffer.from([0, 1, 127, 255]);
    const copied = await copyToSandbox(sandbox, bytes, ".xlsx");
    const multi = Buffer.alloc(1300000, 0xa7);
    const path = join(cwd, copied.relativePath);
    writeFileSync(path, multi);
    const snapshot = await openSandboxSnapshot(sandbox, copied.relativePath, multi.length);
    const chunks: Buffer[] = []; for await (const chunk of snapshot.chunks()) chunks.push(chunk);
    expect(chunks).toHaveLength(2);
    expect(Buffer.concat(chunks)).toEqual(multi);
    await expect(openSandboxSnapshot(sandbox, copied.relativePath, multi.length - 1)).rejects.toThrow("provider_file_too_large");
    const changed = snapshot.chunks();
    await changed.next();
    writeFileSync(path, Buffer.alloc(multi.length, 0x5a));
    await expect(changed.next()).rejects.toThrow("sandbox_snapshot_failed");
    await expect(openSandboxSnapshot(sandbox, "graph-transfer/../other", 1024)).rejects.toThrow("invalid_sandbox_transfer_filename");
    symlinkSync(copied.relativePath.split("/")[1]!, join(cwd, "graph-transfer", "link.xlsx"));
    await expect(openSandboxSnapshot(sandbox, "graph-transfer/link.xlsx", 1024)).rejects.toThrow("sandbox_snapshot_failed");
  } finally { rmSync(cwd, { recursive: true, force: true }); }
}, 30000);
