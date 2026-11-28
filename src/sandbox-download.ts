import { createHash, randomUUID } from "node:crypto";
import { basename, dirname } from "node:path";
import { root } from "@openclaw/fs-safe";
import { resolveMediaBufferPath } from "openclaw/plugin-sdk/media-store";
import type { SandboxRunner } from "./sandbox-transfer.js";

const CHUNK_BYTES = 1024 * 1024;
const program = String.raw`
import base64, hashlib, os, stat, sys
op, name, offset, digest = sys.argv[1:]
offset = int(offset)
try:
    os.mkdir("graph-transfer", 0o700)
except FileExistsError:
    pass
directory = os.open("graph-transfer", os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW)
temp = "." + name + ".partial"
if op == "cleanup":
    try:
        os.unlink(temp, dir_fd=directory)
    except FileNotFoundError:
        pass
elif op == "start":
    os.close(os.open(temp, os.O_WRONLY | os.O_CREAT | os.O_EXCL | os.O_NOFOLLOW, 0o600, dir_fd=directory))
elif op == "append":
    fd = os.open(temp, os.O_WRONLY | os.O_APPEND | os.O_NOFOLLOW | os.O_NONBLOCK, dir_fd=directory)
    with os.fdopen(fd, "ab") as out:
        st = os.fstat(out.fileno())
        if not stat.S_ISREG(st.st_mode) or st.st_nlink != 1 or st.st_size != offset:
            raise ValueError("invalid partial")
        out.write(base64.b64decode(sys.stdin.buffer.read(1398105), validate=True))
elif op == "finish":
    fd = os.open(temp, os.O_RDONLY | os.O_NOFOLLOW | os.O_NONBLOCK, dir_fd=directory)
    with os.fdopen(fd, "rb") as source:
        st = os.fstat(source.fileno())
        if not stat.S_ISREG(st.st_mode) or st.st_nlink != 1 or st.st_size != offset:
            raise ValueError("invalid partial")
        sha = hashlib.sha256()
        while chunk := source.read(4194304):
            sha.update(chunk)
        after = os.fstat(source.fileno())
        if sha.hexdigest() != digest or (st.st_size, st.st_ctime_ns) != (after.st_size, after.st_ctime_ns):
            raise ValueError("integrity")
        current = os.stat(temp, dir_fd=directory, follow_symlinks=False)
        if (current.st_dev, current.st_ino, current.st_ctime_ns) != (after.st_dev, after.st_ino, after.st_ctime_ns):
            raise ValueError("changed")
        os.link(temp, name, src_dir_fd=directory, dst_dir_fd=directory, follow_symlinks=False)
        os.unlink(temp, dir_fd=directory)
else:
    raise ValueError("invalid operation")
`;

type CopyContext = { sandbox?: SandboxRunner; assertInvocationCurrent?: () => void };

export async function stageDownloadedCopy(
  context: CopyContext, requested: unknown,
  artifact: { id: string; name: string; sizeBytes: number },
  sha256: string, signal?: AbortSignal,
) {
  if (requested !== true) return {};
  const { sandbox, assertInvocationCurrent } = context;
  if (!sandbox || !assertInvocationCurrent) return { sandboxCopy: { ok: false, error: "sandbox_unavailable" } };
  const name = randomUUID() + (/\.[A-Za-z0-9]{1,12}$/u.exec(artifact.name)?.[0] ?? ".bin");
  const run = async (op: string, offset: number, input?: string) => {
    if (op !== "cleanup") {
      signal?.throwIfAborted();
      assertInvocationCurrent();
    }
    const outcome = await sandbox.run({
      argv: ["python3", "-I", "-c", program, op, name, String(offset), sha256],
      input, signal: op === "cleanup" ? undefined : signal, maxOutputBytes: 4096,
    });
    if (outcome.code !== 0) throw new Error("sandbox_copy_failed");
  };
  let opened: Awaited<ReturnType<Awaited<ReturnType<typeof root>>["open"]>> | undefined;
  let started = false;
  try {
    assertInvocationCurrent();
    const path = await resolveMediaBufferPath(artifact.id, "inbound");
    const dir = await root(dirname(path), { hardlinks: "reject", symlinks: "reject" });
    opened = await dir.open(basename(path), { hardlinks: "reject", symlinks: "reject" });
    if (!opened.stat.isFile() || opened.stat.size !== artifact.sizeBytes) throw new Error("sandbox_copy_failed");
    const hash = createHash("sha256");
    started = true;
    await run("start", 0);
    for (let offset = 0; offset < artifact.sizeBytes;) {
      const chunk = Buffer.allocUnsafe(Math.min(CHUNK_BYTES, artifact.sizeBytes - offset));
      const { bytesRead } = await opened.handle.read(chunk, 0, chunk.length, offset);
      if (bytesRead !== chunk.length) throw new Error("sandbox_copy_failed");
      hash.update(chunk);
      await run("append", offset, chunk.toString("base64"));
      offset += bytesRead;
    }
    if (hash.digest("hex") !== sha256) throw new Error("sandbox_copy_failed");
    await run("finish", artifact.sizeBytes);
    started = false;
    return { sandboxCopy: { ok: true, relativePath: "graph-transfer/" + name, bytes: artifact.sizeBytes } };
  } catch {
    return { sandboxCopy: { ok: false, error: "sandbox_copy_failed" } };
  } finally {
    await opened?.handle.close().catch(() => undefined);
    if (started) await run("cleanup", 0).catch(() => undefined);
  }
}
