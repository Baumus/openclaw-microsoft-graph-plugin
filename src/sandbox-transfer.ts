import { randomUUID } from "node:crypto";

/** Binary transfer over the host's invocation-bound sandbox command surface.
 * Never expose the encoded bytes as tool parameters or tool results. */
const CHUNK_BYTES = 1024 * 1024;

export type SandboxRunner = {
  run(params: { argv: readonly string[]; input?: string; signal?: AbortSignal; maxOutputBytes?: number }): Promise<{ code: number; stdout: string; stderr: string }>;
};

const script = String.raw`
import base64, os, stat, sys
op, name, limit = sys.argv[1], sys.argv[2], int(sys.argv[3])
root = os.open(".", os.O_RDONLY | os.O_DIRECTORY)
if op == "put":
    try:
        os.mkdir("graph-transfer", 0o700, dir_fd=root)
    except FileExistsError:
        pass
directory = os.open("graph-transfer", os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW, dir_fd=root)
if op == "put":
    encoded_limit = ((limit + 2) // 3) * 4
    data = base64.b64decode(sys.stdin.buffer.read(encoded_limit + 1), validate=True)
    if len(data) > limit:
        raise ValueError("oversize")
    fd = os.open(name, os.O_WRONLY | os.O_CREAT | os.O_EXCL | os.O_NOFOLLOW, 0o600, dir_fd=directory)
    try:
        with os.fdopen(fd, "wb") as output:
            output.write(data)
    except BaseException:
        os.unlink(name, dir_fd=directory)
        raise
elif op in ("inspect", "chunk"):
    fd = os.open(name, os.O_RDONLY | os.O_NOFOLLOW | os.O_NONBLOCK, dir_fd=directory)
    with os.fdopen(fd, "rb") as source:
        def identity(st):
            if not stat.S_ISREG(st.st_mode) or st.st_nlink != 1:
                raise ValueError("unsupported file")
            return ":".join(str(x) for x in (st.st_dev, st.st_ino, st.st_size, st.st_mtime_ns, st.st_ctime_ns))
        before = os.fstat(source.fileno())
        stamp = identity(before)
        if op == "inspect":
            print(f"{before.st_size} {stamp}", end="")
        else:
            expected, start = sys.argv[4], int(sys.argv[5])
            if expected != stamp or not 0 <= start < before.st_size:
                raise ValueError("file changed")
            source.seek(start)
            count = min(limit, before.st_size - start)
            data = source.read(count)
            if len(data) != count or identity(os.fstat(source.fileno())) != stamp:
                raise ValueError("file changed")
            sys.stdout.buffer.write(base64.b64encode(data))
else:
    raise ValueError("invalid operation")
`;

const filename = /^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/u;
const encodedLimit = Math.ceil(CHUNK_BYTES / 3) * 4;

function command(op: "put" | "inspect" | "chunk", name: string) {
  if (!filename.test(name)) throw new Error("invalid_sandbox_transfer_filename");
  return ["python3", "-I", "-c", script, op, name, String(CHUNK_BYTES)];
}

export async function copyToSandbox(sandbox: SandboxRunner, bytes: Uint8Array, extension: string, signal?: AbortSignal) {
  if (bytes.byteLength > CHUNK_BYTES) throw new Error("sandbox_transfer_too_large");
  if (!/^\.[A-Za-z0-9]{1,12}$/u.test(extension)) throw new Error("invalid_sandbox_transfer_filename");
  const name = randomUUID() + extension;
  const result = await sandbox.run({ argv: command("put", name), input: Buffer.from(bytes).toString("base64"), signal, maxOutputBytes: 1024 });
  if (result.code !== 0) throw new Error("sandbox_copy_failed");
  return { relativePath: "graph-transfer/" + name, bytes: bytes.byteLength };
}

/** One source generation, provider-owned total size limit, bounded read messages. */
export async function openSandboxSnapshot(sandbox: SandboxRunner, relativePath: string, maxBytes: number, signal?: AbortSignal) {
  if (!relativePath.startsWith("graph-transfer/")) throw new Error("invalid_sandbox_transfer_filename");
  const name = relativePath.slice("graph-transfer/".length);
  const inspect = async () => {
    signal?.throwIfAborted();
    const result = await sandbox.run({ argv: command("inspect", name), signal, maxOutputBytes: 512 });
    if (result.code !== 0 || !/^\d+ \d+:\d+:\d+:-?\d+:-?\d+$/u.test(result.stdout)) throw new Error("sandbox_snapshot_failed");
    return result.stdout;
  };
  const receipt = await inspect();
  const [sizeText, stamp] = receipt.split(" ");
  const size = Number(sizeText);
  if (!Number.isSafeInteger(size) || size > maxBytes) throw new Error("provider_file_too_large");
  return { size, async *chunks() {
    for (let offset = 0; offset < size;) {
      signal?.throwIfAborted();
      const result = await sandbox.run({ argv: [...command("chunk", name), stamp!, String(offset)], signal, maxOutputBytes: encodedLimit });
      if (result.code !== 0 || result.stdout.length > encodedLimit) throw new Error("sandbox_snapshot_failed");
      const bytes = Buffer.from(result.stdout, "base64");
      if (bytes.length !== Math.min(CHUNK_BYTES, size - offset) || bytes.toString("base64") !== result.stdout) throw new Error("sandbox_snapshot_failed");
      yield bytes;
      offset += bytes.length;
    }
    if (await inspect() !== receipt) throw new Error("sandbox_snapshot_failed");
  } };
}
