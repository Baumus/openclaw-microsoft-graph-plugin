import { createHash, randomUUID } from "node:crypto";
import { constants } from "node:fs";
import { lstat, open, readdir, readlink, rmdir } from "node:fs/promises";
import { hostname } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { root as secureRoot, type OpenResult } from "openclaw/plugin-sdk/infra-runtime";
import { saveMediaStream } from "openclaw/plugin-sdk/media-store";
import { sanitizeAttachmentName } from "./attachment-download.js";
import { ONEDRIVE_WRITE_MAX_BYTES } from "./graph.js";

const CHUNK_BYTES = 4 * 1024 * 1024;
export const WORKSPACE_STAGING_QUOTA_BYTES = 128 * 1024 * 1024;
const WORKSPACE_STAGING_FILE_MAX_BYTES = 64 * 1024 * 1024;
const MAX_WORKSPACE_STAGING_RESERVATIONS = 64;
const STAGING_NAMESPACE = "baumus-msgraph-workspace-staging";
const STAGING_SUBDIR = `inbound/${STAGING_NAMESPACE}`;
const STALE_STAGING_AGE_MS = 8 * 24 * 60 * 60_000;
const RECONCILIATION_INTERVAL_MS = 60 * 60_000;
const OWNER_FILE = ".workspace-staging-owner";
const RUN_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const STAGED_FILE = /^(?:[\p{L}\p{N}._-]{1,60}---)?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}(?:\.[a-z0-9]{1,16})?$/u;
const STAGED_TEMP = /^\.[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.[0-9]+\.[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.tmp$/;
const pidNamespace = readlink("/proc/self/ns/pid").catch(() => undefined);

function processAlive(pid: number): boolean {
  try { process.kill(pid, 0); return true; }
  catch (error) { return (error as NodeJS.ErrnoException).code !== "ESRCH"; }
}

export type WorkspaceStagingLease = { cleanup(): Promise<void> };

export class WorkspaceStagingStore {
  readonly runId = randomUUID();
  private reservedBytes = 0;
  private reservations = 0;
  private readonly pending = new Map<string, { lease: WorkspaceStagingLease; timer: NodeJS.Timeout; sessionId?: string; toolName: string; onExpire: () => void; executing: boolean }>();
  private readonly reconciling = new Map<string, NodeJS.Timeout>();
  private readonly owners = new Map<string, Promise<void>>();
  private readonly failedCleanups = new Set<WorkspaceStagingLease>();

  retryCleanup(lease: WorkspaceStagingLease): void { this.failedCleanups.add(lease); }

  async retryFailedCleanups(): Promise<void> {
    for (const lease of this.failedCleanups) {
      try { await lease.cleanup(); this.failedCleanups.delete(lease); }
      catch { this.failedCleanups.add(lease); }
    }
  }

  async ownRun(stateDir: string): Promise<void> {
    const key = resolve(stateDir);
    let owner = this.owners.get(key);
    if (!owner) {
      owner = (async () => {
        const directory = join(key, "media", STAGING_SUBDIR, this.runId);
        const state = await secureRoot(key, { symlinks: "reject", hardlinks: "reject" });
        if (await lstat(directory).then(() => true, (error: NodeJS.ErrnoException) => { if (error.code === "ENOENT") return false; throw error; }))
          throw new Error("workspace_staging_namespace_invalid");
        await state.mkdir(`media/${STAGING_SUBDIR}/${this.runId}`, { mutationSymlinks: "reject", private: true });
        const identity = await lstat(directory);
        if (!identity.isDirectory() || identity.isSymbolicLink()) throw new Error("workspace_staging_namespace_invalid");
        const run = await secureRoot(directory, { symlinks: "reject", hardlinks: "reject" });
        await run.create(OWNER_FILE, JSON.stringify({ runId: this.runId, pid: process.pid, host: hostname(), pidNamespace: await pidNamespace }), { private: true, mutationSymlinks: "reject" });
        const current = await lstat(directory);
        if (current.dev !== identity.dev || current.ino !== identity.ino) throw new Error("workspace_staging_namespace_invalid");
      })();
      this.owners.set(key, owner);
      void owner.catch(() => { if (this.owners.get(key) === owner) this.owners.delete(key); });
    }
    await owner;
  }

  startReconciliation(stateDir: string, onFailure: (error: unknown) => void = () => undefined): void {
    const key = resolve(stateDir);
    if (this.reconciling.has(key)) return;
    void reconcileWorkspaceStaging(key, this.runId).catch(onFailure);
    const timer = setInterval(() => {
      const activeDirectory = join(key, "media", STAGING_SUBDIR, this.runId);
      if (this.reservations) void open(activeDirectory, constants.O_RDONLY | constants.O_DIRECTORY | constants.O_NOFOLLOW).then(async (handle) => {
        try { await handle.utimes(new Date(), new Date()); }
        finally { await handle.close(); }
      }).catch(() => undefined);
      void this.retryFailedCleanups().catch(onFailure);
      void reconcileWorkspaceStaging(key, this.runId).catch(onFailure);
    }, RECONCILIATION_INTERVAL_MS);
    timer.unref();
    this.reconciling.set(key, timer);
  }

  reserve(size: number): () => void {
    if (!Number.isSafeInteger(size) || size < 0 || this.reservations >= MAX_WORKSPACE_STAGING_RESERVATIONS || size > WORKSPACE_STAGING_FILE_MAX_BYTES
      || size > WORKSPACE_STAGING_QUOTA_BYTES - this.reservedBytes) throw new Error("workspace_staging_quota_exceeded");
    this.reservedBytes += size;
    this.reservations++;
    let released = false;
    return () => { if (!released) { released = true; this.reservedBytes -= size; this.reservations--; } };
  }

  bind(toolCallId: string, toolName: string, lease: WorkspaceStagingLease, sessionId: string | undefined, onExpire: () => void): void {
    if (!toolCallId || this.pending.has(toolCallId)) throw new Error("approval_context_invalid_or_changed");
    const timer = setTimeout(() => { onExpire(); void this.cleanup(toolCallId).catch(() => undefined); }, 15 * 60_000);
    timer.unref();
    this.pending.set(toolCallId, { lease, timer, sessionId, toolName, onExpire, executing: false });
  }

  has(toolCallId: string): boolean { return this.pending.has(toolCallId); }

  beginExecution(toolCallId: string, toolName: string, sessionId?: string): void {
    const entry = this.pending.get(toolCallId);
    if (entry && entry.toolName === toolName && entry.sessionId === sessionId) {
      entry.executing = true;
      clearTimeout(entry.timer);
    }
  }

  async cleanup(toolCallId: string | undefined, toolName?: string, sessionId?: string): Promise<void> {
    if (!toolCallId) return;
    const entry = this.pending.get(toolCallId);
    if (!entry || (toolName && entry.toolName !== toolName) || (sessionId && entry.sessionId !== sessionId)) return;
    this.pending.delete(toolCallId);
    clearTimeout(entry.timer);
    try { await entry.lease.cleanup(); }
    catch (error) { this.retryCleanup(entry.lease); throw error; }
  }

  clearSession(sessionId: string): void {
    for (const [toolCallId, entry] of this.pending) if (entry.sessionId === sessionId && !entry.executing) {
      entry.onExpire();
      void this.cleanup(toolCallId).catch(() => undefined);
    }
  }
}

export async function reconcileWorkspaceStaging(stateDir: string, activeRunId?: string, now = Date.now()): Promise<void> {
  const directory = join(stateDir, "media", STAGING_SUBDIR);
  const namespace = await lstat(directory).catch((error: NodeJS.ErrnoException) => {
    if (error.code === "ENOENT") return undefined;
    throw error;
  });
  if (!namespace) return;
  if (!namespace.isDirectory() || namespace.isSymbolicLink()) throw new Error("workspace_staging_namespace_invalid");
  const failures: unknown[] = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    if (!RUN_ID.test(entry.name) || entry.name === activeRunId || !entry.isDirectory()) continue;
    const path = join(directory, entry.name);
    const identity = await lstat(path).catch(() => undefined);
    if (!identity?.isDirectory()) continue;
    try {
      const owner = await lstat(join(path, OWNER_FILE)).catch(() => undefined);
      if (!owner?.isFile() || owner.nlink !== 1 || owner.size > 512) continue;
      const run = await secureRoot(path, { symlinks: "reject", hardlinks: "reject" });
      const marker = await run.open(OWNER_FILE, { symlinks: "reject", hardlinks: "reject" });
      let lease: { runId?: unknown; pid?: unknown; host?: unknown; pidNamespace?: unknown };
      try { lease = JSON.parse(await marker.handle.readFile({ encoding: "utf8" })); }
      finally { await marker.handle.close(); }
      if (lease?.runId !== entry.name || !Number.isSafeInteger(lease.pid) || (lease.pid as number) <= 0 || typeof lease.host !== "string") continue;
      const sameProcessScope = lease.host === hostname() && !!lease.pidNamespace && lease.pidNamespace === await pidNamespace;
      if (sameProcessScope ? processAlive(lease.pid as number) : now - identity.mtimeMs < STALE_STAGING_AGE_MS) continue;
      const current = await lstat(path);
      if (current.dev !== identity.dev || current.ino !== identity.ino
        || (!sameProcessScope && now - current.mtimeMs < STALE_STAGING_AGE_MS)) continue;
      for (const file of await readdir(path, { withFileTypes: true })) {
        if (file.name === OWNER_FILE || (!STAGED_FILE.test(file.name) && !STAGED_TEMP.test(file.name))) continue;
        if (file.isDirectory() && STAGED_TEMP.test(file.name)) {
          const temporary = await lstat(join(path, file.name));
          if (!temporary.isDirectory()) continue;
          await run.remove(file.name, { recursive: true, mutationSymlinks: "reject", maxEntries: 2048, maxDepth: 2 });
          continue;
        }
        if (!file.isFile()) continue;
        const opened = await run.open(file.name, { symlinks: "reject", hardlinks: "reject" });
        try {
          if (!opened.stat.isFile() || opened.stat.nlink !== 1) continue;
        } finally { await opened.handle.close(); }
        await run.remove(file.name, { mutationSymlinks: "reject" });
      }
      if ((await readdir(path)).length !== 1) continue;
      await run.remove(OWNER_FILE, { mutationSymlinks: "reject" });
      const final = await lstat(path);
      if (final.dev === identity.dev && final.ino === identity.ino) await rmdir(path).catch((error: NodeJS.ErrnoException) => { if (error.code !== "ENOTEMPTY" && error.code !== "ENOENT") throw error; });
    } catch (error) { failures.push(error); }
  }
  if (failures.length) throw new AggregateError(failures, "workspace_staging_reconciliation_failed");
}

const STORE_KEY = Symbol.for("@baumus/openclaw-microsoft-graph/workspace-staging");
export const workspaceStagingStore: WorkspaceStagingStore = ((globalThis as Record<symbol, unknown>)[STORE_KEY] as WorkspaceStagingStore | undefined)
  ?? ((globalThis as Record<symbol, unknown>)[STORE_KEY] = new WorkspaceStagingStore()) as WorkspaceStagingStore;
const EXTENSION_CONTENT_TYPES: Readonly<Record<string, string>> = {
  pdf: "application/pdf", txt: "text/plain", md: "text/markdown", csv: "text/csv",
  json: "application/json", png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

export function workspaceFileContentType(relativePath: string): string {
  const extension = /\.([A-Za-z0-9]+)$/.exec(relativePath)?.[1]?.toLowerCase();
  return (extension && EXTENSION_CONTENT_TYPES[extension]) || "application/octet-stream";
}

export function validateWorkspaceRelativeFilePath(value: unknown): string {
  if (typeof value !== "string" || !value || value.length > 1024
    || value.startsWith("/") || /[\\\u0000-\u001f\u007f]/.test(value)
    || value.split("/").some((part) => !part || part === "." || part === "..")) {
    throw new Error("invalid_workspace_relative_path");
  }
  return value;
}

/** Copy a caller-owned workspace file into the host's private inbound store. No Graph access. */
export async function stageWorkspaceFile(
  workspaceDir: string | undefined,
  sourceRelativePath: unknown,
  contentType = "application/octet-stream",
  save: typeof saveMediaStream = saveMediaStream,
  signal?: AbortSignal,
  store = workspaceStagingStore,
  stateDir?: string,
) {
  const relativePath = validateWorkspaceRelativeFilePath(sourceRelativePath);
  if (typeof workspaceDir !== "string" || !workspaceDir) throw new Error("workspace_context_unavailable");
  if (typeof contentType !== "string" || contentType.length < 3 || contentType.length > 160
    || !/^[A-Za-z0-9!#$&^_.+\-]+\/[A-Za-z0-9!#$&^_.+\-]+$/.test(contentType)) throw new Error("invalid_content_type");
  let opened: OpenResult | undefined;
  let release: (() => void) | undefined;
  let lease: WorkspaceStagingLease | undefined;
  try {
    signal?.throwIfAborted();
    const workspace = await secureRoot(workspaceDir, { symlinks: "reject", hardlinks: "reject" });
    opened = await workspace.open(relativePath, { symlinks: "reject", hardlinks: "reject" });
    const initial = opened.stat;
    if (!initial.isFile() || !Number.isSafeInteger(initial.size) || initial.size < 0) throw new Error("invalid_workspace_file");
    if (initial.size > ONEDRIVE_WRITE_MAX_BYTES) throw new Error("provider_file_too_large");
    release = store.reserve(initial.size);
    const sameIdentity = (current: typeof initial) => current.isFile()
      && current.dev === initial.dev && current.ino === initial.ino
      && current.size === initial.size && current.mtimeMs === initial.mtimeMs
      && current.ctimeMs === initial.ctimeMs && current.nlink === initial.nlink;
    const hash = createHash("sha256");
    let count = 0;
    async function* chunks() {
      for (let offset = 0; offset < initial.size;) {
        signal?.throwIfAborted();
        const length = Math.min(CHUNK_BYTES, initial.size - offset);
        const buffer = Buffer.allocUnsafe(length);
        const { bytesRead } = await opened!.handle.read(buffer, 0, length, offset);
        if (bytesRead !== length) throw new Error("workspace_file_changed");
        hash.update(buffer);
        count += bytesRead;
        offset += bytesRead;
        yield buffer;
      }
      if (!sameIdentity(await opened!.handle.stat())) throw new Error("workspace_file_changed");
    }
    const name = sanitizeAttachmentName(basename(relativePath));
    signal?.throwIfAborted();
    if (stateDir) await store.ownRun(stateDir);
    const saved = await save(chunks(), contentType, `${STAGING_SUBDIR}/${store.runId}`, ONEDRIVE_WRITE_MAX_BYTES, name, name);
    if (typeof saved.path !== "string" || basename(saved.path) !== saved.id || basename(dirname(saved.path)) !== store.runId
      || basename(dirname(dirname(saved.path))) !== STAGING_NAMESPACE
      || (stateDir && resolve(dirname(saved.path)) !== join(resolve(stateDir), "media", STAGING_SUBDIR, store.runId))) throw new Error("workspace_file_unavailable");
    const identity = await lstat(saved.path);
    let cleaning: Promise<void> | undefined;
    lease = { cleanup: () => cleaning ??= (async () => {
      const staging = await secureRoot(dirname(saved.path), { symlinks: "reject", hardlinks: "reject" });
      const current = await staging.open(saved.id, { symlinks: "reject", hardlinks: "reject" }).catch(() => undefined);
      if (!current) {
        const remaining = await lstat(saved.path).catch((error: NodeJS.ErrnoException) => {
          if (error.code === "ENOENT") return undefined;
          throw error;
        });
        if (!remaining) release!();
        return;
      }
      try {
        if (current.stat.dev !== identity.dev || current.stat.ino !== identity.ino || current.stat.nlink !== 1) return;
      } finally { await current.handle.close(); }
      await staging.remove(saved.id, { mutationSymlinks: "reject" });
      release!();
    })().catch((error) => { cleaning = undefined; throw error; }) };
    if (!identity.isFile() || identity.nlink !== 1 || identity.size !== saved.size) throw new Error("workspace_file_unavailable");
    signal?.throwIfAborted();
    if (saved.size !== initial.size || count !== initial.size) throw new Error("workspace_file_changed");
    if (!sameIdentity(await opened.handle.stat())) throw new Error("workspace_file_changed");
    return { sourceMediaUri: `media://${STAGING_SUBDIR}/${store.runId}/${saved.id}`, sourceSha256: hash.digest("hex"), sourceByteSize: saved.size, lease };
  } catch (error) {
    if (lease) await lease.cleanup().catch(() => store.retryCleanup(lease!));
    else release?.();
    if (error instanceof DOMException && error.name === "AbortError") throw error;
    if (error instanceof Error && ["provider_file_too_large", "workspace_staging_quota_exceeded", "workspace_file_changed"].includes(error.message)) throw error;
    throw new Error("workspace_file_unavailable");
  } finally {
    await opened?.handle.close().catch(() => undefined);
  }
}
