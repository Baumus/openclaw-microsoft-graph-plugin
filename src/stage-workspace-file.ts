import { createHash } from "node:crypto";
import { lstat } from "node:fs/promises";
import { basename, dirname } from "node:path";
import { root as secureRoot, type OpenResult } from "openclaw/plugin-sdk/infra-runtime";
import { saveMediaStream } from "openclaw/plugin-sdk/media-store";
import { sanitizeAttachmentName } from "./attachment-download.js";
import { ONEDRIVE_WRITE_MAX_BYTES } from "./graph.js";

const CHUNK_BYTES = 4 * 1024 * 1024;
export const WORKSPACE_STAGING_QUOTA_BYTES = 128 * 1024 * 1024;
const WORKSPACE_STAGING_FILE_MAX_BYTES = 64 * 1024 * 1024;
const MAX_WORKSPACE_STAGING_RESERVATIONS = 64;

export type WorkspaceStagingLease = { cleanup(): Promise<void> };

export class WorkspaceStagingStore {
  private reservedBytes = 0;
  private reservations = 0;
  private readonly pending = new Map<string, { lease: WorkspaceStagingLease; timer: NodeJS.Timeout; sessionId?: string; toolName: string; onExpire: () => void; executing: boolean }>();

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
    await entry.lease.cleanup();
  }

  clearSession(sessionId: string): void {
    for (const [toolCallId, entry] of this.pending) if (entry.sessionId === sessionId && !entry.executing) {
      entry.onExpire();
      void this.cleanup(toolCallId).catch(() => undefined);
    }
  }
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
    const saved = await save(chunks(), contentType, "inbound", ONEDRIVE_WRITE_MAX_BYTES, name, name);
    if (typeof saved.path !== "string" || basename(saved.path) !== saved.id) throw new Error("workspace_file_unavailable");
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
    })() };
    if (!identity.isFile() || identity.nlink !== 1 || identity.size !== saved.size) throw new Error("workspace_file_unavailable");
    signal?.throwIfAborted();
    if (saved.size !== initial.size || count !== initial.size) throw new Error("workspace_file_changed");
    if (!sameIdentity(await opened.handle.stat())) throw new Error("workspace_file_changed");
    return { sourceMediaUri: "media://inbound/" + saved.id, sourceSha256: hash.digest("hex"), sourceByteSize: saved.size, lease };
  } catch (error) {
    if (lease) await lease.cleanup().catch(() => undefined);
    else release?.();
    if (error instanceof DOMException && error.name === "AbortError") throw error;
    if (error instanceof Error && ["provider_file_too_large", "workspace_staging_quota_exceeded", "workspace_file_changed"].includes(error.message)) throw error;
    throw new Error("workspace_file_unavailable");
  } finally {
    await opened?.handle.close().catch(() => undefined);
  }
}
