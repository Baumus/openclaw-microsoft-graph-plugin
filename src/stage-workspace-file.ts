import { createHash } from "node:crypto";
import { basename } from "node:path";
import { root as secureRoot, type OpenResult } from "openclaw/plugin-sdk/infra-runtime";
import { saveMediaStream } from "openclaw/plugin-sdk/media-store";
import { sanitizeAttachmentName } from "./attachment-download.js";
import { ONEDRIVE_WRITE_MAX_BYTES } from "./graph.js";

const CHUNK_BYTES = 4 * 1024 * 1024;
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
) {
  const relativePath = validateWorkspaceRelativeFilePath(sourceRelativePath);
  if (typeof workspaceDir !== "string" || !workspaceDir) throw new Error("workspace_context_unavailable");
  if (typeof contentType !== "string" || contentType.length < 3 || contentType.length > 160
    || !/^[A-Za-z0-9!#$&^_.+\-]+\/[A-Za-z0-9!#$&^_.+\-]+$/.test(contentType)) throw new Error("invalid_content_type");
  let opened: OpenResult | undefined;
  try {
    signal?.throwIfAborted();
    const workspace = await secureRoot(workspaceDir, { symlinks: "reject", hardlinks: "reject" });
    opened = await workspace.open(relativePath, { symlinks: "reject", hardlinks: "reject" });
    const initial = opened.stat;
    if (!initial.isFile() || !Number.isSafeInteger(initial.size) || initial.size < 0) throw new Error("invalid_workspace_file");
    if (initial.size > ONEDRIVE_WRITE_MAX_BYTES) throw new Error("provider_file_too_large");
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
    signal?.throwIfAborted();
    if (saved.size !== initial.size || count !== initial.size) throw new Error("workspace_file_changed");
    if (!sameIdentity(await opened.handle.stat())) throw new Error("workspace_file_changed");
    return { sourceMediaUri: "media://inbound/" + saved.id, sourceSha256: hash.digest("hex"), sourceByteSize: saved.size };
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") throw error;
    if (error instanceof Error && ["provider_file_too_large", "workspace_file_changed"].includes(error.message)) throw error;
    throw new Error("workspace_file_unavailable");
  } finally {
    await opened?.handle.close().catch(() => undefined);
  }
}
