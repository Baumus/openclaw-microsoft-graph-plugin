import { createHash } from "node:crypto";
import { saveMediaStream } from "openclaw/plugin-sdk/media-store";
import { contentTypeAllowed, drivePath, graphRequest, graphStreamRequest, ONEDRIVE_READ_MAX_BYTES, OUTLOOK_ATTACHMENT_MAX_BYTES, safeDriveItem, safeId } from "./graph.js";
import type { AllowedRoot } from "./policy.js";

const MAX_SAFE_FILENAME_BYTES = 180;

export type OutlookAttachmentOwner =
  | { kind: "message"; messageId: string }
  | { kind: "event"; eventId: string; calendarId?: string };

export type DownloadOutlookAttachmentOptions = {
  token: string;
  owner: OutlookAttachmentOwner;
  attachmentId: string;
  maxBytes: number;
  readIdleTimeoutMs: number;
  signal?: AbortSignal;
  fetchFn?: typeof fetch;
};

function fail(message: "invalid_provider_response" | "unsupported_attachment_type" | "file_too_large" | "invalid_attachment_name"): never {
  throw new Error(message);
}

function decodeTraversalCandidate(value: string): string[] {
  const candidates = [value];
  let current = value;
  for (let depth = 0; depth < 4 && current.includes("%"); depth += 1) {
    try {
      const decoded = decodeURIComponent(current);
      if (decoded === current) break;
      candidates.push(decoded);
      current = decoded;
    } catch {
      break;
    }
  }
  return candidates;
}

/** Reject path-bearing provider names and normalize the remaining leaf name. */
export function sanitizeAttachmentName(value: unknown): string {
  if (typeof value !== "string" || !value || value.length > 512) fail("invalid_attachment_name");
  const normalized = value.normalize("NFKC");
  for (const candidate of decodeTraversalCandidate(normalized)) {
    if (candidate === "." || candidate === ".." || candidate.includes("/") || candidate.includes("\\")) fail("invalid_attachment_name");
  }
  let safe = normalized
    .replace(/[\u0000-\u001f\u007f]/g, "_")
    .replace(/[<>:"|?*%]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[ .]+$/g, "");
  if (!safe || safe === "." || safe === "..") safe = "attachment.bin";
  if (Buffer.byteLength(safe, "utf8") > MAX_SAFE_FILENAME_BYTES) {
    const extensionMatch = /\.[A-Za-z0-9]{1,16}$/.exec(safe);
    const extension = extensionMatch?.[0] ?? "";
    const stem = extension ? safe.slice(0, -extension.length) : safe;
    let shortened = "";
    for (const character of stem) {
      if (Buffer.byteLength(shortened + character + extension, "utf8") > MAX_SAFE_FILENAME_BYTES) break;
      shortened += character;
    }
    safe = `${shortened || "attachment"}${extension}`;
  }
  return safe;
}

export function sanitizeAttachmentMime(value: unknown): string {
  const mime = value ?? "application/octet-stream";
  if (typeof mime !== "string" || mime.length < 3 || mime.length > 160 || /[\r\n\u0000-\u001f\u007f]/.test(mime)) fail("invalid_provider_response");
  const [mediaType, ...parameters] = mime.split(";").map((part) => part.trim());
  const token = "[A-Za-z0-9!#$&^_.+\\-]+";
  if (!new RegExp(`^${token}\/${token}$`).test(mediaType)) fail("invalid_provider_response");
  for (const parameter of parameters) if (!new RegExp(`^${token}=(?:${token}|\"[^\"\\r\\n]*\")$`).test(parameter)) fail("invalid_provider_response");
  return [mediaType.toLowerCase(), ...parameters].join("; ");
}

export async function publishPrivateMediaBytes(bytes: Uint8Array, nameValue: unknown, mimeValue: unknown, maxBytes: number) {
  if (!Number.isSafeInteger(maxBytes) || maxBytes < 1 || bytes.byteLength > maxBytes) throw new Error("file_too_large");
  const name = sanitizeAttachmentName(nameValue);
  const contentType = sanitizeAttachmentMime(mimeValue);
  async function* source() { yield bytes; }
  const saved = await saveMediaStream(source(), contentType, "inbound", maxBytes, name, name);
  if (!Number.isSafeInteger(saved.size) || saved.size !== bytes.byteLength) fail("invalid_provider_response");
  const artifact = {
    type: "file" as const,
    id: saved.id,
    uri: `media://inbound/${saved.id}`,
    mimeType: sanitizeAttachmentMime(saved.contentType ?? contentType),
    name,
    sizeBytes: saved.size,
  };
  return { artifact, media: { outbound: false, attachments: [artifact] } };
}

/** Stream one allowlisted OneDrive file directly into OpenClaw private media. */
export async function downloadOneDriveFile(options: { root: AllowedRoot; relativePath: string; token: string; signal?: AbortSignal; fetchFn?: typeof fetch }) {
  const fetchFn = options.fetchFn ?? fetch;
  const metadata = await graphRequest(options.token, `${drivePath(options.root, options.relativePath)}?$select=id,name,size,file,createdDateTime,lastModifiedDateTime,webUrl`, { signal: options.signal }, fetchFn);
  if (metadata?.folder || !metadata?.file || !Number.isSafeInteger(metadata.size) || metadata.size < 0) fail("invalid_provider_response");
  if (metadata.size > ONEDRIVE_READ_MAX_BYTES) throw new Error("provider_file_too_large");
  const name = sanitizeAttachmentName(metadata.name);
  const contentType = sanitizeAttachmentMime(metadata.file.mimeType ?? "application/octet-stream");
  if (!contentTypeAllowed(contentType)) fail("unsupported_attachment_type");
  const response = await graphStreamRequest(options.token, drivePath(options.root, options.relativePath, "/content"), { signal: options.signal }, fetchFn);
  const hash = createHash("sha256");
  let streamedBytes = 0;
  async function* source() {
    if (!response.body && metadata.size !== 0) fail("invalid_provider_response");
    const reader = response.body?.getReader();
    try {
      while (reader) {
        const { done, value } = await reader.read();
        if (done) break;
        streamedBytes += value.byteLength;
        if (streamedBytes > ONEDRIVE_READ_MAX_BYTES) {
          await reader.cancel().catch(() => undefined);
          throw new Error("provider_file_too_large");
        }
        hash.update(value);
        yield value;
      }
    } finally {
      reader?.releaseLock();
    }
  }
  const saved = await saveMediaStream(source(), contentType, "inbound", ONEDRIVE_READ_MAX_BYTES, name, name);
  if (!Number.isSafeInteger(saved.size) || saved.size !== streamedBytes || saved.size !== metadata.size) fail("invalid_provider_response");
  const artifact = {
    type: "file" as const,
    id: saved.id,
    uri: `media://inbound/${saved.id}`,
    mimeType: sanitizeAttachmentMime(saved.contentType ?? contentType),
    name,
    sizeBytes: saved.size,
  };
  return {
    item: safeDriveItem(options.root, metadata),
    bytes: saved.size,
    sha256: hash.digest("hex"),
    artifact,
    media: { outbound: false, attachments: [artifact] },
  };
}

function attachmentBasePath(owner: OutlookAttachmentOwner, attachmentId: string): string {
  const attachment = safeId(attachmentId);
  if (owner.kind === "message") return `/me/messages/${safeId(owner.messageId)}/attachments/${attachment}`;
  const event = safeId(owner.eventId);
  return owner.calendarId === undefined
    ? `/me/events/${event}/attachments/${attachment}`
    : `/me/calendars/${safeId(owner.calendarId)}/events/${event}/attachments/${attachment}`;
}

async function readWithIdleTimeout(reader: ReadableStreamDefaultReader<Uint8Array>, timeoutMs: number) {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      reader.read(),
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => {
          reject(new DOMException("Attachment response body stalled", "TimeoutError"));
          void reader.cancel().catch(() => undefined);
        }, timeoutMs);
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}

async function* boundedResponseBody(response: Response, reportedBytes: number, maximumBytes: number, readIdleTimeoutMs: number): AsyncGenerator<Uint8Array> {
  const rawLength = response.headers.get("content-length");
  let contentLength: number | undefined;
  if (rawLength !== null) {
    if (!/^\d+$/.test(rawLength)) {
      await response.body?.cancel().catch(() => undefined);
      fail("invalid_provider_response");
    }
    const length = Number(rawLength);
    if (!Number.isSafeInteger(length)) {
      await response.body?.cancel().catch(() => undefined);
      fail("invalid_provider_response");
    }
    if (length > maximumBytes) {
      await response.body?.cancel().catch(() => undefined);
      fail("file_too_large");
    }
    contentLength = length;
  }
  if (!response.body && (contentLength ?? reportedBytes) !== 0) fail("invalid_provider_response");
  const reader = response.body?.getReader();
  let bytes = 0;
  try {
    while (reader) {
      const { done, value } = await readWithIdleTimeout(reader, readIdleTimeoutMs);
      if (done) break;
      bytes += value.byteLength;
      if (bytes > maximumBytes) {
        await reader.cancel().catch(() => undefined);
        fail("file_too_large");
      }
      yield value;
    }
    if (contentLength !== undefined && bytes !== contentLength) fail("invalid_provider_response");
    if (bytes === 0 && reportedBytes !== 0) fail("invalid_provider_response");
  } finally {
    reader?.releaseLock();
  }
}

/** Download one exact Graph file attachment into OpenClaw's private media store. */
export async function downloadOutlookFileAttachment(options: DownloadOutlookAttachmentOptions) {
  if (!Number.isSafeInteger(options.maxBytes) || options.maxBytes < 1 || options.maxBytes > OUTLOOK_ATTACHMENT_MAX_BYTES) throw new Error("invalid_byte_limit");
  if (!Number.isSafeInteger(options.readIdleTimeoutMs) || options.readIdleTimeoutMs < 1 || options.readIdleTimeoutMs > 30_000) throw new Error("invalid_request_timeout");
  const fetchFn = options.fetchFn ?? fetch;
  const base = attachmentBasePath(options.owner, options.attachmentId);
  const item = await graphRequest(options.token, `${base}?$select=id,name,contentType,size,isInline`, { signal: options.signal }, fetchFn);
  if (item?.["@odata.type"] !== "#microsoft.graph.fileAttachment") fail("unsupported_attachment_type");
  if (item.id !== options.attachmentId || !Number.isSafeInteger(item.size) || item.size < 0) fail("invalid_provider_response");
  const name = sanitizeAttachmentName(item.name);
  const contentType = sanitizeAttachmentMime(item.contentType);
  const response = await graphStreamRequest(options.token, `${base}/$value`, { signal: options.signal }, fetchFn);
  const hash = createHash("sha256");
  async function* hashingSource() {
    for await (const chunk of boundedResponseBody(response, item.size, options.maxBytes, options.readIdleTimeoutMs)) {
      hash.update(chunk);
      yield chunk;
    }
  }
  const saved = await saveMediaStream(hashingSource(), contentType, "inbound", options.maxBytes, name, name);
  if (!Number.isSafeInteger(saved.size) || saved.size < 0 || saved.size > options.maxBytes) fail("invalid_provider_response");
  const storedContentType = sanitizeAttachmentMime(saved.contentType ?? contentType);
  const sourceSha256 = hash.digest("hex");
  const artifact = {
    type: "file" as const,
    id: saved.id,
    uri: `media://inbound/${saved.id}`,
    mimeType: storedContentType,
    name,
    sizeBytes: saved.size,
  };
  return {
    attachment: {
      id: item.id,
      name,
      contentType: storedContentType,
      size: item.size,
      bytes: saved.size,
      ...(typeof item.isInline === "boolean" ? { isInline: item.isInline } : {}),
    },
    sourceSha256,
    sourceByteSize: saved.size,
    artifact,
    media: {
      outbound: false,
      attachments: [artifact],
    },
  };
}
