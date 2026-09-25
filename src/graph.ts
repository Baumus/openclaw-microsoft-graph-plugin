import { createHash } from "node:crypto";
import type { AllowedRoot } from "./policy.js";

const GRAPH = "https://graph.microsoft.com/v1.0";
const MAX_JSON_BYTES = 512 * 1024;
const DEFAULT_GRAPH_REQUEST_TIMEOUT_MS = 30_000;
const DEFAULT_GRAPH_READ_RETRIES = 2;
const DEFAULT_GRAPH_RETRY_BASE_DELAY_MS = 100;
const DEFAULT_GRAPH_RETRY_MAX_DELAY_MS = 2_000;
const TEXT_MIME = /^(text\/|application\/(json|xml|csv|javascript)(?:$|;))/i;
const BINARY_MIME = new Set([
  "application/pdf",
  "application/octet-stream",
  "application/zip",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "image/jpeg",
  "image/png",
  "video/mp4",
]);

export const ONEDRIVE_READ_MAX_BYTES = 250 * 1024 * 1024 * 1024;
export const ONEDRIVE_SIMPLE_UPLOAD_MAX_BYTES = 250 * 1024 * 1024;
export const ONEDRIVE_WRITE_MAX_BYTES = ONEDRIVE_READ_MAX_BYTES;
export const OUTLOOK_ATTACHMENT_MAX_BYTES = 150 * 1024 * 1024;
export const TODO_ATTACHMENT_MAX_BYTES = 25 * 1024 * 1024;
export const DIRECT_ATTACHMENT_MAX_BYTES = 3 * 1024 * 1024;
export const ATTACHMENT_UPLOAD_CHUNK_BYTES = 10 * 320 * 1024;
export const ONEDRIVE_UPLOAD_CHUNK_BYTES = 16 * 320 * 1024;

type GraphRequestPolicy = {
  requestTimeoutMs: number;
  readRetries: number;
  retryBaseDelayMs: number;
  retryMaxDelayMs: number;
  externalSignal?: AbortSignal;
  deadlineSignal: AbortSignal;
};

const graphRequestPolicies = new WeakMap<AbortSignal, GraphRequestPolicy>();

/** Create one whole-operation signal while retaining an independent timeout for each Graph request. */
export function graphOperationSignal(
  externalSignal: AbortSignal | undefined,
  operationTimeoutMs: number,
  requestTimeoutMs: number,
): AbortSignal {
  if (!Number.isInteger(operationTimeoutMs) || operationTimeoutMs < 1 || operationTimeoutMs > 7 * 24 * 60 * 60 * 1000
    || !Number.isInteger(requestTimeoutMs) || requestTimeoutMs < 1 || requestTimeoutMs > 30_000) throw new Error("invalid_request_timeout");
  const deadlineSignal = AbortSignal.timeout(operationTimeoutMs);
  const signal = externalSignal ? AbortSignal.any([externalSignal, deadlineSignal]) : deadlineSignal;
  graphRequestPolicies.set(signal, {
    requestTimeoutMs,
    readRetries: DEFAULT_GRAPH_READ_RETRIES,
    retryBaseDelayMs: DEFAULT_GRAPH_RETRY_BASE_DELAY_MS,
    retryMaxDelayMs: DEFAULT_GRAPH_RETRY_MAX_DELAY_MS,
    externalSignal,
    deadlineSignal,
  });
  return signal;
}

/** True only for the plugin-owned whole-operation deadline, never caller cancellation. */
export function graphOperationDeadlineReached(signal: AbortSignal | undefined): boolean {
  const policy = signal ? graphRequestPolicies.get(signal) : undefined;
  return Boolean(policy?.deadlineSignal.aborted && !policy.externalSignal?.aborted);
}

function preferredAbortReason(signal: AbortSignal | undefined): unknown {
  const policy = signal ? graphRequestPolicies.get(signal) : undefined;
  if (policy?.externalSignal?.aborted) return policy.externalSignal.reason;
  if (signal?.aborted) return signal.reason;
  return undefined;
}

function throwPreferredAbort(signal: AbortSignal | undefined): void {
  const reason = preferredAbortReason(signal);
  if (reason !== undefined) throw reason;
}

function invalidContinuation(message: "invalid_continuation" | "invalid_provider_response"): never {
  throw new Error(message);
}

function rejectEncodedTraversal(rawPathname: string, message: "invalid_continuation" | "invalid_provider_response"): void {
  let decoded = rawPathname;
  for (let depth = 0; depth < 6; depth += 1) {
    if (decoded.includes("\\") || /[\u0000-\u001f\u007f]/.test(decoded) || decoded.split("/").some((segment) => segment === "." || segment === "..")) invalidContinuation(message);
    if (!decoded.includes("%")) return;
    let next: string;
    try { next = decodeURIComponent(decoded); } catch { invalidContinuation(message); }
    if (next === decoded) return;
    decoded = next;
  }
  if (decoded.includes("%")) invalidContinuation(message);
}

/** Canonicalize a provider-owned Graph nextLink and bind it to one exact collection pathname. */
export function canonicalGraphContinuation(value: unknown, expectedPathname: string, message: "invalid_continuation" | "invalid_provider_response" = "invalid_continuation"): string {
  if (typeof value !== "string" || value.length < 1 || value.length > 8192 || !expectedPathname.startsWith("/") || value.includes("#") || value.includes("\\") || /[\u0000-\u001f\u007f]/.test(value)) invalidContinuation(message);
  let url: URL;
  if (value.startsWith("/")) {
    if (value.startsWith("//")) invalidContinuation(message);
    try { url = new URL(value, "https://graph.microsoft.com"); } catch { invalidContinuation(message); }
  } else {
    if (!/^https:\/\/graph\.microsoft\.com\//i.test(value)) invalidContinuation(message);
    try { url = new URL(value); } catch { invalidContinuation(message); }
    if (url.protocol !== "https:" || url.hostname.toLowerCase() !== "graph.microsoft.com" || url.port || url.username || url.password) invalidContinuation(message);
  }
  const rawPathname = value.startsWith("/") ? value.split("?", 1)[0] : value.slice(value.indexOf("/", "https://".length)).split("?", 1)[0];
  rejectEncodedTraversal(rawPathname, message);
  let decodedExpected: string;
  let decodedRaw: string;
  let decodedUrl: string;
  try {
    decodedExpected = decodeURIComponent(expectedPathname);
    decodedRaw = decodeURIComponent(rawPathname);
    decodedUrl = decodeURIComponent(url.pathname);
  } catch { invalidContinuation(message); }
  if (decodedRaw !== decodedExpected || decodedUrl !== decodedExpected) invalidContinuation(message);
  return `${expectedPathname}${url.search}`;
}

/** Validate canonical RFC 4648 base64 and derive decoded length without decoding. */
export function base64DecodedByteLengthStrict(value: unknown, message = "invalid_write_input"): number {
  if (typeof value !== "string" || value.length === 0 || value.length % 4 !== 0) throw new Error(message);
  const padding = value.endsWith("==") ? 2 : value.endsWith("=") ? 1 : 0;
  const dataLength = value.length - padding;
  const sextet = (code: number): number => {
    if (code >= 65 && code <= 90) return code - 65;
    if (code >= 97 && code <= 122) return code - 71;
    if (code >= 48 && code <= 57) return code + 4;
    if (code === 43) return 62;
    if (code === 47) return 63;
    return -1;
  };
  for (let index = 0; index < dataLength; index += 1) if (sextet(value.charCodeAt(index)) < 0) throw new Error(message);
  for (let index = dataLength; index < value.length; index += 1) if (value.charCodeAt(index) !== 61) throw new Error(message);
  if ((padding === 1 && (sextet(value.charCodeAt(dataLength - 1)) & 0b11) !== 0)
    || (padding === 2 && (sextet(value.charCodeAt(dataLength - 1)) & 0b1111) !== 0)) throw new Error(message);
  return value.length / 4 * 3 - padding;
}

/** Decode only RFC 4648 canonical base64 so malformed writes cannot decode permissively. */
export function decodeBase64Strict(value: unknown, message = "invalid_write_input"): Buffer {
  base64DecodedByteLengthStrict(value, message);
  return Buffer.from(value as string, "base64");
}

export function validateDriveWriteInput(relativePath: string, contentBase64: string, contentType: string, maxBytes?: number): void {
  if (!relativePath || !contentTypeAllowed(contentType)) throw new Error("invalid_write_input");
  const decodedBytes = base64DecodedByteLengthStrict(contentBase64);
  if (maxBytes !== undefined && decodedBytes > maxBytes) throw new Error("file_too_large");
}

export function validateDriveWriteBytes(relativePath: string, content: Uint8Array, contentType: string, maxBytes?: number): void {
  if (!relativePath || !(content instanceof Uint8Array) || !contentTypeAllowed(contentType)) throw new Error("invalid_write_input");
  if (maxBytes !== undefined && content.byteLength > maxBytes) throw new Error("file_too_large");
}

export function validateDriveMetadataInput(relativePath: string, changes: { name?: string; destinationRelativePath?: string; description?: string | null; fileSystemInfo?: { createdDateTime?: string; lastModifiedDateTime?: string } }): void {
  if (!relativePath || !changes || Object.keys(changes).length === 0) throw new Error("invalid_drive_metadata");
  if (changes.name !== undefined && (!changes.name || changes.name.length > 255 || /[\\/:*?"<>|]/.test(changes.name))) throw new Error("invalid_drive_name");
  for (const value of [changes.fileSystemInfo?.createdDateTime, changes.fileSystemInfo?.lastModifiedDateTime]) {
    if (value !== undefined && !isStrictDateTimeOffset(value)) throw new Error("invalid_datetime");
  }
}

function isStrictDateTimeOffset(value: unknown): value is string {
  if (typeof value !== "string") return false;
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d{1,7}))?(Z|([+-])(\d{2}):(\d{2}))$/.exec(value);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const daysInMonth = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (year < 1 || month < 1 || month > 12 || day < 1 || day > daysInMonth[month - 1] || hour > 23 || minute > 59 || second > 59) return false;
  if (match[8] !== "Z") {
    const offsetHour = Number(match[10]);
    const offsetMinute = Number(match[11]);
    if (offsetHour > 14 || offsetMinute > 59 || (offsetHour === 14 && offsetMinute !== 0)) return false;
  }
  return true;
}

export function validateDriveFolderInput(name: string): void {
  if (!name || name.length > 255 || /[\\/:*?"<>|]/.test(name)) throw new Error("invalid_drive_name");
}

export function base64JsonResponseLimit(rawBytes: number): number {
  if (!Number.isSafeInteger(rawBytes) || rawBytes < 0) throw new Error("invalid_byte_limit");
  return 4 * Math.ceil(rawBytes / 3) + 64 * 1024;
}

export type GraphOptions = {
  method?: "GET" | "POST" | "PATCH" | "DELETE";
  body?: unknown;
  headers?: Record<string, string>;
  signal?: AbortSignal;
  maxBytes?: number;
  response?: "json" | "bytes" | "none";
  requestTimeoutMs?: number;
  readRetries?: number;
  retryBaseDelayMs?: number;
  retryMaxDelayMs?: number;
  retryNow?: () => number;
  retrySleep?: (delayMs: number, signal?: AbortSignal) => Promise<void>;
};

export function safeId(value: string): string {
  if (typeof value !== "string" || !value || value.length > 512 || !/^[A-Za-z0-9._~!$&'()*+,;=:@%-]+$/.test(value)) throw new Error("invalid_resource_id");
  return encodeURIComponent(value);
}

async function readBytes(response: Response, maximum: number): Promise<Uint8Array> {
  const length = Number(response.headers.get("content-length"));
  if (Number.isFinite(length) && length > maximum) throw new Error("provider_response_too_large");
  if (!response.body) return new Uint8Array();
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > maximum) { await reader.cancel(); throw new Error("provider_response_too_large"); }
      chunks.push(value);
    }
  } finally { reader.releaseLock(); }
  return new Uint8Array(Buffer.concat(chunks.map((chunk) => Buffer.from(chunk)), total));
}

async function readDigest(response: Response, maximum: number): Promise<{ bytes: number; sha256: string }> {
  const length = Number(response.headers.get("content-length"));
  if (Number.isFinite(length) && length > maximum) throw new Error("provider_response_too_large");
  const hash = createHash("sha256");
  if (!response.body) return { bytes: 0, sha256: hash.digest("hex") };
  const reader = response.body.getReader();
  let total = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > maximum) { await reader.cancel(); throw new Error("provider_response_too_large"); }
      hash.update(value);
    }
  } finally { reader.releaseLock(); }
  return { bytes: total, sha256: hash.digest("hex") };
}

function providerStatusError(status: number): Error {
  if (status === 401 || status === 403) return Object.assign(new Error("provider_access_denied"), { providerStatus: status });
  if (status === 404) return new Error("item_not_found");
  if (status === 409 || status === 412) return new Error("item_conflict");
  if (status === 429) return new Error("provider_throttled");
  return new Error(`provider_error_${status}`);
}

type GraphFetchOptions = Pick<GraphOptions, "requestTimeoutMs" | "readRetries" | "retryBaseDelayMs" | "retryMaxDelayMs" | "retryNow" | "retrySleep">;

function graphFetchPolicy(signal: AbortSignal | undefined, options: GraphFetchOptions) {
  const inherited = signal ? graphRequestPolicies.get(signal) : undefined;
  const requestTimeoutMs = options.requestTimeoutMs ?? inherited?.requestTimeoutMs ?? DEFAULT_GRAPH_REQUEST_TIMEOUT_MS;
  const readRetries = options.readRetries ?? inherited?.readRetries ?? DEFAULT_GRAPH_READ_RETRIES;
  const retryBaseDelayMs = options.retryBaseDelayMs ?? inherited?.retryBaseDelayMs ?? DEFAULT_GRAPH_RETRY_BASE_DELAY_MS;
  const retryMaxDelayMs = options.retryMaxDelayMs ?? inherited?.retryMaxDelayMs ?? DEFAULT_GRAPH_RETRY_MAX_DELAY_MS;
  if (!Number.isInteger(requestTimeoutMs) || requestTimeoutMs < 1 || requestTimeoutMs > 30_000
    || !Number.isInteger(readRetries) || readRetries < 0 || readRetries > 4
    || !Number.isInteger(retryBaseDelayMs) || retryBaseDelayMs < 0 || retryBaseDelayMs > 5_000
    || !Number.isInteger(retryMaxDelayMs) || retryMaxDelayMs < 0 || retryMaxDelayMs > 30_000
    || retryBaseDelayMs > retryMaxDelayMs) throw new Error("invalid_request_timeout");
  return { requestTimeoutMs, readRetries, retryBaseDelayMs, retryMaxDelayMs, now: options.retryNow ?? Date.now, sleep: options.retrySleep ?? abortableDelay };
}

async function abortableDelay(delayMs: number, signal?: AbortSignal): Promise<void> {
  if (delayMs <= 0) { throwPreferredAbort(signal); return; }
  throwPreferredAbort(signal);
  let abort: (() => void) | undefined;
  try {
    await new Promise<void>((resolve, reject) => {
      const timer = setTimeout(resolve, delayMs);
      abort = () => { clearTimeout(timer); reject(preferredAbortReason(signal)); };
      signal?.addEventListener("abort", abort, { once: true });
    });
  } finally {
    if (abort) signal?.removeEventListener("abort", abort);
  }
}

function retryAfterDelay(response: Response, now: () => number): number | undefined {
  const raw = response.headers.get("retry-after")?.trim();
  if (!raw) return undefined;
  if (/^\d+$/.test(raw)) return Number(raw) * 1_000;
  const date = Date.parse(raw);
  return Number.isFinite(date) ? Math.max(0, date - now()) : undefined;
}

function transientTransportError(error: unknown, requestTimeout: AbortSignal): boolean {
  return requestTimeout.aborted
    || error instanceof TypeError
    || (error instanceof DOMException && error.name === "TimeoutError");
}

async function fetchGraphValue<T>(
  url: string,
  init: RequestInit,
  operationSignal: AbortSignal | undefined,
  fetchFn: typeof fetch,
  options: GraphFetchOptions = {},
  consume: (response: Response) => Promise<T>,
): Promise<T> {
  const method = String(init.method ?? "GET").toUpperCase();
  const retryable = method === "GET";
  const policy = graphFetchPolicy(operationSignal, options);
  let retries = 0;
  while (true) {
    throwPreferredAbort(operationSignal);
    const requestTimeout = AbortSignal.timeout(policy.requestTimeoutMs);
    const requestSignal = operationSignal ? AbortSignal.any([operationSignal, requestTimeout]) : requestTimeout;
    let response: Response;
    try {
      response = await fetchFn(url, { ...init, signal: requestSignal });
    } catch (error) {
      throwPreferredAbort(operationSignal);
      if (!retryable || retries >= policy.readRetries || !transientTransportError(error, requestTimeout)) {
        if (requestTimeout.aborted) throw requestTimeout.reason;
        throw new Error("provider_unavailable");
      }
      const delay = Math.min(policy.retryBaseDelayMs * 2 ** retries, policy.retryMaxDelayMs);
      retries += 1;
      await policy.sleep(delay, operationSignal);
      continue;
    }
    if (retryable && new Set([429, 502, 503, 504]).has(response.status) && retries < policy.readRetries) {
      const requestedDelay = retryAfterDelay(response, policy.now);
      if (requestedDelay === undefined || requestedDelay <= policy.retryMaxDelayMs) {
        const delay = requestedDelay ?? Math.min(policy.retryBaseDelayMs * 2 ** retries, policy.retryMaxDelayMs);
        try { await response.body?.cancel(); } catch { /* discard transient response bytes */ }
        retries += 1;
        await policy.sleep(delay, operationSignal);
        continue;
      }
    }
    try {
      return await consume(response);
    } catch (error) {
      throwPreferredAbort(operationSignal);
      if (!retryable || retries >= policy.readRetries || !transientTransportError(error, requestTimeout)) {
        if (requestTimeout.aborted) throw requestTimeout.reason;
        throw error;
      }
      try { await response.body?.cancel(); } catch { /* discard timed-out response bytes */ }
      const delay = Math.min(policy.retryBaseDelayMs * 2 ** retries, policy.retryMaxDelayMs);
      retries += 1;
      await policy.sleep(delay, operationSignal);
    }
  }
}

async function fetchGraphResponse(
  url: string,
  init: RequestInit,
  operationSignal: AbortSignal | undefined,
  fetchFn: typeof fetch,
  options: GraphFetchOptions = {},
): Promise<Response> {
  const method = String(init.method ?? "GET").toUpperCase();
  const retryable = method === "GET";
  const policy = graphFetchPolicy(operationSignal, options);
  let retries = 0;
  while (true) {
    throwPreferredAbort(operationSignal);
    const headerTimeout = new AbortController();
    const requestSignal = operationSignal
      ? AbortSignal.any([operationSignal, headerTimeout.signal])
      : headerTimeout.signal;
    const timer = setTimeout(
      () => headerTimeout.abort(new DOMException("Graph response headers timed out", "TimeoutError")),
      policy.requestTimeoutMs,
    );
    let response: Response;
    try {
      response = await fetchFn(url, { ...init, signal: requestSignal });
    } catch (error) {
      throwPreferredAbort(operationSignal);
      if (!retryable || retries >= policy.readRetries || !transientTransportError(error, headerTimeout.signal)) {
        if (headerTimeout.signal.aborted) throw headerTimeout.signal.reason;
        throw new Error("provider_unavailable");
      }
      const delay = Math.min(policy.retryBaseDelayMs * 2 ** retries, policy.retryMaxDelayMs);
      retries += 1;
      await policy.sleep(delay, operationSignal);
      continue;
    } finally {
      clearTimeout(timer);
    }
    if (retryable && new Set([429, 502, 503, 504]).has(response.status) && retries < policy.readRetries) {
      const requestedDelay = retryAfterDelay(response, policy.now);
      if (requestedDelay === undefined || requestedDelay <= policy.retryMaxDelayMs) {
        const delay = requestedDelay ?? Math.min(policy.retryBaseDelayMs * 2 ** retries, policy.retryMaxDelayMs);
        try { await response.body?.cancel(); } catch { /* discard transient response bytes */ }
        retries += 1;
        await policy.sleep(delay, operationSignal);
        continue;
      }
    }
    return response;
  }
}

export async function graphRequest(token: string, path: string, options: GraphOptions = {}, fetchFn: typeof fetch = fetch): Promise<any> {
  if (!path.startsWith("/") || path.startsWith("//") || path.includes("\\") || /[\u0000-\u001f\u007f]/.test(path)) throw new Error("invalid_graph_path");
  const url = `${GRAPH}${path}`;
  const headers: Record<string, string> = { accept: "application/json", authorization: "Bearer " + token, ...options.headers };
  const body = options.body === undefined ? undefined : JSON.stringify(options.body);
  if (body !== undefined) headers["content-type"] = "application/json";
  return fetchGraphValue(url, { method: options.method ?? "GET", headers, body }, options.signal, fetchFn, options, async (response) => {
    if (!response.ok) throw providerStatusError(response.status);
    if (options.response === "none" || response.status === 204) return null;
    const bytes = await readBytes(response, options.maxBytes ?? MAX_JSON_BYTES);
    if (options.response === "bytes") return bytes;
    try { return bytes.byteLength ? JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)) : {}; }
    catch { throw new Error("invalid_provider_response"); }
  });
}

/** Open one Graph response as a stream without materializing its bytes. */
export async function graphStreamRequest(token: string, path: string, options: Pick<GraphOptions, "signal"> = {}, fetchFn: typeof fetch = fetch): Promise<Response> {
  if (!path.startsWith("/") || path.startsWith("//") || path.includes("\\") || /[\u0000-\u001f\u007f]/.test(path)) throw new Error("invalid_graph_path");
  const response = await fetchGraphResponse(`${GRAPH}${path}`, {
      method: "GET",
      headers: { accept: "application/octet-stream", authorization: "Bearer " + token },
      redirect: "follow",
    }, options.signal, fetchFn);
  if (!response.ok) {
    try { await response.body?.cancel(); } catch { /* discard provider error bytes */ }
    throw providerStatusError(response.status);
  }
  return response;
}

async function graphDigestRequest(token: string, path: string, maxBytes: number, signal?: AbortSignal, fetchFn: typeof fetch = fetch): Promise<{ bytes: number; sha256: string }> {
  if (!path.startsWith("/") || path.startsWith("//") || path.includes("\\") || /[\u0000-\u001f\u007f]/.test(path)) throw new Error("invalid_graph_path");
  return fetchGraphValue(`${GRAPH}${path}`, {
      method: "GET",
      headers: { accept: "application/octet-stream", authorization: "Bearer " + token },
    }, signal, fetchFn, {}, async (response) => {
      if (!response.ok) throw providerStatusError(response.status);
      return readDigest(response, maxBytes);
    });
}

export type AttachmentUploadSessionKind = "outlook" | "todo";

/** Validate a provider-issued upload URL before any request is sent to it. */
export function canonicalAttachmentUploadUrl(value: unknown, kind: AttachmentUploadSessionKind): string {
  if (typeof value !== "string" || value.length < 1 || value.length > 8192 || value.includes("\\") || /[\u0000-\u001f\u007f]/.test(value)) throw new Error("invalid_provider_response");
  let url: URL;
  try { url = new URL(value); } catch { throw new Error("invalid_provider_response"); }
  const expectedHost = kind === "outlook" ? "outlook.office.com" : "graph.microsoft.com";
  if (url.protocol !== "https:" || url.hostname.toLowerCase() !== expectedHost || url.port || url.username || url.password || url.hash) throw new Error("invalid_provider_response");
  const rawPathStart = value.indexOf("/", value.indexOf("://") + 3);
  const rawPathname = rawPathStart < 0 ? "/" : value.slice(rawPathStart).split(/[?#]/, 1)[0];
  rejectEncodedTraversal(rawPathname, "invalid_provider_response");
  rejectEncodedTraversal(url.pathname, "invalid_provider_response");
  if (kind === "todo" && url.search) throw new Error("invalid_provider_response");
  return value;
}

function uploadSessionNextOffset(payload: unknown): number {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) throw new Error("invalid_provider_response");
  const item = payload as Record<string, unknown>;
  const ranges = item.nextExpectedRanges ?? item.NextExpectedRanges;
  if (!Array.isArray(ranges) || ranges.length < 1 || ranges.length > 16) throw new Error("invalid_provider_response");
  let earliest = Number.POSITIVE_INFINITY;
  for (const range of ranges) {
    if (typeof range !== "string" || !/^\d+(?:-\d*)?$/.test(range)) throw new Error("invalid_provider_response");
    const start = Number(range.split("-", 1)[0]);
    if (!Number.isSafeInteger(start) || start < 0) throw new Error("invalid_provider_response");
    earliest = Math.min(earliest, start);
  }
  return earliest;
}

async function boundedUploadJson(response: Response): Promise<unknown> {
  const bytes = await readBytes(response, MAX_JSON_BYTES);
  try { return bytes.byteLength ? JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)) : {}; }
  catch { throw new Error("invalid_provider_response"); }
}

async function cancelAttachmentUploadSession(uploadUrl: string, kind: AttachmentUploadSessionKind, token: string, fetchFn: typeof fetch): Promise<void> {
  const headers = kind === "todo" ? { authorization: "Bearer " + token } : undefined;
  try { await fetchFn(uploadUrl, { method: "DELETE", headers, redirect: "error", signal: AbortSignal.timeout(5000) }); }
  catch { /* best-effort cleanup only; preserve the original upload error */ }
}

/** Upload a validated attachment in bounded sequential chunks through a provider-issued session. */
export async function uploadAttachmentSession(
  token: string,
  createSessionPath: string,
  createSessionBody: Record<string, unknown>,
  contentBase64: string,
  kind: AttachmentUploadSessionKind,
  signal?: AbortSignal,
  fetchFn: typeof fetch = fetch,
  requestTimeoutMs = 30000,
) {
  const content = decodeBase64Strict(contentBase64, "invalid_attachment");
  if (!Number.isInteger(requestTimeoutMs) || requestTimeoutMs < 1000 || requestTimeoutMs > 30000) throw new Error("invalid_request_timeout");
  const stepSignal = () => signal ? AbortSignal.any([signal, AbortSignal.timeout(requestTimeoutMs)]) : AbortSignal.timeout(requestTimeoutMs);
  const session = await graphRequest(token, createSessionPath, { method: "POST", body: createSessionBody, signal: stepSignal() }, fetchFn);
  const baseUploadUrl = canonicalAttachmentUploadUrl(session?.uploadUrl, kind);
  if (uploadSessionNextOffset(session) !== 0) throw new Error("invalid_provider_response");
  const target = new URL(baseUploadUrl);
  if (kind === "todo") target.pathname = `${target.pathname.replace(/\/$/, "")}/content`;
  const uploadUrl = target.href;
  let chunks = 0;
  try {
    for (let start = 0; start < content.byteLength; start += ATTACHMENT_UPLOAD_CHUNK_BYTES) {
      signal?.throwIfAborted();
      const endExclusive = Math.min(content.byteLength, start + ATTACHMENT_UPLOAD_CHUNK_BYTES);
      const end = endExclusive - 1;
      const headers: Record<string, string> = {
        "content-length": String(endExclusive - start),
        "content-range": `bytes ${start}-${end}/${content.byteLength}`,
        "content-type": "application/octet-stream",
      };
      if (kind === "todo") headers.authorization = "Bearer " + token;
      let response: Response;
      const chunkSignal = stepSignal();
      try {
        response = await fetchFn(uploadUrl, { method: "PUT", headers, body: content.subarray(start, endExclusive) as unknown as BodyInit, redirect: "error", signal: chunkSignal });
      } catch (error) {
        if (chunkSignal.aborted) throw chunkSignal.reason;
        throw new Error("provider_unavailable");
      }
      if (!response.ok) {
        try { await response.body?.cancel(); } catch { /* discard provider error content */ }
        throw providerStatusError(response.status);
      }
      chunks += 1;
      const final = endExclusive === content.byteLength;
      if (final) {
        if (response.status !== 201) throw new Error("invalid_provider_response");
        await readBytes(response, MAX_JSON_BYTES);
      } else {
        if (response.status !== 200 && response.status !== 202) throw new Error("invalid_provider_response");
        if (uploadSessionNextOffset(await boundedUploadJson(response)) !== endExclusive) throw new Error("invalid_provider_response");
      }
    }
  } catch (error) {
    await cancelAttachmentUploadSession(baseUploadUrl, kind, token, fetchFn);
    throw error;
  }
  return { bytes: content.byteLength, chunks, upload_mode: "session" as const };
}

export function drivePath(root: AllowedRoot, relativePath: string, suffix = ""): string {
  const base = `/drives/${safeId(root.drive_id)}/items/${safeId(root.item_id)}`;
  if (!relativePath) return `${base}${suffix}`;
  const encoded = relativePath.split("/").map((segment) => encodeURIComponent(segment)).join("/");
  return `${base}:/${encoded}:${suffix}`;
}

export function safeDriveItem(root: AllowedRoot, item: any) {
  const string = (value: unknown, max: number) => typeof value === "string" && value.length <= max ? value : null;
  const webUrl = (() => {
    const value = string(item?.webUrl, 2048);
    if (!value) return null;
    try {
      const url = new URL(value);
      const host = url.hostname.toLowerCase();
      return url.protocol === "https:" && (host === "onedrive.live.com" || host === "1drv.ms" || host.endsWith(".sharepoint.com")) ? value : null;
    } catch { return null; }
  })();
  return {
    root_label: root.label,
    id: string(item?.id, 256),
    name: string(item?.name, 512),
    web_url: webUrl,
    created: string(item?.createdDateTime, 64),
    last_modified: string(item?.lastModifiedDateTime, 64),
    description: string(item?.description, 4096),
    file_system_info: item?.fileSystemInfo && typeof item.fileSystemInfo === "object" ? {
      createdDateTime: string(item.fileSystemInfo.createdDateTime, 64),
      lastModifiedDateTime: string(item.fileSystemInfo.lastModifiedDateTime, 64),
    } : null,
    parent: item?.parentReference && typeof item.parentReference === "object" ? {
      id: string(item.parentReference.id, 256),
      path: string(item.parentReference.path, 2048),
    } : null,
    size: Number.isSafeInteger(item?.size) && item.size >= 0 ? item.size : null,
    mime_type: string(item?.file?.mimeType, 160),
    is_folder: item?.folder !== undefined,
    child_count: Number.isSafeInteger(item?.folder?.childCount) && item.folder.childCount >= 0 ? item.folder.childCount : null,
  };
}

const DRIVE_SELECT = "id,name,webUrl,createdDateTime,lastModifiedDateTime,description,fileSystemInfo,parentReference,size,file,folder,eTag";

function continuationPath(value: unknown, pathname: string): string {
  if (typeof value === "string" && /^https:/i.test(value)) {
    const absolutePath = `/v1.0${pathname}`;
    const canonical = canonicalGraphContinuation(value, absolutePath);
    return canonical.slice("/v1.0".length);
  }
  return canonicalGraphContinuation(value, pathname);
}

/** Validate a provider-owned nextLink against the stable Graph origin and exact collection path. */
export function canonicalProviderContinuation(value: unknown, pathname: string, message: "invalid_continuation" | "invalid_provider_response" = "invalid_continuation"): string {
  if (typeof value === "string" && /^https:/i.test(value)) {
    const absolutePath = `/v1.0${pathname}`;
    const canonical = canonicalGraphContinuation(value, absolutePath, message);
    return canonical.slice("/v1.0".length);
  }
  return canonicalGraphContinuation(value, pathname, message);
}

function providerContinuationUrl(value: unknown, pathname: string, message: "invalid_continuation" | "invalid_provider_response"): string {
  if (typeof value !== "string" || !/^https:/i.test(value)) invalidContinuation(message);
  canonicalProviderContinuation(value, pathname, message);
  return value;
}

async function graphContinuationRequest(token: string, continuation: unknown, pathname: string, signal?: AbortSignal, fetchFn: typeof fetch = fetch): Promise<any> {
  const url = providerContinuationUrl(continuation, pathname, "invalid_continuation");
  const directFetch = fetchFn;
  fetchFn = ((input: string | URL | Request, init?: RequestInit) => fetchGraphResponse(String(input), init ?? {}, signal, directFetch)) as typeof fetch;
  let response: Response;
  try {
    response = await fetchFn(url, { method: "GET", headers: { accept: "application/json", authorization: "Bearer " + token }, signal });
  } catch (error) {
    throwPreferredAbort(signal);
    if (error instanceof DOMException && error.name === "TimeoutError") throw error;
    if (signal?.aborted) throw signal.reason;
    throw new Error("provider_unavailable");
  }
  if (!response.ok) throw providerStatusError(response.status);
  let bytes: Uint8Array;
  try { bytes = await readBytes(response, MAX_JSON_BYTES); }
  catch (error) { throwPreferredAbort(signal); throw error; }
  try { return bytes.byteLength ? JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)) : {}; }
  catch { throw new Error("invalid_provider_response"); }
}

function drivePage(root: AllowedRoot, payload: any, limit: number, prefix: string) {
  const values = Array.isArray(payload?.value) ? payload.value : [];
  const items = values.slice(0, limit).map((item: any) => safeDriveItem(root, item));
  const rawNext = payload?.["@odata.nextLink"];
  const providerNextLink = rawNext === undefined ? undefined : canonicalProviderContinuation(rawNext, prefix, "invalid_provider_response");
  return { items, truncated: values.length > limit || providerNextLink !== undefined, ...(values.length <= limit && providerNextLink ? { providerNextLink } : {}) };
}

const DRIVE_SEARCH_PROVIDER_PAGES = 4;
const DRIVE_SEARCH_ANCESTRY_DEPTH = 64;
const DRIVE_SEARCH_ANCESTRY_REQUESTS = 2048;
const DRIVE_EXACT_FALLBACK_PAGES = 16;
const DRIVE_EXACT_FALLBACK_ITEMS = 200;
const DRIVE_EXACT_FALLBACK_STATE_IDS = 5000;

export type DriveSearchBudgets = {
  pages?: number;
  items?: number;
  stateIds?: number;
  providerPages?: number;
  ancestryDepth?: number;
  ancestryRequests?: number;
};

type DriveAncestryNode = { parentId: string | null; driveMatches: boolean };
type DriveAncestryProof = {
  maxDepth: number;
  maxRequests: number;
  requests: number;
  nodes: Map<string, DriveAncestryNode | null>;
  verdicts: Map<string, boolean>;
};

function boundedBudget(value: number | undefined, fallback: number, minimum: number, maximum: number): number {
  return Math.max(minimum, Math.min(value ?? fallback, maximum));
}

function ancestryProof(budgets: DriveSearchBudgets = {}): DriveAncestryProof {
  return {
    maxDepth: boundedBudget(budgets.ancestryDepth, DRIVE_SEARCH_ANCESTRY_DEPTH, 1, 128),
    maxRequests: boundedBudget(budgets.ancestryRequests, DRIVE_SEARCH_ANCESTRY_REQUESTS, 1, 4096),
    requests: 0,
    nodes: new Map(),
    verdicts: new Map(),
  };
}

function driveItemByIdPath(root: AllowedRoot, itemId: string): string {
  return `/drives/${safeId(root.drive_id)}/items/${safeId(itemId)}`;
}

async function loadDriveAncestryNode(
  root: AllowedRoot,
  itemId: string,
  token: string,
  proof: DriveAncestryProof,
  signal: AbortSignal | undefined,
  fetchFn: typeof fetch,
): Promise<DriveAncestryNode | null> {
  if (proof.nodes.has(itemId)) return proof.nodes.get(itemId)!;
  if (proof.requests >= proof.maxRequests) throw new Error("search_budget_exhausted");
  proof.requests += 1;
  const payload = await graphRequest(token, `${driveItemByIdPath(root, itemId)}?$select=id%2CparentReference`, { signal }, fetchFn);
  if (payload?.id !== itemId) {
    proof.nodes.set(itemId, null);
    return null;
  }
  if (itemId === root.item_id) {
    const node = { parentId: null, driveMatches: true };
    proof.nodes.set(itemId, node);
    return node;
  }
  const parent = payload?.parentReference;
  const node = parent && typeof parent === "object" && validStateId(parent.id) && typeof parent.driveId === "string"
    ? { parentId: parent.id, driveMatches: parent.driveId.toLowerCase() === root.drive_id.toLowerCase() }
    : null;
  proof.nodes.set(itemId, node);
  return node;
}

async function searchItemInsideRoot(
  root: AllowedRoot,
  item: any,
  token: string,
  proof: DriveAncestryProof,
  signal: AbortSignal | undefined,
  fetchFn: typeof fetch,
): Promise<boolean> {
  const itemId = item?.id;
  if (!validStateId(itemId)) return false;
  const known = proof.verdicts.get(itemId);
  if (known !== undefined) return known;
  const chain: string[] = [];
  const seen = new Set<string>();
  let currentId = itemId;
  for (let depth = 0; depth <= proof.maxDepth; depth += 1) {
    const verdict = proof.verdicts.get(currentId);
    if (verdict !== undefined) {
      for (const id of chain) proof.verdicts.set(id, verdict);
      return verdict;
    }
    if (seen.has(currentId)) {
      for (const id of chain) proof.verdicts.set(id, false);
      return false;
    }
    seen.add(currentId);
    chain.push(currentId);
    const node = await loadDriveAncestryNode(root, currentId, token, proof, signal, fetchFn);
    if (!node || !node.driveMatches) {
      for (const id of chain) proof.verdicts.set(id, false);
      return false;
    }
    if (currentId === root.item_id) {
      for (const id of chain) proof.verdicts.set(id, true);
      return true;
    }
    if (depth === proof.maxDepth || !node.parentId) {
      for (const id of chain) proof.verdicts.set(id, false);
      return false;
    }
    currentId = node.parentId;
  }
  return false;
}

async function driveSearchPage(
  root: AllowedRoot,
  payload: any,
  limit: number,
  prefix: string,
  token: string,
  proof: DriveAncestryProof,
  signal: AbortSignal | undefined,
  fetchFn: typeof fetch,
) {
  const values = Array.isArray(payload?.value) ? payload.value : [];
  if (values.length > limit) throw new Error("invalid_provider_response");
  const contained: any[] = [];
  for (const item of values) if (await searchItemInsideRoot(root, item, token, proof, signal, fetchFn)) contained.push(item);
  const items = contained.map((item: any) => safeDriveItem(root, item));
  const rawNext = payload?.["@odata.nextLink"];
  const providerNextLink = rawNext === undefined ? undefined : providerContinuationUrl(rawNext, prefix, "invalid_provider_response");
  return { items, truncated: providerNextLink !== undefined, ...(providerNextLink ? { providerNextLink } : {}), scanned: values.length };
}

export async function driveList(root: AllowedRoot, relativePath: string, token: string, limit: number, signal?: AbortSignal, fetchFn: typeof fetch = fetch) {
  const prefix = drivePath(root, relativePath, "/children");
  const payload = await graphRequest(token, `${prefix}?$top=${limit}&$select=${encodeURIComponent(DRIVE_SELECT)}`, { signal }, fetchFn);
  return drivePage(root, payload, limit, prefix);
}

export async function driveListContinuation(root: AllowedRoot, relativePath: string, token: string, limit: number, continuation: string, signal?: AbortSignal, fetchFn: typeof fetch = fetch) {
  const prefix = drivePath(root, relativePath, "/children");
  const payload = await graphRequest(token, continuationPath(continuation, prefix), { signal }, fetchFn);
  return drivePage(root, payload, limit, prefix);
}

export function driveSearchPath(root: AllowedRoot, query: string): string {
  const escaped = query.replaceAll("'", "''");
  return `/drives/${safeId(root.drive_id)}/items/${safeId(root.item_id)}/search(q='${encodeURIComponent(escaped)}')`;
}

export async function driveSearch(root: AllowedRoot, query: string, token: string, limit: number, signal?: AbortSignal, fetchFn: typeof fetch = fetch) {
  const pathname = driveSearchPath(root, query);
  const payload = await graphRequest(token, `${pathname}?$top=${limit}&$select=${encodeURIComponent(DRIVE_SELECT)}`, { signal }, fetchFn);
  return driveSearchPage(root, payload, limit, pathname, token, ancestryProof(), signal, fetchFn);
}

export async function driveSearchContinuation(root: AllowedRoot, query: string, token: string, limit: number, continuation: string, signal?: AbortSignal, fetchFn: typeof fetch = fetch) {
  const pathname = driveSearchPath(root, query);
  const payload = await graphContinuationRequest(token, continuation, pathname, signal, fetchFn);
  return driveSearchPage(root, payload, limit, pathname, token, ancestryProof(), signal, fetchFn);
}

export type DriveSearchContinuationState =
  | { kind: "provider"; nextLink: string; returnedIds: string[] }
  | { kind: "exact_fallback"; current: { folderId: string; nextLink?: string } | null; queue: string[]; seenFolderIds: string[]; seenItemIds: string[]; pendingItems?: ReturnType<typeof safeDriveItem>[]; scanIncomplete?: true };

export type DriveSearchMode = "provider" | "filename_exact" | "filename_stem" | "filename_contains";

export type NormalizedDriveSearch = {
  query: string;
  mode: DriveSearchMode;
  exhaustive: boolean;
};

export function exactFilenameQuery(query: string): string | null {
  const value = query.trim().normalize("NFC");
  if (!value || value.length > 255 || /[\\/*?<>:"|\u0000-\u001f\u007f]/.test(value)) return null;
  const dot = value.lastIndexOf(".");
  return dot > 0 && dot < value.length - 1 ? value : null;
}

function explicitFilenameQuery(query: string): string {
  const value = query.trim().normalize("NFC");
  if (!value || value.length > 255 || /[\\/*?<>:"|\u0000-\u001f\u007f]/.test(value)) throw new Error("invalid_search");
  return value;
}

/** Normalize routing criteria before credential selection or continuation lookup. */
export function normalizeDriveSearch(query: unknown, mode?: unknown, exhaustive?: unknown): NormalizedDriveSearch {
  if (typeof query !== "string") throw new Error("invalid_search");
  const normalizedQuery = query.trim().normalize("NFC");
  if (!normalizedQuery) throw new Error("invalid_search");
  if (mode !== undefined && !new Set<unknown>(["provider", "filename_exact", "filename_stem", "filename_contains"]).has(mode)) throw new Error("invalid_search");
  if (exhaustive !== undefined && typeof exhaustive !== "boolean") throw new Error("invalid_search");
  const normalizedMode = mode === undefined
    ? exactFilenameQuery(normalizedQuery) ? "filename_exact" : "provider"
    : mode as DriveSearchMode;
  const normalizedExhaustive = exhaustive === true;
  if (normalizedMode !== "provider") explicitFilenameQuery(normalizedQuery);
  else if (normalizedExhaustive) throw new Error("invalid_search");
  return { query: normalizedQuery, mode: normalizedMode, exhaustive: normalizedExhaustive };
}

function filenameStem(value: string): string {
  const dot = value.lastIndexOf(".");
  return dot > 0 ? value.slice(0, dot) : value;
}

function filenameMatches(name: string, query: string, mode: Exclude<DriveSearchMode, "provider">): boolean {
  const normalizedName = name.normalize("NFC").toLocaleLowerCase("en-US");
  const normalizedQuery = query.toLocaleLowerCase("en-US");
  if (mode === "filename_exact") return normalizedName === normalizedQuery;
  if (mode === "filename_stem") return filenameStem(normalizedName) === filenameStem(normalizedQuery);
  return normalizedName.includes(normalizedQuery);
}

function validStateId(value: unknown): value is string {
  return typeof value === "string" && value.length > 0 && value.length <= 256 && !/[\u0000-\u001f\u007f]/.test(value);
}

function validateSearchState(value: unknown, stateIds: number): DriveSearchContinuationState {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("invalid_continuation");
  const state = value as any;
  if (state.kind === "provider") {
    if (typeof state.nextLink !== "string" || !Array.isArray(state.returnedIds) || state.returnedIds.length > stateIds || state.returnedIds.some((id: unknown) => !validStateId(id))) throw new Error("invalid_continuation");
    return structuredClone(state);
  }
  if (state.kind !== "exact_fallback" || !Array.isArray(state.queue) || !Array.isArray(state.seenFolderIds) || !Array.isArray(state.seenItemIds)) throw new Error("invalid_continuation");
  if (state.scanIncomplete !== undefined && state.scanIncomplete !== true) throw new Error("invalid_continuation");
  const allIds = [...state.queue, ...state.seenFolderIds, ...state.seenItemIds];
  if (allIds.length > stateIds || allIds.some((id: unknown) => !validStateId(id))) throw new Error("invalid_continuation");
  if (state.current !== null && (!state.current || typeof state.current !== "object" || !validStateId(state.current.folderId) || (state.current.nextLink !== undefined && typeof state.current.nextLink !== "string"))) throw new Error("invalid_continuation");
  const pendingItems = state.pendingItems ?? [];
  if (!Array.isArray(pendingItems)) throw new Error("invalid_continuation");
  const pendingIds = pendingItems.map((item: any) => item?.id);
  if (pendingIds.some((id: unknown) => !validStateId(id)) || new Set(pendingIds).size !== pendingIds.length || pendingIds.some((id: string) => !state.seenItemIds.includes(id))) throw new Error("invalid_continuation");
  return structuredClone({ ...state, pendingItems });
}

function childrenByIdPath(root: AllowedRoot, folderId: string): string {
  return `/drives/${safeId(root.drive_id)}/items/${safeId(folderId)}/children`;
}

function directDriveChild(root: AllowedRoot, folderId: string, item: any): boolean {
  const parent = item?.parentReference;
  return validStateId(item?.id)
    && parent && typeof parent === "object"
    && parent.id === folderId
    && typeof parent.driveId === "string"
    && parent.driveId.toLowerCase() === root.drive_id.toLowerCase();
}

function skippableExactDescendantError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  return error.message === "invalid_provider_response"
    || (error.message === "provider_access_denied" && (error as Error & { providerStatus?: number }).providerStatus === 403);
}

async function exactFilenameFallback(
  root: AllowedRoot,
  search: NormalizedDriveSearch & { mode: Exclude<DriveSearchMode, "provider"> },
  token: string,
  limit: number,
  inputState: DriveSearchContinuationState | undefined,
  signal: AbortSignal | undefined,
  fetchFn: typeof fetch,
  budgets: DriveSearchBudgets,
) {
  const maxPages = Math.max(1, Math.min(budgets.pages ?? DRIVE_EXACT_FALLBACK_PAGES, 32));
  const maxItems = Math.max(1, Math.min(budgets.items ?? DRIVE_EXACT_FALLBACK_ITEMS, 1000));
  const maxStateIds = Math.max(10, Math.min(budgets.stateIds ?? DRIVE_EXACT_FALLBACK_STATE_IDS, 10000));
  const returnedIds = inputState?.kind === "provider" ? inputState.returnedIds : [];
  const state = inputState?.kind === "exact_fallback" ? validateSearchState(inputState, maxStateIds) as Extract<DriveSearchContinuationState, { kind: "exact_fallback" }> : {
    kind: "exact_fallback" as const,
    current: { folderId: root.item_id },
    queue: [],
    seenFolderIds: [root.item_id],
    seenItemIds: [...returnedIds],
    pendingItems: [],
  };
  state.pendingItems ??= [];
  if (state.pendingItems.some((item) => item.root_label !== root.label)) throw new Error("invalid_continuation");
  const seenFolders = new Set(state.seenFolderIds);
  const seenItems = new Set(state.seenItemIds);
  const items: ReturnType<typeof safeDriveItem>[] = [];
  items.push(...state.pendingItems.splice(0, limit));
  let pages = 0;
  let scanned = 0;
  while (state.current && pages < maxPages && scanned < maxItems && items.length < limit) {
    const folderId = state.current.folderId;
    const prefix = childrenByIdPath(root, folderId);
    // The caller's result limit must not force tiny Graph pages. Exact-name
    // traversal has its own bounded scan budget, so use that budget for the
    // provider page and keep `limit` solely as the return cap.
    const pageTop = Math.max(1, maxItems - scanned);
    pages += 1;
    let values: any[];
    let providerNextLink: string | undefined;
    try {
      const payload = state.current.nextLink
        ? await graphContinuationRequest(token, state.current.nextLink, prefix, signal, fetchFn)
        : await graphRequest(token, `${prefix}?$top=${pageTop}&$select=${encodeURIComponent(DRIVE_SELECT)}`, { signal }, fetchFn);
      if (!Array.isArray(payload?.value) || payload.value.length > pageTop) throw new Error("invalid_provider_response");
      values = payload.value;
      const rawNext = payload?.["@odata.nextLink"];
      providerNextLink = rawNext === undefined ? undefined : providerContinuationUrl(rawNext, prefix, "invalid_provider_response");
    } catch (error) {
      if (signal?.aborted && !graphOperationDeadlineReached(signal)) throwPreferredAbort(signal);
      if (graphOperationDeadlineReached(signal) || (error instanceof DOMException && error.name === "TimeoutError")) break;
      if (folderId === root.item_id || !skippableExactDescendantError(error)) throw error;
      // Continue with reachable siblings, but retain that this traversal can no
      // longer prove whole-root completeness. The flag must survive opaque
      // continuations so a later terminal page cannot claim scan_complete.
      state.scanIncomplete = true;
      state.current = state.queue.length ? { folderId: state.queue.shift()! } : null;
      continue;
    }
    for (let index = 0; index < values.length && scanned < maxItems; index += 1) {
      const item = values[index];
      scanned += 1;
      if (!directDriveChild(root, folderId, item)) continue;
      const id = item?.id;
      if (item?.folder !== undefined) {
        if (!seenFolders.has(id)) { seenFolders.add(id); state.queue.push(id); }
        continue;
      }
      if (item?.file !== undefined && typeof item?.name === "string" && filenameMatches(item.name, search.query, search.mode) && !seenItems.has(id)) {
        seenItems.add(id);
        const safeItem = safeDriveItem(root, item);
        if (items.length < limit) items.push(safeItem);
        else state.pendingItems.push(safeItem);
        if (!search.exhaustive) {
          const scanComplete = state.scanIncomplete !== true && providerNextLink === undefined && index === values.length - 1 && state.queue.length === 0;
          return { items, truncated: false, scan_complete: scanComplete, match_satisfied: true, fallback: true, scanned };
        }
      }
    }
    if (providerNextLink !== undefined) state.current.nextLink = providerNextLink;
    else state.current = state.queue.length ? { folderId: state.queue.shift()! } : null;
    state.seenFolderIds = [...seenFolders];
    state.seenItemIds = [...seenItems];
    if (state.queue.length + state.seenFolderIds.length + state.seenItemIds.length > maxStateIds) throw new Error("search_budget_exhausted");
    // A provider continuation can retain the page size of the request that
    // created it. Do not consume it with only the remainder of this call's
    // item budget: a valid next page may then be larger than `pageTop` and be
    // misclassified as malformed. Resume the opaque nextLink with a fresh
    // per-call scan budget instead.
    if (providerNextLink !== undefined) break;
  }
  const continuationState = state.current || state.pendingItems.length ? state : undefined;
  return {
    items,
    truncated: continuationState !== undefined,
    scan_complete: continuationState === undefined && state.scanIncomplete !== true,
    match_satisfied: seenItems.size > 0,
    ...(continuationState ? { continuationState } : {}),
    fallback: true,
    scanned,
  };
}

export async function driveSearchScoped(
  root: AllowedRoot,
  query: string,
  token: string,
  limit: number,
  inputState?: unknown,
  signal?: AbortSignal,
  fetchFn: typeof fetch = fetch,
  budgets: DriveSearchBudgets = {},
  mode?: DriveSearchMode,
  exhaustive = false,
) {
  const search = normalizeDriveSearch(query, mode, exhaustive);
  const maxStateIds = Math.max(10, Math.min(budgets.stateIds ?? DRIVE_EXACT_FALLBACK_STATE_IDS, 10000));
  const state = inputState === undefined ? undefined : validateSearchState(inputState, maxStateIds);
  if (state && ((search.mode === "provider" && state.kind !== "provider") || (search.mode !== "provider" && state.kind !== "exact_fallback"))) throw new Error("invalid_continuation");
  if (search.mode !== "provider") {
    if (state?.kind === "provider") throw new Error("invalid_continuation");
    return exactFilenameFallback(root, search as NormalizedDriveSearch & { mode: Exclude<DriveSearchMode, "provider"> }, token, limit, state, signal, fetchFn, budgets);
  }

  const pathname = driveSearchPath(root, search.query);
  const proof = ancestryProof(budgets);
  const maxProviderPages = boundedBudget(budgets.providerPages, DRIVE_SEARCH_PROVIDER_PAGES, 1, 16);
  let nextLink = state?.kind === "provider" ? state.nextLink : undefined;
  let page: Awaited<ReturnType<typeof driveSearchPage>>;
  let providerPages = 0;
  do {
    const payload = nextLink
      ? await graphContinuationRequest(token, nextLink, pathname, signal, fetchFn)
      : await graphRequest(token, `${pathname}?$top=${limit}&$select=${encodeURIComponent(DRIVE_SELECT)}`, { signal }, fetchFn);
    page = await driveSearchPage(root, payload, limit, pathname, token, proof, signal, fetchFn);
    providerPages += 1;
    nextLink = page.providerNextLink;
  } while (!page.items.length && nextLink && providerPages < maxProviderPages);

  const { providerNextLink, ...result } = page;
  const returnedIds = [...new Set([
    ...(state?.kind === "provider" ? state.returnedIds : []),
    ...page.items.map((item: any) => item?.id).filter(validStateId),
  ])];
  if (returnedIds.length > maxStateIds) throw new Error("search_budget_exhausted");
  const continuationState = providerNextLink ? { kind: "provider" as const, nextLink: providerNextLink, returnedIds } : undefined;
  return {
    ...result,
    scan_complete: providerNextLink === undefined,
    match_satisfied: returnedIds.length > 0,
    ...(continuationState ? { continuationState } : {}),
  };
}

export function contentTypeAllowed(value: string): boolean { const mime = value.split(";", 1)[0].trim().toLowerCase(); return TEXT_MIME.test(mime) || BINARY_MIME.has(mime); }

export async function driveRead(root: AllowedRoot, relativePath: string, token: string, mode: "text" | "base64" | "digest", maxBytes: number, signal?: AbortSignal, fetchFn: typeof fetch = fetch) {
  if (!relativePath) throw new Error("invalid_relative_path");
  const metadata = await graphRequest(token, `${drivePath(root, relativePath)}?$select=${encodeURIComponent(DRIVE_SELECT)}`, { signal }, fetchFn);
  if (metadata.folder) throw new Error("file_required");
  const mime = String(metadata.file?.mimeType ?? "application/octet-stream").toLowerCase();
  if (!contentTypeAllowed(mime)) throw new Error("unsupported_file_type");
  if (Number(metadata.size) > maxBytes) throw new Error("file_too_large");
  if (mode === "digest") {
    const digest = await graphDigestRequest(token, drivePath(root, relativePath, "/content"), maxBytes, signal, fetchFn);
    return { ok: true, operation: "read", item: safeDriveItem(root, metadata), bytes: digest.bytes, mode, sha256: digest.sha256 };
  }
  const bytes = await graphRequest(token, drivePath(root, relativePath, "/content"), { signal, response: "bytes", maxBytes }, fetchFn);
  const base = { ok: true, operation: "read", item: safeDriveItem(root, metadata), bytes: bytes.byteLength, mode };
  if (mode === "base64") return { ...base, content_base64: Buffer.from(bytes).toString("base64") };
  if (!TEXT_MIME.test(mime)) throw new Error("binary_requires_base64_or_digest");
  try { return { ...base, content_text: new TextDecoder("utf-8", { fatal: true }).decode(bytes) }; }
  catch { throw new Error("invalid_text_encoding"); }
}

/** Read one AGENTS.md candidate with exactly one non-retried Graph GET. */
export async function driveReadInstructionsCandidate(root: AllowedRoot, relativePath: string, token: string, maxBytes: number, signal?: AbortSignal, fetchFn: typeof fetch = fetch): Promise<Uint8Array | null> {
  if (!relativePath || !relativePath.endsWith("AGENTS.md") || !Number.isSafeInteger(maxBytes) || maxBytes < 1) throw new Error("invalid_instruction_candidate");
  try {
    return await graphRequest(token, drivePath(root, relativePath, "/content"), { signal, response: "bytes", maxBytes, readRetries: 0 }, fetchFn);
  } catch (error) {
    if (error instanceof Error && error.message === "item_not_found") return null;
    throw error;
  }
}

export async function driveWriteBytes(root: AllowedRoot, relativePath: string, token: string, content: Uint8Array, contentType: string, maxBytes: number | undefined, update: boolean, signal?: AbortSignal, fetchFn: typeof fetch = fetch) {
  validateDriveWriteBytes(relativePath, content, contentType, maxBytes);
  const sourceByteSize = content.byteLength;
  const sourceSha256 = createHash("sha256").update(content).digest("hex");
  const headers: Record<string, string> = { authorization: "Bearer " + token, "content-type": contentType, [update ? "if-match" : "if-none-match"]: update ? "*" : "*" };
  if (update) {
    const current = await graphRequest(token, `${drivePath(root, relativePath)}?$select=eTag,file,folder`, { signal });
    if (current.folder || typeof current.eTag !== "string") throw new Error("file_required");
    headers["if-match"] = current.eTag;
  }
  let response: Response;
  try { response = await fetchFn(`${GRAPH}${drivePath(root, relativePath, "/content")}`, { method: "PUT", headers, body: content as unknown as BodyInit, signal }); }
  catch { throw new Error("provider_unavailable"); }
  if (!response.ok) {
    if (response.status === 409 || response.status === 412) throw new Error("item_conflict");
    if (response.status === 401 || response.status === 403) throw new Error("provider_access_denied");
    throw new Error(`provider_error_${response.status}`);
  }
  const bytes = await readBytes(response, MAX_JSON_BYTES);
  let item: unknown;
  try { item = bytes.byteLength ? JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)) : {}; }
  catch { throw new Error("invalid_provider_response"); }
  const projected = safeDriveItem(root, item);
  const graphReportedSize = projected.size;
  const sizeMatch = graphReportedSize === sourceByteSize;
  if (!sizeMatch) throw new Error("invalid_provider_response");
  return {
    ok: true,
    operation: update ? "update" : "upload",
    source_byte_size: sourceByteSize,
    source_sha256: sourceSha256,
    graph_reported_size: graphReportedSize,
    size_match: true,
    item: projected,
  };
}

export type DriveUploadSource = {
  size: number;
  sha256: string;
  readChunk(offset: number, maximumBytes: number): Promise<Buffer>;
  assertUnchanged(): Promise<void>;
};

/** Validate a provider-issued OneDrive upload URL before sending file bytes without Graph auth. */
export function canonicalDriveUploadUrl(value: unknown): string {
  if (typeof value !== "string" || value.length < 1 || value.length > 8192 || value.includes("\\") || /[\u0000-\u001f\u007f]/.test(value)) throw new Error("invalid_provider_response");
  let url: URL;
  try { url = new URL(value); } catch { throw new Error("invalid_provider_response"); }
  const hostname = url.hostname.toLowerCase();
  if (url.protocol !== "https:" || !hostname.endsWith(".up.1drv.com") || url.port || url.username || url.password || url.hash) throw new Error("invalid_provider_response");
  const rawPathStart = value.indexOf("/", value.indexOf("://") + 3);
  const rawPathname = rawPathStart < 0 ? "/" : value.slice(rawPathStart).split(/[?#]/, 1)[0];
  rejectEncodedTraversal(rawPathname, "invalid_provider_response");
  rejectEncodedTraversal(url.pathname, "invalid_provider_response");
  return value;
}

function driveWriteReceipt(root: AllowedRoot, update: boolean, source: DriveUploadSource, item: unknown, uploadMode: "simple" | "session", chunks: number) {
  const projected = safeDriveItem(root, item);
  if (projected.size !== source.size) throw new Error("invalid_provider_response");
  return {
    ok: true,
    operation: update ? "update" : "upload",
    source_byte_size: source.size,
    source_sha256: source.sha256,
    graph_reported_size: projected.size,
    size_match: true,
    upload_mode: uploadMode,
    chunks,
    item: projected,
  };
}

async function parseDriveItemResponse(response: Response): Promise<unknown> {
  const bytes = await readBytes(response, MAX_JSON_BYTES);
  try { return bytes.byteLength ? JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)) : {}; }
  catch { throw new Error("invalid_provider_response"); }
}

/** Stream a protected source through simple upload or a sequential OneDrive upload session. */
export async function driveWriteSource(root: AllowedRoot, relativePath: string, token: string, source: DriveUploadSource, contentType: string, update: boolean, signal?: AbortSignal, fetchFn: typeof fetch = fetch, requestTimeoutMs = 30_000, simpleUploadMaxBytes = ONEDRIVE_SIMPLE_UPLOAD_MAX_BYTES, uploadChunkBytes = ONEDRIVE_UPLOAD_CHUNK_BYTES) {
  if (!relativePath || !contentTypeAllowed(contentType) || !Number.isSafeInteger(source.size) || source.size < 0 || !/^[a-f0-9]{64}$/.test(source.sha256)) throw new Error("invalid_write_input");
  if (source.size > ONEDRIVE_WRITE_MAX_BYTES) throw new Error("provider_file_too_large");
  if (!Number.isInteger(requestTimeoutMs) || requestTimeoutMs < 1000 || requestTimeoutMs > 30_000) throw new Error("invalid_request_timeout");
  if (!Number.isSafeInteger(simpleUploadMaxBytes) || simpleUploadMaxBytes < 1 || simpleUploadMaxBytes > ONEDRIVE_SIMPLE_UPLOAD_MAX_BYTES) throw new Error("invalid_byte_limit");
  if (!Number.isSafeInteger(uploadChunkBytes) || uploadChunkBytes < 320 * 1024 || uploadChunkBytes >= 60 * 1024 * 1024 || uploadChunkBytes % (320 * 1024) !== 0) throw new Error("invalid_byte_limit");
  let currentEtag: string | undefined;
  if (update) {
    const current = await graphRequest(token, `${drivePath(root, relativePath)}?$select=eTag,file,folder`, { signal }, fetchFn);
    if (current.folder || typeof current.eTag !== "string") throw new Error("file_required");
    currentEtag = current.eTag;
  }

  if (source.size <= simpleUploadMaxBytes) {
    const sourceFailure: { error?: unknown } = {};
    let streamCompleted = source.size === 0;
    async function* bodyChunks() {
      const hash = createHash("sha256");
      try {
        for (let offset = 0; offset < source.size;) {
          const chunk = await source.readChunk(offset, Math.min(uploadChunkBytes, source.size - offset));
          if (!chunk.byteLength || offset + chunk.byteLength > source.size) throw new Error("invalid_source_media_uri");
          hash.update(chunk);
          offset += chunk.byteLength;
          if (offset === source.size) {
            await source.assertUnchanged();
            if (hash.digest("hex") !== source.sha256) throw new Error("invalid_source_media_uri");
          }
          yield chunk;
        }
        streamCompleted = true;
      } catch (error) {
        sourceFailure.error = error;
        throw error;
      }
    }
    if (source.size === 0) await source.assertUnchanged();
    const headers: Record<string, string> = {
      authorization: "Bearer " + token,
      "content-type": contentType,
      "content-length": String(source.size),
      [update ? "if-match" : "if-none-match"]: update ? currentEtag! : "*",
    };
    let response: Response;
    try {
      const body = source.size === 0 ? Buffer.alloc(0) : bodyChunks() as unknown as BodyInit;
      response = await fetchFn(`${GRAPH}${drivePath(root, relativePath, "/content")}`, { method: "PUT", headers, body, signal, ...(source.size > 0 ? { duplex: "half" } : {}) } as RequestInit);
    } catch {
      if (sourceFailure.error) throw sourceFailure.error;
      if (signal?.aborted) throw signal.reason;
      throw new Error("provider_unavailable");
    }
    if (!streamCompleted) throw new Error("invalid_provider_response");
    if (!response.ok) throw providerStatusError(response.status);
    return driveWriteReceipt(root, update, source, await parseDriveItemResponse(response), "simple", source.size === 0 ? 0 : Math.ceil(source.size / uploadChunkBytes));
  }

  const name = relativePath.split("/").at(-1)!;
  const session = await graphRequest(token, drivePath(root, relativePath, "/createUploadSession"), {
    method: "POST",
    body: { item: { "@microsoft.graph.conflictBehavior": update ? "replace" : "fail", name } },
    ...(currentEtag ? { headers: { "if-match": currentEtag } } : {}),
    signal,
  }, fetchFn);
  const uploadUrl = canonicalDriveUploadUrl(session?.uploadUrl);
  const hash = createHash("sha256");
  let chunks = 0;
  try {
    for (let offset = 0; offset < source.size;) {
      signal?.throwIfAborted();
      const chunk = await source.readChunk(offset, Math.min(uploadChunkBytes, source.size - offset));
      if (!chunk.byteLength || offset + chunk.byteLength > source.size) throw new Error("invalid_source_media_uri");
      hash.update(chunk);
      const endExclusive = offset + chunk.byteLength;
      if (endExclusive === source.size) {
        await source.assertUnchanged();
        if (hash.digest("hex") !== source.sha256) throw new Error("invalid_source_media_uri");
      }
      const requestSignal = signal ? AbortSignal.any([signal, AbortSignal.timeout(requestTimeoutMs)]) : AbortSignal.timeout(requestTimeoutMs);
      let response: Response;
      try {
        response = await fetchFn(uploadUrl, {
          method: "PUT",
          headers: {
            "content-length": String(chunk.byteLength),
            "content-range": `bytes ${offset}-${endExclusive - 1}/${source.size}`,
            "content-type": "application/octet-stream",
          },
          body: chunk as unknown as BodyInit,
          redirect: "error",
          signal: requestSignal,
        });
      } catch {
        if (requestSignal.aborted) throw requestSignal.reason;
        throw new Error("provider_unavailable");
      }
      if (!response.ok) throw providerStatusError(response.status);
      chunks += 1;
      if (endExclusive === source.size) {
        if (response.status !== 200 && response.status !== 201) throw new Error("invalid_provider_response");
        return driveWriteReceipt(root, update, source, await parseDriveItemResponse(response), "session", chunks);
      }
      if (response.status !== 202 || uploadSessionNextOffset(await boundedUploadJson(response)) !== endExclusive) throw new Error("invalid_provider_response");
      offset = endExclusive;
    }
    throw new Error("invalid_provider_response");
  } catch (error) {
    try { await fetchFn(uploadUrl, { method: "DELETE", redirect: "error", signal: AbortSignal.timeout(5000) }); } catch { /* best-effort cleanup */ }
    throw error;
  }
}

export async function driveWrite(root: AllowedRoot, relativePath: string, token: string, contentBase64: string, contentType: string, maxBytes: number | undefined, update: boolean, signal?: AbortSignal, fetchFn: typeof fetch = fetch) {
  validateDriveWriteInput(relativePath, contentBase64, contentType, maxBytes);
  return driveWriteBytes(root, relativePath, token, decodeBase64Strict(contentBase64), contentType, maxBytes, update, signal, fetchFn);
}

export async function driveDelete(root: AllowedRoot, relativePath: string, token: string, signal?: AbortSignal) {
  if (!relativePath) throw new Error("invalid_relative_path");
  await graphRequest(token, drivePath(root, relativePath), { method: "DELETE", response: "none", signal });
  return { ok: true, operation: "delete", deleted: true, root_label: root.label };
}

export async function driveMetadataUpdate(root: AllowedRoot, relativePath: string, token: string, changes: { name?: string; destinationRelativePath?: string; description?: string | null; fileSystemInfo?: { createdDateTime?: string; lastModifiedDateTime?: string } }, signal?: AbortSignal, fetchFn: typeof fetch = fetch) {
  validateDriveMetadataInput(relativePath, changes);
  const body: Record<string, unknown> = {};
  if (changes.name !== undefined) {
    body.name = changes.name;
  }
  if (changes.description !== undefined) body.description = changes.description;
  if (changes.fileSystemInfo !== undefined) body.fileSystemInfo = changes.fileSystemInfo;
  if (changes.destinationRelativePath !== undefined) {
    const destination = await graphRequest(token, `${drivePath(root, changes.destinationRelativePath)}?$select=id,folder`, { signal }, fetchFn);
    if (typeof destination?.id !== "string" || destination.folder === undefined) throw new Error("invalid_destination");
    body.parentReference = { id: destination.id };
  }
  const current = await graphRequest(token, `${drivePath(root, relativePath)}?$select=eTag`, { signal }, fetchFn);
  if (typeof current?.eTag !== "string") throw new Error("invalid_provider_response");
  const item = await graphRequest(token, drivePath(root, relativePath), { method: "PATCH", body, headers: { "if-match": current.eTag }, signal }, fetchFn);
  return { ok: true, operation: "metadata_update", item: safeDriveItem(root, item) };
}

export async function driveCreateFolder(root: AllowedRoot, parentRelativePath: string, name: string, conflictBehavior: "fail" | "rename", token: string, signal?: AbortSignal, fetchFn: typeof fetch = fetch) {
  validateDriveFolderInput(name);
  const item = await graphRequest(token, drivePath(root, parentRelativePath, "/children"), {
    method: "POST",
    body: { name, folder: {}, "@microsoft.graph.conflictBehavior": conflictBehavior },
    signal,
  }, fetchFn);
  return { ok: true, operation: "create_folder", item: safeDriveItem(root, item) };
}
