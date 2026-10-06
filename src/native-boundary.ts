import { createHmac, timingSafeEqual } from "node:crypto";
import { chmod, lstat, mkdir, unlink } from "node:fs/promises";
import { createServer, type Server, type Socket } from "node:net";
import { dirname, isAbsolute } from "node:path";

const MAXIMUM_REQUEST_BYTES = 1024 * 1024;
const MAXIMUM_CLOCK_SKEW_MS = 30_000;
const MAXIMUM_REPLAY_ENTRIES = 4096;

export const NATIVE_CONNECTED_TOOLS = new Set([
  "onedrive_search",
  "onedrive_list",
  "onedrive_read",
  "onedrive_upload_small",
  "onedrive_metadata_update",
  "onedrive_create_folder",
  "onedrive_delete",
  "outlook_calendar_read",
  "outlook_calendar_write",
  "outlook_mail_read",
  "outlook_mail_write",
  "microsoft_todo_read",
  "microsoft_todo_write",
]);

type JsonObject = Record<string, unknown>;

export type NativeBoundaryRequest = {
  format: "openclaw-microsoft-graph-native-request/1";
  operationId: string;
  intentDigest: string;
  tool: string;
  parameters: JsonObject;
  issuedAtMs: number;
  nonce: string;
  signature: string;
};

export type NativeBoundaryExecutor = (
  tool: string,
  parameters: JsonObject,
) => Promise<unknown>;

function canonicalValue(value: unknown): unknown {
  if (value === null || typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") {
    if (!Number.isSafeInteger(value)) throw new Error("native_boundary_noncanonical_json");
    return value;
  }
  if (Array.isArray(value)) return value.map(canonicalValue);
  if (!value || typeof value !== "object" || Object.getPrototypeOf(value) !== Object.prototype) {
    throw new Error("native_boundary_noncanonical_json");
  }
  return Object.fromEntries(
    Object.keys(value as JsonObject).sort().map((key) => [key, canonicalValue((value as JsonObject)[key])]),
  );
}

export function canonicalNativeJson(value: unknown): Buffer {
  return Buffer.from(JSON.stringify(canonicalValue(value)), "utf8");
}

function decodeKey(value: unknown): Buffer {
  if (typeof value !== "string" || !/^[A-Za-z0-9_-]{43}$/.test(value)) {
    throw new Error("native_boundary_key_invalid");
  }
  const key = Buffer.from(value, "base64url");
  if (key.byteLength !== 32 || key.toString("base64url") !== value) {
    throw new Error("native_boundary_key_invalid");
  }
  return key;
}

function requestBody(request: NativeBoundaryRequest): Omit<NativeBoundaryRequest, "signature"> {
  const { signature: _signature, ...body } = request;
  return body;
}

export function signNativeBoundaryRequest(
  request: Omit<NativeBoundaryRequest, "signature">,
  keyValue: string,
): NativeBoundaryRequest {
  const key = decodeKey(keyValue);
  const signature = createHmac("sha256", key).update(canonicalNativeJson(request)).digest("base64url");
  return { ...request, signature };
}

export class NativeBoundaryVerifier {
  readonly #key: Buffer;
  readonly #seen = new Map<string, number>();

  constructor(keyValue: string, private readonly now: () => number = Date.now) {
    this.#key = decodeKey(keyValue);
  }

  verify(value: unknown): NativeBoundaryRequest {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      throw new Error("native_boundary_request_invalid");
    }
    const request = value as NativeBoundaryRequest;
    if (
      Object.keys(request).sort().join(",") !==
        "format,intentDigest,issuedAtMs,nonce,operationId,parameters,signature,tool"
      || request.format !== "openclaw-microsoft-graph-native-request/1"
      || typeof request.operationId !== "string"
      || !/^[A-Za-z0-9_-]{1,128}$/.test(request.operationId)
      || typeof request.intentDigest !== "string"
      || !/^sha256:[a-f0-9]{64}$/.test(request.intentDigest)
      || typeof request.tool !== "string"
      || !NATIVE_CONNECTED_TOOLS.has(request.tool)
      || !request.parameters
      || typeof request.parameters !== "object"
      || Array.isArray(request.parameters)
      || Object.getPrototypeOf(request.parameters) !== Object.prototype
      || !Number.isSafeInteger(request.issuedAtMs)
      || Math.abs(this.now() - request.issuedAtMs) > MAXIMUM_CLOCK_SKEW_MS
      || typeof request.nonce !== "string"
      || !/^[a-f0-9]{32}$/.test(request.nonce)
      || typeof request.signature !== "string"
      || !/^[A-Za-z0-9_-]{43}$/.test(request.signature)
    ) {
      throw new Error("native_boundary_request_invalid");
    }
    canonicalNativeJson(request.parameters);
    const expected = createHmac("sha256", this.#key)
      .update(canonicalNativeJson(requestBody(request)))
      .digest();
    const supplied = Buffer.from(request.signature, "base64url");
    if (supplied.byteLength !== expected.byteLength || !timingSafeEqual(supplied, expected)) {
      throw new Error("native_boundary_signature_invalid");
    }
    const cutoff = this.now() - MAXIMUM_CLOCK_SKEW_MS;
    for (const [nonce, issuedAt] of this.#seen) if (issuedAt < cutoff) this.#seen.delete(nonce);
    if (this.#seen.has(request.nonce)) throw new Error("native_boundary_replay");
    if (this.#seen.size >= MAXIMUM_REPLAY_ENTRIES) throw new Error("native_boundary_replay_capacity");
    this.#seen.set(request.nonce, request.issuedAtMs);
    return request;
  }
}

function response(value: unknown): Buffer {
  const encoded = canonicalNativeJson(value);
  if (encoded.byteLength > MAXIMUM_REQUEST_BYTES) throw new Error("native_boundary_response_too_large");
  return Buffer.concat([encoded, Buffer.from("\n")]);
}

function refusalCode(error: unknown): string {
  const code = error instanceof Error ? error.message : "";
  if (code === "item_not_found") return "ITEM_NOT_FOUND";
  if (code === "item_conflict") return "ITEM_CONFLICT";
  if (code === "access_denied" || code === "provider_access_denied") return "ACCESS_DENIED";
  if (code === "provider_throttled" || code === "provider_unavailable" || code === "request_timeout") {
    return "PROVIDER_UNAVAILABLE";
  }
  return "NATIVE_BOUNDARY_REFUSED";
}

export class NativeBoundaryService {
  readonly #verifier: NativeBoundaryVerifier;
  #server?: Server;

  constructor(
    private readonly socketPath: string,
    keyValue: string,
    private readonly execute: NativeBoundaryExecutor,
    now: () => number = Date.now,
  ) {
    if (!isAbsolute(socketPath) || typeof execute !== "function") {
      throw new Error("native_boundary_configuration_invalid");
    }
    this.#verifier = new NativeBoundaryVerifier(keyValue, now);
  }

  async #handle(socket: Socket): Promise<void> {
    socket.setTimeout(30_000, () => socket.destroy());
    const chunks: Buffer[] = [];
    let total = 0;
    try {
      const raw = await new Promise<Buffer>((resolve, reject) => {
        let settled = false;
        socket.on("data", (chunk: Buffer) => {
          if (settled) return;
          total += chunk.byteLength;
          if (total > MAXIMUM_REQUEST_BYTES + 1) {
            settled = true;
            reject(new Error("native_boundary_request_too_large"));
            return;
          }
          chunks.push(chunk);
          const joined = Buffer.concat(chunks, total);
          const newline = joined.indexOf(10);
          if (newline < 0) return;
          if (joined.subarray(newline + 1).some((byte) => ![9, 10, 13, 32].includes(byte))) {
            settled = true;
            reject(new Error("native_boundary_trailing_data"));
            return;
          }
          settled = true;
          resolve(joined.subarray(0, newline));
        });
        socket.on("error", reject);
        socket.on("end", () => {
          if (!settled) reject(new Error("native_boundary_request_incomplete"));
        });
      });
      const parsed = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(raw)) as unknown;
      const request = this.#verifier.verify(parsed);
      let result: unknown;
      try {
        result = await this.execute(request.tool, request.parameters);
      } catch (error) {
        socket.end(response({
          format: "openclaw-microsoft-graph-native-response/1",
          status: "REFUSED",
          code: refusalCode(error),
        }));
        return;
      }
      socket.end(response({
        format: "openclaw-microsoft-graph-native-response/1",
        status: "OK",
        operationId: request.operationId,
        intentDigest: request.intentDigest,
        result,
      }));
    } catch {
      socket.end(response({
        format: "openclaw-microsoft-graph-native-response/1",
        status: "REFUSED",
        code: "NATIVE_BOUNDARY_REFUSED",
      }));
    }
  }

  async start(): Promise<void> {
    if (this.#server) throw new Error("native_boundary_already_started");
    await mkdir(dirname(this.socketPath), { recursive: true, mode: 0o700 });
    try {
      const existing = await lstat(this.socketPath);
      if (!existing.isSocket() || (typeof process.getuid === "function" && existing.uid !== process.getuid())) {
        throw new Error("native_boundary_socket_unsafe");
      }
      await unlink(this.socketPath);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
    const server = createServer((socket) => { void this.#handle(socket); });
    this.#server = server;
    await new Promise<void>((resolve, reject) => {
      server.once("error", reject);
      server.listen(this.socketPath, () => {
        server.off("error", reject);
        resolve();
      });
    });
    await chmod(this.socketPath, 0o600);
  }

  async stop(): Promise<void> {
    const server = this.#server;
    this.#server = undefined;
    if (server) await new Promise<void>((resolve) => server.close(() => resolve()));
    try { await unlink(this.socketPath); } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
  }
}
