import { createHash, createPublicKey, verify as verifySignature } from "node:crypto";
import { lstat } from "node:fs/promises";
import { createConnection } from "node:net";
import { isAbsolute } from "node:path";

import { canonicalNativeJson } from "./native-boundary.js";

const PERMIT_FORMAT = "native-connected-execution-permit/1";
const ENVELOPE_FORMAT = "native-connected-execution-permit-envelope/1";
const RESPONSE_FORMAT = "native-connected-external-execution-response/1";
const REGISTRY_KEY = Symbol.for("gemacode/native-execution-permits/1");
const MAXIMUM_RESPONSE_BYTES = 1024 * 1024;
const ED25519_SPKI_PREFIX = Buffer.from("302a300506032b6570032100", "hex");
const DIGEST = /^sha256:[0-9a-f]{64}$/;
const CALL_ID = /^[A-Za-z0-9][A-Za-z0-9_.:/-]{0,255}$/;
const TOOL = /^[A-Za-z][A-Za-z0-9_.-]{0,127}$/;

type JsonObject = Record<string, unknown>;
type PreparationEntry = {
  prepare: (params: unknown) => Promise<unknown>;
  sessionKey?: string;
  createdAtMs: number;
};

export type NativeExecutionConfig = {
  nativeExecutionRequired?: boolean;
  nativeExecutionPublicKey?: unknown;
  nativeExecutionSocketPath?: string;
  nativeExecutionSocketOwnerUid?: number;
};

export type VerifiedExecutionPermit = {
  operation_id: string;
  intent_digest: string;
  adapter_id: string;
  action: string;
  effect_class: "READ" | "STATE_MUTATION" | "MESSAGE_DELIVERY";
  tool_name: string;
  tool_call_id: string;
  parameters_digest: string;
  authorization_binding_digest: string;
  issued_at_ms: number;
  expires_at_ms: number;
  nonce: string;
};

export type ConsumedExecutionPermit = {
  permit: JsonObject;
  operationId: string;
  intentDigest: string;
};

function exactKeys(value: JsonObject, expected: string[]): boolean {
  return Object.keys(value).sort().join(",") === [...expected].sort().join(",");
}

/** QEL's established canonical JSON escapes every non-ASCII UTF-16 code unit. */
export function canonicalQelJson(value: unknown): Buffer {
  const native = canonicalNativeJson(value).toString("utf8");
  let ascii = "";
  for (let index = 0; index < native.length; index += 1) {
    const code = native.charCodeAt(index);
    ascii += code > 0x7f ? `\\u${code.toString(16).padStart(4, "0")}` : native[index];
  }
  return Buffer.from(ascii, "ascii");
}

export function qelParametersDigest(value: unknown): string {
  return `sha256:${createHash("sha256").update(canonicalQelJson(value)).digest("hex")}`;
}

function publicKey(value: unknown) {
  if (typeof value !== "string" || !/^b64u:[A-Za-z0-9_-]{43}$/.test(value)) {
    throw new Error("native_execution_public_key_invalid");
  }
  const raw = Buffer.from(value.slice(5), "base64url");
  if (raw.byteLength !== 32 || `b64u:${raw.toString("base64url")}` !== value) {
    throw new Error("native_execution_public_key_invalid");
  }
  return createPublicKey({ key: Buffer.concat([ED25519_SPKI_PREFIX, raw]), format: "der", type: "spki" });
}

function signature(value: unknown): Buffer {
  if (typeof value !== "string" || !/^b64u:[A-Za-z0-9_-]{86}$/.test(value)) {
    throw new Error("native_execution_signature_invalid");
  }
  const raw = Buffer.from(value.slice(5), "base64url");
  if (raw.byteLength !== 64 || `b64u:${raw.toString("base64url")}` !== value) {
    throw new Error("native_execution_signature_invalid");
  }
  return raw;
}

export function verifyExecutionPermit(
  value: unknown,
  keyValue: unknown,
  target: { toolName: string; toolCallId: string; parametersDigest: string },
  nowMs = Date.now(),
): VerifiedExecutionPermit {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("native_execution_permit_invalid");
  const envelope = value as JsonObject;
  if (!exactKeys(envelope, ["format", "body", "body_sha256", "signature"]) || envelope.format !== ENVELOPE_FORMAT) {
    throw new Error("native_execution_permit_invalid");
  }
  if (!envelope.body || typeof envelope.body !== "object" || Array.isArray(envelope.body)) throw new Error("native_execution_permit_invalid");
  const body = envelope.body as JsonObject;
  const fields = [
    "format", "operation_id", "intent_digest", "adapter_id", "action", "effect_class",
    "tool_name", "tool_call_id", "parameters_digest", "authorization_binding_digest",
    "issued_at_ms", "expires_at_ms", "nonce",
  ];
  if (
    !exactKeys(body, fields) || body.format !== PERMIT_FORMAT ||
    typeof body.operation_id !== "string" || !body.operation_id ||
    typeof body.intent_digest !== "string" || !DIGEST.test(body.intent_digest) ||
    typeof body.adapter_id !== "string" || !body.adapter_id ||
    typeof body.action !== "string" || !body.action ||
    !new Set(["READ", "STATE_MUTATION", "MESSAGE_DELIVERY"]).has(String(body.effect_class)) ||
    typeof body.tool_name !== "string" || !TOOL.test(body.tool_name) ||
    typeof body.tool_call_id !== "string" || !CALL_ID.test(body.tool_call_id) ||
    typeof body.parameters_digest !== "string" || !DIGEST.test(body.parameters_digest) ||
    typeof body.authorization_binding_digest !== "string" || !DIGEST.test(body.authorization_binding_digest) ||
    !Number.isSafeInteger(body.issued_at_ms) || !Number.isSafeInteger(body.expires_at_ms) ||
    (body.issued_at_ms as number) < 0 || (body.expires_at_ms as number) <= (body.issued_at_ms as number) ||
    (body.expires_at_ms as number) - (body.issued_at_ms as number) > 30_000 ||
    nowMs < (body.issued_at_ms as number) || nowMs > (body.expires_at_ms as number) ||
    typeof body.nonce !== "string" || !/^[0-9a-f]{32}$/.test(body.nonce) ||
    body.tool_name !== target.toolName || body.tool_call_id !== target.toolCallId ||
    body.parameters_digest !== target.parametersDigest
  ) throw new Error("native_execution_permit_invalid");
  const encoded = canonicalQelJson(body);
  const digest = `sha256:${createHash("sha256").update(encoded).digest("hex")}`;
  if (envelope.body_sha256 !== digest) throw new Error("native_execution_permit_invalid");
  if (!envelope.signature || typeof envelope.signature !== "object" || Array.isArray(envelope.signature)) throw new Error("native_execution_permit_invalid");
  const signed = envelope.signature as JsonObject;
  if (!exactKeys(signed, ["alg", "value"]) || signed.alg !== "Ed25519" ||
      !verifySignature(null, encoded, publicKey(keyValue), signature(signed.value))) {
    throw new Error("native_execution_signature_invalid");
  }
  return body as unknown as VerifiedExecutionPermit;
}

function takePreparation(toolCallId: string): PreparationEntry {
  const registry = (globalThis as Record<symbol, unknown>)[REGISTRY_KEY];
  if (!(registry instanceof Map)) throw new Error("native_execution_preparation_missing");
  const entry = registry.get(toolCallId) as PreparationEntry | undefined;
  registry.delete(toolCallId);
  if (!entry || typeof entry.prepare !== "function") throw new Error("native_execution_preparation_missing");
  return entry;
}

async function unixRequest(path: string, expectedUid: number, value: JsonObject): Promise<JsonObject> {
  if (!isAbsolute(path) || !Number.isSafeInteger(expectedUid) || expectedUid < 0) throw new Error("native_execution_configuration_invalid");
  const metadata = await lstat(path).catch(() => undefined);
  if (!metadata?.isSocket() || metadata.uid !== expectedUid || (metadata.mode & 0o002) !== 0) throw new Error("native_execution_socket_invalid");
  const request = Buffer.concat([canonicalQelJson(value), Buffer.from("\n")]);
  if (request.byteLength > MAXIMUM_RESPONSE_BYTES) throw new Error("native_execution_request_too_large");
  return new Promise((resolve, reject) => {
    const connection = createConnection(path);
    const chunks: Buffer[] = [];
    let size = 0;
    let settled = false;
    const finish = (error?: Error, response?: JsonObject) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      connection.destroy();
      if (error) reject(error); else resolve(response!);
    };
    const timer = setTimeout(() => finish(new Error("native_execution_timeout")), 10_000);
    connection.on("connect", () => connection.end(request));
    connection.on("data", (chunk: Buffer) => {
      size += chunk.byteLength;
      if (size > MAXIMUM_RESPONSE_BYTES) return finish(new Error("native_execution_response_too_large"));
      chunks.push(chunk);
    });
    connection.on("error", () => finish(new Error("native_execution_unavailable")));
    connection.on("end", () => {
      try {
        const raw = Buffer.concat(chunks);
        if (!raw.length || raw[raw.length - 1] !== 10 || raw.subarray(0, -1).includes(10)) throw new Error();
        const parsed = JSON.parse(raw.subarray(0, -1).toString("utf8")) as unknown;
        if (!parsed || typeof parsed !== "object" || Array.isArray(parsed) || !canonicalQelJson(parsed).equals(raw.subarray(0, -1))) throw new Error();
        finish(undefined, parsed as JsonObject);
      } catch { finish(new Error("native_execution_response_invalid")); }
    });
  });
}

export async function consumeNativeExecutionPermit(
  config: NativeExecutionConfig,
  toolName: string,
  toolCallId: string,
  params: unknown,
  dependencies: { now?: () => number; request?: typeof unixRequest } = {},
): Promise<ConsumedExecutionPermit | undefined> {
  if (config.nativeExecutionRequired !== true) return undefined;
  const entry = takePreparation(toolCallId);
  const preparation = await entry.prepare(params);
  if (!preparation || typeof preparation !== "object" || Array.isArray(preparation)) throw new Error("native_execution_preparation_invalid");
  const prepared = preparation as JsonObject;
  if (
    prepared.format !== RESPONSE_FORMAT || prepared.status !== "PREPARED" ||
    typeof prepared.operation_id !== "string" || typeof prepared.intent_digest !== "string" ||
    !prepared.permit || typeof prepared.permit !== "object"
  ) throw new Error("native_execution_preparation_invalid");
  const parametersDigest = qelParametersDigest(params);
  const body = verifyExecutionPermit(
    prepared.permit, config.nativeExecutionPublicKey,
    { toolName, toolCallId, parametersDigest }, dependencies.now?.() ?? Date.now(),
  );
  if (body.operation_id !== prepared.operation_id || body.intent_digest !== prepared.intent_digest) {
    throw new Error("native_execution_preparation_invalid");
  }
  if (typeof config.nativeExecutionSocketPath !== "string") throw new Error("native_execution_configuration_invalid");
  const response = await (dependencies.request ?? unixRequest)(
    config.nativeExecutionSocketPath,
    config.nativeExecutionSocketOwnerUid ?? 0,
    {
      command: "consume", permit: prepared.permit, tool_name: toolName,
      tool_call_id: toolCallId, parameters_digest: parametersDigest,
    },
  );
  if (
    !exactKeys(response, ["format", "status", "operation_id", "intent_digest"]) ||
    response.format !== RESPONSE_FORMAT || response.status !== "CONSUMED" ||
    response.operation_id !== body.operation_id || response.intent_digest !== body.intent_digest
  ) throw new Error("native_execution_consumption_refused");
  return {
    permit: prepared.permit as JsonObject,
    operationId: body.operation_id,
    intentDigest: body.intent_digest,
  };
}

export async function completeNativeExecutionPermit(
  config: NativeExecutionConfig,
  consumed: ConsumedExecutionPermit | undefined,
  result: unknown,
  dependencies: { request?: typeof unixRequest } = {},
): Promise<JsonObject | undefined> {
  if (config.nativeExecutionRequired !== true) return undefined;
  if (!consumed || typeof config.nativeExecutionSocketPath !== "string") {
    throw new Error("native_execution_completion_invalid");
  }
  const response = await (dependencies.request ?? unixRequest)(
    config.nativeExecutionSocketPath,
    config.nativeExecutionSocketOwnerUid ?? 0,
    { command: "complete", permit: consumed.permit, result },
  );
  if (
    !exactKeys(response, ["format", "status", "operation_id", "intent_digest", "reason", "evidence"]) ||
    response.format !== RESPONSE_FORMAT || response.status !== "COMPLETED" ||
    response.operation_id !== consumed.operationId ||
    response.intent_digest !== consumed.intentDigest || response.reason !== null ||
    !response.evidence || typeof response.evidence !== "object" || Array.isArray(response.evidence)
  ) throw new Error("native_execution_completion_unverified");
  return response;
}
