import { createHash, generateKeyPairSync, sign } from "node:crypto";
import { afterEach, describe, expect, it, vi } from "vitest";

import {
  canonicalQelJson,
  completeNativeExecutionPermit,
  consumeNativeExecutionPermit,
  qelParametersDigest,
  verifyExecutionPermit,
} from "./native-execution-permit.js";

const NOW = 1_900_000_000_000;
const REGISTRY_KEY = Symbol.for("gemacode/native-execution-permits/1");
const { privateKey, publicKey } = generateKeyPairSync("ed25519");
const publicDer = publicKey.export({ format: "der", type: "spki" });
const PUBLIC_KEY = `b64u:${publicDer.subarray(-32).toString("base64url")}`;
const PARAMS = { action: "create_task", listId: "list-1", title: "Cita clínica" };

function permit(params: unknown = PARAMS, overrides: Record<string, unknown> = {}) {
  const body = {
    format: "native-connected-execution-permit/1",
    operation_id: "openclaw-operation-1",
    intent_digest: `sha256:${"ab".repeat(32)}`,
    adapter_id: "openclaw.microsoft-graph/1",
    action: "MICROSOFT_TODO_WRITE_CREATE_TASK",
    effect_class: "STATE_MUTATION",
    tool_name: "microsoft_todo_write",
    tool_call_id: "call/1",
    parameters_digest: qelParametersDigest(params),
    authorization_binding_digest: `sha256:${"cd".repeat(32)}`,
    issued_at_ms: NOW,
    expires_at_ms: NOW + 10_000,
    nonce: "12".repeat(16),
    ...overrides,
  };
  const encoded = canonicalQelJson(body);
  return {
    format: "native-connected-execution-permit-envelope/1",
    body,
    body_sha256: `sha256:${createHash("sha256").update(encoded).digest("hex")}`,
    signature: { alg: "Ed25519", value: `b64u:${sign(null, encoded, privateKey).toString("base64url")}` },
  };
}

afterEach(() => {
  delete (globalThis as Record<symbol, unknown>)[REGISTRY_KEY];
});

describe("Native OS external execution permits", () => {
  it("matches QEL canonical Unicode parameter digests", () => {
    expect(qelParametersDigest(PARAMS)).toBe(
      "sha256:b0294bb7a5836192970702557c4874595641ed737c89f456b559bf173d14145a",
    );
  });

  it("verifies an exact signed target and rejects drift or expiry", () => {
    const value = permit();
    expect(verifyExecutionPermit(value, PUBLIC_KEY, {
      toolName: "microsoft_todo_write",
      toolCallId: "call/1",
      parametersDigest: qelParametersDigest(PARAMS),
    }, NOW)).toMatchObject({ operation_id: "openclaw-operation-1" });
    expect(() => verifyExecutionPermit(value, PUBLIC_KEY, {
      toolName: "microsoft_todo_write",
      toolCallId: "call/2",
      parametersDigest: qelParametersDigest(PARAMS),
    }, NOW)).toThrow("native_execution_permit_invalid");
    expect(() => verifyExecutionPermit(value, PUBLIC_KEY, {
      toolName: "microsoft_todo_write",
      toolCallId: "call/1",
      parametersDigest: qelParametersDigest({ ...PARAMS, title: "Otra" }),
    }, NOW)).toThrow("native_execution_permit_invalid");
    expect(() => verifyExecutionPermit(value, PUBLIC_KEY, {
      toolName: "microsoft_todo_write",
      toolCallId: "call/1",
      parametersDigest: qelParametersDigest(PARAMS),
    }, NOW + 10_001)).toThrow("native_execution_permit_invalid");
  });

  it("prepares with the final params and consumes once before execution", async () => {
    const prepare = vi.fn(async (params: unknown) => ({
      format: "native-connected-external-execution-response/1",
      status: "PREPARED",
      operation_id: "openclaw-operation-1",
      intent_digest: `sha256:${"ab".repeat(32)}`,
      permit: permit(params),
    }));
    (globalThis as Record<symbol, unknown>)[REGISTRY_KEY] = new Map([
      ["call/1", { prepare, sessionKey: "session-1", createdAtMs: NOW }],
    ]);
    const request = vi.fn(async (_path: string, _uid: number, _value: Record<string, unknown>) => ({
      format: "native-connected-external-execution-response/1",
      status: "CONSUMED",
      operation_id: "openclaw-operation-1",
      intent_digest: `sha256:${"ab".repeat(32)}`,
    }));
    const consumed = await consumeNativeExecutionPermit({
      nativeExecutionRequired: true,
      nativeExecutionPublicKey: PUBLIC_KEY,
      nativeExecutionSocketPath: "/run/native/external.sock",
      nativeExecutionSocketOwnerUid: 0,
    }, "microsoft_todo_write", "call/1", PARAMS, { now: () => NOW, request });
    expect(prepare).toHaveBeenCalledWith(PARAMS);
    expect(request).toHaveBeenCalledWith(
      "/run/native/external.sock",
      0,
      expect.objectContaining({ command: "consume", parameters_digest: qelParametersDigest(PARAMS) }),
    );
    expect(consumed).toMatchObject({ operationId: "openclaw-operation-1" });
    const completionRequest = vi.fn(async () => ({
      format: "native-connected-external-execution-response/1",
      status: "COMPLETED",
      operation_id: "openclaw-operation-1",
      intent_digest: `sha256:${"ab".repeat(32)}`,
      reason: null,
      evidence: { format: "native-connected-action-evidence/1" },
    }));
    await expect(completeNativeExecutionPermit({
      nativeExecutionRequired: true,
      nativeExecutionSocketPath: "/run/native/external.sock",
      nativeExecutionSocketOwnerUid: 0,
    }, consumed, { ok: true }, { request: completionRequest })).resolves.toMatchObject({
      status: "COMPLETED",
    });
    expect(completionRequest).toHaveBeenCalledWith(
      "/run/native/external.sock", 0,
      expect.objectContaining({ command: "complete", result: { ok: true } }),
    );
    await expect(consumeNativeExecutionPermit({ nativeExecutionRequired: true },
      "microsoft_todo_write", "call/1", PARAMS)).rejects.toThrow("native_execution_preparation_missing");
  });

  it("does nothing when the production gate is disabled", async () => {
    await expect(consumeNativeExecutionPermit({}, "microsoft_todo_write", "missing", PARAMS)).resolves.toBeUndefined();
  });
});
