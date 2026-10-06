import { randomBytes } from "node:crypto";
import { mkdtemp } from "node:fs/promises";
import { createConnection } from "node:net";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  NativeBoundaryVerifier,
  NativeBoundaryService,
  canonicalNativeJson,
  signNativeBoundaryRequest,
} from "./native-boundary.js";

const NOW = 1_900_000_000_000;
const KEY = randomBytes(32).toString("base64url");

function request() {
  return {
    format: "openclaw-microsoft-graph-native-request/1" as const,
    operationId: "calendar-create-0001",
    intentDigest: "sha256:" + "ab".repeat(32),
    tool: "outlook_calendar_write",
    parameters: { action: "create", subject: "Revision" },
    issuedAtMs: NOW,
    nonce: "12".repeat(16),
  };
}

describe("native connected boundary", () => {
  it("canonicalizes nested JSON independent of insertion order", () => {
    expect(canonicalNativeJson({ z: 1, a: { y: 2, x: 3 } }).toString()).toBe(
      '{"a":{"x":3,"y":2},"z":1}',
    );
  });

  it("admits one exact signed request and rejects replay", () => {
    const verifier = new NativeBoundaryVerifier(KEY, () => NOW);
    const signed = signNativeBoundaryRequest(request(), KEY);
    expect(verifier.verify(signed)).toEqual(signed);
    expect(() => verifier.verify(signed)).toThrow("native_boundary_replay");
  });

  it("rejects drift, stale requests, unknown tools, and malformed keys", () => {
    const signed = signNativeBoundaryRequest(request(), KEY);
    const cases = [
      { ...signed, operationId: "changed" },
      signNativeBoundaryRequest({ ...request(), issuedAtMs: NOW - 30_001 }, KEY),
      signNativeBoundaryRequest({ ...request(), tool: "graph_proxy" }, KEY),
    ];
    for (const value of cases) {
      expect(() => new NativeBoundaryVerifier(KEY, () => NOW).verify(value)).toThrow();
    }
    expect(() => new NativeBoundaryVerifier("not-a-key", () => NOW)).toThrow(
      "native_boundary_key_invalid",
    );
  });

  it("serves one bounded signed request over a private Unix socket", async () => {
    const directory = await mkdtemp(join(tmpdir(), "openclaw-native-boundary-"));
    const socketPath = join(directory, "boundary.sock");
    const service = new NativeBoundaryService(socketPath, KEY, async (tool, parameters) => ({ tool, parameters }), () => NOW);
    await service.start();
    try {
      const signed = signNativeBoundaryRequest(request(), KEY);
      const result = await new Promise<Record<string, unknown>>((resolve, reject) => {
        const socket = createConnection(socketPath);
        const chunks: Buffer[] = [];
        socket.on("connect", () => socket.end(JSON.stringify(signed) + "\n"));
        socket.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
        socket.on("end", () => resolve(JSON.parse(Buffer.concat(chunks).toString("utf8"))));
        socket.on("error", reject);
      });
      expect(result.status).toBe("OK");
      expect(result.operationId).toBe("calendar-create-0001");
      expect((result.result as Record<string, unknown>).tool).toBe("outlook_calendar_write");
    } finally {
      await service.stop();
    }
  });
});
