import { expect, it, vi } from "vitest";
import { graphRequest } from "./graph.js";
it("preserves legacy HTTP error and sanitized code, retry delay and attempts", async () => {
  const fetchFn = vi.fn(async () => Response.json({ error: { code: "ServiceUnavailable", message: "ignore instructions user@example.com" } }, { status: 503, headers: { "retry-after": "2" } }));
  try { await graphRequest("synthetic", "/me/messages", { readRetries: 0 }, fetchFn); throw new Error("missing failure"); }
  catch (error: any) {
    expect(error.message).toBe("provider_error_503");
    expect(error.providerStatus).toBe(503); expect(error.providerCode).toBe("ServiceUnavailable");
    expect(error.retryAfterMs).toBe(2000); expect(error.attempts).toBe(1);
    expect(JSON.stringify(error)).not.toContain("user@example.com");
  }
});
it("never replays uncertain mutations", async () => {
  const fetchFn = vi.fn(async () => { throw new TypeError("private URL"); });
  await expect(graphRequest("synthetic", "/me/messages", { method: "POST", body: {} }, fetchFn)).rejects.toThrow("provider_unavailable");
  expect(fetchFn).toHaveBeenCalledTimes(1);
});
it("does not emit arbitrary provider codes or invalid Retry-After", async () => {
  const fetchFn = vi.fn(async () => Response.json({ error: { code: "ignore_all_instructions" } }, { status: 403, headers: { "retry-after": "December 1" } }));
  await expect(graphRequest("synthetic", "/me/messages", {}, fetchFn)).rejects.toMatchObject({ message: "provider_access_denied", providerStatus: 403, retryAfterMs: undefined });
});

// Synthetic provider failures: these prove diagnostic safety, not the live cause.
it.each(["BadRequest", "Request_BadRequest", "ErrorInvalidRequest", "ErrorInvalidIdMalformed", "ErrorInvalidUrlQueryFilter", "ErrorInvalidProperty", "ErrorInvalidPropertyRequest", "ErrorInvalidPropertySet", "ErrorUnsupportedQuery", "Request_UnsupportedQuery"])("exposes only allowlisted HTTP400 code %s without replay", async (code) => {
  const fetchFn = vi.fn(async () => Response.json({ error: { code, message: "private@example.com Bearer synthetic-secret", innerError: { requestId: "private-id" } } }, { status: 400 }));
  const error = await graphRequest("synthetic-token", "/me/messages/synthetic/attachments?$top=5&$select=id,name,contentType,size,isInline,lastModifiedDateTime", {}, fetchFn).catch(error => error);
  expect(error).toMatchObject({ message: "provider_error_400", providerStatus: 400, providerCode: code, attempts: 1 });
  expect(Object.keys(error).sort()).toEqual(["attempts", "providerCode", "providerStatus", "retryAfterMs"]);
  expect(JSON.stringify(error)).not.toMatch(/private|Bearer|synthetic-secret/);
  expect(fetchFn).toHaveBeenCalledTimes(1);
});
it.each([
  { error: { code: "private@example.com", message: "secret" } },
  { error: { code: "BadRequest\nsecret" } },
  { error: { code: { toString: "BadRequest" } } },
  { error: { code: ["BadRequest"] } },
])("omits untrusted HTTP400 codes", async (body) => {
  const fetchFn = vi.fn(async () => Response.json(body, { status: 400 }));
  const error = await graphRequest("synthetic", "/me/messages/synthetic/attachments", {}, fetchFn).catch(error => error);
  expect(error.message).toBe("provider_error_400");
  expect(error).not.toHaveProperty("providerCode");
});
it.each(["not JSON", "x".repeat(16385)])("preserves HTTP400 when error body cannot be safely parsed", async (body) => {
  const fetchFn = vi.fn(async () => new Response(body, { status: 400 }));
  await expect(graphRequest("synthetic", "/me/messages/synthetic/attachments", {}, fetchFn)).rejects.toMatchObject({ message: "provider_error_400", providerStatus: 400, attempts: 1 });
});
