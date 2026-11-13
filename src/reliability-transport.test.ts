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
