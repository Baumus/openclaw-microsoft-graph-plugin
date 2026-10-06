import { describe, expect, it, vi } from "vitest";
import { newerRelease, registerUpdateStatusMethod } from "./update-status.js";

const listing = (version: string, scanStatus = "clean") => ({ package: { name: "@baumus/openclaw-microsoft-graph", latestVersion: version, scanStatus } });

describe("ClawHub update status", () => {
  it("shows only a newer clean stable release of this package", () => {
    expect(newerRelease("3.9.0", listing("3.10.0"))).toEqual({ currentVersion: "3.9.0", latestVersion: "3.10.0", updateAvailable: true });
    expect(newerRelease("3.10.0", listing("3.10.0"))?.updateAvailable).toBe(false);
    expect(newerRelease("3.10.0", listing("3.9.9"))?.updateAvailable).toBe(false);
    expect(newerRelease("3.9.0", listing("3.10.0", "pending"))).toBeUndefined();
    expect(newerRelease("3.9.0", listing("3.10.0-beta.1"))).toBeUndefined();
    expect(newerRelease("3.9.0", { package: { ...listing("3.10.0").package, name: "@other/plugin" } })).toBeUndefined();
  });

  it("uses an admin-only, read-only fixed endpoint and caches the result", async () => {
    let handler: ((arg: { params: unknown; respond: (ok: boolean, payload?: unknown, error?: unknown) => void }) => void | Promise<void>) | undefined;
    const api = { version: "3.9.0", registerGatewayMethod: vi.fn((_name, fn) => { handler = fn; }) };
    const fetcher = vi.fn(async (_url: string, _options?: RequestInit) => Response.json(listing("3.10.0")));
    registerUpdateStatusMethod(api, fetcher as unknown as typeof fetch);
    expect(api.registerGatewayMethod).toHaveBeenCalledWith("microsoft-graph.updateStatus", expect.any(Function), { scope: "operator.admin" });
    const respond = vi.fn();
    await handler!({ params: {}, respond });
    await handler!({ params: {}, respond });
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(fetcher.mock.calls[0][0]).toBe("https://clawhub.ai/api/v1/packages/%40baumus%2Fopenclaw-microsoft-graph");
    expect(respond).toHaveBeenLastCalledWith(true, { currentVersion: "3.9.0", latestVersion: "3.10.0", updateAvailable: true });
    respond.mockClear();
    await handler!({ params: { url: "https://example.org" }, respond });
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(respond).toHaveBeenCalledWith(false, undefined, expect.objectContaining({ code: "INVALID_REQUEST" }));
  });

  it("fails quietly when the registry is unavailable or returns oversized data", async () => {
    for (const fetcher of [vi.fn(async () => { throw new Error("offline"); }), vi.fn(async () => new Response("x".repeat(20_000)))]) {
      let handler: ((arg: { params: unknown; respond: (ok: boolean, payload?: unknown, error?: unknown) => void }) => void | Promise<void>) | undefined;
      registerUpdateStatusMethod({ version: "3.9.0", registerGatewayMethod: (_name, fn) => { handler = fn; } }, fetcher as unknown as typeof fetch);
      const respond = vi.fn();
      await handler!({ params: {}, respond });
      expect(respond).toHaveBeenCalledWith(true, { updateAvailable: false });
    }
  });
});
