import { describe, expect, it, vi } from "vitest";
import { resolveAllowedRoots } from "./resolve-root-pins.mjs";

describe("root pin resolution", () => {
  it("authenticates Graph requests with the acquired token without writing it", async () => {
    const acquiredToken = "synthetic-access-token";
    const readCredentialFn = vi.fn(async () => ({ clientId: "synthetic-client" }));
    const exchangeRefreshTokenFn = vi.fn(async () => acquiredToken);
    const fetchFn = vi.fn(async (_url, init) => {
      expect(init.headers.authorization).toBe(`Bearer ${acquiredToken}`);
      return new Response(JSON.stringify({ id: "synthetic-root-id", name: "Documents" }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    });
    const write = vi.fn();

    await resolveAllowedRoots({
      policy: {
        services: {
          onedrive: {
            allowed_roots: [
              { label: "documents", drive_id: "synthetic-drive", path: "/Documents" },
              { label: "pinned", drive_id: "synthetic-drive", path: "/Pinned", item_id: "already-pinned" },
            ],
          },
        },
      },
      credentialRef: "synthetic/shared",
      readCredentialFn,
      exchangeRefreshTokenFn,
      fetchFn,
      write,
    });

    expect(readCredentialFn).toHaveBeenCalledWith("synthetic/shared");
    expect(exchangeRefreshTokenFn).toHaveBeenCalledWith({ clientId: "synthetic-client" }, ["Files.Read"]);
    expect(fetchFn).toHaveBeenCalledTimes(1);
    expect(write).toHaveBeenCalledWith(`${JSON.stringify({ label: "documents", item_id: "synthetic-root-id", resolved_name: "Documents" })}\n`);
    expect(write.mock.calls.flat().join(" ")).not.toContain(acquiredToken);
  });
});
