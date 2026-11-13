import { randomBytes } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { automaticServiceHealth } from "./service-health.js";
import { createVaultCredential, inspectVaultCredential, readVaultCredential } from "./credential-vault.js";
import { tokenForAuthorizedOperation } from "./credential.js";
import { graphRequest } from "./graph.js";
import type { GraphPolicy } from "./policy.js";

describe("R-001 diagnostic deadline and real rotating vault", () => {
  it.each(["deadline", "caller cancellation", "late endpoint rejection"])("preserves refresh and following data access after %s", async mode => {
    const stateDir = await mkdtemp(join(tmpdir(), "graph-health-vault-"));
    const config = { credentialVaultKey: randomBytes(32).toString("base64url") };
    const policy = {} as GraphPolicy;
    let entered!: () => void;
    const started = new Promise<void>(resolve => { entered = resolve; });
    let release!: () => void;
    const stalled = new Promise<void>(resolve => { release = resolve; });
    let refreshSignal: AbortSignal | undefined;
    const scopes: string[] = [];
    const refreshTokens: string[] = [];
    const fetchFn = vi.fn(async (url: string | URL | Request, init?: RequestInit) => {
      if (String(url).includes("/oauth2/v2.0/token")) {
        const body = new URLSearchParams(String(init?.body));
        scopes.push(body.get("scope")!); refreshTokens.push(body.get("refresh_token")!);
        if (scopes.length === 1) {
          refreshSignal = init?.signal as AbortSignal;
          entered();
          await new Promise<void>((resolve, reject) => {
            stalled.then(resolve);
            refreshSignal!.addEventListener("abort", () => reject(refreshSignal!.reason), { once: true });
          });
        }
        if (scopes.length === 1 && mode === "late endpoint rejection") return new Response(null, { status: 503 });
        return Response.json({ access_token: "synthetic-access", refresh_token: scopes.length === 1 ? "synthetic-rotated" : "synthetic-data-rotated", expires_in: 3600 });
      }
      return Response.json({ value: [] });
    }) as typeof fetch;
    try {
      await createVaultCredential(stateDir, { clientId: "synthetic-client", tenant: "synthetic-org", refreshToken: "synthetic-original", scopes: ["ServiceHealth.Read.All", "Mail.Read", "offline_access"] }, config.credentialVaultKey);
      const controller = new AbortController();
      const health = automaticServiceHealth({ config, policy, stateDir, service: "mail", budgetMs: 200, signal: controller.signal, fetchFn });
      await started;
      if (mode === "caller cancellation") {
        controller.abort(new Error("synthetic-cancel"));
        await expect(health).rejects.toThrow("synthetic-cancel");
        await new Promise(resolve => setTimeout(resolve, 250));
      } else {
        expect(await health).toMatchObject({ status: "unavailable", reason: "diagnostic_timeout" });
        await new Promise(resolve => setTimeout(resolve, 50));
      }
      // Return the late rotated pair after the caller has stopped waiting.
      release();
      const token = await tokenForAuthorizedOperation({ config, policy, stateDir, allowedScopes: ["Mail.Read"], fetchFn });
      expect(refreshSignal?.aborted).toBe(false);
      expect(refreshTokens).toEqual(["synthetic-original", mode === "late endpoint rejection" ? "synthetic-original" : "synthetic-rotated"]);
      expect(scopes[0]).toContain("ServiceHealth.Read.All");
      expect(scopes[1]).toContain("Mail.Read");
      expect(scopes[1]).not.toContain("ServiceHealth.Read.All");
      expect(await inspectVaultCredential(stateDir, config.credentialVaultKey)).toMatchObject({ result: "valid", generation: mode === "late endpoint rejection" ? 2 : 3 });
      expect((await readVaultCredential(stateDir, config.credentialVaultKey)).credential.refreshToken).toBe("synthetic-data-rotated");
      expect(await graphRequest(token, "/me/messages", {}, fetchFn)).toEqual({ value: [] });
      expect(vi.mocked(fetchFn).mock.calls.filter(([url]) => String(url).includes("serviceAnnouncement"))).toHaveLength(0);
    } finally { release(); await rm(stateDir, { recursive: true, force: true }); }
  });
});
