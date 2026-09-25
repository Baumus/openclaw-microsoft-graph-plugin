import { afterEach, describe, expect, it, vi } from "vitest";
import { randomBytes } from "node:crypto";
import { lstat, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { accessTokenCacheStateForTests, assertScopes, clearAccessTokenCacheForTests, exchangeRefreshToken, exchangeRefreshTokenDetailed, parseCredential, readCredential, selectScope, tokenForAuthorizedOperation } from "./credential.js";
import { createVaultCredential, inspectVaultCredential, vaultQuarantinePath } from "./credential-vault.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

const temporary: string[] = [];
const vaultCredential = { clientId: "vault-client", refreshToken: "vault-refresh", tenant: "tenant", scopes: ["Files.Read", "offline_access"] };
function vaultPolicy() { const current = graphPolicyFixture(); return { version: 2 as const, rules: current.rules, services: current.services }; }
function vaultKey(): string { return randomBytes(32).toString("base64url"); }
async function temporaryState(): Promise<string> { const path = await mkdtemp(join(tmpdir(), "credential-cache-test-")); temporary.push(path); return path; }

afterEach(async () => {
  clearAccessTokenCacheForTests();
  vi.restoreAllMocks();
  vi.useRealTimers();
  await Promise.all(temporary.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});

describe("Gateway credential handling", () => {
  it("requires declared scopes and never invents service scopes", () => {
    const credential = parseCredential(JSON.stringify({ clientId: "client", refreshToken: "refresh", scopes: ["Files.ReadWrite"] }));
    expect(credential.scopes).toEqual(["Files.ReadWrite", "offline_access"]);
    expect(() => assertScopes(credential, ["Files.ReadWrite"])).not.toThrow();
    expect(() => assertScopes(credential, ["Mail.Read"])).toThrow("credential_scope_missing");
  });

  it("selects and requests only the operation-specific scope", async () => {
    const credential = parseCredential(JSON.stringify({ clientId: "client", refreshToken: "refresh", scopes: ["Files.ReadWrite"] }));
    expect(selectScope(credential, ["Files.Read"])).toBe("Files.Read");
    let requested = "";
    const fetchFn = async (_url: string | URL | Request, init?: RequestInit) => {
      requested = String(init?.body);
      return new Response(JSON.stringify({ access_token: "bounded-token", expires_in: 3600 }), { status: 200, headers: { "content-type": "application/json" } });
    };
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).resolves.toBe("bounded-token");
    expect(new URLSearchParams(requested).get("scope")).toBe("Files.Read");
  });

  it("rejects oversized OAuth responses", async () => {
    const credential = parseCredential(JSON.stringify({ clientId: "client", refreshToken: "refresh", scopes: ["Files.Read"] }));
    const fetchFn = async () => new Response("x", { status: 200, headers: { "content-length": "65537" } });
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).rejects.toThrow("authentication_failed");
  });

  it("returns a bounded replacement refresh token only to the durable vault caller", async () => {
    const credential = parseCredential(JSON.stringify({ clientId: "client", refreshToken: "synthetic-old", scopes: ["Files.Read"] }));
    const fetchFn = async () => new Response(JSON.stringify({ access_token: "synthetic-access", refresh_token: "synthetic-replacement", expires_in: 120 }), { status: 200 });
    await expect(exchangeRefreshTokenDetailed(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).resolves.toMatchObject({
      accessToken: "synthetic-access",
      replacementRefreshToken: "synthetic-replacement",
    });
    clearAccessTokenCacheForTests();
    const invalid = async () => new Response(JSON.stringify({ access_token: "synthetic-access", refresh_token: "" }), { status: 200 });
    await expect(exchangeRefreshTokenDetailed(credential, ["Files.Read"], undefined, invalid as typeof fetch)).rejects.toThrow("credential_reauthorization_required");
  });

  it.each([
    ["whitespace", "token value"],
    ["control", "token\nvalue"],
    ["Unicode", "tökén"],
    ["header separator", "token:value"],
    ["oversized", "a".repeat(16 * 1024 + 1)],
  ])("rejects an HTTP-200 access token with %s", async (_name, accessToken) => {
    const credential = parseCredential(JSON.stringify({ clientId: "client", refreshToken: "refresh", scopes: ["Files.Read"] }));
    const fetchFn = async () => new Response(JSON.stringify({ access_token: accessToken, expires_in: 3600 }), { status: 200 });
    await expect(exchangeRefreshTokenDetailed(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).rejects.toThrow("authentication_failed");
  });

  it.each([
    ["missing", undefined],
    ["string", "3600"],
    ["zero", 0],
    ["negative", -1],
    ["fractional", 1.5],
    ["unsafe", Number.MAX_SAFE_INTEGER + 1],
    ["over one day", 86_401],
  ])("requires a bounded positive integer expires_in: %s", async (_name, expiresIn) => {
    const credential = parseCredential(JSON.stringify({ clientId: "client", refreshToken: "refresh", scopes: ["Files.Read"] }));
    const payload = { access_token: "synthetic-access", ...(expiresIn === undefined ? {} : { expires_in: expiresIn }) };
    const fetchFn = async () => new Response(JSON.stringify(payload), { status: 200 });
    await expect(exchangeRefreshTokenDetailed(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).rejects.toThrow("authentication_failed");
  });

  it("reuses a scoped access token until its pre-expiry boundary", async () => {
    let now = 1_000;
    vi.spyOn(Date, "now").mockImplementation(() => now);
    const credential = parseCredential(JSON.stringify({ clientId: "cache-client", refreshToken: "cache-refresh", tenant: "tenant", scopes: ["Files.Read"] }));
    const fetchFn = vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({ access_token: "first-token", expires_in: 120 }), { status: 200 }))
      .mockResolvedValueOnce(new Response(JSON.stringify({ access_token: "second-token", expires_in: 120 }), { status: 200 }));

    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).resolves.toBe("first-token");
    now = 60_999;
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).resolves.toBe("first-token");
    expect(fetchFn).toHaveBeenCalledTimes(1);

    now = 61_000;
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).resolves.toBe("second-token");
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it("caps process-local cache residency at one hour while applying expiry skew", async () => {
    let now = 1_000;
    vi.spyOn(Date, "now").mockImplementation(() => now);
    const credential = parseCredential(JSON.stringify({ clientId: "cache-cap-client", refreshToken: "cache-cap-refresh", tenant: "tenant", scopes: ["Files.Read"] }));
    const fetchFn = vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({ access_token: "first-token", expires_in: 86_400 }), { status: 200 }))
      .mockResolvedValueOnce(new Response(JSON.stringify({ access_token: "second-token", expires_in: 86_400 }), { status: 200 }));

    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).resolves.toBe("first-token");
    now += 3_539_999;
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).resolves.toBe("first-token");
    now += 1;
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).resolves.toBe("second-token");
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it("bounds the access-token cache with deterministic LRU eviction and active expiry cleanup", async () => {
    let now = 1_000;
    vi.spyOn(Date, "now").mockImplementation(() => now);
    const fetchFn = vi.fn(async () => new Response(JSON.stringify({ access_token: "synthetic-token", expires_in: 3600 }), { status: 200 }));
    for (let index = 0; index < 140; index += 1) {
      const credential = parseCredential(JSON.stringify({ clientId: `client-${index}`, refreshToken: `refresh-${index}`, scopes: ["Files.Read"] }));
      await exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch);
    }
    expect(accessTokenCacheStateForTests()).toEqual({ entries: 128, refreshes: 0 });
    now += 4_000_000;
    expect(accessTokenCacheStateForTests()).toEqual({ entries: 0, refreshes: 0 });
  });

  it("admits vault tokens only after durable rotation and removes the predecessor identity", async () => {
    const state = await temporaryState();
    const read = vaultKey();
    await createVaultCredential(state, vaultCredential, read);
    const config = { credentialVaultKey: read };
    const fetchFn = vi.fn(async () => new Response(JSON.stringify({ access_token: "durable-access", refresh_token: "rotated-refresh", expires_in: 3600 }), { status: 200 }));
    const params = { config, policy: vaultPolicy(), allowedScopes: ["Files.Read"], stateDir: state, requestTimeoutMs: 1000, fetchFn: fetchFn as typeof fetch };
    await expect(tokenForAuthorizedOperation(params)).resolves.toBe("durable-access");
    expect(accessTokenCacheStateForTests().entries).toBe(1);
    await expect(tokenForAuthorizedOperation(params)).resolves.toBe("durable-access");
    expect(fetchFn).toHaveBeenCalledTimes(1);
    expect(accessTokenCacheStateForTests().entries).toBe(1);
  });

  it("materializes only the shared vault key at runtime", async () => {
    const state = await temporaryState();
    const read = vaultKey();
    await createVaultCredential(state, vaultCredential, read);
    const fetchFn = vi.fn(async () => new Response(JSON.stringify({ access_token: "shared-credential", expires_in: 3600 }), { status: 200 }));
    await expect(tokenForAuthorizedOperation({
      config: { credentialVaultKey: read },
      policy: vaultPolicy(), allowedScopes: ["Files.Read"], stateDir: state, requestTimeoutMs: 1000, fetchFn: fetchFn as typeof fetch,
    })).resolves.toBe("shared-credential");
  });

  it("keeps one cache identity across repeated durable vault rotations", async () => {
    let now = Date.now();
    vi.spyOn(Date, "now").mockImplementation(() => now);
    const state = await temporaryState();
    const read = vaultKey();
    await createVaultCredential(state, vaultCredential, read);
    let rotation = 0;
    const fetchFn = vi.fn(async () => {
      rotation += 1;
      return new Response(JSON.stringify({ access_token: `access-${rotation}`, refresh_token: `refresh-${rotation}`, expires_in: 61 }), { status: 200 });
    });
    const params = { config: { credentialVaultKey: read }, policy: vaultPolicy(), allowedScopes: ["Files.Read"], stateDir: state, requestTimeoutMs: 1000, fetchFn: fetchFn as typeof fetch };
    for (let index = 1; index <= 5; index += 1) {
      await expect(tokenForAuthorizedOperation(params)).resolves.toBe(`access-${index}`);
      expect(accessTokenCacheStateForTests().entries).toBe(1);
      now += 1_001;
    }
    expect(fetchFn).toHaveBeenCalledTimes(5);
  });

  it("does not reuse a provisional vault token after publication failure", async () => {
    const state = await temporaryState();
    const read = vaultKey();
    await createVaultCredential(state, vaultCredential, read);
    const fetchFn = vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({ access_token: "must-not-cache", refresh_token: "lost-rotation", expires_in: 3600 }), { status: 200 }))
      .mockResolvedValueOnce(new Response(null, { status: 401 }));
    const base = { config: { credentialVaultKey: read }, policy: vaultPolicy(), allowedScopes: ["Files.Read"], stateDir: state, requestTimeoutMs: 1000, fetchFn: fetchFn as typeof fetch };
    await expect(tokenForAuthorizedOperation({ ...base, vaultTestHooks: { beforePublication: () => { throw new Error("synthetic_publish_failure"); } } })).rejects.toThrow("credential_refresh_outcome_uncertain");
    expect(accessTokenCacheStateForTests().entries).toBe(0);
    await expect(tokenForAuthorizedOperation(base)).rejects.toThrow("credential_reauthorization_required");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("quarantines an ambiguous dispatched vault refresh and blocks retries across calls", async () => {
    const state = await temporaryState();
    const read = vaultKey();
    await createVaultCredential(state, vaultCredential, read);
    const fetchFn = vi.fn(async () => { throw new Error("synthetic_lost_response"); });
    const params = { config: { credentialVaultKey: read }, policy: vaultPolicy(), allowedScopes: ["Files.Read"], stateDir: state, requestTimeoutMs: 1000, fetchFn: fetchFn as typeof fetch };
    await expect(tokenForAuthorizedOperation(params)).rejects.toThrow("credential_refresh_outcome_uncertain");
    await expect(tokenForAuthorizedOperation(params)).rejects.toThrow("credential_reauthorization_required");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it.each([
    ["malformed JSON", () => new Response("not-json", { status: 200 })],
    ["truncated JSON", () => new Response('{"access_token":', { status: 200 })],
    ["invalid UTF-8", () => new Response(Uint8Array.from([0xff]), { status: 200 })],
    ["oversized body", () => new Response("x".repeat(64 * 1024 + 1), { status: 200 })],
    ["empty body", () => new Response("", { status: 200 })],
    ["missing body", () => new Response(null, { status: 200 })],
    ["non-object body", () => new Response("[]", { status: 200 })],
    ["missing access token", () => new Response("{}", { status: 200 })],
    ["empty access token", () => new Response(JSON.stringify({ access_token: "" }), { status: 200 })],
    ["invalid access token", () => new Response(JSON.stringify({ access_token: 7 }), { status: 200 })],
    ["whitespace access token", () => new Response(JSON.stringify({ access_token: "unsafe token", expires_in: 3600 }), { status: 200 })],
    ["control-character access token", () => new Response(JSON.stringify({ access_token: "unsafe\ntoken", expires_in: 3600 }), { status: 200 })],
    ["Unicode access token", () => new Response(JSON.stringify({ access_token: "tökén", expires_in: 3600 }), { status: 200 })],
    ["header-separator access token", () => new Response(JSON.stringify({ access_token: "unsafe:token", expires_in: 3600 }), { status: 200 })],
    ["oversized access token", () => new Response(JSON.stringify({ access_token: "a".repeat(16 * 1024 + 1), expires_in: 3600 }), { status: 200 })],
    ["missing expiry", () => new Response(JSON.stringify({ access_token: "synthetic-access" }), { status: 200 })],
    ["invalid expiry", () => new Response(JSON.stringify({ access_token: "synthetic-access", expires_in: 86_401 }), { status: 200 })],
    ["invalid replacement refresh token", () => new Response(JSON.stringify({ access_token: "synthetic-access", refresh_token: "" }), { status: 200 })],
  ])("durably quarantines an unusable HTTP-200 OAuth response: %s", async (_name, response) => {
    const state = await temporaryState();
    const read = vaultKey();
    await createVaultCredential(state, vaultCredential, read);
    const fetchFn = vi.fn(async () => response());
    const params = { config: { credentialVaultKey: read }, policy: vaultPolicy(), allowedScopes: ["Files.Read"], stateDir: state, requestTimeoutMs: 1000, fetchFn: fetchFn as typeof fetch };

    await expect(tokenForAuthorizedOperation(params)).rejects.toThrow("credential_refresh_outcome_uncertain");
    expect((await lstat(vaultQuarantinePath(state))).mode & 0o777).toBe(0o600);
    expect(await inspectVaultCredential(state, read)).toMatchObject({ result: "quarantined" });
    await expect(tokenForAuthorizedOperation(params)).rejects.toThrow("credential_reauthorization_required");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it.each([400, 401])("keeps definitive HTTP %i vault responses as authentication failures without quarantine", async (status) => {
    const state = await temporaryState();
    const read = vaultKey();
    await createVaultCredential(state, vaultCredential, read);
    const fetchFn = vi.fn(async () => new Response(null, { status }));
    const params = { config: { credentialVaultKey: read }, policy: vaultPolicy(), allowedScopes: ["Files.Read"], stateDir: state, requestTimeoutMs: 1000, fetchFn: fetchFn as typeof fetch };

    await expect(tokenForAuthorizedOperation(params)).rejects.toThrow("authentication_failed");
    await expect(tokenForAuthorizedOperation(params)).rejects.toThrow("authentication_failed");
    expect(fetchFn).toHaveBeenCalledTimes(2);
    await expect(lstat(vaultQuarantinePath(state))).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("quarantines an HTTP 503 vault response and blocks retry before OAuth dispatch", async () => {
    const state = await temporaryState();
    const read = vaultKey();
    await createVaultCredential(state, vaultCredential, read);
    const fetchFn = vi.fn(async () => new Response(null, { status: 503 }));
    const params = { config: { credentialVaultKey: read }, policy: vaultPolicy(), allowedScopes: ["Files.Read"], stateDir: state, requestTimeoutMs: 1000, fetchFn: fetchFn as typeof fetch };

    await expect(tokenForAuthorizedOperation(params)).rejects.toThrow("credential_refresh_outcome_uncertain");
    expect(await inspectVaultCredential(state, read)).toMatchObject({ result: "quarantined" });
    await expect(tokenForAuthorizedOperation(params)).rejects.toThrow("credential_reauthorization_required");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("bounds OAuth exchange by request timeout and cleans up its timer", async () => {
    vi.useFakeTimers();
    const credential = parseCredential(JSON.stringify({ clientId: "timeout-client", refreshToken: "timeout-refresh", scopes: ["Files.Read"] }));
    const fetchFn = vi.fn((_url: string | URL | Request, init?: RequestInit) => new Promise<Response>((_resolve, reject) => {
      init?.signal?.addEventListener("abort", () => reject(init.signal?.reason), { once: true });
    }));
    const pending = exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch, { requestTimeoutMs: 25 });
    const rejected = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(25);
    await rejected;
    expect(vi.getTimerCount()).toBe(0);
  });

  it("coalesces concurrent refreshes without caching failures", async () => {
    const credential = parseCredential(JSON.stringify({ clientId: "concurrent-client", refreshToken: "concurrent-refresh", scopes: ["Files.Read"] }));
    let release!: (response: Response) => void;
    const fetchFn = vi.fn(() => new Promise<Response>((resolve) => { release = resolve; }));
    const first = exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch);
    const second = exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch);
    expect(fetchFn).toHaveBeenCalledTimes(1);
    release(new Response(JSON.stringify({ access_token: "shared-token", expires_in: 3600 }), { status: 200 }));
    await expect(Promise.all([first, second])).resolves.toEqual(["shared-token", "shared-token"]);

    clearAccessTokenCacheForTests();
    const failingFetch = vi.fn()
      .mockResolvedValueOnce(new Response(null, { status: 503 }))
      .mockResolvedValueOnce(new Response(JSON.stringify({ access_token: "recovered-token", expires_in: 3600 }), { status: 200 }));
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, failingFetch as typeof fetch)).rejects.toThrow("authentication_failed");
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, failingFetch as typeof fetch)).resolves.toBe("recovered-token");
    expect(failingFetch).toHaveBeenCalledTimes(2);
  });

  it("keeps a shared token refresh alive when one concurrent caller cancels", async () => {
    const credential = parseCredential(JSON.stringify({ clientId: "cancel-client", refreshToken: "cancel-refresh", scopes: ["Files.Read"] }));
    let release!: (response: Response) => void;
    const fetchFn = vi.fn((_url: string | URL | Request, init?: RequestInit) => new Promise<Response>((resolve, reject) => {
      release = resolve;
      init?.signal?.addEventListener("abort", () => reject(init.signal?.reason), { once: true });
    }));
    const cancelled = new AbortController();
    const live = new AbortController();
    const first = exchangeRefreshToken(credential, ["Files.Read"], cancelled.signal, fetchFn as typeof fetch);
    const second = exchangeRefreshToken(credential, ["Files.Read"], live.signal, fetchFn as typeof fetch);
    const reason = new Error("caller_cancelled");
    cancelled.abort(reason);
    await expect(first).rejects.toBe(reason);
    release(new Response(JSON.stringify({ access_token: "surviving-token", expires_in: 3600 }), { status: 200 }));
    await expect(second).resolves.toBe("surviving-token");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("separates tokens by requested scope and never retries OAuth", async () => {
    const credential = parseCredential(JSON.stringify({ clientId: "scope-client", refreshToken: "scope-refresh", scopes: ["Files.Read", "Mail.Read"] }));
    const fetchFn = vi.fn(async (_url: string | URL | Request, init?: RequestInit) => {
      const scope = new URLSearchParams(String(init?.body)).get("scope");
      return new Response(JSON.stringify({ access_token: `${scope}-token`, expires_in: 3600 }), { status: 200 });
    });
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, fetchFn as typeof fetch)).resolves.toBe("Files.Read-token");
    await expect(exchangeRefreshToken(credential, ["Mail.Read"], undefined, fetchFn as typeof fetch)).resolves.toBe("Mail.Read-token");
    expect(fetchFn).toHaveBeenCalledTimes(2);

    clearAccessTokenCacheForTests();
    const denied = vi.fn(async () => new Response(null, { status: 503 }));
    await expect(exchangeRefreshToken(credential, ["Files.Read"], undefined, denied as typeof fetch)).rejects.toThrow("authentication_failed");
    expect(denied).toHaveBeenCalledTimes(1);
  });

  it("preserves caller cancellation before credential access", async () => {
    const controller = new AbortController();
    const reason = new Error("caller_cancelled");
    controller.abort(reason);
    await expect(readCredential("workspace/microsoft/test", controller.signal)).rejects.toBe(reason);
  });

});
