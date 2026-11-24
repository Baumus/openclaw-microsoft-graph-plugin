import { randomBytes } from "node:crypto";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { DeviceCodeSignIn } from "./device-code.js";
import { createVaultCredential, inspectVaultCredential, readVaultCredential, refreshVaultCredential, vaultCredentialPath, vaultQuarantinePath } from "./credential-vault.js";
import { requiredPolicyScopes } from "./credential-cli.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

const dirs: string[] = [];
const clientId = "11111111-1111-4111-8111-111111111111";
const tenant = "22222222-2222-4222-8222-222222222222";
async function setup() {
  const state = await mkdtemp(join(tmpdir(), "graph-device-code-")); dirs.push(state);
  const key = randomBytes(32).toString("base64url");
  return { state, key, config: { policy: graphPolicyFixture(), credentialVaultKey: key } };
}
afterEach(async () => { await Promise.all(dirs.splice(0).map((dir) => rm(dir, { recursive: true, force: true }))); vi.restoreAllMocks(); });
const deviceResponse = { device_code: "synthetic-private-device-code", user_code: "ABCD-EFGH", verification_uri: "https://microsoft.com/devicelogin", expires_in: 600, interval: 1 };
const response = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status });
async function terminal(signIn: DeviceCodeSignIn, sessionId: string) {
  // Vault publication includes durable filesystem I/O; wait for its observable
  // terminal state instead of assuming it completes within 30 short polls.
  return vi.waitFor(() => {
    const status = signIn.status(sessionId);
    expect(status.state, "device sign-in must reach a terminal state").not.toBe("pending");
    return status;
  }, { timeout: 2_000, interval: 10 });
}

describe("gateway-side device-code sign-in", () => {
  it("allows named OneDrive setup before grants without changing the policy", async () => {
    const { state, config } = await setup();
    config.policy.services.onedrive.allowed_roots = [];
    for (const service of ["calendar", "contacts", "mail", "todo"] as const) {
      const policy = config.policy.services[service];
      if (policy) policy.agents = {};
    }
    const fetchFn = vi.fn(async (_url: RequestInfo | URL, init?: RequestInit) => {
      const body = new URLSearchParams(init?.body as URLSearchParams);
      if (body.has("scope")) {
        expect(body.get("scope")).toBe("Files.Read offline_access");
        return response(deviceResponse);
      }
      return response({ refresh_token: "synthetic-refresh", scope: "Files.Read offline_access" });
    });
    const legacy = new DeviceCodeSignIn(config, () => state, fetchFn);
    await expect(legacy.start(clientId, tenant)).rejects.toThrow("invalid_policy");
    expect(fetchFn).not.toHaveBeenCalled();
    const named = new DeviceCodeSignIn(config, () => state, fetchFn, async () => undefined, true);
    const started = await named.start(clientId, tenant);
    expect(await terminal(named, started.sessionId)).toEqual({ state: "created", scopes: ["Files.Read", "offline_access"] });
    expect(requiredPolicyScopes(config.policy)).toEqual([]);
  });

  it("stores the exact granted workload scopes and refresh token only in the vault", async () => {
    const { state, key, config } = await setup();
    let requested = ""; let deviceRequests = 0;
    const fetchFn = vi.fn(async (_url: string, init?: RequestInit) => {
      const body = new URLSearchParams(init?.body as URLSearchParams);
      if (body.has("scope")) { deviceRequests++; requested = body.get("scope")!; return response(deviceResponse); }
      return response({ refresh_token: "synthetic-refresh", scope: requested });
    }) as unknown as typeof fetch;
    const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, async () => undefined);
    const started = await signIn.start(clientId, tenant);
    expect(started).toMatchObject({ userCode: "ABCD-EFGH", verificationUri: "https://microsoft.com/devicelogin" });
    expect(started.scopes).toContain("offline_access");
    expect(await signIn.start(clientId, tenant)).toEqual(started);
    expect(deviceRequests).toBe(1);
    expect(await terminal(signIn, started.sessionId)).toEqual({ state: "created", scopes: requested.split(" ") });
    const record = await readVaultCredential(state, key);
    expect(record.credential).toEqual({ clientId, tenant, refreshToken: "synthetic-refresh", scopes: requested.split(" ") });
    expect(JSON.stringify(started)).not.toContain("synthetic-refresh");
  });

  it("accepts Microsoft's alternate device page but rejects arbitrary browser destinations", async () => {
    for (const uri of ["https://login.microsoft.com/device", "https://microsoft.com/devicelogin", "https://www.microsoft.com/devicelogin", "https://www.microsoft.com/link"]) {
      const { state, config } = await setup();
      const fetchFn = vi.fn().mockResolvedValue(response({ ...deviceResponse, verification_uri: uri })) as unknown as typeof fetch;
      const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, () => new Promise<void>(() => undefined));
      expect((await signIn.start(clientId, tenant)).verificationUri).toBe(uri);
    }
    for (const uri of ["https://login.microsoft.com.evil.invalid/device", "https://login.microsoft.com/other", "http://login.microsoft.com/device", "http://www.microsoft.com/link", "https://www.microsoft.com/link/", "https://www.microsoft.com.evil.invalid/link"]) {
      const { state, config } = await setup();
      const fetchFn = vi.fn().mockResolvedValue(response({ ...deviceResponse, verification_uri: uri })) as unknown as typeof fetch;
      const signIn = new DeviceCodeSignIn(config, () => state, fetchFn);
      await expect(signIn.start(clientId, tenant)).rejects.toThrow("device_authorization_invalid_response");
      expect(await inspectVaultCredential(state)).toEqual({ result: "missing" });
    }
  });

  it.each([
    ["transport_failed", async () => { throw new Error("synthetic-private-transport-error"); }],
    ["transport_failed", async () => new Response(new ReadableStream({ start(stream) { stream.error(new Error("synthetic-private-body-error")); } }))],
    ["invalid_response", async () => new Response("synthetic-private-non-json", { status: 502 })],
    ["invalid_response", async () => response([], 400)],
    ["invalid_response", async () => response({ padding: "x".repeat(65536) })],
    ["invalid_response", async () => response({}, 503)],
    ["token_rejected", async () => response({ error: "invalid_grant", error_description: "synthetic-private-error", correlation_id: tenant, error_codes: [70000, 53003], refresh_token: "synthetic-private-refresh", ...deviceResponse }, 400)],
    ["token_rejected_without_codes", async () => response({ error: "invalid_grant", error_codes: [70000, "synthetic-private-error"] }, 400)],
    ["invalid_token_fields", async () => response({ refresh_token: "synthetic-private-refresh" })],
    ["invalid_token_fields", async () => response({ refresh_token: "synthetic-private-refresh", scope: " " })],
  ] as const)("reports %s without exposing response details", async (category, tokenResponse) => {
    const { state, config } = await setup();
    const fetchFn = vi.fn<typeof fetch>().mockResolvedValueOnce(response(deviceResponse)).mockImplementationOnce(tokenResponse);
    const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, async () => undefined);
    const started = await signIn.start(clientId, tenant);
    expect(await terminal(signIn, started.sessionId)).toEqual({ state: "failed", error: `device_authorization_${category === "token_rejected_without_codes" ? "token_rejected" : category}`, ...(category === "token_rejected" ? { microsoftErrorCodes: [70000, 53003] } : {}) });
    expect(await inspectVaultCredential(state)).toEqual({ result: "missing" });
  });

  it("reports an unexpected local vault failure without exposing its message", async () => {
    const { state, config } = await setup();
    const scope = requiredPolicyScopes(config.policy).join(" ");
    const fetchFn = vi.fn<typeof fetch>().mockResolvedValueOnce(response(deviceResponse))
      .mockResolvedValueOnce(response({ refresh_token: "synthetic-private-refresh", scope }));
    const stateDir = vi.fn(() => state).mockReturnValueOnce(state).mockImplementationOnce(() => {
      throw new Error("synthetic-private-vault-error");
    });
    const signIn = new DeviceCodeSignIn(config, stateDir, fetchFn, async () => undefined);
    const started = await signIn.start(clientId, tenant);
    expect(await terminal(signIn, started.sessionId)).toEqual({ state: "failed", error: "credential_vault_write_failed" });
    expect(await inspectVaultCredential(state)).toEqual({ result: "missing" });
  });

  it("never creates a vault on missing grant, consent denial, or malformed response", async () => {
    for (const tokenResponse of [
      response({ refresh_token: "synthetic-refresh", scope: "Files.Read offline_access" }),
      response({ error: "authorization_declined" }, 400),
      response({ access_token: "no-refresh", scope: "Files.Read offline_access" }),
    ]) {
      const { state, config } = await setup();
      const fetchFn = vi.fn().mockResolvedValueOnce(response(deviceResponse)).mockResolvedValueOnce(tokenResponse) as unknown as typeof fetch;
      const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, async () => undefined);
      const started = await signIn.start(clientId, tenant);
      expect((await terminal(signIn, started.sessionId)).state).toBe("failed");
      expect(await inspectVaultCredential(state)).toEqual({ result: "missing" });
    }
  });

  it("cancels a pending browser sign-in without publishing a credential", async () => {
    const { state, config } = await setup();
    let release!: () => void;
    const wait = () => new Promise<void>((resolve) => { release = resolve; });
    const fetchFn = vi.fn().mockResolvedValue(response(deviceResponse)) as unknown as typeof fetch;
    const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, wait);
    const started = await signIn.start(clientId, tenant);
    expect(signIn.cancel(started.sessionId)).toEqual({ state: "failed", error: "device_authorization_cancelled" });
    release();
    await Promise.resolve();
    expect(await inspectVaultCredential(state)).toEqual({ result: "missing" });
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("fails closed before network on an existing vault or invalid client", async () => {
    const { state, key, config } = await setup();
    const fetchFn = vi.fn().mockResolvedValue(response(deviceResponse)) as unknown as typeof fetch;
    const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, async () => undefined);
    await expect(signIn.start("not-a-guid", tenant)).rejects.toThrow("invalid_rpc_parameters");
    await expect(signIn.start("-".repeat(36), tenant)).rejects.toThrow("invalid_rpc_parameters");
    await expect(signIn.start(clientId, "-bad.example.com")).rejects.toThrow("invalid_rpc_parameters");
    await createVaultCredential(state, { clientId, tenant, refreshToken: "existing-secret", scopes: ["Files.Read", "offline_access"] }, key);
    await expect(signIn.start(clientId, tenant)).rejects.toThrow("credential_vault_conflict");
    expect(fetchFn).not.toHaveBeenCalled();
  });;

  it("reauthorizes an existing vault only after all current policy scopes are granted", async () => {
    const { state, key, config } = await setup();
    const original = { clientId, tenant, refreshToken: "original-secret", scopes: [...requiredPolicyScopes(config.policy), "offline_access"] };
    await createVaultCredential(state, original, key);
    config.policy.services.contacts = { agents: { main: { operations: ["read", "create"], resources: ["me"] } } };
    let requested = ""; let release!: (value: Response) => void;
    const token = new Promise<Response>(resolve => { release = resolve; });
    const fetchFn = vi.fn(async (_url: string, init?: RequestInit) => {
      const body = new URLSearchParams(init?.body as URLSearchParams);
      if (body.has("scope")) { requested = body.get("scope")!; return response(deviceResponse); }
      return token;
    }) as unknown as typeof fetch;
    const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, async () => undefined);
    const before = await inspectVaultCredential(state, key);
    const started = await signIn.reauthorize();
    expect(started.scopes).toContain("Contacts.ReadWrite");
    expect((await readVaultCredential(state, key)).credential).toEqual(original);
    // A valid asynchronous response may outlast the old ~150 ms polling budget.
    const releaseTimer = setTimeout(() => release(response({ refresh_token: "renewed-secret", scope: requested })), 250);
    try {
      expect(await terminal(signIn, started.sessionId)).toEqual({ state: "created", scopes: requested.split(" ") });
      const after = await readVaultCredential(state, key);
      expect(after.credential.refreshToken).toBe("renewed-secret");
      expect(after.credential.scopes).toContain("Contacts.ReadWrite");
      expect(after.envelope.generation).toBe(before.generation! + 1);
    } finally { clearTimeout(releaseTimer); }
  });

  it("preserves the original vault on missing consent, wrong app, or concurrent rotation", async () => {
    const { state, key, config } = await setup();
    const original = { clientId, tenant, refreshToken: "original-secret", scopes: [...requiredPolicyScopes(config.policy), "offline_access"] };
    await createVaultCredential(state, original, key);
    config.policy.services.contacts = { agents: { main: { operations: ["read", "create"], resources: ["me"] } } };
    let requested = ""; let tokenScope = "";
    const fetchFn = vi.fn(async (_url: string, init?: RequestInit) => {
      const body = new URLSearchParams(init?.body as URLSearchParams);
      if (body.has("scope")) { requested = body.get("scope")!; return response(deviceResponse); }
      return response({ refresh_token: "candidate-secret", scope: tokenScope });
    }) as unknown as typeof fetch;
    const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, async () => undefined);
    await expect(signIn.reauthorize("33333333-3333-4333-8333-333333333333", tenant)).rejects.toThrow("credential_vault_conflict");
    expect(fetchFn).not.toHaveBeenCalled();
    await expect(signIn.reauthorize(clientId, "33333333-3333-4333-8333-333333333333")).rejects.toThrow("credential_vault_conflict_identity");
    expect(fetchFn).not.toHaveBeenCalled();
    tokenScope = original.scopes.join(" ");
    const denied = await signIn.reauthorize(clientId, tenant);
    expect(await terminal(signIn, denied.sessionId)).toEqual({ state: "failed", error: "credential_scope_missing" });
    expect((await readVaultCredential(state, key)).credential).toEqual(original);
    tokenScope = requested;
    const pending = await signIn.reauthorize(clientId, tenant);
    await refreshVaultCredential({ stateDir: state, key, exchange: async () => ({ accessToken: "synthetic-access", replacementRefreshToken: "rotated-secret" }) });
    expect(await terminal(signIn, pending.sessionId)).toEqual({ state: "failed", error: "credential_vault_conflict_changed" });
    expect((await readVaultCredential(state, key)).credential.refreshToken).toBe("rotated-secret");
  });
  it("retains exact vault bytes on denied, malformed, cancelled or changed-policy renewal", async () => {
    for (const outcome of ["denied", "malformed", "cancelled", "policy"] as const) {
      const { state, key, config } = await setup();
      await createVaultCredential(state, { clientId, tenant, refreshToken: "prior-synthetic", scopes: ["Files.Read", "offline_access"] }, key);
      const before = await readFile(vaultCredentialPath(state));
      let release!: () => void;
      const wait = () => new Promise<void>(resolve => { release = resolve; });
      const fetchFn = vi.fn().mockResolvedValueOnce(response(deviceResponse)).mockResolvedValueOnce(
        outcome === "denied" ? response({ error: "authorization_declined" }, 400) :
        outcome === "malformed" ? response({ refresh_token: 7 }) :
        response({ refresh_token: "candidate-synthetic", scope: [...requiredPolicyScopes(config.policy), "offline_access"].join(" ") })
      ) as unknown as typeof fetch;
      const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, wait);
      const started = await signIn.reauthorize();
      if (outcome === "cancelled") signIn.cancel(started.sessionId);
      if (outcome === "policy") config.policy.services.mail.agents = {};
      release();
      const status = await terminal(signIn, started.sessionId);
      expect(status).toEqual({ state: "failed", error: outcome === "denied" ? "device_authorization_declined" : outcome === "malformed" ? "device_authorization_invalid_token_fields" : outcome === "cancelled" ? "device_authorization_cancelled" : "credential_vault_conflict_policy_changed" });
      expect(await readFile(vaultCredentialPath(state))).toEqual(before);
      if (outcome === "cancelled") expect(fetchFn).toHaveBeenCalledTimes(1);
    }
  });

  it("does not bypass an authenticated quarantine or expose its contents", async () => {
    const { state, key, config } = await setup();
    await createVaultCredential(state, { clientId, tenant, refreshToken: "prior-synthetic", scopes: ["Files.Read", "offline_access"] }, key);
    await expect(refreshVaultCredential({ stateDir: state, key, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow("credential_refresh_outcome_uncertain");
    const before = await readFile(vaultCredentialPath(state)); const marker = await readFile(vaultQuarantinePath(state));
    const fetchFn = vi.fn();
    const signIn = new DeviceCodeSignIn(config, () => state, fetchFn);
    await expect(signIn.reauthorize()).rejects.toThrow("credential_vault_conflict_quarantined");
    expect(fetchFn).not.toHaveBeenCalled();
    expect(await readFile(vaultCredentialPath(state))).toEqual(before);
    expect(await readFile(vaultQuarantinePath(state))).toEqual(marker);
  });

  it.each(["success", "missing scope", "cancelled"])("recovers a quarantined credential only after fully consented sign-in (%s)", async (outcome) => {
    const { state, key, config } = await setup();
    await createVaultCredential(state, { clientId, tenant, refreshToken: "old-synthetic", scopes: ["Files.Read"] }, key);
    await expect(refreshVaultCredential({ stateDir: state, key, exchange: async () => {
      throw new Error("credential_refresh_outcome_uncertain");
    } })).rejects.toThrow("credential_refresh_outcome_uncertain");
    const before = await readFile(vaultCredentialPath(state));
    const marker = await readFile(vaultQuarantinePath(state));
    const fetchFn = vi.fn<typeof fetch>()
      .mockResolvedValueOnce(response(deviceResponse))
      .mockResolvedValueOnce(response({ refresh_token: "new-synthetic", scope: outcome === "missing scope" ? "User.Read" : requiredPolicyScopes(config.policy).join(" ") }));
    let resume!: () => void;
    const signIn = new DeviceCodeSignIn(config, () => state, fetchFn, () => new Promise<void>(resolve => { resume = resolve; }));
    const started = await signIn.recover();
    expect(await readFile(vaultQuarantinePath(state))).toEqual(marker);
    if (outcome === "cancelled") signIn.cancel(started.sessionId);
    resume();
    const status = await terminal(signIn, started.sessionId);
    if (outcome === "success") {
      expect(status.state).toBe("created");
      expect(await inspectVaultCredential(state, key)).toMatchObject({ result: "valid", generation: 2 });
      expect((await readVaultCredential(state, key)).credential.refreshToken).toBe("new-synthetic");
    } else {
      expect(status).toMatchObject({ state: "failed", error: outcome === "cancelled" ? "device_authorization_cancelled" : "credential_scope_missing" });
      expect(await readFile(vaultCredentialPath(state))).toEqual(before);
      expect(await readFile(vaultQuarantinePath(state))).toEqual(marker);
    }
  });

});
