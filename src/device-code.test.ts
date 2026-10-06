import { randomBytes } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { DeviceCodeSignIn } from "./device-code.js";
import { createVaultCredential, inspectVaultCredential, readVaultCredential } from "./credential-vault.js";
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
  for (let index = 0; index < 30; index++) {
    const status = signIn.status(sessionId);
    if (status.state !== "pending") return status;
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  throw new Error("device_test_timeout");
}

describe("gateway-side device-code sign-in", () => {
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
      await expect(signIn.start(clientId, tenant)).rejects.toThrow("device_authorization_failed");
      expect(await inspectVaultCredential(state)).toEqual({ result: "missing" });
    }
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
    await createVaultCredential(state, { clientId, tenant, refreshToken: "existing-secret", scopes: ["Files.Read", "offline_access"] }, key);
    await expect(signIn.start(clientId, tenant)).rejects.toThrow("credential_vault_conflict");
    expect(fetchFn).not.toHaveBeenCalled();
  });;
});
