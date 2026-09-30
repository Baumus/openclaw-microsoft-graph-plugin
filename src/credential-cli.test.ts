import { randomBytes } from "node:crypto";
import { EventEmitter } from "node:events";
import { access, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PassThrough } from "node:stream";
import { afterEach, describe, expect, it, vi } from "vitest";
import { CREDENTIAL_GATEWAY_METHODS, credentialVaultStatus, migrateFromPass, recoverCredential, registerCredentialCli, registerCredentialGatewayMethods, requiredPolicyScopes, resolveHostCliInvocation, restorePass } from "./credential-cli.js";
import { createVaultCredential, inspectVaultCredential, refreshVaultCredential, vaultQuarantinePath } from "./credential-vault.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

const temporary: string[] = [];
const credential = { clientId: "synthetic-client", refreshToken: "synthetic-refresh", tenant: "synthetic-tenant", scopes: ["Files.ReadWrite", "Calendars.ReadWrite", "Mail.ReadWrite", "Mail.Send", "Tasks.ReadWrite", "offline_access"] };
const key = () => randomBytes(32).toString("base64url");
async function stateDir() { const path = await mkdtemp(join(tmpdir(), "microsoft-graph-cli-test-")); temporary.push(path); return path; }
afterEach(async () => { vi.useRealTimers(); await Promise.all(temporary.splice(0).map((path) => rm(path, { recursive: true, force: true }))); vi.restoreAllMocks(); });

class TestCommand {
  children = new Map<string, TestCommand>();
  actionHandler?: (options?: any) => Promise<void>;
  command(spec: string) { const child = new TestCommand(); this.children.set(spec.split(" ")[0], child); return child; }
  description() { return this; }
  requiredOption() { return this; }
  option() { return this; }
  action(handler: (options?: any) => Promise<void>) { this.actionHandler = handler; return this; }
}

function fakeChild(stdout: string | Buffer, options: { stderr?: string | Buffer; code?: number; signal?: NodeJS.Signals | null; error?: Error; close?: boolean } = {}) {
  const child = new EventEmitter() as EventEmitter & { stdout: PassThrough; stderr: PassThrough; kill: ReturnType<typeof vi.fn> };
  child.stdout = new PassThrough();
  child.stderr = new PassThrough();
  child.kill = vi.fn(() => true);
  if (options.close !== false) queueMicrotask(() => {
    if (options.error) { child.emit("error", options.error); return; }
    child.stdout.end(stdout);
    child.stderr.end(options.stderr ?? "");
    child.emit("close", options.code ?? 0, options.signal ?? null);
  });
  return child;
}

function credentialCommands(spawnProcess: ReturnType<typeof vi.fn>, confirm = vi.fn(async (_expected: string) => undefined)) {
  let registrar: ((context: { program: TestCommand }) => void) | undefined;
  const forbiddenRequest = vi.fn(async () => { throw new Error("restricted_external_plugin_gateway_request"); });
  registerCredentialCli({
    registerCli: (callback: any) => { registrar = callback; },
    runtime: { gateway: { request: forbiddenRequest } },
  } as any, {
    confirm,
    spawn: spawnProcess as any,
    invocation: { command: "/runtime/node", argsPrefix: ["/host/openclaw.mjs"] },
  });
  const program = new TestCommand(); registrar!({ program });
  return { commands: program.children.get("microsoft-graph")!.children.get("credentials")!, confirm, forbiddenRequest };
}

const metadata = { generation: 1, keyId: "k".repeat(22), digest: "d".repeat(43), binding: "b".repeat(43) };
const timestamp = "2026-09-25T00:00:00.000Z";

describe("single credential operator surfaces", () => {
  it("derives the all-scope migration requirement from the policy", () => {
    expect(requiredPolicyScopes(graphPolicyFixture())).toEqual(["Calendars.Read", "Calendars.ReadWrite", "Files.Read", "Files.ReadWrite", "Mail.Read", "Mail.ReadWrite", "Mail.Send", "Tasks.Read", "Tasks.ReadWrite"]);
  });

  it("requires one explicit all-scope source and supports dry-run/apply with secret-free receipts", async () => {
    const state = await stateDir(); const active = key(); const config = { policy: graphPolicyFixture(), credentialVaultKey: active }; const readPass = vi.fn(async () => credential);
    await expect(migrateFromPass(config, state, "", false, { readPass })).rejects.toThrow("invalid_secret_reference");
    await expect(migrateFromPass(config, state, "selected/source", false, { readPass: async () => ({ ...credential, scopes: ["Files.ReadWrite", "offline_access"] }) })).rejects.toThrow("credential_scope_missing");
    expect(await migrateFromPass(config, state, "selected/source", false, { readPass })).toMatchObject({ result: "ready" });
    expect(await inspectVaultCredential(state)).toEqual({ result: "missing" });
    await expect(access(join(state, "plugin-data"))).rejects.toThrow();
    const applied = await migrateFromPass(config, state, "selected/source", true, { readPass }); expect(applied).toMatchObject({ result: "created", generation: 1 });
    const serialized = JSON.stringify(applied); for (const secret of [credential.clientId, credential.refreshToken, active, "selected/source"]) expect(serialized).not.toContain(secret);
    await expect(migrateFromPass(config, state, "selected/source", true, { readPass })).rejects.toThrow("credential_vault_conflict");
  });

  it("restores the one current credential to one explicit destination with an uncertainty receipt", async () => {
    const state = await stateDir(); const active = key(); const config = { policy: graphPolicyFixture(), credentialVaultKey: active }; await createVaultCredential(state, { ...credential, refreshToken: "rotated" }, active);
    const stored = new Map<string, typeof credential>();
    expect(await restorePass(config, state, "rollback/destination", false)).toMatchObject({ result: "ready" });
    expect(await restorePass(config, state, "rollback/destination", true, { writePass: async (ref, value) => { stored.set(ref, value); }, readPass: async (ref) => stored.get(ref)! })).toMatchObject({ result: "complete" });
    expect(await restorePass(config, state, "rollback/uncertain", true, { writePass: async () => { throw new Error("uncertain"); } })).toMatchObject({ result: "unknown" });
  });

  it("exposes one quarantine binding and requires it for recovery", async () => {
    const state = await stateDir(); const active = key(); const config = { policy: graphPolicyFixture(), credentialVaultKey: active }; await createVaultCredential(state, credential, active);
    await expect(refreshVaultCredential({ stateDir: state, key: active, exchange: async () => { throw new Error("credential_refresh_outcome_uncertain"); } })).rejects.toThrow("credential_refresh_outcome_uncertain");
    const status = await credentialVaultStatus(config, state); expect(status).toMatchObject({ policyVersion: 2, credential: { result: "quarantined", binding: expect.any(String) } });
    await expect(recoverCredential(config, state, "stale", true)).rejects.toThrow("credential_vault_conflict");
    const marker = await readFile(vaultQuarantinePath(state), "utf8");
    await expect(recoverCredential(config, state, status.credential.binding!, false)).resolves.toMatchObject({ result: "quarantined" });
    expect(await readFile(vaultQuarantinePath(state), "utf8")).toBe(marker);
    await expect(recoverCredential(config, state, status.credential.binding!, true)).resolves.toMatchObject({ result: "recovered" });
  });

  it("registers strict scoped Gateway methods and sanitizes all failures", async () => {
    const registrations = new Map<string, { handler: (context: any) => Promise<void>; options: { scope: string } }>();
    const migrate = vi.fn(async () => ({ result: "ready", timestamp }));
    const api = { registerGatewayMethod: vi.fn((method: string, handler: (context: any) => Promise<void>, options: { scope: string }) => registrations.set(method, { handler, options })) };
    registerCredentialGatewayMethods(api as any, { enabled: false, policy: graphPolicyFixture(), credentialVaultKey: key() } as any, () => "/gateway/state", { migrate });
    expect(Object.fromEntries([...registrations].map(([method, value]) => [method, value.options.scope]))).toEqual({
      [CREDENTIAL_GATEWAY_METHODS.status]: "operator.read",
      [CREDENTIAL_GATEWAY_METHODS.migrate]: "operator.admin",
      [CREDENTIAL_GATEWAY_METHODS.restore]: "operator.admin",
      [CREDENTIAL_GATEWAY_METHODS.recover]: "operator.admin",
      [CREDENTIAL_GATEWAY_METHODS.deviceStart]: "operator.admin",
      [CREDENTIAL_GATEWAY_METHODS.deviceStatus]: "operator.admin",
    });

    const invoke = async (method: string, params: unknown) => {
      const respond = vi.fn(); await registrations.get(method)!.handler({ params, respond }); return respond.mock.calls[0];
    };
    expect(await invoke(CREDENTIAL_GATEWAY_METHODS.migrate, { source: "selected/source", apply: false })).toEqual([true, { ok: true, value: { result: "ready", timestamp } }]);
    expect(migrate).toHaveBeenCalledWith(expect.objectContaining({ enabled: false }), "/gateway/state", "selected/source", false);

    for (const [method, params] of [
      [CREDENTIAL_GATEWAY_METHODS.status, { unknown: true }],
      [CREDENTIAL_GATEWAY_METHODS.migrate, { source: "selected/source", apply: false, unknown: true }],
      [CREDENTIAL_GATEWAY_METHODS.migrate, { source: "bad secret", apply: false }],
      [CREDENTIAL_GATEWAY_METHODS.restore, { destination: "rollback/destination", apply: "false" }],
      [CREDENTIAL_GATEWAY_METHODS.recover, { expectedBinding: "short", apply: false }],
    ] as const) expect(await invoke(method, params)).toEqual([true, { ok: false, error: "invalid_rpc_parameters" }]);

    const secret = "synthetic-refresh-token-must-not-cross-rpc";
    migrate.mockRejectedValueOnce(new Error(`pass failed with ${secret}`));
    const failed = await invoke(CREDENTIAL_GATEWAY_METHODS.migrate, { source: "selected/source", apply: true });
    expect(failed).toEqual([true, { ok: false, error: "internal_error" }]);
    expect(JSON.stringify(failed)).not.toContain(secret);
    migrate.mockResolvedValueOnce({ result: "ready", timestamp, refreshToken: secret } as any);
    const rejectedResult = await invoke(CREDENTIAL_GATEWAY_METHODS.migrate, { source: "selected/source", apply: false });
    expect(rejectedResult).toEqual([true, { ok: false, error: "internal_error" }]);
    expect(JSON.stringify(rejectedResult)).not.toContain(secret);
  });

  it("routes the external-plugin CLI through the host gateway command with exact confirmations", async () => {
    const outputs = [
      { ok: true, value: { policyVersion: 2, credential: { result: "missing" } } },
      { ok: true, value: { result: "created", ...metadata, timestamp } },
      { ok: true, value: { result: "complete", ...metadata, timestamp } },
      { ok: true, value: { result: "recovered", ...metadata, timestamp } },
    ];
    const spawnProcess = vi.fn(() => fakeChild(`${JSON.stringify(outputs.shift())}\n`));
    const { commands, confirm, forbiddenRequest } = credentialCommands(spawnProcess);
    const write = vi.spyOn(process.stdout, "write").mockImplementation(() => true);

    await commands.children.get("status")!.actionHandler!();
    await commands.children.get("migrate-from-pass")!.actionHandler!({ source: "selected/source", apply: true });
    await commands.children.get("restore-pass")!.actionHandler!({ destination: "rollback/destination", apply: true });
    await commands.children.get("recover-refresh")!.actionHandler!({ expectedBinding: "b".repeat(43), apply: true });

    expect(confirm.mock.calls.map(([value]) => value)).toEqual([
      "MIGRATE MICROSOFT GRAPH CREDENTIAL",
      "RESTORE MICROSOFT GRAPH CREDENTIAL",
      "RECOVER MICROSOFT GRAPH REFRESH",
    ]);
    expect(spawnProcess.mock.calls).toEqual([
      ["/runtime/node", ["/host/openclaw.mjs", "gateway", "call", CREDENTIAL_GATEWAY_METHODS.status, "--json", "--params", "{}", "--timeout", "30000"], { shell: false, stdio: ["ignore", "pipe", "pipe"], windowsHide: true }],
      ["/runtime/node", ["/host/openclaw.mjs", "gateway", "call", CREDENTIAL_GATEWAY_METHODS.migrate, "--json", "--params", JSON.stringify({ source: "selected/source", apply: true }), "--timeout", "30000"], { shell: false, stdio: ["ignore", "pipe", "pipe"], windowsHide: true }],
      ["/runtime/node", ["/host/openclaw.mjs", "gateway", "call", CREDENTIAL_GATEWAY_METHODS.restore, "--json", "--params", JSON.stringify({ destination: "rollback/destination", apply: true }), "--timeout", "30000"], { shell: false, stdio: ["ignore", "pipe", "pipe"], windowsHide: true }],
      ["/runtime/node", ["/host/openclaw.mjs", "gateway", "call", CREDENTIAL_GATEWAY_METHODS.recover, "--json", "--params", JSON.stringify({ expectedBinding: "b".repeat(43), apply: true }), "--timeout", "30000"], { shell: false, stdio: ["ignore", "pipe", "pipe"], windowsHide: true }],
    ]);
    expect(forbiddenRequest).not.toHaveBeenCalled();
    expect(write.mock.calls.map(([value]) => value)).toEqual([
      `${JSON.stringify({ policyVersion: 2, credential: { result: "missing" } })}\n`,
      `${JSON.stringify({ result: "created", ...metadata, timestamp })}\n`,
      `${JSON.stringify({ result: "complete", ...metadata, timestamp })}\n`,
      `${JSON.stringify({ result: "recovered", ...metadata, timestamp })}\n`,
    ]);
    for (let index = 0; index < confirm.mock.invocationCallOrder.length; index += 1) expect(confirm.mock.invocationCallOrder[index]).toBeLessThan(spawnProcess.mock.invocationCallOrder[index + 1]);
  });

  it("chooses the current OpenClaw Node entrypoint only when it is valid", async () => {
    const directory = await stateDir();
    const entry = join(directory, "openclaw.mjs");
    const unrelated = join(directory, "runner.mjs");
    await writeFile(entry, "");
    await writeFile(unrelated, "");
    expect(resolveHostCliInvocation(process.execPath, ["node", entry])).toEqual({ command: process.execPath, argsPrefix: [entry] });
    expect(resolveHostCliInvocation(process.execPath, ["node", unrelated])).toEqual({ command: "openclaw", argsPrefix: [] });
    expect(resolveHostCliInvocation("relative-node", ["node", entry])).toEqual({ command: "openclaw", argsPrefix: [] });
  });

  it("rejects non-reference CLI parameters before placing them in child arguments", async () => {
    const spawnProcess = vi.fn();
    const { commands, confirm } = credentialCommands(spawnProcess);
    await expect(commands.children.get("migrate-from-pass")!.actionHandler!({ source: "not a pass reference", apply: true })).rejects.toThrow("invalid_rpc_parameters");
    await expect(commands.children.get("restore-pass")!.actionHandler!({ destination: "refresh token contents", apply: true })).rejects.toThrow("invalid_rpc_parameters");
    await expect(commands.children.get("recover-refresh")!.actionHandler!({ expectedBinding: "short", apply: true })).rejects.toThrow("invalid_rpc_parameters");
    expect(confirm).not.toHaveBeenCalled();
    expect(spawnProcess).not.toHaveBeenCalled();
  });

  it.each([
    ["oversized stdout", Buffer.alloc(64 * 1024 + 1), ""],
    ["oversized stderr", JSON.stringify({ ok: true, value: { policyVersion: 2, credential: { result: "missing" } } }), Buffer.alloc(64 * 1024 + 1)],
  ])("closes on %s without exposing child output", async (_label, stdout, stderr) => {
    const child = fakeChild(stdout, { stderr });
    const spawnProcess = vi.fn(() => child);
    const { commands } = credentialCommands(spawnProcess);
    await expect(commands.children.get("status")!.actionHandler!()).rejects.toThrow("internal_error");
    expect(child.kill).toHaveBeenCalledWith("SIGKILL");
  });

  it("fails closed on malformed output, child errors, and nonzero exits", async () => {
    const secret = "refresh-token-that-must-not-surface";
    const children = [
      () => fakeChild(`not-json-${secret}`),
      () => fakeChild("", { error: new Error(secret) }),
      () => fakeChild("", { stderr: secret, code: 1 }),
      () => fakeChild(JSON.stringify({ ok: true, value: { policyVersion: 2, credential: { result: "missing" } }, extra: secret })),
    ];
    const spawnProcess = vi.fn(() => children.shift()!());
    const { commands } = credentialCommands(spawnProcess);
    const status = commands.children.get("status")!.actionHandler!;
    for (let index = 0; index < 4; index += 1) await expect(status()).rejects.toThrow(/^internal_error$/);
  });

  it("passes through only the allowlisted sanitized operation errors", async () => {
    const secret = "refresh-token-that-must-not-surface";
    const outputs = [
      { ok: false, error: "credential_vault_locked" },
      { ok: false, error: secret },
    ];
    const spawnProcess = vi.fn(() => fakeChild(JSON.stringify(outputs.shift())));
    const { commands } = credentialCommands(spawnProcess);
    const status = commands.children.get("status")!.actionHandler!;
    await expect(status()).rejects.toThrow(/^credential_vault_locked$/);
    await expect(status()).rejects.toThrow(/^internal_error$/);
  });

  it("terminates a stuck host CLI after the bounded process timeout", async () => {
    vi.useFakeTimers();
    const child = fakeChild("", { close: false });
    const spawnProcess = vi.fn(() => child);
    const { commands } = credentialCommands(spawnProcess);
    const pending = commands.children.get("status")!.actionHandler!();
    const rejected = expect(pending).rejects.toThrow("internal_error");
    await vi.advanceTimersByTimeAsync(35_000);
    await rejected;
    expect(child.kill).toHaveBeenCalledWith("SIGKILL");
  });
});
