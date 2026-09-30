import { createHash, timingSafeEqual } from "node:crypto";
import { spawn } from "node:child_process";
import { statSync } from "node:fs";
import { basename, isAbsolute } from "node:path";
import { createInterface } from "node:readline/promises";
import { readCredential, selectScope, type CredentialBackendConfig } from "./credential.js";
import { acquireVaultLock, clearVaultQuarantine, decodeVaultKey, inspectVaultCredential, readVaultCredential, vaultRecordBinding, type VaultCredential, type VaultRecord } from "./credential-vault.js";
import { validatePolicy, type GraphPolicy } from "./policy.js";
import { DeviceCodeSignIn } from "./device-code.js";

type CliConfig = CredentialBackendConfig & { policy?: GraphPolicy };
type Receipt = { result: string; generation?: number; keyId?: string; digest?: string; binding?: string; timestamp: string };
type Dependencies = { readPass?: typeof readCredential; writePass?: typeof writePassCredential };
type CliApi = {
  registerCli(registrar: (context: { program: any }) => void | Promise<void>, options?: Record<string, unknown>): void;
};
type GatewayHandlerContext = { params: unknown; respond(ok: boolean, payload?: unknown, error?: unknown): void };
type GatewayApi = { registerGatewayMethod(method: string, handler: (context: GatewayHandlerContext) => void | Promise<void>, options: { scope: "operator.read" | "operator.admin" }): void };
type GatewayOperations = {
  status?: typeof credentialVaultStatus;
  restore?: (config: CliConfig, stateDir: string, destinationRef: string, apply: boolean) => Promise<Receipt>;
  recover?: typeof recoverCredential;
};
type HostCliInvocation = { command: string; argsPrefix: string[] };
type CliDependencies = {
  confirm?: (expected: string) => Promise<void>;
  spawn?: typeof spawn;
  invocation?: HostCliInvocation;
};

export const CREDENTIAL_GATEWAY_METHODS = {
  status: "microsoft-graph.credentials.status",
  restore: "microsoft-graph.credentials.restore-pass",
  recover: "microsoft-graph.credentials.recover-refresh",
  deviceStart: "microsoft-graph.credentials.device-start",
  deviceStatus: "microsoft-graph.credentials.device-status",
  deviceCancel: "microsoft-graph.credentials.device-cancel",
} as const;

const PASS_REF = /^[A-Za-z0-9._/@+-]+$/;
const BINDING = /^[A-Za-z0-9_-]{43}$/;
const KEY_ID = /^[A-Za-z0-9_-]{22}$/;
const MAX_PASS_REF_LENGTH = 1024;
const GATEWAY_RPC_TIMEOUT_MS = 30_000;
const GATEWAY_PROCESS_TIMEOUT_MS = 35_000;
const MAX_GATEWAY_OUTPUT_BYTES = 64 * 1024;
const SAFE_OPERATION_ERRORS = new Set([
  "credential_scope_missing",
  "credential_unavailable",
  "credential_vault_conflict",
  "credential_vault_locked",
  "credential_vault_unavailable",
  "credential_vault_write_failed",
  "invalid_policy",
  "invalid_rpc_parameters",
  "invalid_secret_reference",
  "device_authorization_failed",
  "device_authorization_declined",
  "device_authorization_expired",
  "device_authorization_in_progress",
  "device_session_unavailable",
  "device_authorization_cancelled",
]);
function policyFor(config: CliConfig): GraphPolicy { return validatePolicy(config.policy); }
function keyFor(config: CliConfig): string {
  if (typeof config.credentialVaultKey !== "string") throw new Error("credential_vault_unavailable");
  decodeVaultKey(config.credentialVaultKey); return config.credentialVaultKey;
}
function receipt(result: string, record?: Pick<VaultRecord, "envelope" | "digest">): Receipt {
  return { result, ...(record ? { generation: record.envelope.generation, keyId: record.envelope.keyId, digest: record.digest, binding: vaultRecordBinding(record).binding } : {}), timestamp: new Date().toISOString() };
}

export function resolveHostCliInvocation(execPath = process.execPath, argv = process.argv): HostCliInvocation {
  const entry = argv[1];
  try {
    if (isAbsolute(execPath) && statSync(execPath).isFile()
      && typeof entry === "string" && isAbsolute(entry)
      && /^openclaw(?:\.(?:mjs|js))?$/i.test(basename(entry)) && statSync(entry).isFile()) {
      return { command: execPath, argsPrefix: [entry] };
    }
  } catch {
    // Fall through to the host command name when this process is not a direct OpenClaw Node launch.
  }
  return { command: "openclaw", argsPrefix: [] };
}
function sameCredential(left: VaultCredential, right: VaultCredential): boolean {
  const canonical = (value: VaultCredential) => JSON.stringify({ clientId: value.clientId, refreshToken: value.refreshToken, tenant: value.tenant, scopes: value.scopes });
  return timingSafeEqual(createHash("sha256").update(canonical(left)).digest(), createHash("sha256").update(canonical(right)).digest());
}

export function requiredPolicyScopes(policy: GraphPolicy): string[] {
  const required = new Set<string>();
  for (const root of policy.services.onedrive.allowed_roots) {
    const grants = Object.values(root.agents).flatMap((entry) => Object.entries(entry.permissions).filter(([, allowed]) => allowed).map(([operation]) => operation));
    if (root.permissions.read && grants.includes("read")) required.add("Files.Read");
    if ((root.permissions.write && grants.includes("write")) || (root.permissions.delete && grants.includes("delete"))) required.add("Files.ReadWrite");
  }
  const mappings: Record<"calendar" | "mail" | "todo", Record<string, string>> = {
    calendar: { read: "Calendars.Read", create: "Calendars.ReadWrite", update: "Calendars.ReadWrite", respond: "Calendars.ReadWrite", attach: "Calendars.ReadWrite", delete: "Calendars.ReadWrite" },
    mail: { read: "Mail.Read", draft: "Mail.ReadWrite", update: "Mail.ReadWrite", move: "Mail.ReadWrite", mark: "Mail.ReadWrite", send: "Mail.Send", delete: "Mail.ReadWrite" },
    todo: { read: "Tasks.Read", create: "Tasks.ReadWrite", update: "Tasks.ReadWrite", delete: "Tasks.ReadWrite" },
  };
  for (const service of ["calendar", "mail", "todo"] as const) for (const grant of Object.values(policy.services[service].agents)) for (const operation of grant.operations) required.add(mappings[service][operation]);
  return [...required].sort();
}

export async function credentialVaultStatus(config: CliConfig, stateDir: string) {
  return { policyVersion: policyFor(config).version, credential: await inspectVaultCredential(stateDir, typeof config.credentialVaultKey === "string" ? config.credentialVaultKey : undefined) };
}

async function writePassCredential(secretRef: string, credential: VaultCredential): Promise<void> {
  if (!validPassRef(secretRef)) throw new Error("invalid_secret_reference");
  const content = `${JSON.stringify({ clientId: credential.clientId, refreshToken: credential.refreshToken, tenant: credential.tenant, scopes: credential.scopes })}\n`;
  await new Promise<void>((resolve, reject) => {
    const child = spawn("pass", ["insert", "-m", "-f", secretRef], { stdio: ["pipe", "ignore", "ignore"] });
    const timer = setTimeout(() => child.kill("SIGTERM"), 10_000);
    child.once("error", () => { clearTimeout(timer); reject(new Error("credential_vault_write_failed")); });
    child.once("exit", (code) => { clearTimeout(timer); code === 0 ? resolve() : reject(new Error("credential_vault_write_failed")); });
    child.stdin.end(content);
  });
}

export async function restorePass(config: CliConfig, stateDir: string, destinationRef: string, apply: boolean, dependencies: Dependencies = {}): Promise<Receipt> {
  policyFor(config); const key = keyFor(config);
  if (!validPassRef(destinationRef)) throw new Error("invalid_secret_reference");
  if (!apply) return receipt("ready", await readVaultCredential(stateDir, key));
  const lock = await acquireVaultLock(stateDir, "restore-pass");
  try {
    const record = await readVaultCredential(stateDir, key);
    if (!await lock.verifyStillHeld()) throw new Error("credential_vault_locked");
    try {
      await (dependencies.writePass ?? writePassCredential)(destinationRef, record.credential);
      const observed = await (dependencies.readPass ?? readCredential)(destinationRef);
      return receipt(sameCredential(observed, record.credential) ? "complete" : "unknown", record);
    } catch { return receipt("unknown", record); }
  } finally { await lock.release().catch(() => undefined); }
}

export async function recoverCredential(config: CliConfig, stateDir: string, expectedBinding: string, apply: boolean): Promise<Receipt> {
  policyFor(config); const record = await clearVaultQuarantine(stateDir, keyFor(config), expectedBinding, apply);
  return receipt(apply ? "recovered" : "quarantined", record);
}

async function exactConfirmation(expected: string): Promise<void> {
  if (!process.stdin.isTTY || !process.stderr.isTTY) throw new Error("interactive_confirmation_required");
  const prompt = createInterface({ input: process.stdin, output: process.stderr });
  try { if (await prompt.question(`Type ${expected} to continue: `) !== expected) throw new Error("interactive_confirmation_failed"); }
  finally { prompt.close(); }
}
function mode(options: { dryRun?: boolean; apply?: boolean }): boolean {
  if (options.dryRun === options.apply) throw new Error("select_exactly_one_mode"); return options.apply === true;
}
function printSanitized(value: unknown): void { process.stdout.write(`${JSON.stringify(value)}\n`); }

function validPassRef(value: unknown): value is string {
  return typeof value === "string" && value.length >= 1 && value.length <= MAX_PASS_REF_LENGTH && PASS_REF.test(value);
}
function cliPassRef(value: unknown): string {
  if (!validPassRef(value)) throw new Error("invalid_rpc_parameters");
  return value;
}
function cliBinding(value: unknown): string {
  if (typeof value !== "string" || !BINDING.test(value)) throw new Error("invalid_rpc_parameters");
  return value;
}
function exactParams(value: unknown, keys: readonly string[]): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("invalid_rpc_parameters");
  const params = value as Record<string, unknown>; const actual = Object.keys(params);
  if (actual.length !== keys.length || actual.some((key) => !keys.includes(key))) throw new Error("invalid_rpc_parameters");
  return params;
}
function applyParam(value: unknown): boolean {
  if (typeof value !== "boolean") throw new Error("invalid_rpc_parameters");
  return value;
}
function sanitizedError(error: unknown): string {
  const code = error instanceof Error ? error.message : "internal_error";
  return SAFE_OPERATION_ERRORS.has(code) ? code : "internal_error";
}
function gatewayHandler<T>(action: (params: unknown) => Promise<T>, validate: (value: unknown) => unknown) {
  return async ({ params, respond }: GatewayHandlerContext) => {
    try { respond(true, { ok: true, value: validate(await action(params)) }); }
    catch (error) { respond(true, { ok: false, error: sanitizedError(error) }); }
  };
}

export function registerCredentialGatewayMethods(api: GatewayApi, config: CliConfig, stateDir: () => string, operations: GatewayOperations = {}): void {
  api.registerGatewayMethod(CREDENTIAL_GATEWAY_METHODS.status, gatewayHandler(async (input) => {
    exactParams(input, []);
    return (operations.status ?? credentialVaultStatus)(config, stateDir());
  }, credentialStatusResult), { scope: "operator.read" });
  api.registerGatewayMethod(CREDENTIAL_GATEWAY_METHODS.restore, gatewayHandler(async (input) => {
    const params = exactParams(input, ["destination", "apply"]);
    if (!validPassRef(params.destination)) throw new Error("invalid_rpc_parameters");
    return (operations.restore ?? restorePass)(config, stateDir(), params.destination, applyParam(params.apply));
  }, (value) => receiptResult(value, ["ready", "complete", "unknown"], true)), { scope: "operator.admin" });
  api.registerGatewayMethod(CREDENTIAL_GATEWAY_METHODS.recover, gatewayHandler(async (input) => {
    const params = exactParams(input, ["expectedBinding", "apply"]);
    if (typeof params.expectedBinding !== "string" || !BINDING.test(params.expectedBinding)) throw new Error("invalid_rpc_parameters");
    return (operations.recover ?? recoverCredential)(config, stateDir(), params.expectedBinding, applyParam(params.apply));
  }, (value) => receiptResult(value, ["quarantined", "recovered"], true)), { scope: "operator.admin" });
  const device = new DeviceCodeSignIn(config, stateDir);
  api.registerGatewayMethod(CREDENTIAL_GATEWAY_METHODS.deviceStart, gatewayHandler(async (input) => {
    const params = exactParams(input, ["clientId", "tenant"]);
    if (typeof params.clientId !== "string" || typeof params.tenant !== "string") throw new Error("invalid_rpc_parameters");
    return device.start(params.clientId, params.tenant);
  }, deviceStartResult), { scope: "operator.admin" });
  api.registerGatewayMethod(CREDENTIAL_GATEWAY_METHODS.deviceStatus, gatewayHandler(async (input) => {
    const params = exactParams(input, ["sessionId"]);
    if (typeof params.sessionId !== "string") throw new Error("invalid_rpc_parameters");
    return device.status(params.sessionId);
  }, deviceStatusResult), { scope: "operator.admin" });
  api.registerGatewayMethod(CREDENTIAL_GATEWAY_METHODS.deviceCancel, gatewayHandler(async (input) => {
    const params = exactParams(input, ["sessionId"]);
    if (typeof params.sessionId !== "string") throw new Error("invalid_rpc_parameters");
    return device.cancel(params.sessionId);
  }, deviceStatusResult), { scope: "operator.admin" });
}

function exactResult(value: unknown, required: readonly string[], optional: readonly string[] = []): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("internal_error");
  const result = value as Record<string, unknown>; const keys = Object.keys(result);
  if (required.some((key) => !keys.includes(key)) || keys.some((key) => !required.includes(key) && !optional.includes(key))) throw new Error("internal_error");
  return result;
}
function validateMetadata(value: Record<string, unknown>): void {
  if (!Number.isSafeInteger(value.generation) || Number(value.generation) < 1
    || typeof value.keyId !== "string" || !KEY_ID.test(value.keyId)
    || typeof value.digest !== "string" || !BINDING.test(value.digest)
    || typeof value.binding !== "string" || !BINDING.test(value.binding)) throw new Error("internal_error");
}
function credentialStatusResult(value: unknown): unknown {
  const status = exactResult(value, ["policyVersion", "credential"]);
  if (status.policyVersion !== 2) throw new Error("internal_error");
  const credential = exactResult(status.credential, ["result"], ["generation", "keyId", "digest", "binding"]);
  if (credential.result === "missing" || credential.result === "unavailable") {
    if (Object.keys(credential).length !== 1) throw new Error("internal_error");
  } else if (credential.result === "valid" || credential.result === "quarantined") validateMetadata(credential);
  else throw new Error("internal_error");
  return value;
}
function receiptResult(value: unknown, allowed: readonly string[], metadataRequired: boolean | ((result: string) => boolean)): unknown {
  const parsed = exactResult(value, ["result", "timestamp"], ["generation", "keyId", "digest", "binding"]);
  if (typeof parsed.result !== "string" || !allowed.includes(parsed.result)
    || typeof parsed.timestamp !== "string" || !Number.isFinite(Date.parse(parsed.timestamp)) || new Date(parsed.timestamp).toISOString() !== parsed.timestamp) throw new Error("internal_error");
  const required = typeof metadataRequired === "function" ? metadataRequired(parsed.result) : metadataRequired;
  if (required) validateMetadata(parsed);
  else if (Object.keys(parsed).some((key) => ["generation", "keyId", "digest", "binding"].includes(key))) throw new Error("internal_error");
  return value;
}
function deviceStartResult(value: unknown): unknown {
  const parsed = exactResult(value, ["sessionId", "userCode", "verificationUri", "expiresAt", "scopes"]);
  if (typeof parsed.sessionId !== "string" || !/^[0-9a-fA-F-]{36}$/.test(parsed.sessionId)
    || typeof parsed.userCode !== "string" || !/^[A-Za-z0-9-]{4,32}$/.test(parsed.userCode)
    || !["https://microsoft.com/devicelogin", "https://www.microsoft.com/devicelogin"].includes(parsed.verificationUri as string)
    || typeof parsed.expiresAt !== "string" || !Number.isFinite(Date.parse(parsed.expiresAt))
    || !Array.isArray(parsed.scopes) || parsed.scopes.length === 0 || parsed.scopes.some((scope) => typeof scope !== "string" || !/^[A-Za-z.]+$/.test(scope))) throw new Error("internal_error");
  return value;
}
function deviceStatusResult(value: unknown): unknown {
  const parsed = exactResult(value, ["state"], ["error", "scopes"]);
  if (!["pending", "created", "failed"].includes(parsed.state as string)
    || (parsed.error !== undefined && (parsed.state !== "failed" || typeof parsed.error !== "string" || !SAFE_OPERATION_ERRORS.has(parsed.error)))
    || (parsed.state === "created" && parsed.scopes === undefined)
    || (parsed.scopes !== undefined && (parsed.state !== "created" || !Array.isArray(parsed.scopes) || parsed.scopes.length === 0 || parsed.scopes.some((scope) => typeof scope !== "string" || !/^[A-Za-z.]+$/.test(scope))))) throw new Error("internal_error");
  return value;
}
async function hostGatewayCall(method: string, params: Record<string, unknown>, invocation: HostCliInvocation, spawnProcess: typeof spawn): Promise<unknown> {
  const args = [
    ...invocation.argsPrefix,
    "gateway", "call", method,
    "--json",
    "--params", JSON.stringify(params),
    "--timeout", String(GATEWAY_RPC_TIMEOUT_MS),
  ];
  const stdout = await new Promise<string>((resolve, reject) => {
    let child: ReturnType<typeof spawn>;
    try {
      child = spawnProcess(invocation.command, args, { shell: false, stdio: ["ignore", "pipe", "pipe"], windowsHide: true });
    } catch {
      reject(new Error("internal_error"));
      return;
    }
    let settled = false;
    let stdoutBytes = 0;
    let stderrBytes = 0;
    const chunks: Buffer[] = [];
    const fail = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      child.kill("SIGKILL");
      reject(new Error("internal_error"));
    };
    const timer = setTimeout(fail, GATEWAY_PROCESS_TIMEOUT_MS);
    if (!child.stdout || !child.stderr) {
      fail();
      return;
    }
    child.stdout.on("data", (chunk: Buffer | string) => {
      if (settled) return;
      const data = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
      stdoutBytes += data.byteLength;
      if (stdoutBytes > MAX_GATEWAY_OUTPUT_BYTES) { fail(); return; }
      chunks.push(data);
    });
    child.stderr.on("data", (chunk: Buffer | string) => {
      if (settled) return;
      stderrBytes += Buffer.isBuffer(chunk) ? chunk.byteLength : Buffer.byteLength(chunk);
      if (stderrBytes > MAX_GATEWAY_OUTPUT_BYTES) fail();
    });
    child.once("error", fail);
    child.once("close", (code, signal) => {
      if (settled) return;
      if (code !== 0 || signal !== null) { fail(); return; }
      settled = true;
      clearTimeout(timer);
      resolve(Buffer.concat(chunks, stdoutBytes).toString("utf8"));
    });
  });
  try { return JSON.parse(stdout); }
  catch { throw new Error("internal_error"); }
}

async function credentialGatewayRequest(method: string, params: Record<string, unknown>, validate: (value: unknown) => unknown, invocation: HostCliInvocation, spawnProcess: typeof spawn): Promise<unknown> {
  const response = exactResult(await hostGatewayCall(method, params, invocation, spawnProcess), ["ok"], ["value", "error"]);
  if (response.ok === true && Object.keys(response).length === 2 && "value" in response) return validate(response.value);
  if (response.ok === false && Object.keys(response).length === 2 && typeof response.error === "string") throw new Error(SAFE_OPERATION_ERRORS.has(response.error) ? response.error : "internal_error");
  throw new Error("internal_error");
}

export function registerCredentialCli(api: CliApi, dependencies: CliDependencies = {}): void {
  const confirm = dependencies.confirm ?? exactConfirmation;
  const invocation = dependencies.invocation ?? resolveHostCliInvocation();
  const spawnProcess = dependencies.spawn ?? spawn;
  api.registerCli(({ program }) => {
    const credentials = program.command("microsoft-graph").description("Microsoft Graph operator commands").command("credentials").description("Manage the encrypted Microsoft Graph credential");
    credentials.command("sign-in").description("Sign in with Microsoft device code and create the encrypted vault").requiredOption("--client-id <app-id>").requiredOption("--tenant <tenant-id>").action(async (options: { clientId: string; tenant: string }) => {
      if (!process.stdin.isTTY || !process.stderr.isTTY) throw new Error("interactive_confirmation_required");
      const started = await credentialGatewayRequest(CREDENTIAL_GATEWAY_METHODS.deviceStart, { clientId: options.clientId, tenant: options.tenant }, deviceStartResult, invocation, spawnProcess) as { sessionId: string; userCode: string; verificationUri: string; expiresAt: string; scopes: string[] };
      process.stderr.write(`Requested scopes: ${started.scopes.join(" ")}\nOpen ${started.verificationUri} and enter code ${started.userCode}. Complete sign-in with the approved client.\n`);
      while (Date.now() < Date.parse(started.expiresAt) + 5000) {
        await new Promise<void>((resolve) => setTimeout(resolve, 5000));
        const status = await credentialGatewayRequest(CREDENTIAL_GATEWAY_METHODS.deviceStatus, { sessionId: started.sessionId }, deviceStatusResult, invocation, spawnProcess) as { state: string; error?: string; scopes?: string[] };
        if (status.state === "created") { printSanitized({ result: "created", scopes: status.scopes }); return; }
        if (status.state === "failed") throw new Error(status.error ?? "device_authorization_failed");
      }
      throw new Error("device_authorization_expired");
    });
    credentials.command("status").description("Show sanitized credential status").action(async () => printSanitized(await credentialGatewayRequest(CREDENTIAL_GATEWAY_METHODS.status, {}, credentialStatusResult, invocation, spawnProcess)));
    credentials.command("restore-pass").requiredOption("--destination <pass-ref>").option("--dry-run").option("--apply").action(async (options: { destination: string; dryRun?: boolean; apply?: boolean }) => { const destination = cliPassRef(options.destination); const apply = mode(options); if (apply) await confirm("RESTORE MICROSOFT GRAPH CREDENTIAL"); printSanitized(await credentialGatewayRequest(CREDENTIAL_GATEWAY_METHODS.restore, { destination, apply }, (value) => receiptResult(value, ["ready", "complete", "unknown"], true), invocation, spawnProcess)); });
    credentials.command("recover-refresh").requiredOption("--expected-binding <binding>").option("--dry-run").option("--apply").action(async (options: { expectedBinding: string; dryRun?: boolean; apply?: boolean }) => { const expectedBinding = cliBinding(options.expectedBinding); const apply = mode(options); if (apply) await confirm("RECOVER MICROSOFT GRAPH REFRESH"); printSanitized(await credentialGatewayRequest(CREDENTIAL_GATEWAY_METHODS.recover, { expectedBinding, apply }, (value) => receiptResult(value, ["quarantined", "recovered"], true), invocation, spawnProcess)); });
  }, { commands: ["microsoft-graph"], descriptors: [{ name: "microsoft-graph", description: "Microsoft Graph operator commands", hasSubcommands: true, machineOutput: () => true }] });
}
