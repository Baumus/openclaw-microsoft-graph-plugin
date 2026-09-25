import { createHash } from "node:crypto";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("./credential.js", async () => {
  const actual = await vi.importActual<typeof import("./credential.js")>("./credential.js");
  const readCredential = vi.fn(async () => ({ clientId: "synthetic", refreshToken: "synthetic", tenant: "common", scopes: ["Files.ReadWrite"] }));
  const exchangeRefreshToken = vi.fn(async () => "synthetic-token");
  return {
    ...actual,
    readCredential,
    exchangeRefreshToken,
    tokenForAuthorizedOperation: vi.fn(async () => { await readCredential(); return exchangeRefreshToken(); }),
  };
});

import { exchangeRefreshToken, readCredential } from "./credential.js";
import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

const owner = { senderIsOwner: true, channel: "telegram" };
let workspaceDir = "";
let sequence = 0;

function digest(bytes: Uint8Array) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function consumeBody(body: unknown): Promise<void> {
  if (body && typeof (body as any)[Symbol.asyncIterator] === "function") {
    for await (const _chunk of body as AsyncIterable<Uint8Array>) { /* consume upload stream */ }
  }
}

function runtime(sessionId: string, agentId = "main", options: { policy?: ReturnType<typeof graphPolicyFixture>; instructionPreflight?: boolean } = {}) {
  const policy = options.policy ?? graphPolicyFixture();
  const hooks: Record<string, (...args: any[]) => Promise<any> | any> = {};
  entry.register({
    pluginConfig: { enabled: options.instructionPreflight === true, policy },
    registerTool: vi.fn(),
    on: (name: string, handler: any) => { hooks[name] = handler; },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);

  const factories: Array<(context: any) => any> = [];
  entry.register({
    pluginConfig: { enabled: true, policy },
    registerTool: (factory: any) => factories.push(factory),
    on: vi.fn(),
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  const toolContext = { agentId, sessionId, workspaceDir };
  const tools = Object.fromEntries(factories.map((factory) => {
    const tool = factory(toolContext);
    return [tool.name, tool];
  }));
  return { hooks, tools, context: { ...toolContext, requester: owner } };
}

function tokenFrom(blocked: any): string {
  const token = blocked?.blockReason?.match(/chatConfirmationToken="(mgw1_[A-Za-z0-9_-]{43})"/)?.[1];
  if (!token) throw new Error("missing confirmation token");
  return token;
}

async function challenge(hooks: Record<string, (...args: any[]) => any>, context: any, toolName: string, params: Record<string, unknown>) {
  const blocked = await hooks.before_tool_call({ toolName, params }, context);
  expect(blocked).toMatchObject({ block: true, blockReason: expect.stringContaining("chat_confirmation_required") });
  return tokenFrom(blocked);
}

async function arm(hooks: Record<string, (...args: any[]) => any>, context: any, toolName: string, params: Record<string, unknown>, token: string) {
  const confirmed = { ...params, chatConfirmed: true, chatConfirmationToken: token };
  expect(await hooks.before_tool_call({ toolName, params: confirmed }, context)).toBeUndefined();
  return confirmed;
}

function graphSuccess(bytes: Uint8Array, update = false) {
  return vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
    if (update && init?.method !== "PUT") return Response.json({ eTag: "before", file: {} });
    await consumeBody(init?.body);
    return Response.json({ id: "item-1", name: "invoice.pdf", size: bytes.byteLength, file: { mimeType: "application/pdf" } }, { status: update ? 200 : 201 });
  });
}

beforeEach(async () => {
  sequence += 1;
  workspaceDir = await mkdtemp(join(tmpdir(), "msgraph-content-identity-"));
  await mkdir(join(workspaceDir, "media", "inbound"), { recursive: true });
  vi.mocked(readCredential).mockClear();
  vi.mocked(exchangeRefreshToken).mockClear();
});

afterEach(async () => {
  vi.restoreAllMocks();
  await rm(workspaceDir, { recursive: true, force: true });
});

describe.each([
  ["onedrive_upload", false],
  ["onedrive_update", true],
] as const)("content-identity confirmation for %s", (toolName, update) => {
  it("accepts identical bytes under a refreshed URI and claims before Graph", async () => {
    const bytes = Buffer.from(`invoice-${toolName}`);
    await writeFile(join(workspaceDir, "media", "inbound", "fresh.pdf"), bytes);
    const { hooks, tools, context } = runtime(`same-bytes-${sequence}-${toolName}`);
    const intended = {
      rootLabel: "synthetic_documents",
      relativePath: "invoices/invoice.pdf",
      sourceMediaUri: "media://inbound/expired.pdf",
      sourceSha256: digest(bytes),
      sourceByteSize: bytes.byteLength,
      contentType: "application/pdf",
    };
    const token = await challenge(hooks, context, toolName, intended);
    const refreshed = await arm(hooks, context, toolName, { ...intended, sourceMediaUri: "media://inbound/fresh.pdf" }, token);
    const fetchSpy = graphSuccess(bytes, update);
    const response = await tools[toolName].execute("write", refreshed);
    expect(response.details).toMatchObject({ ok: true, source_sha256: digest(bytes), source_byte_size: bytes.byteLength });
    expect(readCredential).toHaveBeenCalledTimes(1);
    expect(fetchSpy).toHaveBeenCalledTimes(update ? 2 : 1);

    const replay = await hooks.before_tool_call({ toolName, params: refreshed }, context);
    expect(replay).toMatchObject({ block: true, blockReason: expect.stringContaining("chat_confirmation_invalid_or_changed") });
  });

  it("keeps the token usable after a missing or wrong artifact and reads no credential", async () => {
    const bytes = Buffer.from(`retry-${toolName}`);
    const { hooks, tools, context } = runtime(`artifact-retry-${sequence}-${toolName}`);
    const intended = {
      rootLabel: "synthetic_documents",
      relativePath: "invoices/retry.pdf",
      sourceMediaUri: "media://inbound/original.pdf",
      sourceSha256: digest(bytes),
      sourceByteSize: bytes.byteLength,
      contentType: "application/pdf",
    };
    const token = await challenge(hooks, context, toolName, intended);
    const missing = await arm(hooks, context, toolName, { ...intended, sourceMediaUri: "media://inbound/missing.pdf" }, token);
    expect((await tools[toolName].execute("missing", missing)).details).toEqual({ ok: false, error: "invalid_source_media_uri" });
    expect(readCredential).not.toHaveBeenCalled();

    await writeFile(join(workspaceDir, "media", "inbound", "wrong.pdf"), Buffer.alloc(bytes.byteLength, 0x7a));
    const wrong = await arm(hooks, context, toolName, { ...intended, sourceMediaUri: "media://inbound/wrong.pdf" }, token);
    expect((await tools[toolName].execute("wrong", wrong)).details).toEqual({ ok: false, error: "invalid_source_fingerprint" });
    expect(readCredential).not.toHaveBeenCalled();

    await writeFile(join(workspaceDir, "media", "inbound", "correct.pdf"), bytes);
    const correct = await arm(hooks, context, toolName, { ...intended, sourceMediaUri: "media://inbound/correct.pdf" }, token);
    graphSuccess(bytes, update);
    expect((await tools[toolName].execute("correct", correct)).details).toMatchObject({ ok: true });
    expect(readCredential).toHaveBeenCalledTimes(1);
  });
});

it("rejects changed semantic bindings across digest, size, destination, root, content type, tool, agent, and session", async () => {
  const bytes = Buffer.from("binding");
  const { hooks, context } = runtime(`bindings-${sequence}`);
  const intended = {
    rootLabel: "synthetic_documents",
    relativePath: "invoice.pdf",
    sourceMediaUri: "media://inbound/original.pdf",
    sourceSha256: digest(bytes),
    sourceByteSize: bytes.byteLength,
    contentType: "application/pdf",
  };
  const token = await challenge(hooks, context, "onedrive_upload", intended);
  const changes = [
    { toolName: "onedrive_upload", params: { ...intended, sourceSha256: "0".repeat(64) }, context, error: "chat_confirmation_invalid_or_changed" },
    { toolName: "onedrive_upload", params: { ...intended, sourceByteSize: bytes.byteLength + 1 }, context, error: "chat_confirmation_invalid_or_changed" },
    { toolName: "onedrive_upload", params: { ...intended, relativePath: "other.pdf" }, context, error: "chat_confirmation_invalid_or_changed" },
    { toolName: "onedrive_upload", params: { ...intended, rootLabel: "other_root" }, context, error: "access_denied" },
    { toolName: "onedrive_upload", params: { ...intended, contentType: "text/plain" }, context, error: "chat_confirmation_invalid_or_changed" },
    { toolName: "onedrive_update", params: intended, context, error: "chat_confirmation_invalid_or_changed" },
    { toolName: "onedrive_upload", params: intended, context: { ...context, agentId: "other" }, error: "access_denied" },
    { toolName: "onedrive_upload", params: intended, context: { ...context, sessionId: "other" }, error: "chat_confirmation_invalid_or_changed" },
  ];
  for (const changed of changes) {
    const result = await hooks.before_tool_call({ toolName: changed.toolName, params: { ...changed.params, chatConfirmed: true, chatConfirmationToken: token } }, changed.context);
    expect(result).toMatchObject({ block: true, blockReason: expect.stringContaining(changed.error) });
  }
});

it("rejects a same-label root re-pin before instruction discovery, artifact opening, credentials, or Graph", async () => {
  const bytes = Buffer.from("root-pin");
  const original = graphPolicyFixture();
  delete original.services.onedrive.allowed_roots[0].agents_instructions;
  original.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
  const repinned = structuredClone(original);
  repinned.services.onedrive.allowed_roots[0].drive_id = "repinned-drive";
  repinned.services.onedrive.allowed_roots[0].item_id = "repinned-root";
  repinned.services.onedrive.allowed_roots[0].agents_instructions = "trusted";
  const activePolicy = structuredClone(original);
  const { hooks, tools, context } = runtime(`root-repin-${sequence}`, "main", { policy: activePolicy, instructionPreflight: true });
  const intended = {
    rootLabel: "synthetic_documents",
    relativePath: "invoice.pdf",
    sourceMediaUri: "media://inbound/missing.pdf",
    sourceSha256: digest(bytes),
    sourceByteSize: bytes.byteLength,
    contentType: "application/pdf",
  };

  const staleBeforeHook = await challenge(hooks, context, "onedrive_upload", intended);
  Object.assign(activePolicy.services.onedrive.allowed_roots[0], repinned.services.onedrive.allowed_roots[0]);
  vi.mocked(readCredential).mockClear();
  vi.mocked(exchangeRefreshToken).mockClear();
  const fetchSpy = vi.spyOn(globalThis, "fetch").mockRejectedValue(new Error("unexpected Graph access"));
  const rejected = await hooks.before_tool_call({
    toolName: "onedrive_upload",
    params: { ...intended, chatConfirmed: true, chatConfirmationToken: staleBeforeHook },
  }, context);
  expect(rejected).toMatchObject({ block: true, blockReason: expect.stringContaining("chat_confirmation_invalid_or_changed") });
  expect(readCredential).not.toHaveBeenCalled();
  expect(exchangeRefreshToken).not.toHaveBeenCalled();
  expect(fetchSpy).not.toHaveBeenCalled();

  Object.assign(activePolicy.services.onedrive.allowed_roots[0], original.services.onedrive.allowed_roots[0]);
  delete activePolicy.services.onedrive.allowed_roots[0].agents_instructions;
  const staleAtExecution = await challenge(hooks, context, "onedrive_upload", intended);
  const armed = await arm(hooks, context, "onedrive_upload", intended, staleAtExecution);
  Object.assign(activePolicy.services.onedrive.allowed_roots[0], repinned.services.onedrive.allowed_roots[0]);
  vi.mocked(readCredential).mockClear();
  vi.mocked(exchangeRefreshToken).mockClear();
  fetchSpy.mockClear();
  expect((await tools.onedrive_upload.execute("re-pinned", armed)).details).toEqual({ ok: false, error: "chat_confirmation_invalid_or_changed" });
  expect(readCredential).not.toHaveBeenCalled();
  expect(exchangeRefreshToken).not.toHaveBeenCalled();
  expect(fetchSpy).not.toHaveBeenCalled();
});

it("allows exactly one concurrent claim", async () => {
  const bytes = Buffer.from("concurrent");
  await Promise.all(["one.pdf", "two.pdf"].map((name) => writeFile(join(workspaceDir, "media", "inbound", name), bytes)));
  const { hooks, tools, context } = runtime(`concurrent-${sequence}`);
  const intended = { rootLabel: "synthetic_documents", relativePath: "invoice.pdf", sourceMediaUri: "media://inbound/original.pdf", sourceSha256: digest(bytes), sourceByteSize: bytes.byteLength, contentType: "application/pdf" };
  const token = await challenge(hooks, context, "onedrive_upload", intended);
  const calls = await Promise.all(["one.pdf", "two.pdf"].map((name) => arm(hooks, context, "onedrive_upload", { ...intended, sourceMediaUri: `media://inbound/${name}` }, token)));
  const fetchSpy = graphSuccess(bytes);
  const results = await Promise.all(calls.map((params, index) => tools.onedrive_upload.execute(`concurrent-${index}`, params).then((response: any) => response.details)));
  expect(results.filter((value) => value.ok)).toHaveLength(1);
  expect(results.filter((value) => value.error === "chat_confirmation_invalid_or_changed")).toHaveLength(1);
  expect(readCredential).toHaveBeenCalledTimes(1);
  expect(fetchSpy).toHaveBeenCalledTimes(1);
});

it("authorizes before opening and claims before credentials", async () => {
  const bytes = Buffer.from("ordered");
  await writeFile(join(workspaceDir, "media", "inbound", "ordered.pdf"), bytes);
  const unauthorized = runtime(`unauthorized-${sequence}`, "unauthorized-agent");
  const params = { rootLabel: "synthetic_documents", relativePath: "ordered.pdf", sourceMediaUri: "media://inbound/missing.pdf", sourceSha256: digest(bytes), sourceByteSize: bytes.byteLength, contentType: "application/pdf" };
  const unauthorizedCall = { ...params, chatConfirmed: true, chatConfirmationToken: "mgw1_" + "A".repeat(43) };
  expect((await unauthorized.tools.onedrive_upload.execute("unauthorized", unauthorizedCall)).details).toEqual({ ok: false, error: "access_denied" });
  expect(readCredential).not.toHaveBeenCalled();

  const authorized = runtime(`unarmed-${sequence}`);
  const unarmed = { ...params, sourceMediaUri: "media://inbound/ordered.pdf", chatConfirmed: true, chatConfirmationToken: "mgw1_" + "A".repeat(43) };
  expect((await authorized.tools.onedrive_upload.execute("unarmed", unarmed)).details).toEqual({ ok: false, error: "chat_confirmation_invalid_or_changed" });
  expect(readCredential).not.toHaveBeenCalled();
});

it("handles multiple confirmed invoice intents independently", async () => {
  const invoices = [Buffer.from("invoice-a"), Buffer.from("invoice-b")];
  await Promise.all(invoices.map((bytes, index) => writeFile(join(workspaceDir, "media", "inbound", `fresh-${index}.pdf`), bytes)));
  const { hooks, tools, context } = runtime(`batch-${sequence}`);
  const armed: Array<Record<string, unknown>> = [];
  for (let index = 0; index < invoices.length; index += 1) {
    const bytes = invoices[index];
    const intended = { rootLabel: "synthetic_documents", relativePath: `invoices/${index}.pdf`, sourceMediaUri: `media://inbound/expired-${index}.pdf`, sourceSha256: digest(bytes), sourceByteSize: bytes.byteLength, contentType: "application/pdf" };
    const token = await challenge(hooks, context, "onedrive_upload", intended);
    armed.push(await arm(hooks, context, "onedrive_upload", { ...intended, sourceMediaUri: `media://inbound/fresh-${index}.pdf` }, token));
  }
  const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
    await consumeBody(init?.body);
    const length = Number((init?.headers as Record<string, string>)["content-length"]);
    return Response.json({ id: `item-${length}`, name: "invoice.pdf", size: length, file: { mimeType: "application/pdf" } }, { status: 201 });
  });
  const results = await Promise.all(armed.map((params, index) => tools.onedrive_upload.execute(`batch-${index}`, params)));
  expect(results.every((response: any) => response.details.ok)).toBe(true);
  expect(fetchSpy).toHaveBeenCalledTimes(2);
});

it("keeps legacy URI-bound confirmation behavior without fingerprints", async () => {
  const { hooks, context } = runtime(`legacy-${sequence}`);
  const intended = { rootLabel: "synthetic_documents", relativePath: "legacy.pdf", sourceMediaUri: "media://inbound/old.pdf", contentType: "application/pdf" };
  const token = await challenge(hooks, context, "onedrive_upload", intended);
  expect(await hooks.before_tool_call({ toolName: "onedrive_upload", params: { ...intended, chatConfirmed: true, chatConfirmationToken: token } }, context)).toBeUndefined();
  const refreshed = await hooks.before_tool_call({ toolName: "onedrive_upload", params: { ...intended, sourceMediaUri: "media://inbound/new.pdf", chatConfirmed: true, chatConfirmationToken: token } }, context);
  expect(refreshed).toMatchObject({ block: true, blockReason: expect.stringContaining("chat_confirmation_invalid_or_changed") });
});

it("does not reuse a claimed receipt after an uncertain Graph failure", async () => {
  const bytes = Buffer.from("uncertain");
  await writeFile(join(workspaceDir, "media", "inbound", "uncertain.pdf"), bytes);
  const { hooks, tools, context } = runtime(`uncertain-${sequence}`);
  const intended = { rootLabel: "synthetic_documents", relativePath: "uncertain.pdf", sourceMediaUri: "media://inbound/uncertain.pdf", sourceSha256: digest(bytes), sourceByteSize: bytes.byteLength, contentType: "application/pdf" };
  const token = await challenge(hooks, context, "onedrive_upload", intended);
  const confirmed = await arm(hooks, context, "onedrive_upload", intended, token);
  const fetchSpy = vi.spyOn(globalThis, "fetch").mockRejectedValue(new Error("network uncertainty"));
  expect((await tools.onedrive_upload.execute("uncertain", confirmed)).details).toEqual({ ok: false, error: "provider_unavailable" });
  expect(fetchSpy).toHaveBeenCalledTimes(1);
  const replay = await hooks.before_tool_call({ toolName: "onedrive_upload", params: confirmed }, context);
  expect(replay).toMatchObject({ block: true, blockReason: expect.stringContaining("chat_confirmation_invalid_or_changed") });
});
