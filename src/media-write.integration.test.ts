import { createHash } from "node:crypto";
import { link, mkdir, mkdtemp, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("./credential.js", async () => {
  const actual = await vi.importActual<typeof import("./credential.js")>("./credential.js");
  const readCredential = vi.fn(async () => ({ clientId: "synthetic", refreshToken: "synthetic", tenant: "common", scopes: ["Files.ReadWrite", "Calendars.ReadWrite", "Mail.ReadWrite", "Tasks.ReadWrite"] }));
  const exchangeRefreshToken = vi.fn(async () => "synthetic-token");
  return {
    ...actual,
    readCredential,
    exchangeRefreshToken,
    tokenForAuthorizedOperation: vi.fn(async () => { await readCredential(); return exchangeRefreshToken(); }),
  };
});

import entry from "./index.js";
import { exchangeRefreshToken, readCredential } from "./credential.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

let workspaceDir = "";
const fixtureBytes = Buffer.from([0, 255, 1, 254]);
const mp4Bytes = Buffer.from([0, 0, 0, 24, 102, 116, 121, 112, 109, 112, 52, 50]);
const largeMp4Bytes = Buffer.alloc(1024 * 1024 + 17, 0x5a);
const digest = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex");

async function requestBodyBytes(body: unknown): Promise<Buffer> {
  if (body && typeof (body as any)[Symbol.asyncIterator] === "function") {
    const chunks: Buffer[] = [];
    for await (const chunk of body as AsyncIterable<Uint8Array>) chunks.push(Buffer.from(chunk));
    return Buffer.concat(chunks);
  }
  return Buffer.from(body as Uint8Array);
}

function registeredTools(agentId = "main", runtimeWorkspaceDir: string | null = workspaceDir) {
  const factories: Array<(context: any) => any> = [];
  const hooks: Record<string, (...args: any[]) => Promise<any> | any> = {};
  const policy = graphPolicyFixture();
  delete policy.services.onedrive.allowed_roots[0].agents_instructions;
  policy.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
  entry.register({
    pluginConfig: { enabled: true, policy },
    registerTool: (factory: any) => factories.push(factory),
    on: (name: string, handler: any) => { hooks[name] = handler; },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  const context = { agentId, sessionId: `media-write-${agentId}`, workspaceDir: runtimeWorkspaceDir ?? undefined };
  return Object.fromEntries(factories.map((factory) => {
    const tool = factory(context);
    return [tool.name, { ...tool, async execute(toolCallId: string, params: Record<string, unknown>, signal?: AbortSignal) {
      const gate = await hooks.before_tool_call({ toolName: tool.name, toolCallId, params }, context);
      if (gate?.block) return { details: { ok: false, error: gate.blockReason } };
      gate?.requireApproval?.onResolution("allow-once");
      return tool.execute(toolCallId, gate?.params ?? params, signal);
    } }];
  }));
}

beforeEach(async () => {
  vi.mocked(readCredential).mockClear();
  vi.mocked(exchangeRefreshToken).mockClear();
  workspaceDir = await mkdtemp(join(tmpdir(), "msgraph-v2-media-write-"));
  await mkdir(join(workspaceDir, "media", "inbound"), { recursive: true });
  await writeFile(join(workspaceDir, "media", "inbound", "fixture.bin"), fixtureBytes);
  await writeFile(join(workspaceDir, "media", "inbound", "clip.mp4"), mp4Bytes);
  await writeFile(join(workspaceDir, "media", "inbound", "large.mp4"), largeMp4Bytes);
});

afterEach(async () => {
  vi.restoreAllMocks();
  await rm(workspaceDir, { recursive: true, force: true });
});

describe("protected media write inputs", () => {
  it.each(["onedrive_upload", "onedrive_update"])("authorizes %s before opening protected media", async (toolName) => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const response = await registeredTools("unauthorized-agent")[toolName].execute("unauthorized", {
      rootLabel: "synthetic_documents",
      relativePath: "missing.mp4",
      sourceMediaUri: "media://inbound/missing.mp4",
      contentType: "video/mp4",
    });
    expect(response.details).toMatchObject({ ok: false, error: "access_denied" });
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it.each(["symlink", "hardlink", "missing-workspace"])("rejects unsafe OneDrive %s sources before credentials or network", async (kind) => {
    const inbound = join(workspaceDir, "media", "inbound");
    let uri = "media://inbound/fixture.bin";
    let runtimeWorkspace: string | null = workspaceDir;
    if (kind === "symlink") {
      await symlink(join(inbound, "fixture.bin"), join(inbound, "linked.bin"));
      uri = "media://inbound/linked.bin";
    } else if (kind === "hardlink") {
      await link(join(inbound, "fixture.bin"), join(inbound, "hardlinked.bin"));
      uri = "media://inbound/hardlinked.bin";
    } else {
      runtimeWorkspace = null;
    }
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const response = await registeredTools("main", runtimeWorkspace).onedrive_upload.execute("unsafe", {
      rootLabel: "synthetic_documents",
      relativePath: "unsafe.bin",
      sourceMediaUri: uri,
      sourceSha256: digest(fixtureBytes),
      sourceByteSize: fixtureBytes.byteLength,
    });
    expect(response.details).toMatchObject({ ok: false, error: "invalid_source_media_uri" });
    expect(readCredential).not.toHaveBeenCalled();
    expect(exchangeRefreshToken).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it.each([
    ["outlook_calendar_write", { action: "attach", eventId: "event-1", attachmentName: "missing.bin", attachmentMediaUri: "media://inbound/missing.bin" }],
    ["outlook_mail_write", { action: "add_attachment", messageId: "message-1", attachmentName: "missing.bin", attachmentMediaUri: "media://inbound/missing.bin" }],
    ["microsoft_todo_write", { action: "add_attachment", listId: "list-1", taskId: "task-1", attachmentName: "missing.bin", attachmentMediaUri: "media://inbound/missing.bin" }],
  ])("authorizes %s before opening protected media", async (toolName, params) => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const response = await registeredTools("unauthorized-agent")[toolName].execute("unauthorized", params);
    expect(response.details).toMatchObject({ ok: false, error: "access_denied" });
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it.each([
    ["outlook_calendar_write", { action: "attach", eventId: "event-1", attachmentName: "fixture.bin", attachmentMediaUri: "media://inbound/fixture.bin" }],
    ["outlook_mail_write", { action: "add_attachment", messageId: "message-1", attachmentName: "fixture.bin", attachmentMediaUri: "media://inbound/fixture.bin" }],
  ])("uses private media for %s and converts only at the Graph wire boundary", async (toolName, params) => {
    let wireBody: any;
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      wireBody = JSON.parse(String(init?.body));
      return Response.json({ id: "attachment-1", name: "fixture.bin", contentType: "application/octet-stream", size: fixtureBytes.byteLength }, { status: 201 });
    });
    const response = await registeredTools()[toolName].execute("attach", params);
    expect(response.details).toMatchObject({ ok: true, upload_mode: "direct", attachment: { id: "attachment-1", size: fixtureBytes.byteLength } });
    expect(wireBody.contentBytes).toBe(fixtureBytes.toString("base64"));
    expect(JSON.stringify(response.details)).not.toContain(wireBody.contentBytes);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it("uses private media for To Do attachments after the existing ownership check", async () => {
    let wireBody: any;
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      if (init?.method !== "POST") return Response.json({ id: "list-1", isOwner: true, isShared: false });
      wireBody = JSON.parse(String(init.body));
      return Response.json({ id: "attachment-1", name: "fixture.bin", contentType: "application/octet-stream", size: fixtureBytes.byteLength }, { status: 201 });
    });
    const response = await registeredTools().microsoft_todo_write.execute("attach", { action: "add_attachment", listId: "list-1", taskId: "task-1", attachmentName: "fixture.bin", attachmentMediaUri: "media://inbound/fixture.bin" });
    expect(response.details).toMatchObject({ ok: true, upload_mode: "direct", attachment: { id: "attachment-1", size: fixtureBytes.byteLength } });
    expect(wireBody.contentBytes).toBe(fixtureBytes.toString("base64"));
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });

  it("returns a self-verifying OneDrive receipt from one content PUT", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      expect(await requestBodyBytes(init?.body)).toEqual(fixtureBytes);
      return Response.json({ id: "item-1", name: "fixture.bin", webUrl: "https://onedrive.live.com/example", size: fixtureBytes.byteLength, file: { mimeType: "application/octet-stream" } }, { status: 201 });
    });
    const response = await registeredTools().onedrive_upload.execute("upload", { rootLabel: "synthetic_documents", relativePath: "fixture.bin", sourceMediaUri: "media://inbound/fixture.bin", sourceSha256: digest(fixtureBytes), sourceByteSize: fixtureBytes.byteLength });
    expect(response.details).toMatchObject({ ok: true, source_byte_size: fixtureBytes.byteLength, source_sha256: "5d8d910591d272938aef5f966e0816e374beaf7b5adf02cca5f8f770596c2ce3", graph_reported_size: fixtureBytes.byteLength, size_match: true, item: { id: "item-1", web_url: "https://onedrive.live.com/example" } });
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it.each([
    ["onedrive_upload", false],
    ["onedrive_update", true],
  ] as const)("preserves protected MP4 bytes, MIME, and %s write semantics", async (toolName, update) => {
    let wireBytes: Uint8Array = new Uint8Array();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      if (init?.method !== "PUT") return Response.json({ eTag: "synthetic-etag", file: {} }, { status: 200 });
      wireBytes = await requestBodyBytes(init.body);
      return Response.json({ id: "item-mp4", name: "clip.mp4", size: mp4Bytes.byteLength, file: { mimeType: "video/mp4" } }, { status: update ? 200 : 201 });
    });
    const response = await registeredTools()[toolName].execute(update ? "update" : "upload", {
      rootLabel: "synthetic_documents",
      relativePath: "clip.mp4",
      sourceMediaUri: "media://inbound/clip.mp4",
      sourceSha256: digest(mp4Bytes),
      sourceByteSize: mp4Bytes.byteLength,
      contentType: "video/mp4",
    });
    expect(response.details).toMatchObject({
      ok: true,
      operation: update ? "update" : "upload",
      source_byte_size: mp4Bytes.byteLength,
      graph_reported_size: mp4Bytes.byteLength,
      size_match: true,
      item: { id: "item-mp4", name: "clip.mp4", mime_type: "video/mp4" },
    });
    const put = fetchSpy.mock.calls.find(([, init]) => init?.method === "PUT");
    expect(put).toBeDefined();
    const headers = put![1]?.headers as Record<string, string>;
    expect(headers.authorization).toBe("Bearer synthetic-token");
    expect(headers["content-type"]).toBe("video/mp4");
    expect(headers[update ? "if-match" : "if-none-match"]).toBe(update ? "synthetic-etag" : "*");
    expect(headers[update ? "if-none-match" : "if-match"]).toBeUndefined();
    expect(wireBytes).toEqual(mp4Bytes);
    expect(JSON.stringify(response.details)).not.toContain(mp4Bytes.toString("base64"));
    expect(JSON.stringify(response.details)).not.toContain(workspaceDir);
    expect(fetchSpy).toHaveBeenCalledTimes(update ? 2 : 1);
  });

  it.each([
    ["onedrive_upload", false],
    ["onedrive_update", true],
  ] as const)("streams an above-1-MiB MP4 with digest integrity for %s", async (toolName, update) => {
    const requestChunks: number[] = [];
    const wireHash = createHash("sha256");
    let wireByteLength = 0;
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      if (init?.method !== "PUT") return Response.json({ eTag: "large-etag", file: {} });
      const chunks: Buffer[] = [];
      for await (const chunk of init.body as unknown as AsyncIterable<Uint8Array>) {
        const bytes = Buffer.from(chunk);
        requestChunks.push(bytes.byteLength);
        chunks.push(bytes);
        wireHash.update(bytes);
        wireByteLength += bytes.byteLength;
      }
      expect(Buffer.concat(chunks).byteLength).toBe(largeMp4Bytes.byteLength);
      return Response.json({ id: "large-item", name: "large.mp4", size: largeMp4Bytes.byteLength, file: { mimeType: "video/mp4" } }, { status: update ? 200 : 201 });
    });
    const response = await registeredTools()[toolName].execute(update ? "large-update" : "large-upload", {
      rootLabel: "synthetic_documents",
      relativePath: "large.mp4",
      sourceMediaUri: "media://inbound/large.mp4",
      sourceSha256: digest(largeMp4Bytes),
      sourceByteSize: largeMp4Bytes.byteLength,
      contentType: "video/mp4",
    });
    expect(response.details).toMatchObject({
      ok: true,
      upload_mode: "simple",
      source_byte_size: largeMp4Bytes.byteLength,
      source_sha256: createHash("sha256").update(largeMp4Bytes).digest("hex"),
      graph_reported_size: largeMp4Bytes.byteLength,
      size_match: true,
    });
    expect(requestChunks).toEqual([1024 * 1024 + 17]);
    expect(wireByteLength).toBe(largeMp4Bytes.byteLength);
    expect(wireHash.digest("hex")).toBe(createHash("sha256").update(largeMp4Bytes).digest("hex"));
    expect(JSON.stringify(response.details)).not.toContain(largeMp4Bytes.subarray(0, 64).toString("base64"));
    expect(JSON.stringify(response.details)).not.toContain(workspaceDir);
    expect(fetchSpy).toHaveBeenCalledTimes(update ? 2 : 1);
  });

  it("rejects an unsupported MIME from a protected source before Graph", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const response = await registeredTools().onedrive_upload.execute("unsupported", {
      rootLabel: "synthetic_documents",
      relativePath: "clip.webm",
      sourceMediaUri: "media://inbound/clip.mp4",
      sourceSha256: digest(mp4Bytes),
      sourceByteSize: mp4Bytes.byteLength,
      contentType: "video/webm",
    });
    expect(response.details).toMatchObject({ ok: false, error: "invalid_write_input" });
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});
