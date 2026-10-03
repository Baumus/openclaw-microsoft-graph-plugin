import { createHash } from "node:crypto";
import { chmod, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

let mediaRoot = "";

vi.mock("openclaw/plugin-sdk/media-store", () => ({
  saveMediaStream: vi.fn(async (stream: AsyncIterable<unknown>, contentType: string, _subdir: string, maxBytes: number, originalFilename: string) => {
    const chunks: Buffer[] = [];
    let size = 0;
    for await (const chunk of stream) {
      const bytes = Buffer.from(chunk as Uint8Array);
      size += bytes.byteLength;
      if (size > maxBytes) throw new Error("file_too_large");
      chunks.push(bytes);
    }
    const path = join(mediaRoot, originalFilename);
    await writeFile(path, Buffer.concat(chunks), { mode: 0o600 });
    await chmod(path, 0o600);
    return { id: `synthetic-${originalFilename}`, path, size, contentType };
  }),
}));

vi.mock("./credential.js", () => ({
  readCredential: vi.fn(async () => ({ clientId: "synthetic", refreshToken: "synthetic", tenant: "common", scopes: ["Mail.Read", "Calendars.Read", "Files.Read", "Tasks.Read"] })),
  selectScope: (_credential: unknown, allowed: string[]) => allowed[0],
  exchangeRefreshToken: vi.fn(async () => "synthetic-access-token"),
  tokenForAuthorizedOperation: vi.fn(async () => "synthetic-access-token"),
}));

import { sanitizeAttachmentMime, sanitizeAttachmentName } from "./attachment-download.js";
import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function registeredTool(name: string, config: Record<string, unknown> = {}, agentId = "main") {
  const factories: Array<(context: any) => any> = [];
  const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
  entry.register({
    pluginConfig: { enabled: true, policy: graphPolicyFixture(), ...config },
    registerTool: (factory: any) => factories.push(factory),
    on: vi.fn(),
    logger,
  } as any);
  const tools = factories.map((factory) => factory({ agentId, sessionId: "session-1", workspaceDir: mediaRoot })).filter(Boolean);
  const tool = tools.find((candidate) => candidate.name === name);
  if (!tool) throw new Error(`missing tool ${name}`);
  return { tool, logger };
}

beforeEach(async () => {
  mediaRoot = await mkdtemp(join(tmpdir(), "msgraph-outlook-download-"));
});

afterEach(async () => {
  vi.restoreAllMocks();
  await rm(mediaRoot, { recursive: true, force: true });
});

describe.each([
  {
    label: "message",
    toolName: "outlook_mail_read",
    params: { action: "download_attachment", messageId: "message-1", attachmentId: "attachment-1" },
    metadataPath: "/me/messages/message-1/attachments/attachment-1?$select=id,name,contentType,size,isInline",
    binaryPath: "/me/messages/message-1/attachments/attachment-1/$value",
  },
  {
    label: "event",
    toolName: "outlook_calendar_read",
    params: { action: "download_attachment", eventId: "event-1", attachmentId: "attachment-1" },
    metadataPath: "/me/events/event-1/attachments/attachment-1?$select=id,name,contentType,size,isInline",
    binaryPath: "/me/events/event-1/attachments/attachment-1/$value",
  },
])("Outlook $label attachment download", ({ toolName, params, metadataPath, binaryPath }) => {
  it("performs one bounded binary download and returns a private OpenClaw file artifact without Base64", async () => {
    const bytes = Buffer.from("synthetic complete attachment");
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = String(input);
      if (url.endsWith(binaryPath)) return new Response(bytes, { headers: { "content-length": String(bytes.byteLength), "content-type": "application/pdf" } });
      if (url.endsWith(metadataPath)) return new Response(JSON.stringify({
        "@odata.type": "#microsoft.graph.fileAttachment",
        id: "attachment-1",
        name: "report?.pdf",
        contentType: "Application/PDF",
        size: bytes.byteLength,
        isInline: false,
      }), { headers: { "content-type": "application/json" } });
      throw new Error(`unexpected synthetic URL: ${url}`);
    });

    const { tool, logger } = registeredTool(toolName, { maxAttachmentDownloadBytes: 1024 });
    const response = await tool.execute("download", params);

    expect(response.details).toMatchObject({
      ok: true,
      action: "download_attachment",
      attachment: { id: "attachment-1", name: "report_.pdf", contentType: "application/pdf", size: bytes.byteLength, bytes: bytes.byteLength, isInline: false },
      sourceSha256: createHash("sha256").update(bytes).digest("hex"),
      sourceByteSize: bytes.byteLength,
      artifact: { type: "file", id: "synthetic-report_.pdf", uri: "media://inbound/synthetic-report_.pdf", mimeType: "application/pdf", name: "report_.pdf", sizeBytes: bytes.byteLength },
      media: {
        outbound: false,
        attachments: [{ type: "file", id: "synthetic-report_.pdf", uri: "media://inbound/synthetic-report_.pdf", mimeType: "application/pdf", name: "report_.pdf", sizeBytes: bytes.byteLength }],
      },
    });
    const artifactPath = join(mediaRoot, "report_.pdf");
    expect(await readFile(artifactPath)).toEqual(bytes);
    expect((await stat(artifactPath)).mode & 0o777).toBe(0o600);
    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(fetchSpy.mock.calls.filter(([input]) => String(input).endsWith("/$value"))).toHaveLength(1);
    expect(JSON.stringify(response.details)).not.toContain("contentBytes");
    expect(JSON.stringify(response.details)).not.toContain(bytes.toString("base64"));
    expect(JSON.stringify(response.details)).not.toContain(mediaRoot);
    expect(logger.info.mock.calls.flat().join(" ")).not.toContain(bytes.toString("base64"));
  });

  it("rejects item and reference attachments before any binary request", async () => {
    for (const type of ["#microsoft.graph.itemAttachment", "#microsoft.graph.referenceAttachment"]) {
      const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(new Response(JSON.stringify({
        "@odata.type": type,
        id: "attachment-1",
        name: "unsafe.bin",
        contentType: "application/octet-stream",
        size: 1,
      })));
      const { tool } = registeredTool(toolName);
      const response = await tool.execute("reject", params);
      expect(response.details).toMatchObject({ ok: false, error: "unsupported_attachment_type" });
      expect(fetchSpy).toHaveBeenCalledTimes(1);
      fetchSpy.mockRestore();
    }
  });
});

describe("private media download consistency", () => {
  it("publishes OneDrive MP4 downloads as private artifacts without model-facing file bytes", async () => {
    const bytes = Buffer.alloc(1024 * 1024 + 17, 0x61);
    Buffer.from([0, 0, 0, 24, 102, 116, 121, 112, 109, 112, 52, 50]).copy(bytes);
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(bytes.subarray(0, 700_000));
        controller.enqueue(bytes.subarray(700_000));
        controller.close();
      },
    });
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = String(input);
      if (url.includes("/content")) return new Response(body, { headers: { "content-type": "video/mp4" } });
      return Response.json({ id: "drive-item", name: "clip.mp4", size: bytes.byteLength, file: { mimeType: "video/mp4" } });
    });
    const { tool } = registeredTool("onedrive_download", { maxReadOutputBytes: 1024 }, "fixture-reader");
    const response = await tool.execute("drive-download", { rootLabel: "synthetic_documents", relativePath: "clip.mp4" });
    expect(response.details).toMatchObject({ ok: true, operation: "download", bytes: bytes.byteLength, sha256: createHash("sha256").update(bytes).digest("hex"), artifact: { uri: "media://inbound/synthetic-clip.mp4", mimeType: "video/mp4", sizeBytes: bytes.byteLength }, media: { outbound: false } });
    expect(await readFile(join(mediaRoot, "clip.mp4"))).toEqual(bytes);
    expect(JSON.stringify(response.details)).not.toContain(bytes.toString("base64"));
    expect(JSON.stringify(response.details)).not.toContain(mediaRoot);
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  }, 15_000);

  it("publishes To Do attachment reads as private artifacts and removes Graph wire bytes", async () => {
    const bytes = Buffer.from("synthetic todo attachment");
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(Response.json({ id: "attachment-1", name: "todo.txt", contentType: "text/plain", size: bytes.byteLength, contentBytes: bytes.toString("base64") }));
    const { tool } = registeredTool("microsoft_todo_read", { maxReadOutputBytes: 1024 });
    const response = await tool.execute("todo-download", { action: "get_attachment", listId: "list-1", taskId: "task-1", attachmentId: "attachment-1" });
    expect(response.details).toMatchObject({ ok: true, action: "get_attachment", attachment: { id: "attachment-1", size: bytes.byteLength }, artifact: { uri: "media://inbound/synthetic-todo.txt", mimeType: "text/plain", sizeBytes: bytes.byteLength }, media: { outbound: false } });
    expect(JSON.stringify(response.details)).not.toContain("contentBytes");
    expect(JSON.stringify(response.details)).not.toContain(bytes.toString("base64"));
    expect(JSON.stringify(response.details)).not.toContain(mediaRoot);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });
});

describe("Outlook attachment transport boundaries", () => {
  const params = { action: "download_attachment", messageId: "message-1", attachmentId: "attachment-1" };
  const metadataPath = "/me/messages/message-1/attachments/attachment-1?$select=id,name,contentType,size,isInline";
  const binaryPath = "/me/messages/message-1/attachments/attachment-1/$value";

  function metadata(overrides: Record<string, unknown> = {}) {
    return {
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "attachment-1",
      name: "report.pdf",
      contentType: "application/pdf",
      size: 3,
      ...overrides,
    };
  }

  it("accepts provider size metadata that differs from the actual bounded stream", async () => {
    const bytes = Buffer.from("pdf");
    vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = String(input);
      if (url.endsWith(metadataPath)) return new Response(JSON.stringify(metadata({ size: 99_999 })));
      if (url.endsWith(binaryPath)) return new Response(bytes, { headers: { "content-length": String(bytes.byteLength) } });
      throw new Error(`unexpected synthetic URL: ${url}`);
    });
    const { tool } = registeredTool("outlook_mail_read", { maxAttachmentDownloadBytes: 1024 });
    const response = await tool.execute("mismatch", params);
    expect(response.details).toMatchObject({ ok: true, attachment: { size: 99_999, bytes: 3 } });
  });

  it("rejects an oversized Content-Length before consuming the binary body", async () => {
    const cancel = vi.fn();
    const body = new ReadableStream<Uint8Array>({ cancel });
    vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = String(input);
      if (url.endsWith(metadataPath)) return new Response(JSON.stringify(metadata()));
      if (url.endsWith(binaryPath)) return new Response(body, { headers: { "content-length": "1025" } });
      throw new Error(`unexpected synthetic URL: ${url}`);
    });
    const { tool } = registeredTool("outlook_mail_read", { maxAttachmentDownloadBytes: 1024 });
    const response = await tool.execute("oversized-length", params);
    expect(response.details).toMatchObject({ ok: false, error: "file_too_large" });
    expect(cancel).toHaveBeenCalledTimes(1);
  });

  it("rejects a body that exceeds the byte limit when Content-Length is absent", async () => {
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(1024));
        controller.enqueue(new Uint8Array([1]));
        controller.close();
      },
    });
    vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = String(input);
      if (url.endsWith(metadataPath)) return new Response(JSON.stringify(metadata()));
      if (url.endsWith(binaryPath)) return new Response(body);
      throw new Error(`unexpected synthetic URL: ${url}`);
    });
    const { tool } = registeredTool("outlook_mail_read", { maxAttachmentDownloadBytes: 1024 });
    const response = await tool.execute("oversized-stream", params);
    expect(response.details).toMatchObject({ ok: false, error: "file_too_large" });
  });

  it("aborts a stalled response body at the read-idle timeout", async () => {
    const body = new ReadableStream<Uint8Array>({ start() { /* intentionally idle */ } });
    vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = String(input);
      if (url.endsWith(metadataPath)) return new Response(JSON.stringify(metadata()));
      if (url.endsWith(binaryPath)) return new Response(body);
      throw new Error(`unexpected synthetic URL: ${url}`);
    });
    const { tool } = registeredTool("outlook_mail_read", { maxAttachmentDownloadBytes: 1024, requestTimeoutMs: 10 });
    const response = await tool.execute("stalled-stream", params);
    expect(response.details).toMatchObject({ ok: false, error: "request_timeout" });
  });

  it("rejects a truncated body against an explicit Content-Length", async () => {
    const bytes = Buffer.from("short");
    vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = String(input);
      if (url.endsWith(metadataPath)) return new Response(JSON.stringify(metadata({ size: 12 })));
      if (url.endsWith(binaryPath)) return new Response(bytes, { headers: { "content-length": "12" } });
      throw new Error(`unexpected synthetic URL: ${url}`);
    });
    const { tool } = registeredTool("outlook_mail_read", { maxAttachmentDownloadBytes: 1024 });
    const response = await tool.execute("truncated", params);
    expect(response.details).toMatchObject({ ok: false, error: "invalid_provider_response" });
  });

  it("rejects path-bearing names and malformed MIME values before the binary request", async () => {
    for (const overrides of [{ name: "%252e%252e%252fsecret.pdf" }, { contentType: "application/pdf\r\nx-test: injected" }]) {
      const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(new Response(JSON.stringify(metadata(overrides))));
      const { tool } = registeredTool("outlook_mail_read");
      const response = await tool.execute("unsafe-metadata", params);
      expect(response.details.ok).toBe(false);
      expect(fetchSpy).toHaveBeenCalledTimes(1);
      fetchSpy.mockRestore();
    }
  });

  it("builds the exact calendar-scoped attachment route", async () => {
    const bytes = Buffer.from("pdf");
    const metadataScoped = "/me/calendars/synthetic-calendar/events/event-1/attachments/attachment-1?$select=id,name,contentType,size,isInline";
    const binaryScoped = "/me/calendars/synthetic-calendar/events/event-1/attachments/attachment-1/$value";
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = String(input);
      if (url.endsWith(metadataScoped)) return new Response(JSON.stringify(metadata()));
      if (url.endsWith(binaryScoped)) return new Response(bytes, { headers: { "content-length": "3" } });
      throw new Error(`unexpected synthetic URL: ${url}`);
    });
    const { tool } = registeredTool("outlook_calendar_read", { maxAttachmentDownloadBytes: 1024 });
    const response = await tool.execute("calendar-scoped", { action: "download_attachment", calendarId: "synthetic-calendar", eventId: "event-1", attachmentId: "attachment-1" });
    expect(response.details).toMatchObject({ ok: true });
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });

  it("sanitizes leaf names while rejecting traversal and header-injection MIME values", () => {
    expect(sanitizeAttachmentName(" quarterly?.pdf ")).toBe("quarterly_.pdf");
    expect(() => sanitizeAttachmentName("..%252fsecret.pdf")).toThrow("invalid_attachment_name");
    expect(sanitizeAttachmentMime(undefined)).toBe("application/octet-stream");
    expect(() => sanitizeAttachmentMime("text/plain\r\nx-test: injected")).toThrow("invalid_provider_response");
  });
});
