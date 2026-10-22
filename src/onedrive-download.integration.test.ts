import { createHash } from "node:crypto";
import { mkdtemp, readFile, readdir, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("./credential.js", () => ({
  readCredential: vi.fn(),
  selectScope: vi.fn(),
  exchangeRefreshToken: vi.fn(),
  tokenForAuthorizedOperation: vi.fn(async () => "synthetic-access-token"),
}));

import { downloadOneDriveFile } from "./attachment-download.js";
import { exchangeRefreshToken, readCredential, tokenForAuthorizedOperation } from "./credential.js";
import { contentTypeAllowed, driveRead, ONEDRIVE_READ_MAX_BYTES, validTransferContentType, validateDriveWriteBytes } from "./graph.js";
import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

// Deliberately use the real pinned SDK media store, isolated from host state.
let stateRoot = "";
const root = graphPolicyFixture().services.onedrive.allowed_roots[0];
const params = { rootLabel: root.label, relativePath: "report.docx" };

function registeredDownload(policy = graphPolicyFixture(), agentId = "fixture-reader") {
  const factories: Array<(context: any) => any> = [];
  const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
  entry.register({
    pluginConfig: { enabled: true, policy },
    registerTool: (factory: any) => factories.push(factory),
    on: vi.fn(),
    logger,
  } as any);
  const tool = factories.map((factory) => factory({ agentId, sessionId: "synthetic-download-session" }))
    .find((candidate) => candidate?.name === "onedrive_download");
  if (!tool) throw new Error("missing download tool");
  return { tool, logger };
}

function transport(metadata: Record<string, unknown>, bytes = Buffer.from("opaque synthetic payload")) {
  return vi.spyOn(globalThis, "fetch").mockImplementation(async (input, init) => {
    expect(String(input)).toMatch(/^https:\/\/graph\.microsoft\.com\/v1\.0\/drives\/synthetic-drive\/items\/synthetic-root:/);
    expect(new Headers(init?.headers).get("authorization")).toBe("Bearer synthetic-access-token");
    if (String(input).endsWith("/content")) {
      expect(init?.redirect).toBe("follow");
      return new Response(bytes);
    }
    return Response.json(metadata);
  });
}

beforeEach(async () => {
  stateRoot = await mkdtemp(join(tmpdir(), "msgraph-generic-download-"));
  vi.stubEnv("OPENCLAW_STATE_DIR", stateRoot);
  vi.clearAllMocks();
});

afterEach(async () => {
  vi.restoreAllMocks();
  vi.unstubAllEnvs();
  await rm(stateRoot, { recursive: true, force: true });
});

describe("OneDrive generic opaque downloads through the real private-media SDK", () => {
  it.each([
    ["report.doc", "application/msword"],
    ["report?.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
    ["slides.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"],
    ["opaque.custom", "application/x-synthetic-custom-binary"],
    ["page.html", "Text/HTML; charset=UTF-8"],
    ["unknown.custom", undefined],
    ["empty.custom", "application/x-empty"],
  ])("downloads %s without type eligibility filtering or content execution", async (name, mimeType) => {
    const bytes = name === "empty.custom" ? Buffer.alloc(0)
      : name === "page.html" ? Buffer.from("<html><script>globalThis.syntheticDownloadExecuted = true</script>untrusted synthetic bytes</html>")
      : Buffer.concat([Buffer.from([0, 255, 128, 1]), Buffer.from("opaque synthetic binary payload")]);
    const metadata = { id: "synthetic-item", name, size: bytes.byteLength, file: { mimeType } };
    const fetchSpy = transport(metadata, bytes);
    const { tool, logger } = registeredDownload();
    const response = await tool.execute("generic-download", { ...params, relativePath: name });

    expect(response.details).toMatchObject({
      ok: true,
      operation: "download",
      bytes: bytes.byteLength,
      sha256: createHash("sha256").update(bytes).digest("hex"),
      artifact: { type: "file", name: name.replace("?", "_"), sizeBytes: bytes.byteLength },
      media: { outbound: false },
    });
    const artifact = response.details.artifact;
    expect(artifact.uri).toBe("media://inbound/" + artifact.id);
    expect(artifact.mimeType).toBe((mimeType ?? "application/octet-stream").split(";")[0].toLowerCase());
    expect(response.details.media.attachments).toEqual([artifact]);
    expect(await readFile(join(stateRoot, "media", "inbound", artifact.id))).toEqual(bytes);
    expect((await stat(join(stateRoot, "media", "inbound"))).mode & 0o077).toBe(0);
    const modelOutput = JSON.stringify(response);
    const logOutput = JSON.stringify(logger.info.mock.calls);
    for (const output of [modelOutput, logOutput]) {
      expect(output).not.toContain(stateRoot);
      expect(output).not.toContain("synthetic-access-token");
      expect(output).not.toContain("syntheticDownloadExecuted");
      expect(output).not.toContain("contentBytes");
      if (bytes.byteLength) expect(output).not.toContain(bytes.toString("base64"));
    }
    expect((globalThis as any).syntheticDownloadExecuted).toBeUndefined();
    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(tokenForAuthorizedOperation).toHaveBeenCalledTimes(1);
    expect(tokenForAuthorizedOperation).toHaveBeenCalledWith(expect.objectContaining({ allowedScopes: ["Files.Read"] }));
  });

  it.each([
    ["empty MIME", { file: { mimeType: "" } }, "invalid_provider_response"],
    ["malformed MIME", { file: { mimeType: "not-a-mime" } }, "invalid_provider_response"],
    ["MIME header injection", { file: { mimeType: "text/html\r\nx-injected: yes" } }, "invalid_provider_response"],
    ["non-string MIME", { file: { mimeType: 42 } }, "invalid_provider_response"],
    ["unsafe MIME parameter", { file: { mimeType: "text/html; bad" } }, "invalid_provider_response"],
    ["folder", { folder: {} }, "invalid_provider_response"],
    ["missing file facet", { file: undefined }, "invalid_provider_response"],
    ["negative size", { size: -1 }, "invalid_provider_response"],
    ["fractional size", { size: 1.5 }, "invalid_provider_response"],
    ["oversized metadata", { size: ONEDRIVE_READ_MAX_BYTES + 1 }, "provider_file_too_large"],
    ["encoded path name", { name: "%252e%252e%252fescape.docx" }, "invalid_attachment_name"],
    ["path-bearing name", { name: "../escape.docx" }, "invalid_attachment_name"],
  ])("rejects %s before requesting binary bytes", async (_label, overrides, error) => {
    const fetchSpy = transport({ id: "synthetic-item", name: "report.docx", size: 24, file: { mimeType: "application/msword" }, ...overrides });
    const { tool } = registeredDownload();
    expect((await tool.execute("negative-download", params)).details).toMatchObject({ ok: false, error });
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(await readdir(stateRoot)).toEqual([]);
  });

  it.each(["unknown agent", "root read disabled", "agent read disabled", "unknown root"])("keeps %s denied before credentials or transport", async (kind) => {
    const policy = graphPolicyFixture();
    const allowed = policy.services.onedrive.allowed_roots[0];
    if (kind === "root read disabled") allowed.permissions.read = false;
    if (kind === "agent read disabled") allowed.agents["fixture-reader"].permissions.read = false;
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockRejectedValue(new Error("unexpected network"));
    const { tool } = registeredDownload(policy, kind === "unknown agent" ? "ungranted" : "fixture-reader");
    const response = await tool.execute("denied-download", { ...params, ...(kind === "unknown root" ? { rootLabel: "ungranted" } : {}) });
    expect(response.details).toMatchObject({ ok: false, error: "access_denied" });
    expect(tokenForAuthorizedOperation).not.toHaveBeenCalled();
    expect(readCredential).not.toHaveBeenCalled();
    expect(exchangeRefreshToken).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it.each(["../escape.docx", "/absolute.docx", "nested/../../escape.docx", "nested\\escape.docx", "bad\u0000.docx"])("retains path rejection for %s before transport", async (relativePath) => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockRejectedValue(new Error("unexpected network"));
    const { tool } = registeredDownload();
    expect((await tool.execute("invalid-path", { ...params, relativePath })).details)
      .toMatchObject({ ok: false, error: "invalid_relative_path" });
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("rejects a truncated generic file rather than returning a successful artifact", async () => {
    transport({ id: "synthetic-item", name: "report.doc", size: 100, file: { mimeType: "application/msword" } }, Buffer.from("short"));
    const { tool } = registeredDownload();
    expect((await tool.execute("truncated-download", params)).details).toMatchObject({ ok: false, error: "invalid_provider_response" });
  });

  it.each([
    "application/msword",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "application/x-synthetic-custom-binary",
  ])("keeps content-reading eligibility separate from opaque transfer for %s", async (mimeType) => {
    expect(contentTypeAllowed(mimeType)).toBe(false);
    expect(validTransferContentType(mimeType)).toBe(true);
    expect(() => validateDriveWriteBytes("report.bin", new Uint8Array([1]), mimeType, 1024)).not.toThrow();
    const fetchFn = vi.fn(async () => Response.json({ name: "report.bin", size: 1, file: { mimeType } }));
    await expect(driveRead(root, "report.bin", "synthetic-access-token", "digest", 1024, undefined, fetchFn)).rejects.toThrow("unsupported_file_type");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("retains caller cancellation before any download request", async () => {
    const controller = new AbortController();
    controller.abort(new DOMException("synthetic cancellation", "AbortError"));
    const fetchFn = vi.fn<typeof fetch>();
    await expect(downloadOneDriveFile({ root, relativePath: "report.doc", token: "synthetic-access-token", signal: controller.signal, fetchFn })).rejects.toThrow();
    expect(fetchFn).not.toHaveBeenCalled();
  });
});
