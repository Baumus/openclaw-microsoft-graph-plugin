import { createHash } from "node:crypto";
import { beforeEach, describe, expect, it, vi } from "vitest";

// Scale only this isolated test module's provider limit to exercise overflow
// without allocating the production 250 GiB ceiling. Transport code stays real.
vi.mock("./graph.js", async (importOriginal) => ({
  ...await importOriginal<typeof import("./graph.js")>(),
  ONEDRIVE_READ_MAX_BYTES: 16,
}));

vi.mock("openclaw/plugin-sdk/media-store", () => ({
  saveMediaStream: vi.fn(async (source: AsyncIterable<Uint8Array>, contentType: string) => {
    let size = 0;
    for await (const chunk of source) size += chunk.byteLength;
    return { id: "synthetic-id", path: "/synthetic/private/not-for-model", size, contentType };
  }),
}));

import { saveMediaStream } from "openclaw/plugin-sdk/media-store";
import { downloadOneDriveFile } from "./attachment-download.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

const root = graphPolicyFixture().services.onedrive.allowed_roots[0];
const options = { root, relativePath: "opaque.custom", token: "synthetic-token" };

function fetchSource(size: number, body: ReadableStream<Uint8Array> | null) {
  return vi.fn<typeof fetch>(async (input) => String(input).endsWith("/content")
    ? new Response(body)
    : Response.json({ id: "synthetic-item", name: "opaque.custom", size, file: { mimeType: "application/x-custom" } }));
}

beforeEach(() => vi.clearAllMocks());

describe("OneDrive opaque stream integrity guards", () => {
  it("hashes all chunks at the exact byte ceiling and omits the store path", async () => {
    const bytes = Uint8Array.from({ length: 16 }, (_, i) => i);
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(bytes.subarray(0, 5));
        controller.enqueue(bytes.subarray(5));
        controller.close();
      },
    });
    const result = await downloadOneDriveFile({ ...options, fetchFn: fetchSource(16, body) });
    expect(result).toMatchObject({ bytes: 16, sha256: createHash("sha256").update(bytes).digest("hex"), artifact: { mimeType: "application/x-custom", sizeBytes: 16 }, media: { outbound: false } });
    expect(JSON.stringify(result)).not.toContain("/synthetic/private");
    expect(saveMediaStream).toHaveBeenCalledWith(expect.anything(), "application/x-custom", "inbound", 16, "opaque.custom", "opaque.custom");
  });

  it("cancels an oversized stream even when provider metadata understates its size", async () => {
    const cancel = vi.fn();
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(16));
        controller.enqueue(new Uint8Array([1]));
      },
      cancel,
    });
    await expect(downloadOneDriveFile({ ...options, fetchFn: fetchSource(1, body) })).rejects.toThrow("provider_file_too_large");
    expect(cancel).toHaveBeenCalledTimes(1);
  });

  it("rejects provider metadata over the limit without opening a content stream", async () => {
    const fetchFn = fetchSource(17, null);
    await expect(downloadOneDriveFile({ ...options, fetchFn })).rejects.toThrow("provider_file_too_large");
    expect(fetchFn).toHaveBeenCalledTimes(1);
    expect(saveMediaStream).not.toHaveBeenCalled();
  });

  it("rejects a missing nonempty body", async () => {
    await expect(downloadOneDriveFile({ ...options, fetchFn: fetchSource(1, null) })).rejects.toThrow("invalid_provider_response");
  });

  it("rejects metadata/stream length mismatch", async () => {
    const body = new ReadableStream<Uint8Array>({ start(controller) { controller.enqueue(new Uint8Array([1])); controller.close(); } });
    await expect(downloadOneDriveFile({ ...options, fetchFn: fetchSource(2, body) })).rejects.toThrow("invalid_provider_response");
  });

  it("rejects a media-store byte count that differs from the streamed bytes", async () => {
    vi.mocked(saveMediaStream).mockImplementationOnce(async (source) => {
      for await (const _chunk of source as AsyncIterable<Uint8Array>) { /* consume the real source */ }
      return { id: "synthetic-id", path: "/synthetic/private", size: 0, contentType: "application/x-custom" };
    });
    const body = new ReadableStream<Uint8Array>({ start(controller) { controller.enqueue(new Uint8Array([1])); controller.close(); } });
    await expect(downloadOneDriveFile({ ...options, fetchFn: fetchSource(1, body) })).rejects.toThrow("invalid_provider_response");
  });
});
