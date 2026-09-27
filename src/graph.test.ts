import { createHash } from "node:crypto";
import { describe, expect, it, vi } from "vitest";
import { ContinuationStore, normalizedCriteria } from "./continuation.js";
import { ATTACHMENT_UPLOAD_CHUNK_BYTES, base64JsonResponseLimit, canonicalAttachmentUploadUrl, canonicalGraphContinuation, decodeBase64Strict, driveCreateFolder, driveList, driveListContinuation, driveMetadataUpdate, driveRead, driveReadInstructionsCandidate, driveSearch, driveSearchContinuation, driveSearchPath, driveSearchScoped, exactFilenameQuery, graphOperationSignal, graphRequest, graphStreamRequest, driveWrite, driveWriteBytes, driveWriteSource, normalizeDriveSearch, uploadAttachmentSession, validateDriveWriteBytes, validateDriveWriteInput } from "./graph.js";
import type { AllowedRoot } from "./policy.js";

const root: AllowedRoot = {
  label: "pinned_root",
  path: "/Mutable/Display/Path",
  drive_id: "drive-id",
  item_id: "stable-root-id",
  include_descendants: true,
  permissions: { read: true, write: false, delete: false },
  agents: { main: { permissions: { read: true, write: false, delete: false } } },
};

describe("bounded Graph read retries", () => {
  it("uses the supplied access token in buffered and streaming Graph requests", async () => {
    const bufferedFetch = vi.fn().mockResolvedValue(Response.json({ ok: true }));
    await expect(graphRequest("test-access-token", "/me", {}, bufferedFetch as typeof fetch)).resolves.toEqual({ ok: true });
    expect((bufferedFetch.mock.calls[0][1]?.headers as Record<string, string>).authorization).toBe("Bearer test-access-token");

    const streamFetch = vi.fn().mockResolvedValue(new Response("bytes"));
    await expect(graphStreamRequest("test-access-token", "/me/photo/$value", {}, streamFetch as typeof fetch)).resolves.toBeInstanceOf(Response);
    expect((streamFetch.mock.calls[0][1]?.headers as Record<string, string>).authorization).toBe("Bearer test-access-token");
    expect(streamFetch.mock.calls[0][1]?.redirect).toBe("follow");
  });

  it("reads an instruction candidate in one non-retried request and treats only 404 as absent", async () => {
    const presentFetch = vi.fn().mockResolvedValue(new Response("# Root", { status: 200 }));
    await expect(driveReadInstructionsCandidate(root, "a/AGENTS.md", "token", 1024, undefined, presentFetch as typeof fetch)).resolves.toEqual(new TextEncoder().encode("# Root"));
    expect(String(presentFetch.mock.calls[0][0])).toContain("/drives/drive-id/items/stable-root-id:/a/AGENTS.md:/content");
    expect(presentFetch).toHaveBeenCalledTimes(1);

    const missingFetch = vi.fn().mockResolvedValue(new Response(null, { status: 404 }));
    await expect(driveReadInstructionsCandidate(root, "AGENTS.md", "token", 1024, undefined, missingFetch as typeof fetch)).resolves.toBeNull();
    const transientFetch = vi.fn().mockResolvedValue(new Response(null, { status: 503 }));
    await expect(driveReadInstructionsCandidate(root, "AGENTS.md", "token", 1024, undefined, transientFetch as typeof fetch)).rejects.toThrow("provider_error_503");
    expect(transientFetch).toHaveBeenCalledTimes(1);
  });

  it("clears the header timeout after a streaming response begins", async () => {
    const operation = graphOperationSignal(undefined, 1_000, 10);
    const fetchFn = vi.fn((_input: string | URL | Request, init?: RequestInit) => Promise.resolve(new Response(new ReadableStream<Uint8Array>({
      start(controller) {
        init?.signal?.addEventListener("abort", () => controller.error(init.signal?.reason), { once: true });
        setTimeout(() => { controller.enqueue(new Uint8Array([1, 2, 3])); controller.close(); }, 25);
      },
    }))));
    const response = await graphStreamRequest("token", "/me/messages/id/attachments/id/$value", { signal: operation }, fetchFn as typeof fetch);
    await expect(response.arrayBuffer()).resolves.toEqual(new Uint8Array([1, 2, 3]).buffer);
    expect(operation.aborted).toBe(false);
  });

  it.each([429, 502, 503, 504])("retries transient GET status %s only within the read budget", async (status) => {
    const delays: number[] = [];
    const fetchFn = vi.fn()
      .mockResolvedValueOnce(new Response(null, { status }))
      .mockResolvedValueOnce(new Response(JSON.stringify({ ok: true }), { status: 200 }));
    await expect(graphRequest("token", "/me", { retrySleep: async (delay) => { delays.push(delay); } }, fetchFn as typeof fetch)).resolves.toEqual({ ok: true });
    expect(fetchFn).toHaveBeenCalledTimes(2);
    expect(delays).toEqual([100]);
  });

  it("honors a bounded Retry-After and declines an over-budget delay", async () => {
    const delays: number[] = [];
    const retryingFetch = vi.fn()
      .mockResolvedValueOnce(new Response(null, { status: 429, headers: { "retry-after": "1" } }))
      .mockResolvedValueOnce(new Response(JSON.stringify({ ok: true }), { status: 200 }));
    await graphRequest("token", "/me", { retryMaxDelayMs: 2_000, retrySleep: async (delay) => { delays.push(delay); } }, retryingFetch as typeof fetch);
    expect(delays).toEqual([1_000]);

    const boundedFetch = vi.fn(async () => new Response(null, { status: 429, headers: { "retry-after": "3" } }));
    await expect(graphRequest("token", "/me", { retryMaxDelayMs: 2_000, retrySleep: async () => { throw new Error("must not sleep"); } }, boundedFetch as typeof fetch)).rejects.toThrow("provider_throttled");
    expect(boundedFetch).toHaveBeenCalledTimes(1);
  });

  it("retries narrowly classified transport failures but not arbitrary exceptions", async () => {
    const transientFetch = vi.fn().mockRejectedValueOnce(new TypeError("socket reset")).mockResolvedValueOnce(new Response(JSON.stringify({ ok: true }), { status: 200 }));
    await expect(graphRequest("token", "/me", { retrySleep: async () => undefined }, transientFetch as typeof fetch)).resolves.toEqual({ ok: true });
    expect(transientFetch).toHaveBeenCalledTimes(2);

    const permanentFetch = vi.fn(async () => { throw new Error("fixture failure"); });
    await expect(graphRequest("token", "/me", { retrySleep: async () => undefined }, permanentFetch as typeof fetch)).rejects.toThrow("provider_unavailable");
    expect(permanentFetch).toHaveBeenCalledTimes(1);
  });

  it("keeps each request timeout shorter than the whole operation deadline", async () => {
    const operation = graphOperationSignal(undefined, 1_000, 10);
    const fetchFn = vi.fn((_input: string | URL | Request, init?: RequestInit) => new Promise<Response>((_resolve, reject) => {
      const abort = () => reject(init?.signal?.reason);
      if (init?.signal?.aborted) abort();
      else init?.signal?.addEventListener("abort", abort, { once: true });
    }));
    await expect(graphRequest("token", "/me", { signal: operation, readRetries: 0 }, fetchFn as typeof fetch)).rejects.toMatchObject({ name: "TimeoutError" });
    expect(operation.aborted).toBe(false);
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("retries one per-request timeout while the whole operation remains live", async () => {
    const operation = graphOperationSignal(undefined, 1_000, 10);
    let request = 0;
    const fetchFn = vi.fn((_input: string | URL | Request, init?: RequestInit) => {
      request += 1;
      if (request === 2) return Promise.resolve(new Response(JSON.stringify({ ok: true }), { status: 200 }));
      return new Promise<Response>((_resolve, reject) => {
        const abort = () => reject(init?.signal?.reason);
        if (init?.signal?.aborted) abort();
        else init?.signal?.addEventListener("abort", abort, { once: true });
      });
    });
    await expect(graphRequest("token", "/me", { signal: operation, readRetries: 1, retrySleep: async () => undefined }, fetchFn as typeof fetch)).resolves.toEqual({ ok: true });
    expect(operation.aborted).toBe(false);
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it("retries a GET whose buffered response body times out", async () => {
    const operation = graphOperationSignal(undefined, 1_000, 10);
    let request = 0;
    const fetchFn = vi.fn((_input: string | URL | Request, init?: RequestInit) => {
      request += 1;
      if (request === 2) return Promise.resolve(new Response(JSON.stringify({ ok: true }), { status: 200 }));
      return Promise.resolve(new Response(new ReadableStream<Uint8Array>({
        start(controller) {
          const abort = () => controller.error(init?.signal?.reason);
          if (init?.signal?.aborted) abort();
          else init?.signal?.addEventListener("abort", abort, { once: true });
        },
      }), { status: 200 }));
    });
    await expect(graphRequest("token", "/me", { signal: operation, readRetries: 1, retrySleep: async () => undefined }, fetchFn as typeof fetch)).resolves.toEqual({ ok: true });
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it("lets external cancellation win over request retry handling", async () => {
    const controller = new AbortController();
    const reason = new Error("caller_cancelled");
    const operation = graphOperationSignal(controller.signal, 1_000, 500);
    const fetchFn = vi.fn((_input: string | URL | Request, init?: RequestInit) => new Promise<Response>((_resolve, reject) => {
      init?.signal?.addEventListener("abort", () => reject(init.signal?.reason), { once: true });
    }));
    const pending = graphRequest("token", "/me", { signal: operation }, fetchFn as typeof fetch);
    controller.abort(reason);
    await expect(pending).rejects.toBe(reason);
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("never retries writes", async () => {
    const fetchFn = vi.fn(async () => new Response(null, { status: 503 }));
    await expect(graphRequest("token", "/me/messages", { method: "POST", body: {}, retrySleep: async () => undefined }, fetchFn as typeof fetch)).rejects.toThrow("provider_error_503");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });
});

describe("pinned OneDrive addressing", () => {
  it("lists descendants from the pinned item id, never the mutable root path", async () => {
    const fetchFn = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => new Response(JSON.stringify({ value: [] }), { status: 200, headers: { "content-type": "application/json" } }));
    await driveList(root, "SYNTHETIC_FOLDER", "token", 5, undefined, fetchFn as typeof fetch);
    const url = String(fetchFn.mock.calls[0][0]);
    expect(url).toContain("/drives/drive-id/items/stable-root-id:/SYNTHETIC_FOLDER:/children");
    expect(url).not.toContain("Mutable");
  });

  it("reads metadata and content under the same pinned item id", async () => {
    let request = 0;
    const fetchFn = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => ++request === 1
      ? new Response(JSON.stringify({ id: "file", name: "a.txt", size: 2, file: { mimeType: "text/plain" } }), { status: 200 })
      : new Response("ok", { status: 200, headers: { "content-type": "text/plain" } }));
    await driveRead(root, "a.txt", "token", "text", 16, undefined, fetchFn as typeof fetch);
    for (const call of fetchFn.mock.calls) expect(String(call[0])).toContain("/drives/drive-id/items/stable-root-id:/a.txt:");
  });

  it("supports MP4 binary reads while keeping text mode and unsupported MIME fail-closed", async () => {
    const content = Buffer.from([0, 0, 0, 24, 102, 116, 121, 112, 109, 112, 52, 50]);
    const mp4Fetch = () => {
      let request = 0;
      return vi.fn(async () => ++request % 2 === 1
        ? new Response(JSON.stringify({ id: "mp4", name: "clip.mp4", size: content.byteLength, file: { mimeType: "video/mp4" } }), { status: 200 })
        : new Response(content, { status: 200, headers: { "content-type": "video/mp4" } }));
    };

    const base64Fetch = mp4Fetch();
    await expect(driveRead(root, "clip.mp4", "token", "base64", content.byteLength, undefined, base64Fetch as typeof fetch)).resolves.toMatchObject({
      ok: true,
      bytes: content.byteLength,
      mode: "base64",
      content_base64: content.toString("base64"),
    });

    const digestFetch = mp4Fetch();
    await expect(driveRead(root, "clip.mp4", "token", "digest", content.byteLength, undefined, digestFetch as typeof fetch)).resolves.toMatchObject({
      ok: true,
      bytes: content.byteLength,
      mode: "digest",
      sha256: createHash("sha256").update(content).digest("hex"),
    });

    const textFetch = mp4Fetch();
    await expect(driveRead(root, "clip.mp4", "token", "text", content.byteLength, undefined, textFetch as typeof fetch)).rejects.toThrow("binary_requires_base64_or_digest");

    const unsupportedFetch = vi.fn(async () => new Response(JSON.stringify({ id: "unsupported", name: "clip.webm", size: content.byteLength, file: { mimeType: "video/webm" } }), { status: 200 }));
    await expect(driveRead(root, "clip.webm", "token", "digest", content.byteLength, undefined, unsupportedFetch as typeof fetch)).rejects.toThrow("unsupported_file_type");
    expect(unsupportedFetch).toHaveBeenCalledTimes(1);
  });

  it("streams an XLSX digest without returning or buffering file content", async () => {
    const content = Buffer.from([0, 255, 1, 254, 2, 253]);
    let request = 0;
    const fetchFn = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => ++request === 1
      ? new Response(JSON.stringify({ id: "xlsx", name: "SYNTHETIC_RECORD.xlsx", size: content.byteLength, file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } }), { status: 200 })
      : new Response(new ReadableStream({
        start(controller) {
          controller.enqueue(content.subarray(0, 2));
          controller.enqueue(content.subarray(2));
          controller.close();
        },
      }), { status: 200 }));
    const value = await driveRead(root, "SYNTHETIC_RECORD.xlsx", "token", "digest", 250 * 1024 * 1024 * 1024, undefined, fetchFn as typeof fetch);
    expect(value).toMatchObject({
      ok: true,
      operation: "read",
      bytes: content.byteLength,
      mode: "digest",
      sha256: createHash("sha256").update(content).digest("hex"),
    });
    expect(value).not.toHaveProperty("content_base64");
    expect(value).not.toHaveProperty("content_text");
  });

  it("stops a digest stream that exceeds the configured provider read ceiling", async () => {
    let request = 0;
    const fetchFn = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => ++request === 1
      ? new Response(JSON.stringify({ id: "xlsx", name: "SYNTHETIC_RECORD.xlsx", size: 2, file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } }), { status: 200 })
      : new Response(Buffer.from([1, 2, 3, 4]), { status: 200 }));
    await expect(driveRead(root, "SYNTHETIC_RECORD.xlsx", "token", "digest", 3, undefined, fetchFn as typeof fetch)).rejects.toThrow("provider_response_too_large");
  });

  it("returns and validates honest list continuation metadata", async () => {
    const next = "https://graph.microsoft.com/v1.0/drives/drive-id/items/stable-root-id:/SYNTHETIC_FOLDER:/children?$skiptoken=next";
    const fetchFn = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => new Response(JSON.stringify({ value: [{ id: "one", name: "one.txt", file: { mimeType: "text/plain" } }], "@odata.nextLink": next }), { status: 200 }));
    const page = await driveList(root, "SYNTHETIC_FOLDER", "token", 1, undefined, fetchFn as typeof fetch);
    expect(page).toMatchObject({ items: [expect.objectContaining({ id: "one" })], truncated: true, providerNextLink: "/drives/drive-id/items/stable-root-id:/SYNTHETIC_FOLDER:/children?$skiptoken=next" });
    await driveListContinuation(root, "SYNTHETIC_FOLDER", "token", 1, page.providerNextLink!, undefined, fetchFn as typeof fetch);
    expect(String(fetchFn.mock.calls[1][0])).toContain("?$skiptoken=next");
    await expect(driveListContinuation(root, "SYNTHETIC_FOLDER", "token", 1, "https://evil.invalid/v1.0/drives/drive-id/items/stable-root-id:/SYNTHETIC_FOLDER:/children?$skiptoken=x", undefined, fetchFn as typeof fetch)).rejects.toThrow("invalid_continuation");
  });

  it("binds list and search continuations to the exact root, endpoint, and query pathname", async () => {
    const listPath = "/drives/drive-id/items/stable-root-id:/SYNTHETIC_FOLDER:/children";
    expect(canonicalGraphContinuation(`${listPath}?$skiptoken=x`, listPath)).toBe(`${listPath}?$skiptoken=x`);
    expect(canonicalGraphContinuation(`https://graph.microsoft.com/v1.0${listPath}?$skiptoken=x`, `/v1.0${listPath}`)).toBe(`/v1.0${listPath}?$skiptoken=x`);
    const rejected = [
      `${listPath}Extra?$skiptoken=x`,
      "/drives/other/items/stable-root-id:/SYNTHETIC_FOLDER:/children?$skiptoken=x",
      "/drives/drive-id/items/other-root:/SYNTHETIC_FOLDER:/children?$skiptoken=x",
      "/drives/drive-id/items/stable-root-id:/Other:/children?$skiptoken=x",
      "/drives/drive-id/items/stable-root-id:/SYNTHETIC_FOLDER:/search?$skiptoken=x",
      `/drives/drive-id/items/stable-root-id:/SYNTHETIC_FOLDER:/children/../children?$skiptoken=x`,
      `/drives/drive-id/items/stable-root-id:/SYNTHETIC_FOLDER:/children/%2e%2e?$skiptoken=x`,
      `/drives/drive-id/items/stable-root-id:/SYNTHETIC_FOLDER:/children/%252e%252e?$skiptoken=x`,
      `https://user@graph.microsoft.com/v1.0${listPath}?$skiptoken=x`,
      `https://graph.microsoft.com:444/v1.0${listPath}?$skiptoken=x`,
      `http://graph.microsoft.com/v1.0${listPath}?$skiptoken=x`,
      `https://evil.invalid/v1.0${listPath}?$skiptoken=x`,
      `https://graph.microsoft.com/v1.0${listPath}#fragment`,
    ];
    for (const value of rejected) {
      const expected = value.startsWith("http") ? `/v1.0${listPath}` : listPath;
      expect(() => canonicalGraphContinuation(value, expected), value).toThrow("invalid_continuation");
    }

    const searchPath = "/drives/drive-id/items/stable-root-id/search(q='quarterly')";
    const next = `https://graph.microsoft.com/v1.0${searchPath}?$skiptoken=next`;
    const fetchFn = vi.fn(async () => new Response(JSON.stringify({ value: [], "@odata.nextLink": next }), { status: 200 }));
    const first = await driveSearch(root, "quarterly", "token", 5, undefined, fetchFn as typeof fetch);
    expect(first.providerNextLink).toBe(next);
    await driveSearchContinuation(root, "quarterly", "token", 5, first.providerNextLink!, undefined, fetchFn as typeof fetch);
    await expect(driveSearchContinuation(root, "different", "token", 5, first.providerNextLink!, undefined, fetchFn as typeof fetch)).rejects.toThrow("invalid_continuation");

    const apostrophePath = "/drives/drive-id/items/stable-root-id/search(q='O''Brien.md')";
    const encodedNext = `https://graph.microsoft.com/v1.0/drives/drive-id/items/stable-root-id/search(q=%27O%27%27Brien.md%27)?$skiptoken=opaque%252Bstate`;
    expect(canonicalGraphContinuation(encodedNext, `/v1.0${apostrophePath}`)).toBe(`/v1.0${apostrophePath}?$skiptoken=opaque%252Bstate`);
    const encodedFetch = vi.fn(async (_input: string | URL | Request) => new Response(JSON.stringify({ value: [] }), { status: 200 }));
    await driveSearchContinuation(root, "O'Brien.md", "token", 5, encodedNext, undefined, encodedFetch as typeof fetch);
    expect(String(encodedFetch.mock.calls[0][0])).toBe(encodedNext);
  });

  it("proves every provider hit reaches the pinned root by stable IDs, including the root itself", async () => {
    const metadata: Record<string, unknown> = {
      inside: { id: "inside", parentReference: { id: "inside-folder", driveId: "drive-id" } },
      "inside-folder": { id: "inside-folder", parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      direct: { id: "direct", parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      outside: { id: "outside", parentReference: { id: "drive-root", driveId: "drive-id" } },
      "drive-root": { id: "drive-root" },
      "wrong-drive": { id: "wrong-drive", parentReference: { id: "stable-root-id", driveId: "other-drive" } },
      "missing-parent": { id: "missing-parent" },
      "stable-root-id": { id: "stable-root-id" },
    };
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) return new Response(JSON.stringify({ value: [
        { id: "inside", name: "same.xlsx", file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }, parentReference: { driveId: "drive-id", path: "/drive/root:/Untrusted" } },
        { id: "direct", name: "direct.xlsx", file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
        { id: "outside", name: "same.xlsx", file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }, parentReference: { driveId: "drive-id", path: "/drive/root:/Mutable/Display/Path" } },
        { id: "wrong-drive", name: "wrong.xlsx", file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
        { id: "missing-parent", name: "missing.xlsx", file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
        { id: "stable-root-id", name: "Pinned Root", folder: {} },
      ] }), { status: 200 });
      const itemId = decodeURIComponent(url.pathname.split("/").at(-1)!);
      return new Response(JSON.stringify(metadata[itemId] ?? {}), { status: 200 });
    });
    const page = await driveSearch(root, "xlsx", "token", 10, undefined, fetchFn as typeof fetch);
    expect(page.items.map((item) => item.id)).toEqual(["inside", "direct", "stable-root-id"]);
    expect(page.items.some((item) => item.id === "outside")).toBe(false);
  });

  it("recognizes filename-like queries without claiming field or path filtering", () => {
    for (const value of ["README.md", "METADATA.md", "AGENTS.md", "SYNTHETIC_ROOT_DOCUMENT.pdf", "résumé.final.PDF"]) expect(exactFilenameQuery(value)).toBe(value);
    for (const value of ["README", "folder/name.md", "*.pdf", ".env", "trailing."]) expect(exactFilenameQuery(value)).toBeNull();
  });

  it("normalizes automatic and explicit search modes and rejects path-like filename criteria", () => {
    expect(normalizeDriveSearch(" README.md ")).toEqual({ query: "README.md", mode: "filename_exact", exhaustive: false });
    expect(normalizeDriveSearch(" README ")).toEqual({ query: "README", mode: "provider", exhaustive: false });
    expect(normalizeDriveSearch(" README ", "filename_stem", true)).toEqual({ query: "README", mode: "filename_stem", exhaustive: true });
    expect(() => normalizeDriveSearch("folder/README.md", "filename_exact")).toThrow("invalid_search");
    expect(() => normalizeDriveSearch("README", "provider", true)).toThrow("invalid_search");
  });

  it.each([
    ["filename_stem", "README", "README.md"],
    ["filename_exact", ".env", ".ENV"],
    ["filename_contains", "synthetic_record_token", "SYNTHETIC_RECORD_TOKEN_final.pdf"],
  ] as const)("supports deterministic %s matching for %s", async (mode, query, filename) => {
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) throw new Error("provider search must not be called");
      return new Response(JSON.stringify({ value: [
        { id: "match", name: filename, file: { mimeType: "text/plain" }, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ] }), { status: 200 });
    });
    const result = await driveSearchScoped(root, query, "token", 5, undefined, undefined, fetchFn as typeof fetch, {}, mode);
    expect(result).toMatchObject({
      items: [expect.objectContaining({ id: "match", name: filename })],
      truncated: false,
      scan_complete: true,
      match_satisfied: true,
    });
    expect(result).not.toHaveProperty("continuationState");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("honors explicit provider mode for a filename-looking query", async () => {
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) return new Response(JSON.stringify({ value: [
        { id: "provider-hit", name: "README.md", file: { mimeType: "text/markdown" } },
      ] }), { status: 200 });
      const itemId = decodeURIComponent(url.pathname.split("/").at(-1)!);
      return new Response(JSON.stringify(itemId === "provider-hit"
        ? { id: itemId, parentReference: { id: "stable-root-id", driveId: "drive-id" } }
        : { id: "stable-root-id" }), { status: 200 });
    });
    const result = await driveSearchScoped(root, "README.md", "token", 5, undefined, undefined, fetchFn as typeof fetch, {}, "provider");
    expect(result).toMatchObject({
      items: [expect.objectContaining({ id: "provider-hit" })],
      truncated: false,
      scan_complete: true,
      match_satisfied: true,
    });
    expect(fetchFn.mock.calls.some(([input]) => String(input).includes("/search("))).toBe(true);
    expect(fetchFn.mock.calls.some(([input]) => String(input).includes("/children"))).toBe(false);
  });

  it("routes an exact README filename directly to the pinned root children", async () => {
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) throw new Error("provider search must not be called");
      return new Response(JSON.stringify({ value: [
        { id: "readme", name: "README.md", file: { mimeType: "text/markdown" }, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "synthetic-folder", name: "SYNTHETIC_FOLDER", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ] }), { status: 200 });
    });

    const result = await driveSearchScoped(root, "README.md", "token", 5, undefined, undefined, fetchFn as typeof fetch);
    expect(result).toMatchObject({ items: [expect.objectContaining({ id: "readme", name: "README.md" })], truncated: false, scan_complete: false, match_satisfied: true, fallback: true });
    expect(result).not.toHaveProperty("continuationState");
    expect(fetchFn).toHaveBeenCalledTimes(1);
    expect(String(fetchFn.mock.calls[0][0])).toContain("/drives/drive-id/items/stable-root-id/children");
  });

  it("uses the scan budget rather than the return limit for an exact root page", async () => {
    const rootChildren = [
      { id: "synthetic-folder-a", name: "SYNTHETIC_FOLDER_A", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ...Array.from({ length: 12 }, (_, index) => ({
        id: `root-${index}`,
        name: `root-${index}.txt`,
        file: { mimeType: "text/plain" },
        parentReference: { id: "stable-root-id", driveId: "drive-id" },
      })),
      { id: "readme", name: "README.md", file: { mimeType: "text/markdown" }, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
    ];
    const rootPrefix = "/v1.0/drives/drive-id/items/stable-root-id/children";
    const next = `https://graph.microsoft.com${rootPrefix}?$skiptoken=second`;
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.searchParams.has("$skiptoken")) return new Response(null, { status: 403 });
      const top = Number(url.searchParams.get("$top"));
      return new Response(JSON.stringify(top < rootChildren.length
        ? { value: rootChildren.slice(0, top), "@odata.nextLink": next }
        : { value: rootChildren }), { status: 200 });
    });

    const result = await driveSearchScoped(root, "README.md", "token", 5, undefined, undefined, fetchFn as typeof fetch);
    expect(result).toMatchObject({
      items: [expect.objectContaining({ id: "readme", name: "README.md" })],
      fallback: true,
    });
    expect(result.items).toHaveLength(1);
    expect(fetchFn).toHaveBeenCalledTimes(1);
    const request = new URL(String(fetchFn.mock.calls[0][0]));
    expect(request.pathname).toBe(rootPrefix);
    expect(request.searchParams.get("$top")).toBe("200");
  });

  it.each([1, 10])("keeps exhaustive filename_contains provider pages independent of return limit %i", async (limit) => {
    const rootPrefix = "/v1.0/drives/drive-id/items/stable-root-id/children";
    const rootChildren = Array.from({ length: limit + 2 }, (_, index) => ({
      id: `match-${index}`,
      name: `synthetic-match-${index}.pdf`,
      file: { mimeType: "application/pdf" },
      parentReference: { id: "stable-root-id", driveId: "drive-id" },
    }));
    const deniedNext = `https://graph.microsoft.com${rootPrefix}?$skiptoken=denied`;
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.searchParams.has("$skiptoken")) return new Response(null, { status: 403 });
      const top = Number(url.searchParams.get("$top"));
      return new Response(JSON.stringify(top < rootChildren.length
        ? { value: rootChildren.slice(0, top), "@odata.nextLink": deniedNext }
        : { value: rootChildren }), { status: 200 });
    });

    const result = await driveSearchScoped(root, "synthetic-match", "token", limit, undefined, undefined, fetchFn as typeof fetch, {}, "filename_contains", true);
    expect(result.items.map((item) => item.id)).toEqual(rootChildren.slice(0, limit).map((item) => item.id));
    expect(result).toMatchObject({ truncated: true, scan_complete: false, match_satisfied: true, fallback: true });
    expect(result.continuationState).toBeDefined();
    expect(fetchFn).toHaveBeenCalledTimes(1);
    const request = new URL(String(fetchFn.mock.calls[0][0]));
    expect(request.pathname).toBe(rootPrefix);
    expect(request.searchParams.get("$top")).toBe("200");
  });

  it("resumes buffered matches and remaining traversal across real exhaustive continuations", async () => {
    const calls: string[] = [];
    const matchingFile = (id: string, parentId: string) => ({
      id,
      name: `${id}-needle.pdf`,
      file: { mimeType: "application/pdf" },
      parentReference: { id: parentId, driveId: "drive-id" },
    });
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      calls.push(`${url.pathname}${url.search}`);
      if (url.pathname.endsWith("/items/stable-root-id/children")) return new Response(JSON.stringify({ value: [
        matchingFile("match-one", "stable-root-id"),
        matchingFile("match-two", "stable-root-id"),
        matchingFile("match-three", "stable-root-id"),
        { id: "child", name: "Child", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ] }), { status: 200 });
      if (url.pathname.endsWith("/items/child/children")) return new Response(JSON.stringify({ value: [
        matchingFile("match-four", "child"),
      ] }), { status: 200 });
      return new Response(null, { status: 403 });
    });

    const pages = [];
    let continuationState: unknown;
    do {
      const page = await driveSearchScoped(root, "needle", "token", 1, continuationState, undefined, fetchFn as typeof fetch, { items: 6 }, "filename_contains", true);
      pages.push(page);
      continuationState = page.continuationState;
    } while (continuationState !== undefined);

    expect(pages.map((page) => page.items.map((item) => item.id))).toEqual([
      ["match-one"],
      ["match-two"],
      ["match-three"],
      ["match-four"],
    ]);
    expect(pages.map((page) => page.truncated)).toEqual([true, true, true, false]);
    expect(pages.map((page) => page.match_satisfied)).toEqual([true, true, true, true]);
    expect(pages.slice(0, -1).every((page) => page.scan_complete === false && page.continuationState !== undefined)).toBe(true);
    expect(pages.at(-1)).toMatchObject({ truncated: false, scan_complete: true, match_satisfied: true });
    expect(pages.at(-1)).not.toHaveProperty("continuationState");
    expect(pages.flatMap((page) => page.items.map((item) => item.id))).toEqual(["match-one", "match-two", "match-three", "match-four"]);
    expect(fetchFn).toHaveBeenCalledTimes(2);
    expect(calls).toEqual([
      "/v1.0/drives/drive-id/items/stable-root-id/children?$top=6&$select=id%2Cname%2CwebUrl%2CcreatedDateTime%2ClastModifiedDateTime%2Cdescription%2CfileSystemInfo%2CparentReference%2Csize%2Cfile%2Cfolder%2CeTag",
      "/v1.0/drives/drive-id/items/child/children?$top=6&$select=id%2Cname%2CwebUrl%2CcreatedDateTime%2ClastModifiedDateTime%2Cdescription%2CfileSystemInfo%2CparentReference%2Csize%2Cfile%2Cfolder%2CeTag",
    ]);
  });

  it("defers a provider nextLink until a fresh scan budget can accept its original page size", async () => {
    const rootPrefix = "/v1.0/drives/drive-id/items/stable-root-id/children";
    const next = `https://graph.microsoft.com${rootPrefix}?$skiptoken=second`;
    const firstValues = Array.from({ length: 150 }, (_, index) => ({
      id: `nonmatch-${index}`,
      name: `other-${index}.txt`,
      file: { mimeType: "text/plain" },
      parentReference: { id: "stable-root-id", driveId: "drive-id" },
    }));
    const secondValues = Array.from({ length: 100 }, (_, index) => ({
      id: `match-${index}`,
      name: `needle-${index}.txt`,
      file: { mimeType: "text/plain" },
      parentReference: { id: "stable-root-id", driveId: "drive-id" },
    }));
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      return new Response(JSON.stringify(url.searchParams.has("$skiptoken")
        ? { value: secondValues }
        : { value: firstValues, "@odata.nextLink": next }), { status: 200 });
    });

    const first = await driveSearchScoped(root, "needle", "token", 1, undefined, undefined, fetchFn as typeof fetch, { items: 200 }, "filename_contains", true);
    expect(first).toMatchObject({ items: [], truncated: true, scan_complete: false, match_satisfied: false });
    expect(first.continuationState).toMatchObject({ current: { folderId: "stable-root-id", nextLink: next } });
    expect(fetchFn).toHaveBeenCalledTimes(1);

    const second = await driveSearchScoped(root, "needle", "token", 1, first.continuationState, undefined, fetchFn as typeof fetch, { items: 200 }, "filename_contains", true);
    expect(second.items).toEqual([expect.objectContaining({ id: "match-0" })]);
    expect(second).toMatchObject({ truncated: true, scan_complete: false, match_satisfied: true });
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it("stores metadata-rich overflow above the former 256 KiB continuation ceiling", async () => {
    const richItems = Array.from({ length: 50 }, (_, index) => ({
      id: `rich-${index}`,
      name: `needle-${index}.txt`,
      webUrl: `https://onedrive.live.com/${"w".repeat(1800)}${index}`,
      description: "d".repeat(4096),
      file: { mimeType: "text/plain" },
      parentReference: { id: "stable-root-id", driveId: "drive-id", path: `/${"p".repeat(2000)}` },
    }));
    const fetchFn = vi.fn(async () => new Response(JSON.stringify({ value: richItems }), { status: 200 }));
    const page = await driveSearchScoped(root, "needle", "token", 1, undefined, undefined, fetchFn as typeof fetch, {}, "filename_contains", true);
    expect(Buffer.byteLength(JSON.stringify(page.continuationState), "utf8")).toBeGreaterThan(256 * 1024);

    const store = new ContinuationStore();
    const path = driveSearchPath(root, "needle");
    const binding = { agentId: "main", service: "onedrive" as const, action: "search", resource: root.label, criteria: normalizedCriteria({ query: "needle", mode: "filename_contains", exhaustive: true, limit: 1 }) };
    const handle = store.issueState(binding, path, page.continuationState!);
    expect(store.continuationState(store.verify(handle, binding), path)).toEqual(page.continuationState);
  });

  it("falls back to pinned children traversal for root and nested PDF/Markdown files", async () => {
    const responses = (filename: string) => vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) throw new Error("provider search must not be called");
      if (url.pathname.endsWith("/items/stable-root-id/children")) return new Response(JSON.stringify({ value: [
        { id: "root-cycle", name: "loop", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "stable-root-id", name: "self", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "synthetic-nested", name: "SYNTHETIC_NESTED_FOLDER", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "root-pdf", name: "SYNTHETIC_ROOT_DOCUMENT.pdf", file: { mimeType: "application/pdf" }, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "collision-folder", name: filename, folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "wrong-drive", name: filename, file: { mimeType: "text/plain" }, parentReference: { id: "stable-root-id", driveId: "other-drive" } },
        { id: "missing-parent", name: filename, file: { mimeType: "text/plain" } },
      ] }), { status: 200 });
      if (url.pathname.endsWith("/items/synthetic-nested/children")) return new Response(JSON.stringify({ value: [
        { id: "deep-folder", name: "SYNTHETIC_DEEP_FOLDER", folder: {}, parentReference: { id: "synthetic-nested", driveId: "drive-id" } },
      ] }), { status: 200 });
      if (url.pathname.endsWith("/items/deep-folder/children")) return new Response(JSON.stringify({ value: [
        { id: "nested-md", name: "README.md", file: { mimeType: "text/markdown" }, parentReference: { id: "deep-folder", driveId: "drive-id" } },
      ] }), { status: 200 });
      return new Response(JSON.stringify({ value: [] }), { status: 200 });
    });

    const pdfFetch = responses("SYNTHETIC_ROOT_DOCUMENT.pdf");
    const pdf = await driveSearchScoped(root, "SYNTHETIC_ROOT_DOCUMENT.pdf", "token", 10, undefined, undefined, pdfFetch as typeof fetch);
    expect(pdf.items).toEqual([expect.objectContaining({ id: "root-pdf", name: "SYNTHETIC_ROOT_DOCUMENT.pdf", mime_type: "application/pdf" })]);
    expect(pdf).toMatchObject({ fallback: true, truncated: false, scan_complete: false, match_satisfied: true });
    expect(pdf).not.toHaveProperty("continuationState");
    expect(pdfFetch).toHaveBeenCalledTimes(1);
    expect(String(pdfFetch.mock.calls[0][0])).toContain("/drives/drive-id/items/stable-root-id/children");
    expect(pdfFetch.mock.calls.some(([input]) => String(input).includes("/search("))).toBe(false);

    const markdownFetch = responses("README.md");
    const markdown = await driveSearchScoped(root, "readme.MD", "token", 10, undefined, undefined, markdownFetch as typeof fetch);
    expect(markdown.items).toEqual([expect.objectContaining({ id: "nested-md", name: "README.md", mime_type: "text/markdown" })]);
    expect(markdown).toMatchObject({ truncated: false, scan_complete: true, match_satisfied: true });
    expect(markdownFetch).toHaveBeenCalledTimes(5);
    expect(markdown.items.some((item: any) => item.is_folder)).toBe(false);
    expect(markdownFetch.mock.calls.some(([input]) => String(input).includes("/search("))).toBe(false);
  });

  it("returns continuation before a deep exact filename traversal can exhaust the explicit page budget", async () => {
    const calls: string[] = [];
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      calls.push(url.pathname);
      if (url.pathname.includes("/search(")) throw new Error("provider search must not be called");
      if (url.pathname.endsWith("/items/stable-root-id/children")) return new Response(JSON.stringify({ value: [
        { id: "synthetic-level-one", name: "SYNTHETIC_LEVEL_ONE", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ] }), { status: 200 });
      if (url.pathname.endsWith("/items/synthetic-level-one/children")) return new Response(JSON.stringify({ value: [
        { id: "synthetic-level-two", name: "SYNTHETIC_LEVEL_TWO", folder: {}, parentReference: { id: "synthetic-level-one", driveId: "drive-id" } },
      ] }), { status: 200 });
      if (url.pathname.endsWith("/items/synthetic-level-two/children")) return new Response(JSON.stringify({ value: [
        { id: "synthetic-document", name: "2099-01-01_SYNTHETIC_RECORD.pdf", file: { mimeType: "application/pdf" }, parentReference: { id: "synthetic-level-two", driveId: "drive-id" } },
      ] }), { status: 200 });
      throw new Error(`unexpected request: ${url.pathname}`);
    });

    const query = "2099-01-01_SYNTHETIC_RECORD.pdf";
    const first = await driveSearchScoped(root, query, "token", 5, undefined, undefined, fetchFn as typeof fetch, { pages: 2 });
    expect(first).toMatchObject({ items: [], truncated: true, scan_complete: false, match_satisfied: false, fallback: true });
    expect(first.continuationState).toBeDefined();
    expect(fetchFn).toHaveBeenCalledTimes(2);

    const second = await driveSearchScoped(root, query, "token", 5, first.continuationState, undefined, fetchFn as typeof fetch, { pages: 2 });
    expect(second).toMatchObject({
      items: [expect.objectContaining({ id: "synthetic-document", name: query, mime_type: "application/pdf" })],
      truncated: false,
      scan_complete: true,
      match_satisfied: true,
      fallback: true,
    });
    expect(fetchFn).toHaveBeenCalledTimes(3);
    expect(calls).toEqual([
      "/v1.0/drives/drive-id/items/stable-root-id/children",
      "/v1.0/drives/drive-id/items/synthetic-level-one/children",
      "/v1.0/drives/drive-id/items/synthetic-level-two/children",
    ]);
  });

  it("preserves exact-filename progress when the whole-operation deadline is exhausted", async () => {
    const operation = graphOperationSignal(undefined, 20, 100);
    const fetchFn = vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
      const url = new URL(String(input));
      if (url.pathname.endsWith("/items/stable-root-id/children")) return new Response(JSON.stringify({ value: [
        { id: "child", name: "Child", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ] }), { status: 200 });
      return new Promise<Response>((_resolve, reject) => {
        const abort = () => reject(init?.signal?.reason);
        if (init?.signal?.aborted) abort();
        else init?.signal?.addEventListener("abort", abort, { once: true });
      });
    });
    const page = await driveSearchScoped(root, "README.md", "token", 5, undefined, operation, fetchFn as typeof fetch);
    expect(page).toMatchObject({
      items: [],
      truncated: true,
      continuationState: { kind: "exact_fallback", current: { folderId: "child" } },
    });
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it.each([
    ["denied", () => new Response(null, { status: 403 })],
    ["malformed", () => new Response("not-json", { status: 200 })],
    ["malformed nextLink", () => new Response(JSON.stringify({
      value: [{ id: "unproven", name: "2099-01-01_SYNTHETIC_RECORD.pdf", file: { mimeType: "application/pdf" }, parentReference: { id: "unreadable", driveId: "drive-id" } }],
      "@odata.nextLink": "https://evil.invalid/children?$skiptoken=x",
    }), { status: 200 })],
  ])("skips a %s descendant and preserves bounded BFS progress to a reachable deep sibling target", async (_failure, failedResponse) => {
    const calls: string[] = [];
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      calls.push(url.pathname);
      if (url.pathname.endsWith("/items/stable-root-id/children")) return new Response(JSON.stringify({ value: [
        { id: "unreadable", name: "SYNTHETIC_UNREADABLE_FOLDER", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "reachable", name: "SYNTHETIC_REACHABLE_FOLDER", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ] }), { status: 200 });
      if (url.pathname.endsWith("/items/unreadable/children")) return failedResponse();
      if (url.pathname.endsWith("/items/reachable/children")) return new Response(JSON.stringify({ value: [
        { id: "deep", name: "SYNTHETIC_DEEP_FOLDER", folder: {}, parentReference: { id: "reachable", driveId: "drive-id" } },
      ] }), { status: 200 });
      if (url.pathname.endsWith("/items/deep/children")) return new Response(JSON.stringify({ value: [
        { id: "target", name: "2099-01-01_SYNTHETIC_RECORD.pdf", file: { mimeType: "application/pdf" }, parentReference: { id: "deep", driveId: "drive-id" } },
      ] }), { status: 200 });
      throw new Error(`unexpected request: ${url.pathname}`);
    });

    const query = "2099-01-01_SYNTHETIC_RECORD.pdf";
    const page = await driveSearchScoped(root, query, "token", 5, undefined, undefined, fetchFn as typeof fetch);
    expect(page).toMatchObject({
      items: [expect.objectContaining({ id: "target", name: query, mime_type: "application/pdf" })],
      truncated: false,
      scan_complete: false,
      match_satisfied: true,
      fallback: true,
    });
    expect(fetchFn).toHaveBeenCalledTimes(4);
    expect(calls).toEqual([
      "/v1.0/drives/drive-id/items/stable-root-id/children",
      "/v1.0/drives/drive-id/items/unreadable/children",
      "/v1.0/drives/drive-id/items/reachable/children",
      "/v1.0/drives/drive-id/items/deep/children",
    ]);
  });

  it.each([
    ["denied", () => new Response(null, { status: 403 })],
    ["malformed", () => new Response("not-json", { status: 200 })],
  ])("does not claim a complete exhaustive scan after a %s descendant was skipped", async (_failure, failedResponse) => {
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.endsWith("/items/stable-root-id/children")) return new Response(JSON.stringify({ value: [
        { id: "unreadable", name: "Unreadable", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "reachable", name: "Reachable", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ] }), { status: 200 });
      if (url.pathname.endsWith("/items/unreadable/children")) return failedResponse();
      if (url.pathname.endsWith("/items/reachable/children")) return new Response(JSON.stringify({ value: [] }), { status: 200 });
      throw new Error(`unexpected request: ${url.pathname}`);
    });

    const result = await driveSearchScoped(root, "missing.pdf", "token", 5, undefined, undefined, fetchFn as typeof fetch, {}, "filename_exact", true);
    expect(result).toMatchObject({
      items: [],
      truncated: false,
      scan_complete: false,
      match_satisfied: false,
      fallback: true,
    });
    expect(result).not.toHaveProperty("continuationState");
    expect(fetchFn).toHaveBeenCalledTimes(3);
  });

  it("preserves incomplete-scan truth across an opaque traversal continuation", async () => {
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.endsWith("/items/stable-root-id/children")) return new Response(JSON.stringify({ value: [
        { id: "unreadable", name: "Unreadable", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "reachable", name: "Reachable", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ] }), { status: 200 });
      if (url.pathname.endsWith("/items/unreadable/children")) return new Response(null, { status: 403 });
      if (url.pathname.endsWith("/items/reachable/children")) return new Response(JSON.stringify({ value: [] }), { status: 200 });
      throw new Error(`unexpected request: ${url.pathname}`);
    });

    const first = await driveSearchScoped(root, "missing.pdf", "token", 5, undefined, undefined, fetchFn as typeof fetch, { pages: 2 }, "filename_exact", true);
    expect(first).toMatchObject({
      items: [],
      truncated: true,
      scan_complete: false,
      match_satisfied: false,
      continuationState: { kind: "exact_fallback", current: { folderId: "reachable" }, scanIncomplete: true },
    });

    const second = await driveSearchScoped(root, "missing.pdf", "token", 5, first.continuationState, undefined, fetchFn as typeof fetch, { pages: 2 }, "filename_exact", true);
    expect(second).toMatchObject({ items: [], truncated: false, scan_complete: false, match_satisfied: false });
    expect(second).not.toHaveProperty("continuationState");
    expect(fetchFn).toHaveBeenCalledTimes(3);
  });

  it.each([
    ["denied", () => new Response(null, { status: 403 }), "provider_access_denied"],
    ["malformed", () => new Response("not-json", { status: 200 }), "invalid_provider_response"],
  ])("keeps a %s pinned-root child page fatal", async (_failure, response, expected) => {
    const fetchFn = vi.fn(async () => response());
    await expect(driveSearchScoped(root, "README.md", "token", 5, undefined, undefined, fetchFn as typeof fetch)).rejects.toThrow(expected);
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("does not treat a descendant 401 as a folder-local denial", async () => {
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      return url.pathname.endsWith("/items/stable-root-id/children")
        ? new Response(JSON.stringify({ value: [
          { id: "child", name: "Child", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        ] }), { status: 200 })
        : new Response(null, { status: 401 });
    });
    await expect(driveSearchScoped(root, "README.md", "token", 5, undefined, undefined, fetchFn as typeof fetch)).rejects.toThrow("provider_access_denied");
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it("preserves provider nextLink bytes and exact criteria across continuation", async () => {
    const searchPath = "/v1.0/drives/drive-id/items/stable-root-id/search(q='quarterly')";
    const next = `https://graph.microsoft.com${searchPath}?$skiptoken=A%2BB%252F&$top=5`;
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) return url.searchParams.has("$skiptoken")
        ? new Response(JSON.stringify({ value: [{ id: "two", name: "quarterly two", file: { mimeType: "text/plain" } }] }), { status: 200 })
        : new Response(JSON.stringify({ value: [{ id: "one", name: "quarterly", file: { mimeType: "text/plain" } }], "@odata.nextLink": next }), { status: 200 });
      const itemId = decodeURIComponent(url.pathname.split("/").at(-1)!);
      return new Response(JSON.stringify(itemId === "stable-root-id"
        ? { id: itemId }
        : { id: itemId, parentReference: { id: "stable-root-id", driveId: "drive-id" } }), { status: 200 });
    });
    const first = await driveSearchScoped(root, "quarterly", "token", 5, undefined, undefined, fetchFn as typeof fetch);
    expect(first).toMatchObject({ scan_complete: false, match_satisfied: true });
    expect(first.continuationState).toEqual({ kind: "provider", nextLink: next, returnedIds: ["one"] });
    const second = await driveSearchScoped(root, "quarterly", "token", 5, first.continuationState, undefined, fetchFn as typeof fetch);
    expect(fetchFn.mock.calls.some(([input]) => String(input) === next)).toBe(true);
    expect(second.items).toEqual([expect.objectContaining({ id: "two" })]);
    expect(second).toMatchObject({ scan_complete: true, match_satisfied: true });
  });

  it("advances filtered-empty provider pages with the untouched bound nextLink", async () => {
    const searchPath = "/v1.0/drives/drive-id/items/stable-root-id/search(q='quarterly')";
    const next = `https://graph.microsoft.com${searchPath}?$skiptoken=A%2BB%252F&$top=5`;
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) return url.searchParams.has("$skiptoken")
        ? new Response(JSON.stringify({ value: [{ id: "inside", name: "quarterly in root", file: { mimeType: "text/plain" } }] }), { status: 200 })
        : new Response(JSON.stringify({ value: [{ id: "outside", name: "quarterly elsewhere", file: { mimeType: "text/plain" } }], "@odata.nextLink": next }), { status: 200 });
      const itemId = decodeURIComponent(url.pathname.split("/").at(-1)!);
      const metadata: Record<string, unknown> = {
        outside: { id: "outside", parentReference: { id: "drive-root", driveId: "drive-id" } },
        "drive-root": { id: "drive-root" },
        inside: { id: "inside", parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        "stable-root-id": { id: "stable-root-id" },
      };
      return new Response(JSON.stringify(metadata[itemId] ?? {}), { status: 200 });
    });
    const first = await driveSearchScoped(root, "quarterly", "token", 5, undefined, undefined, fetchFn as typeof fetch, { providerPages: 1 });
    expect(first).toMatchObject({ items: [], truncated: true, continuationState: { kind: "provider", nextLink: next } });
    const second = await driveSearchScoped(root, "quarterly", "token", 5, first.continuationState, undefined, fetchFn as typeof fetch, { providerPages: 1 });
    expect(second.items).toEqual([expect.objectContaining({ id: "inside" })]);
    expect(fetchFn.mock.calls.some(([input]) => String(input) === next)).toBe(true);
  });

  it("fails closed for wrong-drive, cyclic, over-depth, and unavailable ancestry proofs", async () => {
    const metadata: Record<string, unknown> = {
      "wrong-drive": { id: "wrong-drive", parentReference: { id: "stable-root-id", driveId: "other-drive" } },
      cycle: { id: "cycle", parentReference: { id: "cycle-parent", driveId: "drive-id" } },
      "cycle-parent": { id: "cycle-parent", parentReference: { id: "cycle", driveId: "drive-id" } },
      deep: { id: "deep", parentReference: { id: "deep-parent", driveId: "drive-id" } },
      "deep-parent": { id: "deep-parent", parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      "stable-root-id": { id: "stable-root-id" },
    };
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) return new Response(JSON.stringify({ value: [
        { id: "wrong-drive", name: "wrong.txt", file: { mimeType: "text/plain" } },
        { id: "cycle", name: "cycle.txt", file: { mimeType: "text/plain" } },
        { id: "deep", name: "deep.txt", file: { mimeType: "text/plain" } },
      ] }), { status: 200 });
      const itemId = decodeURIComponent(url.pathname.split("/").at(-1)!);
      return new Response(JSON.stringify(metadata[itemId] ?? {}), { status: 200 });
    });
    const page = await driveSearchScoped(root, "txt", "token", 3, undefined, undefined, fetchFn as typeof fetch, { ancestryDepth: 1 });
    expect(page).toMatchObject({ items: [], truncated: false });

    const unavailable = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) return new Response(JSON.stringify({ value: [{ id: "network", name: "network.txt", file: { mimeType: "text/plain" } }] }), { status: 200 });
      throw new TypeError("synthetic network failure");
    });
    await expect(driveSearchScoped(root, "txt", "token", 1, undefined, undefined, unavailable as typeof fetch)).rejects.toThrow("provider_unavailable");
  });

  it("never returns empty truncated exact results without usable traversal state", async () => {
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      return url.pathname.includes("/search(")
        ? new Response(JSON.stringify({ value: [], "@odata.nextLink": `${url.origin}${url.pathname}?$skiptoken=broken` }), { status: 200 })
        : new Response(JSON.stringify({ value: [] }), { status: 200 });
    });
    const result = await driveSearchScoped(root, "README.md", "token", 5, undefined, undefined, fetchFn as typeof fetch);
    expect(result).toMatchObject({ items: [], truncated: false, fallback: true });
    expect(result).not.toHaveProperty("continuationState");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("bounds traversal, preserves more work, and detects folder cycles and duplicate ids", async () => {
    const calls: string[] = [];
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      calls.push(url.pathname);
      if (url.pathname.includes("/search(")) return new Response(JSON.stringify({ value: [] }), { status: 200 });
      if (url.pathname.endsWith("/items/stable-root-id/children")) return new Response(JSON.stringify({ value: [
        { id: "stable-root-id", name: "cycle", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "synthetic-duplicate", name: "SYNTHETIC_FOLDER", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
        { id: "synthetic-duplicate", name: "duplicate", folder: {}, parentReference: { id: "stable-root-id", driveId: "drive-id" } },
      ] }), { status: 200 });
      return new Response(JSON.stringify({ value: [{ id: "target", name: "SYNTHETIC_RECORD.pdf", file: { mimeType: "application/pdf" }, parentReference: { id: "synthetic-duplicate", driveId: "drive-id" } }] }), { status: 200 });
    });
    const first = await driveSearchScoped(root, "SYNTHETIC_RECORD.pdf", "token", 5, undefined, undefined, fetchFn as typeof fetch, { pages: 1, items: 10 });
    expect(first).toMatchObject({ items: [], truncated: true, fallback: true });
    expect(first.continuationState).toBeDefined();
    const second = await driveSearchScoped(root, "SYNTHETIC_RECORD.pdf", "token", 5, first.continuationState, undefined, fetchFn as typeof fetch, { pages: 1, items: 10 });
    expect(second.items).toEqual([expect.objectContaining({ id: "target" })]);
    expect(calls.filter((path) => path.endsWith("/items/stable-root-id/children"))).toHaveLength(1);
    expect(calls.filter((path) => path.endsWith("/items/synthetic-duplicate/children"))).toHaveLength(1);
    expect(calls.some((path) => path.includes("/search("))).toBe(false);
  });

  it("stops a non-exhaustive exact search at the first hit without an unnecessary continuation", async () => {
    const prefix = "/v1.0/drives/drive-id/items/stable-root-id/children";
    const fetchFn = vi.fn(async () => new Response(JSON.stringify({
      value: [{ id: "first", name: "SYNTHETIC_RECORD.pdf", file: { mimeType: "application/pdf" }, parentReference: { id: "stable-root-id", driveId: "drive-id" } }],
      "@odata.nextLink": `https://graph.microsoft.com${prefix}?$skiptoken=more`,
    }), { status: 200 }));
    const result = await driveSearchScoped(root, "SYNTHETIC_RECORD.pdf", "token", 1, undefined, undefined, fetchFn as typeof fetch, {}, "filename_exact");
    expect(result).toMatchObject({
      items: [expect.objectContaining({ id: "first" })],
      truncated: false,
      scan_complete: false,
      match_satisfied: true,
    });
    expect(result).not.toHaveProperty("continuationState");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("continues an exhaustive filename search to return duplicate names without skipping matches", async () => {
    const prefix = "/v1.0/drives/drive-id/items/stable-root-id/children";
    const next = `https://graph.microsoft.com${prefix}?$skiptoken=second`;
    const fetchFn = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      const second = url.searchParams.has("$skiptoken");
      return new Response(JSON.stringify({
        value: [{ id: second ? "duplicate-two" : "duplicate-one", name: second ? "synthetic_record.PDF" : "SYNTHETIC_RECORD.pdf", file: { mimeType: "application/pdf" }, parentReference: { id: "stable-root-id", driveId: "drive-id" } }],
        ...(!second ? { "@odata.nextLink": next } : {}),
      }), { status: 200 });
    });
    const first = await driveSearchScoped(root, "SYNTHETIC_RECORD.pdf", "token", 1, undefined, undefined, fetchFn as typeof fetch, {}, "filename_exact", true);
    expect(first).toMatchObject({
      items: [expect.objectContaining({ id: "duplicate-one" })],
      truncated: true,
      scan_complete: false,
      match_satisfied: true,
    });
    expect(first.continuationState).toBeDefined();
    const second = await driveSearchScoped(root, "SYNTHETIC_RECORD.pdf", "token", 1, first.continuationState, undefined, fetchFn as typeof fetch, {}, "filename_exact", true);
    expect(second).toMatchObject({
      items: [expect.objectContaining({ id: "duplicate-two" })],
      truncated: false,
      scan_complete: true,
      match_satisfied: true,
    });
    expect(second).not.toHaveProperty("continuationState");
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it("bypasses a failing provider search only for exact filenames", async () => {
    const exactFetch = vi.fn(async (input: string | URL | Request) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) return new Response(null, { status: 500 });
      return new Response(JSON.stringify({ value: [{ id: "target", name: "README.md", file: { mimeType: "text/markdown" }, parentReference: { id: "stable-root-id", driveId: "drive-id" } }] }), { status: 200 });
    });
    await expect(driveSearchScoped(root, "README.md", "token", 5, undefined, undefined, exactFetch as typeof fetch)).resolves.toMatchObject({ items: [expect.objectContaining({ id: "target" })], fallback: true });
    expect(exactFetch.mock.calls.some(([input]) => String(input).includes("/search("))).toBe(false);
    const broadFetch = vi.fn(async () => new Response(null, { status: 500 }));
    await expect(driveSearchScoped(root, "README", "token", 5, undefined, undefined, broadFetch as typeof fetch)).rejects.toThrow("provider_error_500");
  });

  it("rejects malformed or non-canonical base64 before any OneDrive mutation", async () => {
    expect(decodeBase64Strict("YQ==").toString("utf8")).toBe("a");
    for (const value of ["YQ", "YQ=", "YQ===", "YQ==junk", "YR==", "####"]) expect(() => decodeBase64Strict(value), value).toThrow("invalid_write_input");
    for (const value of ["YQ", "YR==", "YQ==junk"]) {
      const fetchFn = vi.fn();
      await expect(driveWrite(root, "a.txt", "token", value, "text/plain", 10, false, undefined, fetchFn as typeof fetch)).rejects.toThrow("invalid_write_input");
      expect(fetchFn).not.toHaveBeenCalled();
    }
  });

  it("validates the exact 4 MiB decoded boundary without stack overflow or full-copy canonicalization", () => {
    const exact = Buffer.alloc(4 * 1024 * 1024, 97).toString("base64");
    const over = Buffer.alloc(4 * 1024 * 1024 + 1, 97).toString("base64");
    expect(decodeBase64Strict(exact).byteLength).toBe(4 * 1024 * 1024);
    expect(() => validateDriveWriteInput("large.bin", exact, "application/octet-stream", 4 * 1024 * 1024)).not.toThrow();
    expect(() => validateDriveWriteInput("large.bin", over, "application/octet-stream", 4 * 1024 * 1024)).toThrow("file_too_large");
    expect(() => validateDriveWriteInput("large.bin", over, "application/octet-stream")).not.toThrow();
  });

  it("creates uploads conditionally without overwrite", async () => {
    const fetchFn = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => new Response(JSON.stringify({ id: "item-id", name: "a.txt", size: 1, file: { mimeType: "text/plain" } }), { status: 201 }));
    await driveWrite(root, "a.txt", "token", "YQ==", "text/plain", undefined, false, undefined, fetchFn as typeof fetch);
    expect((fetchFn.mock.calls[0][1]?.headers as Record<string, string>)["if-none-match"]).toBe("*");
    expect((fetchFn.mock.calls[0][1]?.headers as Record<string, string>)["if-match"]).toBeUndefined();
  });

  it("uploads raw staged bytes without base64 re-encoding", async () => {
    const content = Buffer.from([0, 255, 1, 254]);
    const fetchFn = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => new Response(JSON.stringify({ id: "item-id", name: "binary.bin", size: content.byteLength, file: { mimeType: "application/octet-stream" } }), { status: 201 }));
    expect(() => validateDriveWriteBytes("binary.bin", content, "application/octet-stream", content.byteLength)).not.toThrow();
    expect(() => validateDriveWriteBytes("binary.bin", content, "application/octet-stream", content.byteLength - 1)).toThrow("file_too_large");
    const receipt = await driveWriteBytes(root, "binary.bin", "token", content, "application/octet-stream", content.byteLength, false, undefined, fetchFn as typeof fetch);
    expect(Buffer.from(fetchFn.mock.calls[0][1]?.body as Uint8Array)).toEqual(content);
    expect(receipt).toMatchObject({
      source_byte_size: content.byteLength,
      source_sha256: "5d8d910591d272938aef5f966e0816e374beaf7b5adf02cca5f8f770596c2ce3",
      graph_reported_size: content.byteLength,
      size_match: true,
      item: { id: "item-id", name: "binary.bin", size: content.byteLength },
    });
    expect(() => validateDriveWriteBytes("unsupported.webm", content, "video/webm", content.byteLength)).toThrow("invalid_write_input");
  });

  it.each([false, true])("uses sequential Graph upload sessions above the simple-upload threshold (update=%s)", async (update) => {
    const testChunkBytes = 320 * 1024;
    const content = Buffer.alloc(testChunkBytes + 17, 0x4d);
    const digest = createHash("sha256").update(content).digest("hex");
    const assertUnchanged = vi.fn(async () => undefined);
    const source = {
      size: content.byteLength,
      sha256: digest,
      readChunk: vi.fn(async (offset: number, maximumBytes: number) => content.subarray(offset, offset + maximumBytes)),
      assertUnchanged,
    };
    const fetchFn = vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
      const url = String(input);
      if (update && init?.method === "GET") return Response.json({ eTag: "stable-etag", file: {} });
      if (url.startsWith("https://graph.microsoft.com/") && init?.method === "POST") {
        return Response.json({ uploadUrl: "https://synthetic.up.1drv.com/up/opaque?token=private" });
      }
      if (url.startsWith("https://synthetic.up.1drv.com/")) {
        const range = (init?.headers as Record<string, string>)["content-range"];
        if (range.startsWith("bytes 0-")) return Response.json({ nextExpectedRanges: [`${testChunkBytes}-`] }, { status: 202 });
        return Response.json({ id: "session-item", name: "large.bin", size: content.byteLength, file: { mimeType: "application/octet-stream" } }, { status: update ? 200 : 201 });
      }
      throw new Error(`unexpected synthetic URL: ${url}`);
    });
    const receipt = await driveWriteSource(root, "large.bin", "token", source, "application/octet-stream", update, undefined, fetchFn as typeof fetch, 30_000, testChunkBytes, testChunkBytes);
    expect(receipt).toMatchObject({ upload_mode: "session", chunks: 2, source_byte_size: content.byteLength, source_sha256: digest, graph_reported_size: content.byteLength, size_match: true });
    const createCall = fetchFn.mock.calls.find(([input, init]) => String(input).startsWith("https://graph.microsoft.com/") && init?.method === "POST")!;
    expect(JSON.parse(String(createCall[1]?.body)).item["@microsoft.graph.conflictBehavior"]).toBe(update ? "replace" : "fail");
    expect(createCall[1]?.headers).toMatchObject({ authorization: "Bearer token", ...(update ? { "if-match": "stable-etag" } : {}) });
    const uploadCalls = fetchFn.mock.calls.filter(([input, init]) => String(input).startsWith("https://synthetic.up.1drv.com/") && init?.method === "PUT");
    expect(uploadCalls).toHaveLength(2);
    expect((uploadCalls[0][1]?.headers as Record<string, string>)["content-length"]).toBe(String(testChunkBytes));
    expect(testChunkBytes % (320 * 1024)).toBe(0);
    expect(testChunkBytes).toBeLessThan(60 * 1024 * 1024);
    for (const [, init] of uploadCalls) expect(init?.headers).not.toHaveProperty("authorization");
    const uploaded = Buffer.concat(uploadCalls.map(([, init]) => Buffer.from(init?.body as Uint8Array)));
    expect(uploaded.byteLength).toBe(content.byteLength);
    expect(createHash("sha256").update(uploaded).digest("hex")).toBe(digest);
    expect(assertUnchanged).toHaveBeenCalledTimes(1);
  });

  it("fails closed when Graph reports a different OneDrive size without a post-read", async () => {
    const content = Buffer.from("synthetic");
    const fetchFn = vi.fn(async () => new Response(JSON.stringify({ id: "item-id", name: "binary.bin", size: content.byteLength + 1, file: { mimeType: "application/octet-stream" } }), { status: 201 }));
    await expect(driveWriteBytes(root, "binary.bin", "token", content, "application/octet-stream", content.byteLength, false, undefined, fetchFn as typeof fetch)).rejects.toThrow("invalid_provider_response");
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it("budgets attachment JSON for encoded overhead while preserving the raw-byte bound", () => {
    expect(base64JsonResponseLimit(1024)).toBe(4 * Math.ceil(1024 / 3) + 64 * 1024);
    expect(base64JsonResponseLimit(1024)).toBeGreaterThan(Buffer.from(Buffer.alloc(1024)).toString("base64").length);
  });

  it("updates root-confined drive metadata and creates folders", async () => {
    let request = 0;
    const fetchFn = vi.fn(async (_input: string | URL | Request, init?: RequestInit) => {
      request += 1;
      if (request === 1) return new Response(JSON.stringify({ id: "folder-id", folder: {} }), { status: 200 });
      if (request === 2) return new Response(JSON.stringify({ eTag: "etag-1" }), { status: 200 });
      return new Response(JSON.stringify({ id: "item-id", name: "renamed.txt", description: "SYNTHETIC_DESCRIPTION", fileSystemInfo: { lastModifiedDateTime: "2026-09-07T12:00:00Z" }, file: { mimeType: "text/plain" } }), { status: init?.method === "POST" ? 201 : 200 });
    });
    const updated = await driveMetadataUpdate(root, "old.txt", "token", { name: "renamed.txt", destinationRelativePath: "SYNTHETIC_DESTINATION", description: "SYNTHETIC_DESCRIPTION", fileSystemInfo: { lastModifiedDateTime: "2026-09-07T12:00:00Z" } }, undefined, fetchFn as typeof fetch);
    expect(updated.item).toMatchObject({ name: "renamed.txt", description: "SYNTHETIC_DESCRIPTION", file_system_info: { lastModifiedDateTime: "2026-09-07T12:00:00Z" } });
    expect(JSON.parse(String(fetchFn.mock.calls[2][1]?.body))).toMatchObject({ name: "renamed.txt", parentReference: { id: "folder-id" }, description: "SYNTHETIC_DESCRIPTION" });
    expect((fetchFn.mock.calls[2][1]?.headers as Record<string, string>)["if-match"]).toBe("etag-1");

    const folderFetch = vi.fn(async (_input: string | URL | Request, init?: RequestInit) => new Response(JSON.stringify({ id: "new-folder", name: "SYNTHETIC_NEW_FOLDER", folder: {} }), { status: 201 }));
    await driveCreateFolder(root, "SYNTHETIC_PARENT", "SYNTHETIC_NEW_FOLDER", "fail", "token", undefined, folderFetch as typeof fetch);
    expect(JSON.parse(String(folderFetch.mock.calls[0][1]?.body))).toEqual({ name: "SYNTHETIC_NEW_FOLDER", folder: {}, "@microsoft.graph.conflictBehavior": "fail" });
  });
});

describe("attachment upload sessions", () => {
  it("accepts only kind-specific Microsoft upload URLs", () => {
    expect(canonicalAttachmentUploadUrl("https://outlook.office.com/api/v2.0/AttachmentSessions('id')?authtoken=opaque", "outlook")).toContain("outlook.office.com");
    expect(canonicalAttachmentUploadUrl("https://graph.microsoft.com/v1.0/users/u/todo/lists/l/tasks/t/attachmentSessions/s", "todo")).toContain("graph.microsoft.com");
    for (const value of [
      "http://outlook.office.com/session",
      "https://user@outlook.office.com/session",
      "https://outlook.office.com:444/session",
      "https://evil.invalid/session",
      "https://sub.outlook.office.com/session",
      "https://outlook.office.com/a/../session",
      "https://outlook.office.com/a/%252e%252e/session",
      "https://outlook.office.com/session#fragment",
    ]) expect(() => canonicalAttachmentUploadUrl(value, "outlook"), value).toThrow("invalid_provider_response");
    expect(() => canonicalAttachmentUploadUrl("https://graph.microsoft.com/v1.0/session?token=unexpected", "todo")).toThrow("invalid_provider_response");
  });

  it("uploads Outlook attachments sequentially without forwarding bearer auth", async () => {
    const content = Buffer.alloc(ATTACHMENT_UPLOAD_CHUNK_BYTES + 17, 7);
    let call = 0;
    const fetchFn = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => {
      call += 1;
      if (call === 1) return new Response(JSON.stringify({ uploadUrl: "https://outlook.office.com/api/v2.0/AttachmentSessions('id')?authtoken=opaque", nextExpectedRanges: ["0-"] }), { status: 201 });
      if (call === 2) return new Response(JSON.stringify({ nextExpectedRanges: [`${ATTACHMENT_UPLOAD_CHUNK_BYTES}-`] }), { status: 200 });
      return new Response(null, { status: 201 });
    });
    const result = await uploadAttachmentSession("token", "/me/messages/m/attachments/createUploadSession", { AttachmentItem: { attachmentType: "file", name: "large.bin", size: content.byteLength } }, content.toString("base64"), "outlook", undefined, fetchFn as typeof fetch);
    expect(result).toEqual({ bytes: content.byteLength, chunks: 2, upload_mode: "session" });
    const firstChunk = fetchFn.mock.calls[1];
    const finalChunk = fetchFn.mock.calls[2];
    expect(String(firstChunk[0])).toContain("authtoken=opaque");
    expect((firstChunk[1]?.headers as Record<string, string>).authorization).toBeUndefined();
    expect((firstChunk[1]?.headers as Record<string, string>)["content-range"]).toBe(`bytes 0-${ATTACHMENT_UPLOAD_CHUNK_BYTES - 1}/${content.byteLength}`);
    expect((finalChunk[1]?.headers as Record<string, string>)["content-range"]).toBe(`bytes ${ATTACHMENT_UPLOAD_CHUNK_BYTES}-${content.byteLength - 1}/${content.byteLength}`);
    expect(firstChunk[1]?.redirect).toBe("error");
  });

  it("uses the authorized To Do content endpoint and cancels failed sessions safely", async () => {
    const todoFetch = vi.fn(async (_input: string | URL | Request, _init?: RequestInit) => todoFetch.mock.calls.length === 1
      ? new Response(JSON.stringify({ uploadUrl: "https://graph.microsoft.com/v1.0/users/u/todo/lists/l/tasks/t/attachmentSessions/s", nextExpectedRanges: ["0-"] }), { status: 200 })
      : new Response(null, { status: 201 }));
    await uploadAttachmentSession("token", "/me/todo/lists/l/tasks/t/attachments/createUploadSession", { attachmentInfo: { attachmentType: "file", name: "a.bin", size: 1 } }, "YQ==", "todo", undefined, todoFetch as typeof fetch);
    expect(String(todoFetch.mock.calls[1][0])).toMatch(/attachmentSessions\/s\/content$/);
    expect((todoFetch.mock.calls[1][1]?.headers as Record<string, string>).authorization).toBe("Bearer token");

    const outlookFetch = vi.fn(async (_input: string | URL | Request, init?: RequestInit) => {
      if (outlookFetch.mock.calls.length === 1) return new Response(JSON.stringify({ uploadUrl: "https://outlook.office.com/api/v2.0/AttachmentSessions('failed')?authtoken=opaque", nextExpectedRanges: ["0-"] }), { status: 201 });
      if (init?.method === "PUT") return new Response(null, { status: 500 });
      return new Response(null, { status: 204 });
    });
    await expect(uploadAttachmentSession("token", "/me/messages/m/attachments/createUploadSession", { AttachmentItem: { attachmentType: "file", name: "a.bin", size: 1 } }, "YQ==", "outlook", undefined, outlookFetch as typeof fetch)).rejects.toThrow("provider_error_500");
    expect(outlookFetch.mock.calls[2][1]?.method).toBe("DELETE");
    expect((outlookFetch.mock.calls[2][1]?.headers as Record<string, string> | undefined)?.authorization).toBeUndefined();
  });
});
