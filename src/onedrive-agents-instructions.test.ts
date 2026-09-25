import { describe, expect, it, vi } from "vitest";
import {
  ONEDRIVE_AGENTS_MAX_DEPTH,
  ONEDRIVE_AGENTS_MAX_FILE_BYTES,
  ONEDRIVE_AGENTS_MAX_SERIALIZED_OUTPUT_BYTES,
  OneDriveAgentsSessionCache,
} from "./onedrive-agents-instructions.js";

const encoded = (value: string) => new TextEncoder().encode(value);
const base = { agentId: "agent-a", sessionId: "session-a", rootPin: "drive:item", rootLabel: "workspace" };

describe("OneDrive AGENTS.md session cache", () => {
  it("requires stable session and agent identities", async () => {
    const cache = new OneDriveAgentsSessionCache();
    const load = async () => encoded("root");
    await expect(cache.discover({ ...base, sessionId: "", relativeDirectory: "", load })).rejects.toThrow("trusted_session_identity_required");
    await expect(cache.discover({ ...base, agentId: "", relativeDirectory: "", load })).rejects.toThrow("trusted_session_identity_required");
  });

  it("returns a compact unmanaged result and negatively caches the absent root", async () => {
    const cache = new OneDriveAgentsSessionCache();
    let reads = 0;
    const load = async () => { reads += 1; return null; };
    const first = await cache.discover({ ...base, relativeDirectory: "a/b", load });
    expect(first).toEqual({ ok: true, managed: false, rootLabel: "workspace", relativeDirectory: "a/b", cacheHit: false, coalesced: false, instructionsIncluded: false, acknowledgementRequired: false, chain: [] });
    const second = await cache.discover({ ...base, relativeDirectory: "other", load });
    expect(second).toMatchObject({ managed: false, cacheHit: true, instructionsIncluded: false, chain: [] });
    expect(reads).toBe(1);
  });

  it("orders and deduplicates root-to-leaf instructions while tolerating nested absence", async () => {
    const cache = new OneDriveAgentsSessionCache();
    const files = new Map([
      ["AGENTS.md", encoded("root")],
      ["a/b/AGENTS.md", encoded("leaf")],
    ]);
    const calls: string[] = [];
    const result = await cache.discover({ ...base, relativeDirectory: "a/b", load: async (path) => { calls.push(path); return files.get(path) ?? null; } });
    expect(calls).toEqual(["AGENTS.md", "a/AGENTS.md", "a/b/AGENTS.md"]);
    expect(result.chain.map((entry) => entry.relativePath)).toEqual(["AGENTS.md", "a/b/AGENTS.md"]);
    expect(result.instructions?.map((entry) => entry.content)).toEqual(["root", "leaf"]);
    expect(new Set(result.chain.map((entry) => entry.relativePath)).size).toBe(result.chain.length);
  });

  it("combines multiple directory chains into one deduplicated receipt", async () => {
    const cache = new OneDriveAgentsSessionCache();
    const calls: string[] = [];
    const load = async (path: string) => { calls.push(path); return encoded(path); };
    const first = await cache.discoverMany({ ...base, relativeDirectories: ["a/folder", "b"], load });
    expect(first.relativeDirectories).toEqual(["a/folder", "b"]);
    expect(first.instructions?.map((entry) => entry.relativePath)).toEqual([
      "AGENTS.md", "a/AGENTS.md", "a/folder/AGENTS.md", "b/AGENTS.md",
    ]);
    await expect(cache.discoverMany({
      ...base,
      relativeDirectories: ["a/folder", "b"],
      acknowledgement: first.acknowledgement,
      load,
    })).resolves.toMatchObject({ instructionsIncluded: false, acknowledgementRequired: false, cacheHit: true });
    expect(calls).toEqual(["AGENTS.md", "a/AGENTS.md", "a/folder/AGENTS.md", "b/AGENTS.md"]);
  });

  it("keeps unacknowledged chain receipts stateless and invalidates them with the root cache", async () => {
    const cache = new OneDriveAgentsSessionCache({ maxSessions: 2, maxRoots: 2, maxEntries: 20, maxBytes: 1024, ttlMs: 60_000 });
    const load = async (path: string) => encoded(path);
    const receipts = new Set<string>();
    for (const directory of ["a", "b", "c", "d", "e"]) {
      const result = await cache.discover({ ...base, relativeDirectory: directory, load });
      receipts.add(result.acknowledgement!);
    }
    expect(receipts.size).toBe(5);
    expect(cache.stats()).toEqual({ sessions: 1, roots: 1, entries: 6, bytes: 64 });

    const staleReceipt = [...receipts][0];
    cache.clearSession(base.sessionId);
    await expect(cache.discover({ ...base, relativeDirectory: "a", acknowledgement: staleReceipt, load }))
      .rejects.toThrow("instruction_acknowledgement_invalid");
  });

  it("uses snapshot semantics for repeated, deeper, and sibling discovery", async () => {
    const cache = new OneDriveAgentsSessionCache();
    const calls: string[] = [];
    const load = async (path: string) => { calls.push(path); return path === "a/missing/AGENTS.md" ? null : encoded(path); };
    const first = await cache.discover({ ...base, relativeDirectory: "a", load });
    expect(first.instructions?.map((entry) => entry.relativePath)).toEqual(["AGENTS.md", "a/AGENTS.md"]);
    expect(first.acknowledgement).toMatch(/^[A-Za-z0-9_-]{32}$/);
    const repeated = await cache.discover({ ...base, relativeDirectory: "a", acknowledgement: first.acknowledgement, load });
    expect(repeated).toMatchObject({ cacheHit: true, instructionsIncluded: false });
    expect(repeated).not.toHaveProperty("instructions");
    const deeper = await cache.discover({ ...base, relativeDirectory: "a/missing/b", load });
    expect(deeper.instructions?.map((entry) => entry.relativePath)).toEqual(["a/missing/b/AGENTS.md"]);
    const belowNegative = await cache.discover({ ...base, relativeDirectory: "a/missing/c", load });
    expect(belowNegative.instructions?.map((entry) => entry.relativePath)).toEqual(["a/missing/c/AGENTS.md"]);
    const sibling = await cache.discover({ ...base, relativeDirectory: "a/c", load });
    expect(sibling.instructions?.map((entry) => entry.relativePath)).toEqual(["a/c/AGENTS.md"]);
    expect(calls).toEqual(["AGENTS.md", "a/AGENTS.md", "a/missing/AGENTS.md", "a/missing/b/AGENTS.md", "a/missing/c/AGENTS.md", "a/c/AGENTS.md"]);
  });

  it("isolates sessions and agents", async () => {
    const cache = new OneDriveAgentsSessionCache();
    let reads = 0;
    const load = async () => { reads += 1; return encoded(`body-${reads}`); };
    await cache.discover({ ...base, relativeDirectory: "", load });
    await cache.discover({ ...base, sessionId: "session-b", relativeDirectory: "", load });
    await cache.discover({ ...base, agentId: "agent-b", relativeDirectory: "", load });
    await cache.discover({ ...base, rootPin: "other-drive:item", relativeDirectory: "", load });
    expect(reads).toBe(4);
    expect(cache.stats()).toMatchObject({ sessions: 2, roots: 4, entries: 4 });
  });

  it("caps one deepest discovery at 17 requests and four parallel nested reads", async () => {
    const cache = new OneDriveAgentsSessionCache();
    let reads = 0;
    let active = 0;
    let peak = 0;
    const load = async (path: string) => {
      reads += 1;
      if (path === "AGENTS.md") return encoded("root");
      active += 1;
      peak = Math.max(peak, active);
      await new Promise((resolve) => setTimeout(resolve, 2));
      active -= 1;
      return null;
    };
    const deepest = Array.from({ length: ONEDRIVE_AGENTS_MAX_DEPTH }, (_, index) => `d${index}`).join("/");
    await cache.discover({ ...base, relativeDirectory: deepest, load, parallelism: 99 });
    expect(reads).toBe(ONEDRIVE_AGENTS_MAX_DEPTH + 1);
    expect(peak).toBe(4);
  });

  it("coalesces concurrent duplicate reads and suppresses bodies only after receipt acknowledgement", async () => {
    const cache = new OneDriveAgentsSessionCache();
    let reads = 0;
    let release!: (value: Uint8Array) => void;
    const pending = new Promise<Uint8Array>((resolve) => { release = resolve; });
    const load = async () => { reads += 1; return pending; };
    const first = cache.discover({ ...base, relativeDirectory: "", load });
    const second = cache.discover({ ...base, relativeDirectory: "", load });
    const controller = new AbortController();
    const cancelledWaiter = cache.discover({ ...base, relativeDirectory: "", load, signal: controller.signal });
    controller.abort(new DOMException("cancelled", "AbortError"));
    await expect(cancelledWaiter).rejects.toMatchObject({ name: "AbortError" });
    release(encoded("root"));
    const results = await Promise.all([first, second]);
    expect(reads).toBe(1);
    expect(results.filter((result) => result.instructionsIncluded)).toHaveLength(2);
    expect(results[0].acknowledgement).toBe(results[1].acknowledgement);
    expect(results.some((result) => result.coalesced)).toBe(true);
    const acknowledged = await cache.discover({ ...base, relativeDirectory: "", acknowledgement: results[0].acknowledgement, load });
    expect(acknowledged).toMatchObject({ instructionsIncluded: false, acknowledgementRequired: false, cacheHit: true });
  });

  it("keeps a shared load alive when the leading caller aborts", async () => {
    const cache = new OneDriveAgentsSessionCache();
    let reads = 0;
    let release!: (value: Uint8Array) => void;
    const pending = new Promise<Uint8Array>((resolve) => { release = resolve; });
    const load = async () => { reads += 1; return pending; };
    const leaderController = new AbortController();
    const leader = cache.discover({ ...base, relativeDirectory: "", load, signal: leaderController.signal });
    const follower = cache.discover({ ...base, relativeDirectory: "", load });
    leaderController.abort(new DOMException("cancelled", "AbortError"));
    await expect(leader).rejects.toMatchObject({ name: "AbortError" });
    release(encoded("root"));
    await expect(follower).resolves.toMatchObject({ managed: true, instructionsIncluded: true, coalesced: true });
    expect(reads).toBe(1);
  });

  it("aborts shared provider work when its final waiter cancels", async () => {
    const cache = new OneDriveAgentsSessionCache();
    let reads = 0;
    let providerSignal: AbortSignal | undefined;
    const load = async (_path: string, signal?: AbortSignal) => {
      reads += 1;
      providerSignal = signal;
      return new Promise<Uint8Array>((_resolve, reject) => signal?.addEventListener("abort", () => reject(signal.reason), { once: true }));
    };
    const controller = new AbortController();
    const discovery = cache.discover({ ...base, relativeDirectory: "", load, signal: controller.signal });
    await Promise.resolve();
    controller.abort(new DOMException("cancelled", "AbortError"));
    await expect(discovery).rejects.toMatchObject({ name: "AbortError" });
    expect(providerSignal?.aborted).toBe(true);
    await new Promise((resolve) => setTimeout(resolve, 0));
    const retry = await cache.discover({ ...base, relativeDirectory: "", load: async () => { reads += 1; return encoded("root"); } });
    expect(retry.managed).toBe(true);
    expect(reads).toBe(2);
  });

  it("counts pending candidates against hard cache capacity", async () => {
    const cache = new OneDriveAgentsSessionCache({ maxSessions: 2, maxRoots: 2, maxEntries: 1, maxBytes: 1024, ttlMs: 60_000 });
    let release!: (value: Uint8Array) => void;
    const pending = new Promise<Uint8Array>((resolve) => { release = resolve; });
    const first = cache.discover({ ...base, relativeDirectory: "", load: async () => pending });
    await Promise.resolve();
    const secondLoad = vi.fn(async () => encoded("other"));
    await expect(cache.discover({ ...base, sessionId: "session-b", rootPin: "drive:other", relativeDirectory: "", load: secondLoad })).rejects.toThrow("instruction_cache_capacity_exceeded");
    expect(secondLoad).not.toHaveBeenCalled();
    release(encoded("root"));
    await expect(first).resolves.toMatchObject({ managed: true });
  });

  it("aborts in-flight work and prevents cache insertion when the session ends", async () => {
    const cache = new OneDriveAgentsSessionCache();
    let observedSignal: AbortSignal | undefined;
    const load = async (_path: string, signal?: AbortSignal) => new Promise<Uint8Array>((resolve, reject) => {
      observedSignal = signal;
      signal?.addEventListener("abort", () => reject(signal.reason), { once: true });
      void resolve;
    });
    const discovery = cache.discover({ ...base, relativeDirectory: "", load });
    await Promise.resolve();
    cache.clearSession(base.sessionId);
    expect(observedSignal?.aborted).toBe(true);
    await expect(discovery).rejects.toMatchObject({ name: "AbortError" });
    expect(cache.stats()).toEqual({ sessions: 0, roots: 0, entries: 0, bytes: 0 });
  });

  it("rejects cached discovery when the session ends during an await continuation", async () => {
    const cache = new OneDriveAgentsSessionCache();
    const load = vi.fn(async () => encoded("root"));
    await cache.discover({ ...base, relativeDirectory: "", load });

    const cached = cache.discover({ ...base, relativeDirectory: "", load });
    cache.clearSession(base.sessionId);

    await expect(cached).rejects.toThrow("instruction_session_ended");
    expect(load).toHaveBeenCalledTimes(1);
    expect(cache.stats()).toEqual({ sessions: 0, roots: 0, entries: 0, bytes: 0 });
  });

  it("actively expires idle session content without a later cache access", async () => {
    const cache = new OneDriveAgentsSessionCache({ maxSessions: 2, maxRoots: 2, maxEntries: 20, maxBytes: 1024, ttlMs: 5 });
    await cache.discover({ ...base, relativeDirectory: "", load: async () => encoded("root") });
    expect(cache.stats().roots).toBe(1);
    await new Promise((resolve) => setTimeout(resolve, 15));
    expect(cache.stats()).toEqual({ sessions: 0, roots: 0, entries: 0, bytes: 0 });
  });

  it("enforces depth, per-file, aggregate, cancellation, and fail-closed capacity bounds", async () => {
    const cache = new OneDriveAgentsSessionCache({ maxSessions: 2, maxRoots: 2, maxEntries: 20, maxBytes: 1024 * 1024, ttlMs: 60_000 });
    const tooDeep = Array.from({ length: ONEDRIVE_AGENTS_MAX_DEPTH + 1 }, (_, index) => `d${index}`).join("/");
    await expect(cache.discover({ ...base, relativeDirectory: tooDeep, load: async () => encoded("x") })).rejects.toThrow("instruction_depth_exceeded");
    await expect(cache.discover({ ...base, relativeDirectory: "", load: async () => new Uint8Array(ONEDRIVE_AGENTS_MAX_FILE_BYTES + 1) })).rejects.toThrow("instruction_file_too_large");

    const aggregate = new OneDriveAgentsSessionCache();
    await expect(aggregate.discover({ ...base, relativeDirectory: "a/b/c/d", load: async () => new Uint8Array(ONEDRIVE_AGENTS_MAX_FILE_BYTES) })).rejects.toThrow("instruction_output_too_large");

    const escaping = new OneDriveAgentsSessionCache();
    const escapingContent = "\u0000".repeat(Math.ceil(ONEDRIVE_AGENTS_MAX_SERIALIZED_OUTPUT_BYTES / 6));
    await expect(escaping.discover({ ...base, relativeDirectory: "", load: async () => encoded(escapingContent) })).rejects.toThrow("instruction_serialized_output_too_large");

    const capacity = new OneDriveAgentsSessionCache({ maxSessions: 1, maxRoots: 1, maxEntries: 1, maxBytes: 1024, ttlMs: 60_000 });
    await expect(capacity.discover({ ...base, relativeDirectory: "a", load: async () => encoded("x") })).rejects.toThrow("instruction_cache_capacity_exceeded");
    expect(capacity.stats()).toMatchObject({ roots: 1, entries: 1, bytes: 1 });

    const controller = new AbortController();
    controller.abort(new DOMException("cancelled", "AbortError"));
    await expect(cache.discover({ ...base, sessionId: "cancelled", relativeDirectory: "", signal: controller.signal, load: async () => encoded("x") })).rejects.toMatchObject({ name: "AbortError" });

    let reads = 0;
    const load = async () => { reads += 1; return encoded("root"); };
    const bounded = new OneDriveAgentsSessionCache({ maxSessions: 2, maxRoots: 2, maxEntries: 20, maxBytes: 1024 * 1024, ttlMs: 60_000 });
    await bounded.discover({ ...base, sessionId: "lru-1", relativeDirectory: "", load });
    await bounded.discover({ ...base, sessionId: "lru-2", relativeDirectory: "", load });
    await expect(bounded.discover({ ...base, sessionId: "lru-3", relativeDirectory: "", load })).rejects.toThrow("instruction_cache_capacity_exceeded");
    expect(bounded.stats()).toMatchObject({ sessions: 2, roots: 2 });
    await bounded.discover({ ...base, sessionId: "lru-1", relativeDirectory: "", load });
    expect(reads).toBe(2);
  });
});
