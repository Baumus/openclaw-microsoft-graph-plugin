import { createHash, createHmac, randomBytes, timingSafeEqual } from "node:crypto";

export const ONEDRIVE_AGENTS_MAX_DEPTH = 16;
export const ONEDRIVE_AGENTS_MAX_REQUESTS = ONEDRIVE_AGENTS_MAX_DEPTH + 1;
export const ONEDRIVE_AGENTS_MAX_BATCH_DIRECTORIES = 3;
export const ONEDRIVE_AGENTS_MAX_FILE_BYTES = 32 * 1024;
export const ONEDRIVE_AGENTS_MAX_OUTPUT_BYTES = 32 * 1024;
export const ONEDRIVE_AGENTS_MAX_SERIALIZED_OUTPUT_BYTES = 40 * 1024;
export const ONEDRIVE_AGENTS_MAX_PARALLEL = 4;

export type OneDriveAgentsCacheLimits = {
  maxSessions: number;
  maxRoots: number;
  maxEntries: number;
  maxBytes: number;
  ttlMs: number;
};

export const ONEDRIVE_AGENTS_CACHE_LIMITS: OneDriveAgentsCacheLimits = {
  maxSessions: 64,
  maxRoots: 128,
  maxEntries: 2_048,
  maxBytes: 4 * 1024 * 1024,
  ttlMs: 12 * 60 * 60 * 1_000,
};

type PresentEntry = {
  present: true;
  relativePath: string;
  content: string;
  bytes: number;
  sha256: string;
};
type AbsentEntry = { present: false; relativePath: string };
type CacheEntry = PresentEntry | AbsentEntry;
type InFlightCandidate = {
  promise: Promise<CacheEntry>;
  controller: AbortController;
  waiters: number;
};
type RootCache = {
  key: string;
  sessionId: string;
  agentId: string;
  rootPin: string;
  entries: Map<string, CacheEntry>;
  inFlight: Map<string, InFlightCandidate>;
  bytes: number;
  lastAccess: number;
  closed: boolean;
  abortController: AbortController;
  expiryTimer?: ReturnType<typeof setTimeout>;
  acknowledgedEntries: Set<string>;
  receiptKey: Buffer;
};

export type OneDriveAgentsInstruction = {
  relativePath: string;
  bytes: number;
  sha256: string;
  content: string;
};

export type OneDriveAgentsChainEntry = Omit<OneDriveAgentsInstruction, "content">;

export type OneDriveAgentsInstructionsResult = {
  ok: true;
  managed: boolean;
  rootLabel: string;
  relativeDirectory: string;
  relativeDirectories?: string[];
  cacheHit: boolean;
  coalesced: boolean;
  instructionsIncluded: boolean;
  acknowledgementRequired: boolean;
  chain: OneDriveAgentsChainEntry[];
  acknowledgement?: string;
  instructions?: OneDriveAgentsInstruction[];
};

type LoadCandidate = (relativePath: string, signal?: AbortSignal) => Promise<Uint8Array | null>;

function assertPositiveInteger(value: number): void {
  if (!Number.isSafeInteger(value) || value < 1) throw new Error("invalid_cache_limits");
}

function rootCacheKey(sessionId: string, agentId: string, rootPin: string): string {
  return JSON.stringify([sessionId, agentId, rootPin]);
}

function receiptForChain(root: RootCache, chainHash: string): string {
  return createHmac("sha256", root.receiptKey).update(chainHash).digest().subarray(0, 24).toString("base64url");
}

function receiptMatches(expected: string, acknowledgement: string): boolean {
  if (!/^[A-Za-z0-9_-]{32}$/.test(acknowledgement)) return false;
  return timingSafeEqual(Buffer.from(expected, "ascii"), Buffer.from(acknowledgement, "ascii"));
}

function candidatePaths(relativeDirectory: string): string[] {
  const segments = relativeDirectory ? relativeDirectory.split("/") : [];
  if (segments.length > ONEDRIVE_AGENTS_MAX_DEPTH) throw new Error("instruction_depth_exceeded");
  const paths = ["AGENTS.md"];
  for (let depth = 1; depth <= segments.length; depth += 1) paths.push(`${segments.slice(0, depth).join("/")}/AGENTS.md`);
  if (paths.length > ONEDRIVE_AGENTS_MAX_REQUESTS) throw new Error("instruction_request_limit_exceeded");
  return paths;
}

function candidatePathsForDirectories(relativeDirectories: string[]): string[] {
  if (relativeDirectories.length < 1 || relativeDirectories.length > ONEDRIVE_AGENTS_MAX_BATCH_DIRECTORIES) {
    throw new Error("instruction_directory_limit_exceeded");
  }
  const paths = new Set<string>();
  for (const relativeDirectory of relativeDirectories) {
    for (const path of candidatePaths(relativeDirectory)) paths.add(path);
  }
  return [...paths];
}

function decodeInstruction(relativePath: string, bytes: Uint8Array | null): CacheEntry {
  if (bytes === null) return { present: false, relativePath };
  if (bytes.byteLength > ONEDRIVE_AGENTS_MAX_FILE_BYTES) throw new Error("instruction_file_too_large");
  let content: string;
  try { content = new TextDecoder("utf-8", { fatal: true }).decode(bytes); }
  catch { throw new Error("invalid_instruction_encoding"); }
  return {
    present: true,
    relativePath,
    content,
    bytes: bytes.byteLength,
    sha256: createHash("sha256").update(bytes).digest("hex"),
  };
}

async function mapBounded<T>(values: string[], limit: number, action: (value: string) => Promise<T>): Promise<T[]> {
  const results = new Array<T>(values.length);
  let cursor = 0;
  const workers = Array.from({ length: Math.min(limit, values.length) }, async () => {
    while (cursor < values.length) {
      const index = cursor;
      cursor += 1;
      results[index] = await action(values[index]);
    }
  });
  await Promise.all(workers);
  return results;
}

async function waitForCandidate<T>(promise: Promise<T>, signal?: AbortSignal): Promise<T> {
  signal?.throwIfAborted();
  if (!signal) return promise;
  let abort: (() => void) | undefined;
  try {
    return await Promise.race([
      promise,
      new Promise<never>((_resolve, reject) => {
        abort = () => reject(signal.reason);
        signal.addEventListener("abort", abort, { once: true });
      }),
    ]);
  } finally {
    if (abort) signal.removeEventListener("abort", abort);
  }
}

export class OneDriveAgentsSessionCache {
  private readonly roots = new Map<string, RootCache>();

  constructor(private readonly limits: OneDriveAgentsCacheLimits = ONEDRIVE_AGENTS_CACHE_LIMITS) {
    assertPositiveInteger(limits.maxSessions);
    assertPositiveInteger(limits.maxRoots);
    assertPositiveInteger(limits.maxEntries);
    assertPositiveInteger(limits.maxBytes);
    assertPositiveInteger(limits.ttlMs);
  }

  clear(): void {
    for (const root of [...this.roots.values()]) this.dispose(root, "instruction_cache_cleared");
  }

  clearSession(sessionId: string): void {
    for (const root of [...this.roots.values()]) {
      if (root.sessionId === sessionId) this.dispose(root, "instruction_session_ended");
    }
  }

  stats(): { sessions: number; roots: number; entries: number; bytes: number } {
    const sessions = new Set<string>();
    let entries = 0;
    let bytes = 0;
    for (const root of this.roots.values()) {
      sessions.add(root.sessionId);
      entries += root.entries.size;
      bytes += root.bytes;
    }
    return { sessions: sessions.size, roots: this.roots.size, entries, bytes };
  }

  /** Approval preflight: inspect only previously discovered instructions; never fetch. */
  cachedAcknowledgement(params: { sessionId: string; agentId: string; rootPin: string; relativeDirectories: string[]; acknowledgement?: string; now?: number }): "ready" | "discover" | "invalid" {
    const root = this.roots.get(rootCacheKey(params.sessionId, params.agentId, params.rootPin));
    if (!root || root.closed || (params.now ?? Date.now()) - root.lastAccess >= this.limits.ttlMs) return "discover";
    const paths = candidatePathsForDirectories([...new Set(params.relativeDirectories)]);
    const entries = paths.map((path) => root.entries.get(path));
    if (entries.some((entry) => !entry)) return "discover";
    if (!entries[0]!.present) return "ready";
    const chain = entries.flatMap((entry) => entry!.present ? [{ relativePath: entry!.relativePath, bytes: entry!.bytes, sha256: entry!.sha256 }] : []);
    if (params.acknowledgement !== undefined) {
      const chainHash = createHash("sha256").update(JSON.stringify(chain)).digest("hex");
      return receiptMatches(receiptForChain(root, chainHash), params.acknowledgement) ? "ready" : "invalid";
    }
    return chain.every((entry) => root.acknowledgedEntries.has(`${entry.relativePath}\u0000${entry.sha256}`)) ? "ready" : "discover";
  }

  private prune(now: number): void {
    for (const root of [...this.roots.values()]) {
      if (root.inFlight.size === 0 && now - root.lastAccess >= this.limits.ttlMs) this.dispose(root, "instruction_cache_expired");
    }
  }

  private dispose(root: RootCache, reason: string): void {
    if (root.closed) return;
    root.closed = true;
    if (root.expiryTimer) clearTimeout(root.expiryTimer);
    root.abortController.abort(new DOMException(reason, "AbortError"));
    root.receiptKey.fill(0);
    if (this.roots.get(root.key) === root) this.roots.delete(root.key);
  }

  private assertActive(root: RootCache): void {
    if (root.closed || this.roots.get(root.key) !== root) {
      throw new Error("instruction_session_ended");
    }
  }

  private scheduleExpiry(root: RootCache): void {
    if (root.closed) return;
    if (root.expiryTimer) clearTimeout(root.expiryTimer);
    const remaining = this.limits.ttlMs - (Date.now() - root.lastAccess);
    root.expiryTimer = setTimeout(() => {
      if (root.closed || this.roots.get(root.key) !== root) return;
      if (root.inFlight.size > 0) return;
      if (Date.now() - root.lastAccess >= this.limits.ttlMs) this.dispose(root, "instruction_cache_expired");
      else this.scheduleExpiry(root);
    }, Math.max(1, remaining));
    root.expiryTimer.unref?.();
  }

  private touch(root: RootCache, now: number): void {
    root.lastAccess = now;
    this.roots.delete(root.key);
    this.roots.set(root.key, root);
    this.scheduleExpiry(root);
  }

  private enforceLimits(additionalPending = 0): void {
    const over = () => {
      const stats = this.stats();
      return stats.sessions > this.limits.maxSessions || stats.roots > this.limits.maxRoots
        || stats.entries + [...this.roots.values()].reduce((total, root) => total + root.inFlight.size, 0) + additionalPending > this.limits.maxEntries
        || stats.bytes > this.limits.maxBytes;
    };
    if (over()) throw new Error("instruction_cache_capacity_exceeded");
  }

  private root(sessionId: string, agentId: string, rootPin: string, now: number): RootCache {
    this.prune(now);
    const key = rootCacheKey(sessionId, agentId, rootPin);
    const existing = this.roots.get(key);
    if (existing) {
      this.touch(existing, now);
      return existing;
    }
    const created: RootCache = {
      key, sessionId, agentId, rootPin, entries: new Map(), inFlight: new Map(), bytes: 0, lastAccess: now,
      closed: false, abortController: new AbortController(), acknowledgedEntries: new Set(), receiptKey: randomBytes(32),
    };
    this.roots.set(key, created);
    this.scheduleExpiry(created);
    try { this.enforceLimits(); }
    catch (error) { this.dispose(created, "instruction_cache_capacity_exceeded"); throw error; }
    return created;
  }

  private async candidate(root: RootCache, relativePath: string, load: LoadCandidate, signal: AbortSignal | undefined): Promise<{ entry: CacheEntry; source: "cache" | "coalesced" | "load" }> {
    signal?.throwIfAborted();
    const cached = root.entries.get(relativePath);
    if (cached) return { entry: cached, source: "cache" };
    const inFlight = root.inFlight.get(relativePath);
    if (inFlight) return { entry: await this.waitForFlight(root, relativePath, inFlight, signal), source: "coalesced" };
    this.enforceLimits(1);
    const controller = new AbortController();
    const onRootAbort = () => controller.abort(root.abortController.signal.reason);
    if (root.abortController.signal.aborted) onRootAbort();
    else root.abortController.signal.addEventListener("abort", onRootAbort, { once: true });
    const pending = (async () => {
      const entry = decodeInstruction(relativePath, await load(relativePath, controller.signal));
      if (root.closed) throw new Error("instruction_session_ended");
      root.entries.set(relativePath, entry);
      if (entry.present) root.bytes += entry.bytes;
      try { this.enforceLimits(-1); }
      catch (error) {
        root.entries.delete(relativePath);
        if (entry.present) root.bytes -= entry.bytes;
        throw error;
      }
      return entry;
    })();
    const flight: InFlightCandidate = { promise: pending, controller, waiters: 0 };
    root.inFlight.set(relativePath, flight);
    void pending.finally(() => {
      root.abortController.signal.removeEventListener("abort", onRootAbort);
      if (root.inFlight.get(relativePath) === flight) root.inFlight.delete(relativePath);
      this.scheduleExpiry(root);
    }).catch(() => undefined);
    return { entry: await this.waitForFlight(root, relativePath, flight, signal), source: "load" };
  }

  private async waitForFlight(root: RootCache, relativePath: string, flight: InFlightCandidate, signal?: AbortSignal): Promise<CacheEntry> {
    flight.waiters += 1;
    try {
      return await waitForCandidate(flight.promise, signal);
    } finally {
      flight.waiters -= 1;
      if (flight.waiters === 0 && root.inFlight.get(relativePath) === flight && !flight.controller.signal.aborted) {
        flight.controller.abort(new DOMException("instruction_no_waiters", "AbortError"));
      }
    }
  }

  async discover(params: {
    sessionId: string;
    agentId: string;
    rootPin: string;
    rootLabel: string;
    relativeDirectory: string;
    load: LoadCandidate;
    acknowledgement?: string;
    signal?: AbortSignal;
    parallelism?: number;
    now?: number;
  }): Promise<OneDriveAgentsInstructionsResult> {
    return this.discoverMany({ ...params, relativeDirectories: [params.relativeDirectory] });
  }

  async discoverMany(params: {
    sessionId: string;
    agentId: string;
    rootPin: string;
    rootLabel: string;
    relativeDirectories: string[];
    load: LoadCandidate;
    acknowledgement?: string;
    signal?: AbortSignal;
    parallelism?: number;
    now?: number;
  }): Promise<OneDriveAgentsInstructionsResult> {
    const { sessionId, agentId, rootPin, rootLabel, relativeDirectories, load, signal, acknowledgement } = params;
    if (!sessionId || !agentId || !rootPin) throw new Error("trusted_session_identity_required");
    const directories = [...new Set(relativeDirectories)];
    const paths = candidatePathsForDirectories(directories);
    const relativeDirectory = directories[0];
    const root = this.root(sessionId, agentId, rootPin, params.now ?? Date.now());
    const rootResult = await this.candidate(root, paths[0], load, signal);
    this.assertActive(root);
    if (!rootResult.entry.present) {
      return {
        ok: true, managed: false, rootLabel, relativeDirectory,
        ...(directories.length > 1 ? { relativeDirectories: directories } : {}),
        cacheHit: rootResult.source === "cache", coalesced: rootResult.source === "coalesced",
        instructionsIncluded: false, acknowledgementRequired: false, chain: [],
      };
    }
    const parallelism = Math.max(1, Math.min(params.parallelism ?? ONEDRIVE_AGENTS_MAX_PARALLEL, ONEDRIVE_AGENTS_MAX_PARALLEL));
    const nested = await mapBounded(paths.slice(1), parallelism, (path) => this.candidate(root, path, load, signal));
    this.assertActive(root);
    const resolved = [rootResult, ...nested];
    const outputBytes = resolved.reduce((total, { entry }) => total + (entry.present ? entry.bytes : 0), 0);
    if (outputBytes > ONEDRIVE_AGENTS_MAX_OUTPUT_BYTES) throw new Error("instruction_output_too_large");
    const chain = resolved.flatMap(({ entry }) => entry.present ? [{ relativePath: entry.relativePath, bytes: entry.bytes, sha256: entry.sha256 }] : []);
    const chainHash = createHash("sha256").update(JSON.stringify(chain)).digest("hex");
    const expectedReceipt = receiptForChain(root, chainHash);
    if (acknowledgement !== undefined) {
      if (!receiptMatches(expectedReceipt, acknowledgement)) {
        throw new Error("instruction_acknowledgement_invalid");
      }
      for (const entry of chain) root.acknowledgedEntries.add(`${entry.relativePath}\u0000${entry.sha256}`);
    }
    const instructions: OneDriveAgentsInstruction[] = resolved.flatMap(({ entry }) => entry.present
      && !root.acknowledgedEntries.has(`${entry.relativePath}\u0000${entry.sha256}`)
        ? [{ relativePath: entry.relativePath, bytes: entry.bytes, sha256: entry.sha256, content: entry.content }]
        : []);
    const acknowledged = instructions.length === 0;
    let receipt: string | undefined;
    if (!acknowledged) {
      receipt = expectedReceipt;
    }
    const sources = resolved.map(({ source }) => source);
    const result: OneDriveAgentsInstructionsResult = {
      ok: true, managed: true, rootLabel, relativeDirectory,
      ...(directories.length > 1 ? { relativeDirectories: directories } : {}),
      cacheHit: sources.every((source) => source === "cache"),
      coalesced: sources.some((source) => source === "coalesced"),
      instructionsIncluded: instructions.length > 0,
      acknowledgementRequired: !acknowledged,
      chain,
      ...(receipt ? { acknowledgement: receipt } : {}),
      ...(instructions.length ? { instructions } : {}),
    };
    if (Buffer.byteLength(JSON.stringify(result), "utf8") > ONEDRIVE_AGENTS_MAX_SERIALIZED_OUTPUT_BYTES) {
      throw new Error("instruction_serialized_output_too_large");
    }
    return result;
  }
}

export const oneDriveAgentsSessionCache = new OneDriveAgentsSessionCache();
