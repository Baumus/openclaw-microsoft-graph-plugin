import { randomBytes } from "node:crypto";
import { canonicalProviderContinuation } from "./graph.js";

export const CONTINUATION_TTL_MS = 5 * 60 * 1000;
export const MAX_CONTINUATIONS = 1024;
export const MAX_CONTINUATION_STATE_BYTES = 4 * 1024 * 1024;
export const MAX_CONTINUATION_STORE_BYTES = 32 * 1024 * 1024;

export type ContinuationBinding = {
  agentId: string;
  service: "calendar" | "mail" | "onedrive" | "todo";
  action: string;
  resource: string;
  criteria: string;
};

type ContinuationState = { resultCount?: number; [key: string]: unknown };

type ContinuationRecord = ContinuationBinding & {
  providerNextLink?: string;
  expectedPath: string;
  state?: ContinuationState;
  createdAt: number;
  expiresAt: number;
  storageBytes: number;
};

export type VerifiedContinuation = Readonly<ContinuationRecord>;

function normalizedValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(normalizedValue);
  if (value !== null && typeof value === "object") {
    return Object.fromEntries(Object.entries(value as Record<string, unknown>)
      .filter(([, entry]) => entry !== undefined)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, entry]) => [key, normalizedValue(entry)]));
  }
  return value;
}

/** Stable serialization for the effective, continuation-free read criteria. */
export function normalizedCriteria(value: unknown): string {
  return JSON.stringify(normalizedValue(value));
}

export class ContinuationStore {
  readonly #records = new Map<string, ContinuationRecord>();
  #storedBytes = 0;

  constructor(
    private readonly ttlMs = CONTINUATION_TTL_MS,
    private readonly maximum = MAX_CONTINUATIONS,
    private readonly now: () => number = Date.now,
    private readonly maximumBytes = MAX_CONTINUATION_STORE_BYTES,
    private readonly maximumStateBytes = MAX_CONTINUATION_STATE_BYTES,
  ) {
    if (!Number.isSafeInteger(ttlMs) || ttlMs < 1
      || !Number.isSafeInteger(maximum) || maximum < 1
      || !Number.isSafeInteger(maximumBytes) || maximumBytes < 1
      || !Number.isSafeInteger(maximumStateBytes) || maximumStateBytes < 1
      || maximumStateBytes > maximumBytes) throw new Error("invalid_continuation_store");
  }

  get size(): number { this.prune(); return this.#records.size; }

  prune(): void {
    const now = this.now();
    for (const [handle, record] of this.#records) if (record.expiresAt <= now) this.#delete(handle);
  }

  #delete(handle: string): void {
    const record = this.#records.get(handle);
    if (!record) return;
    this.#storedBytes -= record.storageBytes;
    this.#records.delete(handle);
  }

  #retain(handle: string, value: Omit<ContinuationRecord, "storageBytes">): void {
    // Count an upper-bound representation including the largest possible
    // decimal storageBytes field, so the aggregate cap is never understated.
    let storageBytes: number;
    try { storageBytes = Buffer.byteLength(JSON.stringify({ ...value, storageBytes: Number.MAX_SAFE_INTEGER }), "utf8"); }
    catch { throw new Error("invalid_continuation"); }
    if (storageBytes > this.maximumBytes) throw new Error("invalid_continuation");
    while (this.#records.size >= this.maximum || this.#storedBytes + storageBytes > this.maximumBytes) {
      const oldest = this.#records.keys().next().value;
      if (oldest === undefined) throw new Error("invalid_continuation");
      this.#delete(oldest);
    }
    this.#records.set(handle, { ...value, storageBytes });
    this.#storedBytes += storageBytes;
  }

  issue(binding: ContinuationBinding, providerNextLink: unknown, expectedPath: string, state?: { resultCount?: number }): string {
    this.prune();
    const canonical = canonicalProviderContinuation(providerNextLink, expectedPath, "invalid_provider_response");
    if (state?.resultCount !== undefined && (!Number.isSafeInteger(state.resultCount) || state.resultCount < 0)) throw new Error("invalid_continuation");
    let handle: string;
    do { handle = `mgc1_${randomBytes(32).toString("base64url")}`; } while (this.#records.has(handle));
    const createdAt = this.now();
    this.#retain(handle, { ...binding, providerNextLink: canonical, expectedPath, ...(state ? { state: { ...state } } : {}), createdAt, expiresAt: createdAt + this.ttlMs });
    return handle;
  }

  issueState(binding: ContinuationBinding, expectedPath: string, state: ContinuationState): string {
    this.prune();
    let encoded: string;
    try { encoded = JSON.stringify(state); } catch { throw new Error("invalid_continuation"); }
    if (!encoded || Buffer.byteLength(encoded, "utf8") > this.maximumStateBytes) throw new Error("invalid_continuation");
    let handle: string;
    do { handle = `mgc1_${randomBytes(32).toString("base64url")}`; } while (this.#records.has(handle));
    const createdAt = this.now();
    this.#retain(handle, { ...binding, expectedPath, state: structuredClone(state), createdAt, expiresAt: createdAt + this.ttlMs });
    return handle;
  }

  verify(handle: unknown, binding: ContinuationBinding): VerifiedContinuation {
    this.prune();
    if (typeof handle !== "string" || !/^mgc1_[A-Za-z0-9_-]{43}$/.test(handle)) throw new Error("invalid_continuation");
    const record = this.#records.get(handle);
    if (!record || record.expiresAt <= this.now()) {
      if (record) this.#delete(handle);
      throw new Error("invalid_continuation");
    }
    for (const key of ["agentId", "service", "action", "resource", "criteria"] as const) {
      if (record[key] !== binding[key]) throw new Error("invalid_continuation");
    }
    return record;
  }

  providerPath(record: VerifiedContinuation, expectedPath: string): string {
    if (record.expectedPath !== expectedPath) throw new Error("invalid_continuation");
    if (record.providerNextLink === undefined) throw new Error("invalid_continuation");
    return canonicalProviderContinuation(record.providerNextLink, expectedPath, "invalid_provider_response");
  }

  continuationState(record: VerifiedContinuation, expectedPath: string): ContinuationState {
    if (record.expectedPath !== expectedPath || record.providerNextLink !== undefined || record.state === undefined) throw new Error("invalid_continuation");
    return structuredClone(record.state);
  }

  resolve(handle: unknown, binding: ContinuationBinding, expectedPath: string): string {
    return this.providerPath(this.verify(handle, binding), expectedPath);
  }
}
