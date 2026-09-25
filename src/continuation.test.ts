import { describe, expect, it } from "vitest";
import { ContinuationStore, MAX_CONTINUATION_STATE_BYTES, normalizedCriteria, type ContinuationBinding } from "./continuation.js";

const cases: Array<{ binding: ContinuationBinding; path: string }> = [
  { binding: { agentId: "main", service: "calendar", action: "list_events", resource: "me", criteria: normalizedCriteria({ action: "list_events", startDateTime: "2026-09-01T00:00:00Z", endDateTime: "2026-10-01T00:00:00Z", limit: 25 }) }, path: "/me/calendarView" },
  { binding: { agentId: "main", service: "mail", action: "list_messages", resource: "me", criteria: normalizedCriteria({ action: "list_messages", folder: "inbox", limit: 25 }) }, path: "/me/mailFolders/inbox/messages" },
  { binding: { agentId: "fixture-reader", service: "onedrive", action: "list", resource: "synthetic_documents", criteria: normalizedCriteria({ rootLabel: "synthetic_documents", relativePath: "Reports", limit: 25 }) }, path: "/drives/synthetic-drive/items/synthetic-root:/Reports:/children" },
  { binding: { agentId: "main", service: "todo", action: "list_tasks", resource: "me", criteria: normalizedCriteria({ action: "list_tasks", listId: "list", limit: 25 }) }, path: "/me/todo/lists/list/tasks" },
];

describe("opaque continuation store", () => {
  it.each(cases)("issues reusable opaque handles for $binding.service without exposing provider URLs", ({ binding, path }) => {
    const store = new ContinuationStore();
    const handle = store.issue(binding, `https://graph.microsoft.com/v1.0${path}?$skiptoken=provider-secret`, path);
    expect(handle).toMatch(/^mgc1_[A-Za-z0-9_-]{43}$/);
    expect(handle).not.toContain("graph.microsoft.com");
    expect(store.resolve(handle, binding, path)).toBe(`${path}?$skiptoken=provider-secret`);
    expect(store.resolve(handle, binding, path)).toBe(`${path}?$skiptoken=provider-secret`);
  });

  it("rejects the reproduced modified calendarView date-range attack", () => {
    const store = new ContinuationStore();
    const original = cases[0];
    const handle = store.issue(original.binding, `${original.path}?$skiptoken=next`, original.path);
    const modified = { ...original.binding, criteria: normalizedCriteria({ action: "list_events", startDateTime: "1900-01-01T00:00:00Z", endDateTime: "2100-01-01T00:00:00Z", limit: 25 }) };
    expect(() => store.resolve(handle, modified, original.path)).toThrow("invalid_continuation");
  });

  it("fails closed for every binding mismatch, fabricated values, raw URLs, expiry, and provider path drift", () => {
    let now = 1_000;
    const store = new ContinuationStore(100, 8, () => now);
    const original = cases[0];
    const handle = store.issue(original.binding, `${original.path}?$skiptoken=next`, original.path);
    const mismatches: ContinuationBinding[] = [
      { ...original.binding, agentId: "other" },
      { ...original.binding, service: "mail" },
      { ...original.binding, action: "list_calendars" },
      { ...original.binding, resource: "other-calendar" },
      { ...original.binding, criteria: normalizedCriteria({ changed: true }) },
    ];
    for (const binding of mismatches) expect(() => store.resolve(handle, binding, original.path)).toThrow("invalid_continuation");
    for (const invalid of ["mgc1_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", `${original.path}?$skiptoken=raw`, `https://graph.microsoft.com/v1.0${original.path}?$skiptoken=raw`]) {
      expect(() => store.resolve(invalid, original.binding, original.path)).toThrow("invalid_continuation");
    }
    expect(() => store.resolve(handle, original.binding, "/me/events")).toThrow("invalid_continuation");
    now += 101;
    expect(() => store.resolve(handle, original.binding, original.path)).toThrow("invalid_continuation");
    expect(store.size).toBe(0);
  });

  it("validates provider nextLinks on issue and bounds retained state", () => {
    const store = new ContinuationStore(1_000, 2, () => 1_000);
    const original = cases[3];
    expect(() => store.issue(original.binding, "https://evil.invalid/v1.0/me/todo/lists/list/tasks?$skiptoken=x", original.path)).toThrow("invalid_provider_response");
    const handles = ["one", "two", "three"].map((value) => store.issue({ ...original.binding, criteria: value }, `${original.path}?$skiptoken=${value}`, original.path));
    expect(store.size).toBe(2);
    expect(() => store.resolve(handles[0], { ...original.binding, criteria: "one" }, original.path)).toThrow("invalid_continuation");
    expect(store.resolve(handles[2], { ...original.binding, criteria: "three" }, original.path)).toContain("$skiptoken=three");
  });

  it("round-trips bounded internal state without treating it as a provider URL", () => {
    const store = new ContinuationStore();
    const original = cases[2];
    const state = { kind: "exact_fallback", current: { folderId: "root-id" }, queue: ["folder-id"], seenFolderIds: ["root-id"], seenItemIds: [] };
    const handle = store.issueState(original.binding, original.path, state);
    const verified = store.verify(handle, original.binding);
    expect(store.continuationState(verified, original.path)).toEqual(state);
    expect(() => store.providerPath(verified, original.path)).toThrow("invalid_continuation");
    expect(() => store.continuationState(verified, "/drives/other/items/root/search(q='x')")).toThrow("invalid_continuation");
  });

  it("accepts large bounded state and rejects state above the per-record cap", () => {
    const store = new ContinuationStore();
    const original = cases[2];
    const state = { kind: "exact_fallback", blob: "x".repeat(300 * 1024) };
    const handle = store.issueState(original.binding, original.path, state);
    expect(store.continuationState(store.verify(handle, original.binding), original.path)).toEqual(state);
    expect(() => store.issueState(original.binding, original.path, { blob: "x".repeat(MAX_CONTINUATION_STATE_BYTES + 1) })).toThrow("invalid_continuation");
  });

  it("evicts oldest state when the aggregate byte budget is reached", () => {
    const original = cases[2];
    const store = new ContinuationStore(1_000, 10, () => 1_000, 1_800, 1_400);
    const firstBinding = { ...original.binding, criteria: "first" };
    const secondBinding = { ...original.binding, criteria: "second" };
    const first = store.issueState(firstBinding, original.path, { blob: "a".repeat(1_024) });
    const second = store.issueState(secondBinding, original.path, { blob: "b".repeat(1_024) });
    expect(store.size).toBe(1);
    expect(() => store.verify(first, firstBinding)).toThrow("invalid_continuation");
    expect(store.continuationState(store.verify(second, secondBinding), original.path)).toEqual({ blob: "b".repeat(1_024) });
  });
});
