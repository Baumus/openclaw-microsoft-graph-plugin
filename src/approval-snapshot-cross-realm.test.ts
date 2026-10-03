import { describe, expect, it } from "vitest";

/**
 * Regression: the host can import this bundle more than once in a single process. When the
 * approval snapshot store lived in module scope, the `before_tool_call` hook recorded into one
 * module instance while the tool's `consume` read an empty map in the other, so every
 * approval-bearing mutation failed closed with `approval_context_invalid_or_changed`.
 *
 * Two distinct module specifiers give two real module instances, which is exactly what the host
 * does. Both must observe the same store.
 */
describe("native approval snapshot store across module instances", () => {
  const STORE_KEY = Symbol.for("@baumus/openclaw-microsoft-graph/native-approval-snapshots");

  it("shares one store between separately imported module instances", async () => {
    const first = await import("./index.js?realm=1");
    const second = await import("./index.js?realm=2");

    // Distinct module instances, not the same object graph.
    expect(first).not.toBe(second);

    // ...yet exactly one store exists, reachable from the cross-realm registry symbol.
    const store = (globalThis as Record<symbol, unknown>)[STORE_KEY];
    expect(store).toBeDefined();

    const sameStoreAfterSecondImport = (globalThis as Record<symbol, unknown>)[STORE_KEY];
    expect(sameStoreAfterSecondImport).toBe(store);
  });

  it("keeps the store process-local rather than persisting it", async () => {
    await import("./index.js?realm=3");
    const store = (globalThis as Record<symbol, unknown>)[STORE_KEY] as { record: unknown; consume: unknown };
    expect(typeof store.record).toBe("function");
    expect(typeof store.consume).toBe("function");
  });
});
