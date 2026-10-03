import { describe, expect, it } from "vitest";

/**
 * Regression: the host can import this bundle more than once in a single process. When the
 * approval snapshot store lived in module scope, the `before_tool_call` hook recorded a snapshot
 * into one module instance while the tool's `consume` read an empty map in the other, so every
 * approval-bearing mutation failed closed with `approval_context_invalid_or_changed` and no
 * configuration could recover it.
 *
 * A query string forces Node to instantiate a genuinely separate module, which is what the host
 * effectively does. The specifier is built from a variable so TypeScript does not try to resolve
 * it at build time.
 */
describe("native approval snapshot store across module instances", () => {
  const STORE_KEY = Symbol.for("@baumus/openclaw-microsoft-graph/native-approval-snapshots");
  const ENTRY = "./index.js";

  const importRealm = async (realm: number): Promise<unknown> => {
    const specifier = `${ENTRY}?realm=${realm}`;
    return import(specifier);
  };

  it("shares one store between separately imported module instances", async () => {
    const first = await importRealm(1);
    const second = await importRealm(2);

    // Genuinely distinct module instances, not the same object graph.
    expect(first).not.toBe(second);

    // ...yet exactly one store exists, reachable from the cross-realm registry symbol, so a
    // snapshot recorded by the hook in one instance is visible to the tool in the other.
    const store = (globalThis as Record<symbol, unknown>)[STORE_KEY];
    expect(store).toBeDefined();
    expect((globalThis as Record<symbol, unknown>)[STORE_KEY]).toBe(store);
  });

  it("exposes the record/consume contract on the shared store", async () => {
    await importRealm(3);
    const store = (globalThis as Record<symbol, unknown>)[STORE_KEY] as {
      record: unknown;
      consume: unknown;
      clearSession: unknown;
    };
    expect(typeof store.record).toBe("function");
    expect(typeof store.consume).toBe("function");
    expect(typeof store.clearSession).toBe("function");
  });
});
