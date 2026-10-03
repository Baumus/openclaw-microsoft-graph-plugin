import { describe, expect, it, vi } from "vitest";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

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
  }, 20_000);

  it("passes a resolved hook approval to a tool from another module instance exactly once", async () => {
    const hookModule = await importRealm(4) as { default: { register: (api: unknown) => void } };
    const toolModule = await importRealm(5) as { default: { register: (api: unknown) => void } };
    const hooks: Record<string, (event: unknown, context: unknown) => Promise<any>> = {};
    const factories: Array<(context: unknown) => any> = [];
    const config = { enabled: true, policy: graphPolicyFixture() };
    const api = {
      pluginConfig: config,
      on: (name: string, handler: (event: unknown, context: unknown) => Promise<any>) => { hooks[name] = handler; },
      registerTool: (factory: (context: unknown) => any) => factories.push(factory),
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    };
    hookModule.default.register({ ...api, registerTool: vi.fn() });
    toolModule.default.register({ ...api, on: vi.fn() });
    const context = { agentId: "main", sessionId: "cross-module-approval" };
    const tool = factories.map((factory) => factory(context)).find((candidate) => candidate.name === "outlook_mail_write");
    expect(tool).toBeDefined();
    const params = { action: "mark_read", messageId: "message-1", isRead: true };
    const callId = "cross-module-approval-call";
    const approval = await hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: callId, params }, context);
    expect(approval.requireApproval).toMatchObject({ severity: "warning" });
    approval.requireApproval.onResolution("allow-once");

    // A different module's tool consumes the hook's snapshot, then reaches the credential gate.
    // No Microsoft request is possible without the intentionally absent test vault key.
    expect((await tool.execute(callId, params)).details).toMatchObject({ ok: false, error: "credential_vault_unavailable" });
    expect((await tool.execute(callId, params)).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
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
