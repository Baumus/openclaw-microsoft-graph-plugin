import { describe, expect, it, vi } from "vitest";
import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function candidate() {
  let storedId: string | undefined = "generation-A";
  let delayNextRead = false;
  let release: (() => void) | undefined;
  let rejectRead = false;
  const lookup = vi.fn(async () => {
    if (delayNextRead) {
      delayNextRead = false;
      await new Promise<void>((resolve) => { release = resolve; });
    }
    if (rejectRead) throw new Error("synthetic lookup unavailable");
    return storedId ? { sessionId: storedId } : undefined;
  });
  const hooks: Record<string, (...args: any[]) => any> = {};
  const factories: Array<(ctx: any) => any> = [];
  entry.register({
    pluginConfig: { enabled: true, policy: graphPolicyFixture() },
    runtime: { agent: { session: { getSessionEntryAsync: lookup, getSessionEntry: () => { throw new Error("deprecated getter"); } } } },
    on: (name: string, handler: (...args: any[]) => any) => { hooks[name] = handler; },
    registerTool: (factory: (ctx: any) => any) => factories.push(factory),
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as never);
  const context = { agentId: "main", sessionKey: "agent:main:main", sessionId: "generation-A" };
  const tool = factories.map((factory) => factory(context)).find((tool) => tool.name === "outlook_calendar_write");
  const params = { action: "delete", eventId: "synthetic-event", calendarId: "synthetic-calendar", timeoutMs: 180000 };
  const fresh = async (id = crypto.randomUUID(), overrides = {}) => {
    const event = { toolName: tool.name, toolCallId: id, params: { ...params, ...overrides } };
    return { event, gate: await hooks.before_tool_call(event, context) };
  };
  return { tool, fresh, end: () => hooks.session_end({}, context), delay: () => { delayNextRead = true; }, release: () => release!(), setId: (id?: string) => { storedId = id; }, fail: () => { rejectRead = true; } };
}

describe("native host fire-and-forget approval resolution", () => {
  it.each(["delete", "update"])("waits for async approval binding before consuming the exact calendar %s once", async (action) => {
    const c = candidate();
    const { event, gate } = await c.fresh(undefined, action === "update" ? { action, subject: "Synthetic update" } : {});
    expect(gate.requireApproval.severity).toBe(action === "delete" ? "critical" : "warning");
    c.delay();
    // Host notifyPluginApprovalResolution intentionally does NOT await this callback.
    const resolution = Promise.resolve(gate.requireApproval.onResolution("allow-once"));
    const execution = c.tool.execute(event.toolCallId, event.params);
    let finished = false;
    void execution.then(() => { finished = true; });
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(finished).toBe(false);
    c.release();
    await resolution;
    // Reaching the intentionally unavailable vault proves approval passed, without Microsoft IO.
    expect((await execution).details).toMatchObject({ ok: false, error: "credential_vault_unavailable" });
    expect((await c.tool.execute(event.toolCallId, event.params)).details.error).toBe("approval_context_invalid_or_changed");
    const retry = await c.fresh();
    void Promise.resolve(retry.gate.requireApproval.onResolution("allow-once"));
    expect((await c.tool.execute(retry.event.toolCallId, retry.event.params)).details.error).toBe("credential_vault_unavailable");
  });

  it.each(["reset", "missing", "rejected"])("fails closed for a %s session read while resolution is pending", async (failure) => {
    const c = candidate();
    const { event, gate } = await c.fresh();
    c.delay();
    const resolution = Promise.resolve(gate.requireApproval.onResolution("allow-once")).catch(() => undefined);
    const execution = c.tool.execute(event.toolCallId, event.params);
    if (failure === "reset") c.setId("generation-B");
    else if (failure === "missing") c.setId(undefined);
    else c.fail();
    c.release();
    await resolution;
    expect((await execution).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed", mutationApplied: false });
  });

  it("rejects changed effective target while resolution is pending", async () => {
    const c = candidate();
    const { event, gate } = await c.fresh();
    c.delay();
    const resolution = Promise.resolve(gate.requireApproval.onResolution("allow-once"));
    const execution = c.tool.execute(event.toolCallId, { ...event.params, eventId: "different-event" });
    c.release();
    await resolution;
    expect((await execution).details.error).toBe("approval_context_invalid_or_changed");
  });

  it.each(["deny", "timeout", "cancelled", "allow-always"])("never authorizes a critical deletion for %s", async (decision) => {
    const c = candidate();
    const { event, gate } = await c.fresh();
    void Promise.resolve(gate.requireApproval.onResolution(decision));
    expect((await c.tool.execute(event.toolCallId, event.params)).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed", mutationApplied: false });
  });

  it("does not wait for or accept approval of another call", async () => {
    const c = candidate();
    const { event, gate } = await c.fresh();
    c.delay();
    const resolution = Promise.resolve(gate.requireApproval.onResolution("allow-once"));
    expect((await c.tool.execute("unapproved-call", event.params)).details.error).toBe("approval_context_invalid_or_changed");
    c.release();
    await resolution;
    expect((await c.tool.execute(event.toolCallId, event.params)).details.error).toBe("credential_vault_unavailable");
  });
  it("consumes a pending approved call at most once under concurrent execution", async () => {
    const c = candidate();
    const { event, gate } = await c.fresh();
    c.delay();
    const resolution = Promise.resolve(gate.requireApproval.onResolution("allow-once"));
    const executions = [c.tool.execute(event.toolCallId, event.params), c.tool.execute(event.toolCallId, event.params)];
    c.release();
    await resolution;
    expect((await Promise.all(executions)).map((result) => result.details.error).sort()).toEqual(["approval_context_invalid_or_changed", "credential_vault_unavailable"]);
  });

  it.each(["session-end", "abort", "already-aborted"])("cancels the barrier for %s without resurrecting approval after the lookup settles", async (reason) => {
    const c = candidate();
    const { event, gate } = await c.fresh();
    c.delay();
    const resolution = Promise.resolve(gate.requireApproval.onResolution("allow-once")).catch(() => undefined);
    const controller = new AbortController();
    if (reason === "already-aborted") controller.abort();
    const execution = c.tool.execute(event.toolCallId, event.params, controller.signal);
    if (reason === "session-end") c.end();
    else controller.abort();
    // Cancellation releases execution even though the store read remains pending.
    expect((await execution).details.error).toBe("approval_context_invalid_or_changed");
    c.release();
    await resolution;
    expect((await c.tool.execute(event.toolCallId, event.params)).details.error).toBe("approval_context_invalid_or_changed");
    const retry = await c.fresh();
    void Promise.resolve(retry.gate.requireApproval.onResolution("allow-once"));
    expect((await c.tool.execute(retry.event.toolCallId, retry.event.params)).details.error).toBe("credential_vault_unavailable");
  });

});
