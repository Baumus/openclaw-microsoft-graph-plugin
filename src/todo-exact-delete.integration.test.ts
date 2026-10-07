import { describe, expect, it, vi } from "vitest";

const credentials = vi.hoisted(() => ({ token: vi.fn(async () => "synthetic-token") }));
vi.mock("./credential.js", async () => {
  const actual = await vi.importActual<typeof import("./credential.js")>("./credential.js");
  return { ...actual, tokenForAuthorizedOperation: credentials.token };
});

import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function setup(policy = graphPolicyFixture()) {
  const factories: Array<(context: any) => any> = [];
  const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
  entry.register({
    pluginConfig: { enabled: true, policy },
    registerTool: (factory: any) => factories.push(factory),
    on: (name: string, handler: any) => { hooks[name] = handler; },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  const context = { agentId: "main", sessionId: "todo-exact-test" };
  const tool = factories.map((factory) => factory(context)).find((candidate) => candidate.name === "microsoft_todo_write");
  const before = (id: string, params: Record<string, unknown>) => hooks.before_tool_call({ toolName: tool.name, toolCallId: id, params }, context);
  return { tool, before };
}

function graphFixture(options: { secondPage?: boolean; secondList?: boolean; malformed?: boolean; cycle?: boolean; noMatch?: boolean; currentTitle?: string; deletedBeforeExecute?: boolean; versionConflict?: boolean; missingVersion?: boolean } = {}) {
  const calls: string[] = [];
  const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input, init) => {
    const url = new URL(String(input));
    const path = url.pathname.replace(/^\/v1\.0/, "");
    calls.push(`${init?.method ?? "GET"} ${path}`);
    const json = (value: unknown) => new Response(JSON.stringify(value), { status: 200 });
    if (init?.method === "DELETE") return options.versionConflict ? new Response(JSON.stringify({ error: { code: "PreconditionFailed" } }), { status: 412 }) : new Response(null, { status: 204 });
    if (path === "/me/todo/lists") return json(options.malformed ? {} : {
      value: [{ id: "list-1", isOwner: true, isShared: false }, ...(options.secondList ? [{ id: "list-2", isOwner: true, isShared: false }] : [])],
    });
    if (path === "/me/todo/lists/list-1/tasks") return json(url.searchParams.has("$skiptoken")
      ? { value: options.cycle ? [] : [{ id: "task-2", title: "Exact" }], ...(options.cycle ? { "@odata.nextLink": `https://graph.microsoft.com/v1.0${path}?$skiptoken=page2` } : {}) }
      : { value: [{ id: "task-1", title: options.noMatch ? "Different" : "Exact" }], ...(options.secondPage || options.cycle ? { "@odata.nextLink": `https://graph.microsoft.com/v1.0${path}?$skiptoken=page2` } : {}) });
    if (path === "/me/todo/lists/list-2/tasks") return json({ value: [{ id: "task-3", title: "Exact" }] });
    if (path === "/me/todo/lists/list-1") return json({ id: "list-1", isOwner: true, isShared: false });
    if (path === "/me/todo/lists/list-1/tasks/task-1") return options.deletedBeforeExecute ? new Response(JSON.stringify({ error: { code: "ErrorItemNotFound" } }), { status: 404 }) : json({ id: "task-1", title: options.currentTitle ?? "Exact", ...(!options.missingVersion ? { "@odata.etag": 'W/"version-1"' } : {}) });
    throw new Error(`unexpected path ${path}`);
  });
  return { calls, fetchSpy };
}

describe("To Do exact-title delete", () => {
  it.each([
    ["duplicate on later page", { secondPage: true }, "exact_search_ambiguous"],
    ["duplicate in another list", { secondList: true }, "exact_search_ambiguous"],
    ["malformed collection", { malformed: true }, "invalid_provider_response"],
    ["cyclic continuation", { cycle: true }, "exact_search_incomplete"],
    ["zero matches", { noMatch: true }, "exact_search_no_match"],
  ] as const)("blocks %s before approval and DELETE", async (_name, options, error) => {
    const { before } = setup();
    const { calls, fetchSpy } = graphFixture(options);
    try {
      const gate = await before(`ambiguous-${_name}`, { action: "delete_task_exact", title: "Exact" });
      expect(gate).toEqual({ block: true, blockReason: error });
      expect(calls.filter((call) => call.startsWith("DELETE"))).toHaveLength(0);
    } finally { fetchSpy.mockRestore(); }
  });

  it("binds resolved IDs to native approval and rejects altered execution params", async () => {
    const { before, tool } = setup();
    const { calls, fetchSpy } = graphFixture();
    try {
      const gate = await before("bind-ids", { action: "delete_task_exact", title: "Exact" });
      expect(gate.params).toMatchObject({ listId: "list-1", taskId: "task-1" });
      expect(gate.requireApproval).toMatchObject({ severity: "critical", allowedDecisions: ["allow-once", "deny"] });
      expect(gate.requireApproval.description).toContain('list ID "list-1", task ID "task-1"');
      await gate.requireApproval.onResolution("allow-once");
      expect((await tool.execute("bind-ids", { ...gate.params, taskId: "task-2" })).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
      expect(calls.filter((call) => call.startsWith("DELETE"))).toHaveLength(0);
    } finally { fetchSpy.mockRestore(); }
  });

  it("rejects name-only direct execution and denied native approval", async () => {
    const { before, tool } = setup();
    const fixture = graphFixture();
    try {
      expect((await tool.execute("direct", { action: "delete_task_exact", title: "Exact" })).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
      const gate = await before("deny", { action: "delete_task_exact", title: "Exact" });
      await gate.requireApproval.onResolution("deny");
      expect((await tool.execute("deny", gate.params)).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
      expect(fixture.calls.some((call) => call.startsWith("DELETE"))).toBe(false);
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it("rechecks stale title after approval and does not DELETE", async () => {
    const { before, tool } = setup();
    const fixture = graphFixture({ currentTitle: "Changed" });
    try {
      const gate = await before("stale", { action: "delete_task_exact", title: "Exact" });
      await gate.requireApproval.onResolution("allow-once");
      expect((await tool.execute("stale", gate.params)).details).toMatchObject({ ok: false, error: "exact_target_changed", mutationApplied: false, retrySafety: "safe_after_correction" });
      expect(fixture.calls.filter((call) => call.startsWith("DELETE"))).toHaveLength(0);
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it("does not DELETE when the approved task disappears before execution", async () => {
    const { before, tool } = setup();
    const fixture = graphFixture({ deletedBeforeExecute: true });
    try {
      const gate = await before("deleted", { action: "delete_task_exact", title: "Exact" });
      await gate.requireApproval.onResolution("allow-once");
      expect((await tool.execute("deleted", gate.params)).details).toMatchObject({ ok: false });
      expect(fixture.calls.filter((call) => call.startsWith("DELETE"))).toHaveLength(0);
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it("deletes only the revalidated approved task and returns an identity receipt", async () => {
    const { before, tool } = setup();
    const fixture = graphFixture();
    try {
      const gate = await before("success", { action: "delete_task_exact", title: "Exact" });
      await gate.requireApproval.onResolution("allow-once");
      expect((await tool.execute("success", gate.params)).details).toMatchObject({ ok: true, deleted: true, listId: "list-1", taskId: "task-1", title: "Exact" });
      expect(fixture.calls.filter((call) => call.startsWith("DELETE"))).toEqual(["DELETE /me/todo/lists/list-1/tasks/task-1"]);
      expect(fixture.fetchSpy).toHaveBeenCalledWith(expect.anything(), expect.objectContaining({ method: "DELETE", headers: expect.objectContaining({ "If-Match": 'W/"version-1"' }) }));
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it.each([
    ["stale version", { versionConflict: true }, "exact_target_changed", 1],
    ["missing version", { missingVersion: true }, "exact_target_version_unavailable", 0],
  ] as const)("fails closed for %s", async (_name, options, error, deletes) => {
    const { before, tool } = setup();
    const fixture = graphFixture(options);
    try {
      const gate = await before("version-" + _name, { action: "delete_task_exact", title: "Exact" });
      await gate.requireApproval.onResolution("allow-once");
      expect((await tool.execute("version-" + _name, gate.params)).details).toMatchObject({ ok: false, error });
      expect(fixture.calls.filter((call) => call.startsWith("DELETE"))).toHaveLength(deletes);
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it("denied read scope blocks before credentials or network", async () => {
    const policy = graphPolicyFixture();
    policy.services.todo.agents.main.operations = ["delete"];
    const { before } = setup(policy);
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    credentials.token.mockClear();
    try {
      expect(await before("denied", { action: "delete_task_exact", title: "Exact" })).toEqual({ block: true, blockReason: "access_denied" });
      expect(credentials.token).not.toHaveBeenCalled();
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally { fetchSpy.mockRestore(); }
  });
});
