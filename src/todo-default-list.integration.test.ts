import { readFile, readdir } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
const mocks = vi.hoisted(() => ({ tokenForAuthorizedOperation: vi.fn() }));
vi.mock("./credential.js", async (importOriginal) => ({
  ...await importOriginal<typeof import("./credential.js")>(), tokenForAuthorizedOperation: mocks.tokenForAuthorizedOperation,
}));
import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

let nextCall = 0;
const defaultList = (id = "actual-default") => ({ id, displayName: "Aufgaben", wellknownListName: "defaultList", isOwner: true, isShared: false });
const input = { action: "create_task", listSelector: "default", title: "Synthetic task" };
function runtime(options: { agentId?: string; operations?: string[]; enabled?: boolean; warning?: boolean; override?: boolean } = {}) {
  const policy = graphPolicyFixture();
  if (options.operations) policy.services.todo.agents.main.operations = options.operations;
  if (options.override !== undefined) policy.rules.warningApprovalsByService = { todo: options.override };
  const hooks: Record<string, any> = {};
  const factories: any[] = [];
  entry.register({
    pluginConfig: { enabled: options.enabled ?? true, policy, warningApprovalsRequired: options.warning ?? true },
    registerTool: (factory: any) => factories.push(factory),
    on: (name: string, hook: any) => { hooks[name] = hook; },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  const ctx = { agentId: options.agentId ?? "main", sessionId: "synthetic-default-list-session" };
  const tool = factories.map(factory => factory(ctx)).find(tool => tool.name === "microsoft_todo_write");
  const event = (params = input) => ({ toolName: tool.name, toolCallId: `default-list-${++nextCall}`, params });
  return { tool, ctx, event, hook: hooks.before_tool_call };
}
function provider(pages: any[]) {
  let changed = false;
  const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (url, init) => {
    const path = String(url).replace("https://graph.microsoft.com/v1.0", "");
    if (path.startsWith("/me/todo/lists?")) return Response.json(changed ? { value: [defaultList("changed-default")] } : pages.shift());
    if (init?.method === "POST") return Response.json({ id: "new-task", title: "Synthetic task" });
    return Response.json({ isOwner: true, isShared: false });
  });
  return { fetchSpy, changeDefault: () => { changed = true; } };
}
beforeEach(() => mocks.tokenForAuthorizedOperation.mockReset().mockResolvedValue("synthetic-token"));
afterEach(() => vi.restoreAllMocks());

describe("existing To Do write default selector", () => {
  it("keeps known-ID writes discovery-free and requires no separate read grant", async () => {
    const { hook, tool, ctx, event } = runtime({ operations: ["create"] });
    const { fetchSpy } = provider([]);
    const call = event({ action: "create_task", listId: "known-list", title: "Synthetic task" } as any);
    const gate = await hook(call, ctx);
    expect(fetchSpy).not.toHaveBeenCalled();
    await gate.requireApproval.onResolution("allow-once");
    expect((await tool.execute(call.toolCallId, { ...call.params, ...gate.params })).details.ok).toBe(true);
    expect(fetchSpy.mock.calls.map(call => String(call[0]))).toEqual([
      "https://graph.microsoft.com/v1.0/me/todo/lists/known-list",
      "https://graph.microsoft.com/v1.0/me/todo/lists/known-list/tasks",
    ]);
  });
  it("resolves the later-page actual default before approval, preserves opaque query, and executes the same target despite a changed default", async () => {
    const { hook, tool, ctx, event } = runtime();
    const continuation = "https://graph.microsoft.com/v1.0/me/todo/lists?$skiptoken=a%2Bb%2F%3D&$top=50";
    const { fetchSpy, changeDefault } = provider([
      { value: [{ ...defaultList("arbitrary"), wellknownListName: "none", displayName: "Tasks" }], "@odata.nextLink": continuation },
      { value: [defaultList()] },
    ]);
    const call = event();
    const gate = await hook(call, ctx);
    expect(gate.params).toMatchObject({ listId: "actual-default", listSelector: undefined });
    expect(gate.requireApproval.description).toContain('list "actual-default"');
    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(String(fetchSpy.mock.calls[1][0])).toBe(continuation);
    expect(fetchSpy.mock.calls.every(call => !call[1]?.method || call[1].method === "GET")).toBe(true);
    expect(mocks.tokenForAuthorizedOperation.mock.calls[0][0].allowedScopes).toEqual(["Tasks.Read"]);
    changeDefault();
    await gate.requireApproval.onResolution("allow-once");
    const merged = { ...call.params, ...gate.params };
    expect((await tool.execute(call.toolCallId, merged)).details.ok).toBe(true);
    expect(fetchSpy).toHaveBeenCalledTimes(4); // discovery x2, existing owner GET, POST; no new readback
    expect(String(fetchSpy.mock.calls[3][0])).toContain("/actual-default/tasks");
  });
  it.each([
    { pages: [{ value: [{ ...defaultList(), wellknownListName: "none" }] }], error: "invalid_default_todo_list_missing" },
    { pages: [{ value: [defaultList("a")], "@odata.nextLink": "/me/todo/lists?$skiptoken=next" }, { value: [defaultList("b")] }], error: "invalid_default_todo_list_ambiguous" },
    { pages: [{ value: [{ ...defaultList(), isShared: true }] }], error: "access_denied" },
    { pages: [{ value: [{ ...defaultList(), isOwner: false }] }], error: "access_denied" },
    { pages: [{ value: [{ ...defaultList(), isShared: undefined }] }], error: "access_denied" },
    { pages: [{ value: [defaultList()], "@odata.nextLink": "https://evil.invalid/me/todo/lists" }], error: "invalid_provider_response" },
    { pages: [{ value: [defaultList()], "@odata.nextLink": "/me/todo/lists/other/tasks?$skiptoken=x" }], error: "invalid_provider_response" },
    { pages: [{ value: [defaultList()], "@odata.nextLink": "/me/todo/lists?$top=50" }], error: "invalid_default_todo_list_incomplete" },
    { pages: [{}], error: "invalid_provider_response" },
    { pages: [{ value: Array.from({ length: 51 }, (_, n) => defaultList(String(n))) }], error: "invalid_provider_response" },
    { pages: [{ value: [defaultList()], "@odata.nextLink": "/me/todo/lists?$skiptoken=next" }, { value: [defaultList()] }], error: "invalid_default_todo_list_incomplete" },
    { pages: Array.from({ length: 20 }, (_, n) => ({ value: [], "@odata.nextLink": `/me/todo/lists?$skiptoken=${n}` })), error: "invalid_default_todo_list_incomplete" },
    { pages: Array.from({ length: 11 }, (_, page) => ({ value: Array.from({ length: 50 }, (_, n) => ({ ...defaultList(`id-${page}-${n}`), wellknownListName: page === 0 && n === 0 ? "defaultList" : "none" })), "@odata.nextLink": `/me/todo/lists?$skiptoken=${page}` })), error: "invalid_default_todo_list_incomplete" },
  ])("fails closed without approval or mutation: $error", async ({ pages, error }) => {
    const { hook, ctx, event } = runtime();
    const { fetchSpy } = provider(pages);
    expect(await hook(event(), ctx)).toEqual({ block: true, blockReason: error });
    expect(fetchSpy.mock.calls.every(call => !call[1]?.method || call[1].method === "GET")).toBe(true);
  });
  it.each([
    { options: { operations: ["create"] }, params: input, error: "access_denied" },
    { options: { operations: ["read"] }, params: input, error: "access_denied" },
    { options: { agentId: "untrusted" }, params: input, error: "access_denied" },
    { options: { agentId: "" }, params: input, error: "trusted_agent_identity_required" },
    { options: { enabled: false }, params: input, error: "connector_disabled" },
    { options: {}, params: { ...input, title: undefined }, error: "invalid_title" },
    { options: {}, params: { ...input, title: 123 }, error: "invalid_tool_parameters" },
    { options: {}, params: { ...input, dueDateTime: "2026-02-30T12:00:00" }, error: "invalid_datetime" },
    { options: {}, params: { ...input, listId: "explicit" }, error: "invalid_todo_list_target" },
    { options: {}, params: { ...input, listSelector: "any" }, error: "invalid_todo_list_target" },
    { options: {}, params: { action: "delete_task", listId: "list", taskId: "task", listSelector: "default" }, error: "invalid_write_parameter" },
    { options: {}, params: { ...input, listSelector: undefined }, error: "invalid_resource_id" },
  ])("rejects $error before credentials or provider access", async ({ options, params, error }) => {
    const { hook, ctx, event } = runtime(options);
    const { fetchSpy } = provider([]);
    expect(await hook(event(params as any), ctx)).toEqual({ block: true, blockReason: error });
    expect(mocks.tokenForAuthorizedOperation).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
  });
  it("displays full long default-list IDs with distinguishable suffixes, preserving exact binding and private task text", async () => {
    const { hook, tool, ctx, event } = runtime();
    const firstId = "A".repeat(100) + "+first=";
    const secondId = "A".repeat(100) + "+second=";
    const { fetchSpy } = provider([{ value: [defaultList(firstId)] }, { value: [defaultList(secondId)] }]);
    const first = event();
    const firstGate = await hook(first, ctx);
    const second = event();
    const secondGate = await hook(second, ctx);
    expect(firstGate.requireApproval.description).toContain(`list "${firstId}"`);
    expect(secondGate.requireApproval.description).toContain(`list "${secondId}"`);
    expect(firstGate.requireApproval.description).not.toContain(secondId);
    expect(firstGate.requireApproval.description).not.toContain(input.title);
    expect(firstGate.params.listId).toBe(firstId);
    await firstGate.requireApproval.onResolution("allow-once");
    expect((await tool.execute(first.toolCallId, { ...first.params, ...firstGate.params })).details.ok).toBe(true);
    expect(String(fetchSpy.mock.calls[3][0])).toContain(`/${encodeURIComponent(firstId)}/tasks`);
    await secondGate.requireApproval.onResolution("deny");
  });
  it("uses the actual installed host runner to clear the semantic selector and bind the exact target despite a prior rewrite", async () => {
    // Same private-host discovery convention as hook-composition.integration.test.ts.
    const dist = process.env.OPENCLAW_TEST_HOST_DIST ?? dirname(fileURLToPath(import.meta.resolve("openclaw")));
    const files = (await readdir(dist)).filter(name => /^hooks-[A-Za-z0-9_-]+\.mjs$/.test(name)).sort();
    let createHookRunner: any;
    for (const name of files) {
      if ((await readFile(join(dist, name), "utf8")).includes("createHookRunner as t")) {
        createHookRunner = (await import(pathToFileURL(join(dist, name)).href)).t;
        break;
      }
    }
    expect(createHookRunner).toBeTypeOf("function");
    const { hook, tool, ctx, event } = runtime();
    const { fetchSpy } = provider([{ value: [defaultList()] }]);
    const call = event();
    const runner = createHookRunner({ hooks: [], plugins: [], trustedToolPolicies: [], typedHooks: [
      { pluginId: "synthetic-prior-hook", hookName: "before_tool_call", priority: 0, source: "fixture",
        handler: () => ({ params: { action: "delete_task", listId: "wrong-list", taskId: "wrong-task" } }) },
      { pluginId: "microsoft-graph", hookName: "before_tool_call", priority: Number.MIN_SAFE_INTEGER, source: "candidate", handler: hook },
    ] });
    const gate = await runner.runBeforeToolCall(call, ctx);
    expect(gate.params).toMatchObject({ action: "create_task", listSelector: undefined, listId: "actual-default" });
    expect(gate.requireApproval.severity).toBe("warning");
    await gate.requireApproval.onResolution("allow-once");
    expect((await tool.execute(call.toolCallId, { ...call.params, ...gate.params })).details.ok).toBe(true);
    expect(fetchSpy).toHaveBeenCalledTimes(3);
  });
  it("preserves the existing owner GET as a mutation gate", async () => {
    const { hook, tool, ctx, event } = runtime();
    const { fetchSpy } = provider([{ value: [defaultList()] }]);
    const call = event();
    const gate = await hook(call, ctx);
    await gate.requireApproval.onResolution("allow-once");
    fetchSpy.mockResolvedValueOnce(Response.json({ isOwner: false, isShared: false }));
    expect((await tool.execute(call.toolCallId, { ...call.params, ...gate.params })).details.error).toBe("access_denied");
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });
  it("denial never supplies an executable snapshot or writes a task", async () => {
    const { hook, tool, ctx, event } = runtime();
    const { fetchSpy } = provider([{ value: [defaultList()] }]);
    const call = event();
    const gate = await hook(call, ctx);
    await gate.requireApproval.onResolution("deny");
    expect((await tool.execute(call.toolCallId, { ...call.params, ...gate.params })).details.error).toBe("approval_context_invalid_or_changed");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });
  it("normalizes new calls under existing allow-always trust without sharing a target snapshot", async () => {
    const { hook, tool, ctx, event } = runtime();
    const { fetchSpy } = provider([{ value: [defaultList("first")] }, { value: [defaultList("second")] }]);
    const first = event();
    const gate = await hook(first, ctx);
    await gate.requireApproval.onResolution("allow-always");
    const second = event();
    const trusted = await hook(second, ctx);
    expect(trusted.requireApproval).toBeUndefined();
    expect(trusted.params.listId).toBe("second");
    expect((await tool.execute(second.toolCallId, { ...second.params, ...trusted.params })).details.ok).toBe(true);
    expect((await tool.execute(first.toolCallId, { ...first.params, ...gate.params })).details.ok).toBe(true);
    expect(fetchSpy).toHaveBeenCalledTimes(6);
  });
  it("rejects execution target rewrites before credentials and cannot execute the unresolved selector", async () => {
    const { hook, tool, ctx, event } = runtime();
    const { fetchSpy } = provider([{ value: [defaultList()] }]);
    const call = event();
    const gate = await hook(call, ctx);
    await gate.requireApproval.onResolution("allow-once");
    mocks.tokenForAuthorizedOperation.mockClear();
    fetchSpy.mockClear();
    const result = await tool.execute(call.toolCallId, { ...call.params, ...gate.params, listId: "changed-list" });
    expect(result.details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
    expect((await tool.execute("unresolved", input)).details.error).toBe("approval_context_invalid_or_changed");
    expect(mocks.tokenForAuthorizedOperation).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
  });
  it.each([
    { warning: false, override: undefined, approval: false },
    { warning: true, override: false, approval: false },
    { warning: false, override: true, approval: true },
  ])("uses the same normalized target with warning policy $warning/$override", async ({ warning, override, approval }) => {
    const { hook, tool, ctx, event } = runtime({ warning, override });
    const { fetchSpy } = provider([{ value: [defaultList()] }]);
    const call = event();
    const gate = await hook(call, ctx);
    expect(!!gate.requireApproval).toBe(approval);
    await gate.requireApproval?.onResolution("allow-once");
    expect((await tool.execute(call.toolCallId, { ...call.params, ...gate.params })).details.ok).toBe(true);
    expect(fetchSpy).toHaveBeenCalledTimes(3);
  });
});
