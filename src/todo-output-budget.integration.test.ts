import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("./credential.js", async () => ({
  ...await vi.importActual<typeof import("./credential.js")>("./credential.js"),
  readCredential: vi.fn(async () => ({ clientId: "synthetic", refreshToken: "synthetic", tenant: "common", scopes: ["Tasks.Read"] })),
  exchangeRefreshToken: vi.fn(async () => "synthetic-token"),
  tokenForAuthorizedOperation: vi.fn(async () => "synthetic-token"),
}));

import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function todoTool(maxReadOutputBytes?: number, agentId = "main") {
  const factories: Array<(context: any) => any> = [];
  entry.register({
    pluginConfig: { enabled: true, policy: graphPolicyFixture(), ...(maxReadOutputBytes ? { maxReadOutputBytes } : {}) },
    registerTool: (factory: any) => factories.push(typeof factory === "function" ? factory : (factory as { create: any }).create), on: vi.fn(),
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  return factories.map((factory) => factory({ agentId })).find((tool) => tool.name === "microsoft_todo_read");
}

function expectBounded(response: any, maximum = 8192) {
  const text = response.content[0].text;
  expect(JSON.parse(text)).toEqual(response.details);
  expect(Buffer.byteLength(JSON.stringify(text), "utf8")).toBeLessThanOrEqual(maximum);
}

afterEach(() => vi.restoreAllMocks());

describe("To Do collection output budgets", () => {
  it("summarizes oversized bodies under a smaller configured ceiling without changing exact-ID detail", async () => {
    const task = { id: "task", title: "Short", status: "inProgress", importance: "high", dueDateTime: { dateTime: "2026-11-25T00:00:00", timeZone: "UTC" }, body: { contentType: "text", content: "x".repeat(20_000) }, categories: ["retained in detail"] };
    vi.spyOn(globalThis, "fetch").mockImplementation(async (input) =>
      Response.json(new URL(String(input)).pathname.endsWith("/tasks/task") ? task : { value: [task] }));
    const tool = todoTool(1024);
    const listed = await tool.execute("list", { action: "list_tasks", listId: "list", limit: 50 });
    expect(listed.details).toMatchObject({
      ok: true, view: "summary", detailAction: "get_task", completeness: "complete", truncated: false,
      items: [{ id: task.id, title: task.title, status: task.status, importance: task.importance, dueDateTime: task.dueDateTime }],
      omittedFields: expect.arrayContaining(["body", "categories"]),
    });
    expect(listed.details.items[0]).not.toHaveProperty("body");
    expectBounded(listed, 1024);
    const detail = await tool.execute("detail", { action: "get_task", listId: "list", taskId: "task" });
    expect(detail.details.item).toEqual(task);
  });

  it.each([undefined, 2048])("pages whole Unicode/escaped summaries without skipping retained rows or provider pages (budget %s)", async (maximum) => {
    const tasks = Array.from({ length: 51 }, (_, index) => ({
      id: `${index}-${"a".repeat(140)}`, title: '\u{1f30d}"\\\n'.repeat(30),
      status: index % 2 ? "completed" : "notStarted", importance: "normal",
      dueDateTime: { dateTime: "2026-11-25T00:00:00", timeZone: "UTC" },
      body: { content: "detail".repeat(100) },
    }));
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = new URL(String(input));
      return Response.json(url.searchParams.has("$skiptoken") ? { value: tasks.slice(50) } : {
        value: tasks.slice(0, 50), "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists/list/tasks?$skiptoken=second",
      });
    });
    const tool = todoTool(maximum);
    const params = { action: "list_tasks", listId: "list", limit: 50 };
    const seen: unknown[] = [];
    let continuation: string | undefined;
    for (let page = 0; page < 60; page += 1) {
      const response = await tool.execute(`page-${page}`, { ...params, ...(continuation ? { continuation } : {}) });
      expect(response.details.ok).toBe(true);
      expectBounded(response, maximum ?? 8192);
      expect(response.details.items.length).toBeGreaterThan(0);
      for (const item of response.details.items) {
        const source = tasks[seen.length];
        expect(item).toEqual({ id: source.id, title: source.title, status: source.status, importance: source.importance, dueDateTime: source.dueDateTime });
        seen.push(item.id);
      }
      continuation = response.details.continuation;
      if (!continuation) {
        expect(response.details).toMatchObject({ completeness: "complete", truncated: false });
        break;
      }
      expect(response.details).toMatchObject({ completeness: "partial", truncated: true });
      expect(continuation).toMatch(/^mgc1_[A-Za-z0-9_-]{43}$/);
      expect(response.content[0].text).not.toContain("$skiptoken");
      if (seen.length < 50) expect(fetchSpy).toHaveBeenCalledTimes(1);
    }
    expect(seen).toEqual(tasks.map((task) => task.id));
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });

  it("matches full bodies/categories before projection and explicitly reports byte-capped searches", async () => {
    const tasks = Array.from({ length: 50 }, (_, index) => ({
      id: `task-${index}`, title: "long summary ".repeat(20), status: "waitingOnOthers", importance: "low",
      body: { content: `hidden needle ${"x".repeat(1000)}` }, categories: ["category match"],
    }));
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async () => Response.json({ value: tasks }));
    const tool = todoTool();
    for (const [search, searchFields] of [["needle", ["body"]], ["category match", ["categories"]]] as const) {
      const response = await tool.execute("search", { action: "search_tasks", listId: "list", search, searchFields, limit: 50 });
      expect(response.details).toMatchObject({ ok: true, outputLimited: true, truncated: true, completeness: "partial", scanned: 50 });
      expect(response.details.items.length).toBeGreaterThan(0);
      expect(response.details.items.length).toBeLessThan(50);
      expect(response.details.items[0]).not.toHaveProperty("body");
      expect(response.details).not.toHaveProperty("continuation");
      expectBounded(response);
    }
    const empty = await tool.execute("empty", { action: "search_tasks", listId: "list", search: "absent", limit: 50 });
    expect(empty.details).toMatchObject({ ok: true, items: [], completeness: "complete", noResults: true });
    expectBounded(empty);
    fetchSpy.mockImplementation(async (input) => {
      const page = Number(new URL(String(input)).searchParams.get("$skiptoken") ?? "0");
      return Response.json({ value: tasks, "@odata.nextLink": `https://graph.microsoft.com/v1.0/me/todo/lists/list/tasks?$skiptoken=${page + 1}` });
    });
    const incomplete = await tool.execute("incomplete", { action: "search_tasks", listId: "list", search: "absent", limit: 50 });
    expect(incomplete.details).toMatchObject({ ok: true, items: [], scanned: 500, completeness: "partial", noResults: false });
    expect(incomplete.details).not.toHaveProperty("continuation");
    expectBounded(incomplete);
  });

  it("fails explicitly if one whole summary cannot fit instead of clipping it or returning false absence", async () => {
    vi.spyOn(globalThis, "fetch").mockImplementation(async () => Response.json({ value: [{ id: "task", title: "\u{1f30d}".repeat(3000) }] }));
    const response = await todoTool().execute("oversize", { action: "list_tasks", listId: "list" });
    expect(response.details).toMatchObject({ ok: false, error: "invalid_collection_item_too_large", phase: "failed" });
    expect(response.details).not.toHaveProperty("items");
    expect(response.details).not.toHaveProperty("noResults");
    expectBounded(response);
  });

  it("binds retained rows to list/limit and rejects unsafe provider links before exposing any page", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async () => Response.json({
      value: Array.from({ length: 50 }, (_, index) => ({ id: `task-${index}`, title: "x".repeat(250) })),
    }));
    const tool = todoTool();
    const first = await tool.execute("first", { action: "list_tasks", listId: "list", limit: 50 });
    expect(first.details.continuation).toMatch(/^mgc1_/);
    for (const changed of [{ listId: "other", limit: 50 }, { listId: "list", limit: 49 }]) {
      const response = await tool.execute("changed", { action: "list_tasks", ...changed, continuation: first.details.continuation });
      expect(response.details).toMatchObject({ ok: false, error: "invalid_continuation" });
    }
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    fetchSpy.mockImplementation(async () => Response.json({ value: [{ id: "task" }], "@odata.nextLink": "https://example.invalid/v1.0/me/todo/lists/list/tasks" }));
    const unsafe = await tool.execute("unsafe", { action: "list_tasks", listId: "list" });
    expect(unsafe.details).toMatchObject({ ok: false, error: "invalid_provider_response" });
  });
});
