import { setTimeout as delay } from "node:timers/promises";
import { describe, expect, it, vi } from "vitest";

vi.mock("./credential.js", async () => {
  const actual = await vi.importActual<typeof import("./credential.js")>("./credential.js");
  return {
    ...actual,
    readCredential: vi.fn(async (secretRef: string) => ({
      clientId: "synthetic",
      refreshToken: "synthetic",
      tenant: "common",
      scopes: secretRef.endsWith("/read") ? ["Mail.Read", "Tasks.Read", "offline_access"] : ["Mail.ReadWrite", "Tasks.ReadWrite", "offline_access"],
    })),
    exchangeRefreshToken: vi.fn(async () => "synthetic-token"),
    tokenForAuthorizedOperation: vi.fn(async () => "synthetic-token"),
  };
});

import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function registeredTools(policy = graphPolicyFixture(), limits = {}) {
  const factories: Array<(context: any) => any> = [];
  const hooks: Record<string, (...args: any[]) => Promise<any> | any> = {};
  entry.register({
    pluginConfig: {
      enabled: true,
      policy,
      ...limits,
    },
    registerTool: (factory: any) => factories.push(factory),
    on: (name: string, handler: any) => { hooks[name] = handler; },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  const context = { agentId: "main", sessionId: "execution-contracts" };
  return Object.fromEntries(factories.map((factory) => {
    const tool = factory(context);
    return [tool.name, { ...tool, async execute(toolCallId: string, params: Record<string, unknown>, signal?: AbortSignal) {
      const gate = await hooks.before_tool_call({ toolName: tool.name, toolCallId, params }, context);
      if (gate?.block) return { details: { ok: false, error: gate.blockReason } };
      await gate?.requireApproval?.onResolution("allow-once");
      return tool.execute(toolCallId, gate?.params ?? params, signal);
    } }];
  }));
}

describe("execution contracts", () => {
  it("reads, searches, and mutates only own default contacts with bounded fields and exact IDs", async () => {
    const policy = graphPolicyFixture();
    policy.services.contacts = { agents: { main: { operations: ["read", "create", "update", "delete"], resources: ["me"] } } };
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input, init) => {
      const url = new URL(String(input));
      if (init?.method === "DELETE") return new Response(null, { status: 204 });
      if (init?.method === "POST" || init?.method === "PATCH") return Response.json({ id: "A=B", displayName: "Ada Lovelace", providerOnly: "hidden" });
      if (url.pathname.endsWith("/contacts/A%3DB")) return Response.json({ id: "A=B", displayName: "Ada Lovelace", providerOnly: "hidden" });
      return Response.json({ value: [{ id: "A=B", displayName: "Ada Lovelace", emailAddresses: [{ address: "ada@example.invalid" }], providerOnly: "hidden" }] });
    });
    try {
      const tools = registeredTools(policy);
      const listed = (await tools.outlook_contacts_read.execute("contacts-list", { action: "list_contacts", limit: 10 })).details;
      const found = (await tools.outlook_contacts_read.execute("contacts-search", { action: "search_contacts", search: "ada", limit: 10 })).details;
      const one = (await tools.outlook_contacts_read.execute("contacts-get", { action: "get_contact", contactId: "A=B" })).details;
      const created = (await tools.outlook_contacts_write.execute("contacts-create", { action: "create", displayName: "Ada Lovelace" })).details;
      const updated = (await tools.outlook_contacts_write.execute("contacts-update", { action: "update", contactId: "A=B", companyName: "Analytical Engines" })).details;
      const deleted = (await tools.outlook_contacts_write.execute("contacts-delete", { action: "delete", contactId: "A=B" })).details;
      expect(listed).toMatchObject({ ok: true, items: [{ id: "A=B", displayName: "Ada Lovelace" }] });
      expect(found).toMatchObject({ ok: true, completeness: "complete", items: [{ id: "A=B" }] });
      expect(one).toMatchObject({ ok: true, item: { id: "A=B" } });
      expect(created).toMatchObject({ ok: true, item: { id: "A=B" }, mutationApplied: true });
      expect(updated).toMatchObject({ ok: true, item: { id: "A=B" } });
      expect(deleted).toMatchObject({ ok: true, deleted: true });
      for (const result of [listed.items[0], found.items[0], one.item, created.item, updated.item]) expect(result).not.toHaveProperty("providerOnly");
      expect(fetchSpy.mock.calls.map(([input]) => new URL(String(input)).pathname)).toEqual(expect.arrayContaining(["/v1.0/me/contacts", "/v1.0/me/contacts/A%3DB"]));
      expect(fetchSpy.mock.calls.map(([, init]) => init?.method)).toEqual(expect.arrayContaining(["POST", "PATCH", "DELETE"]));
    } finally { fetchSpy.mockRestore(); }
  });
  it("does not claim a contact is absent when bounded client-side search is incomplete", async () => {
    const policy = graphPolicyFixture();
    policy.services.contacts = { agents: { main: { operations: ["read"], resources: ["me"] } } };
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = new URL(String(input));
      const page = Number(url.searchParams.get("$skiptoken") ?? "0");
      return Response.json({ value: Array.from({ length: 50 }, (_, index) => ({ id: `contact-${page * 50 + index}`, displayName: "Unrelated" })), "@odata.nextLink": `https://graph.microsoft.com/v1.0/me/contacts?$skiptoken=${page + 1}` });
    });
    try {
      const response = (await registeredTools(policy).outlook_contacts_read.execute("search-cap", { action: "search_contacts", search: "missing" })).details;
      expect(response).toMatchObject({ ok: true, items: [], scanned: 500, truncated: true, completeness: "partial", noResults: false });
      expect(response).not.toHaveProperty("continuation");
      expect(fetchSpy).toHaveBeenCalledTimes(10);
    } finally { fetchSpy.mockRestore(); }
  });
  it("encodes opaque To Do resource IDs exactly once across mutation paths", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      if (init?.method === "PATCH") return Response.json({ id: "BBTask=" }, { status: 200 });
      return Response.json({ id: "AAList=", isOwner: true, isShared: false }, { status: 200 });
    });
    try {
      const response = await registeredTools().microsoft_todo_write.execute("update", {
        action: "update_task",
        listId: "AAList=",
        taskId: "BBTask=",
        title: "Updated",
      });
      expect(response.details).toMatchObject({ ok: true, action: "update_task" });
      expect(String(fetchSpy.mock.calls[0][0])).toContain("/todo/lists/AAList%3D");
      expect(String(fetchSpy.mock.calls[1][0])).toContain("/todo/lists/AAList%3D/tasks/BBTask%3D");
      for (const [input] of fetchSpy.mock.calls) expect(String(input)).not.toContain("%253D");
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("avoids unsupported To Do $select queries and sanitizes provider fields locally", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = new URL(String(input));
      if (url.pathname.endsWith("/tasks/task")) {
        return Response.json({ id: "task", title: "Canary", providerOnly: "hidden" }, { status: 200 });
      }
      if (url.pathname.endsWith("/tasks")) {
        return Response.json({ value: [{ id: "task", title: "Canary", providerOnly: "hidden" }] }, { status: 200 });
      }
      return Response.json({ value: [{ id: "list", displayName: "Tasks", isOwner: true, isShared: false, wellknownListName: "defaultList", providerOnly: "hidden" }] }, { status: 200 });
    });
    try {
      const tool = registeredTools().microsoft_todo_read;
      const listResult = (await tool.execute("lists", { action: "list_lists", limit: 1 })).details;
      const taskResult = (await tool.execute("tasks", { action: "list_tasks", listId: "list", limit: 1 })).details;
      const itemResult = (await tool.execute("task", { action: "get_task", listId: "list", taskId: "task" })).details;
      expect(listResult).toMatchObject({ ok: true, items: [{ id: "list", displayName: "Tasks" }] });
      expect(taskResult).toMatchObject({ ok: true, items: [{ id: "task", title: "Canary" }] });
      expect(itemResult).toMatchObject({ ok: true, item: { id: "task", title: "Canary" } });
      expect(listResult.items[0]).not.toHaveProperty("providerOnly");
      expect(taskResult.items[0]).not.toHaveProperty("providerOnly");
      expect(itemResult.item).not.toHaveProperty("providerOnly");
      for (const [input] of fetchSpy.mock.calls) expect(new URL(String(input)).searchParams.has("$select")).toBe(false);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("reads ten lists through the registered To Do action and filters before global ranking", async () => {
    let active = 0, peak = 0;
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = new URL(String(input));
      if (url.pathname.endsWith("/lists")) return Response.json({
        value: Array.from({ length: 10 }, (_, i) => ({ id: `list-${i}=`, displayName: `List ${i}` })),
      });
      const id = Number(/list-(\d+)/.exec(url.pathname)?.[1]);
      peak = Math.max(peak, ++active);
      await delay(2);
      active--;
      if (!url.searchParams.has("$skiptoken")) return Response.json({
        value: Array.from({ length: 25 }, (_, i) => ({ id: `done-${i}`, status: "completed" })),
        "@odata.nextLink": `https://graph.microsoft.com${url.pathname}?$skiptoken=2`,
      });
      return Response.json({ value: [{
        id: `task-${id}`, title: `Open ${id}`, status: "waitingOnOthers",
        dueDateTime: { dateTime: `2030-01-${20 - id}T10:00:00.0000000`, timeZone: "UTC" },
        body: { content: "not projected" },
      }] });
    });
    try {
      const reply = await registeredTools(graphPolicyFixture(), { maxConcurrent: 2 }).microsoft_todo_read
        .execute("aggregate", { action: "aggregate_open_tasks" });
      expect(reply.details).toMatchObject({ ok: true, completeness: "complete", listsComplete: 10, matchedInScan: 10 });
      expect(reply.details.items.map((entry: { taskId: string }) => entry.taskId)).toEqual(["task-9", "task-8", "task-7", "task-6", "task-5"]);
      expect(reply.details.items[0]).toMatchObject({ listId: "list-9=", dueDateTime: "2030-01-11T10:00:00.000Z" });
      expect(reply.content[0].text).not.toContain("not projected");
      expect(fetchSpy).toHaveBeenCalledTimes(21);
      expect(peak).toBe(2);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("reports partial aggregate coverage, rejects cross-list continuation and never retries 429", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = new URL(String(input));
      if (url.pathname.endsWith("/lists")) return Response.json({
        value: [{ id: "limited" }, { id: "foreign-link" }, { id: "throttled" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists?$skiptoken=more",
      });
      if (url.pathname.includes("throttled")) return new Response(null, { status: 429 });
      const next = url.pathname.includes("foreign-link") ? "/v1.0/me/todo/lists/other/tasks" : url.pathname;
      return Response.json({ value: [], "@odata.nextLink": `https://graph.microsoft.com${next}?$skiptoken=more` });
    });
    try {
      const { details } = await registeredTools().microsoft_todo_read.execute("partial", { action: "aggregate_open_tasks" });
      expect(details).toMatchObject({ ok: true, completeness: "partial", inventoryComplete: false, noResults: false,
        coverage: [
          { listId: "limited", state: "scan_limit", scanned: 0 },
          { listId: "foreign-link", state: "read_error" },
          { listId: "throttled", state: "read_error" },
        ] });
      expect(fetchSpy).toHaveBeenCalledTimes(7);
      expect(fetchSpy.mock.calls.some(([url]) => String(url).includes("/other/"))).toBe(false);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("preserves To Do permission checks and reports matches omitted by the byte cap", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) =>
      Response.json({ value: String(input).includes("/tasks")
        ? [{ id: "large", title: "💡".repeat(1000), status: "deferred" }]
        : [{ id: "list" }] }));
    try {
      const policy = graphPolicyFixture();
      policy.services.todo = { agents: { main: { operations: ["create"], resources: ["me"] } } };
      const denied = await registeredTools(policy).microsoft_todo_read
        .execute("denied", { action: "aggregate_open_tasks" });
      expect(denied.details).toMatchObject({ ok: false, error: "access_denied" });
      expect(fetchSpy).not.toHaveBeenCalled();
      const reply = await registeredTools(graphPolicyFixture(), { maxReadOutputBytes: 1024 }).microsoft_todo_read
        .execute("bounded", { action: "aggregate_open_tasks" });
      expect(reply.details).toMatchObject({ ok: true, completeness: "complete", matchedInScan: 1,
        items: [], outputLimited: true, noResults: false });
      expect(Buffer.byteLength(JSON.stringify(reply.content[0].text))).toBeLessThanOrEqual(1024);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("forwards with an explicit body in one mutation", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(Response.json({ id: "forward-draft" }, { status: 201 }));
    try {
      const result = await registeredTools().outlook_mail_write.execute("forward", { action: "forward_draft", messageId: "original", bodyHtml: "<p>Forward</p>", to: ["person@example.invalid"] });
      expect(result.details).toMatchObject({ ok: true, item: { id: "forward-draft" } });
      expect(fetchSpy).toHaveBeenCalledTimes(1);
      expect(JSON.parse(String(fetchSpy.mock.calls[0][1]?.body))).toHaveProperty("message.body");
    } finally { fetchSpy.mockRestore(); }
  });

  it.each(["reply_draft", "reply_all_draft"])("%s creates native history then prepends only new HTML", async (action) => {
    const history = '<html><head><style>.sig{color:blue}</style></head><body class="mail"><div>Original</div><img src="https://images.example.invalid/signature.png"><img src="cid:original-logo"></body></html>';
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      if (init?.method === "POST") return Response.json({ id: "native-draft", isDraft: true });
      if (init?.method === "PATCH") return Response.json({ id: "native-draft", isDraft: true });
      return Response.json({ id: "native-draft", isDraft: true, body: { contentType: "html", content: history } });
    });
    try {
      const result = await registeredTools().outlook_mail_write.execute("reply", { action, messageId: "original", bodyHtml: "<p>New answer</p>" });
      expect(result.details).toMatchObject({ ok: true, item: { id: "native-draft" } });
      expect(fetchSpy).toHaveBeenCalledTimes(3);
      expect(JSON.parse(String(fetchSpy.mock.calls[0][1]?.body))).toEqual({});
      expect(String(fetchSpy.mock.calls[0][0])).toContain(action === "reply_draft" ? "/createReply" : "/createReplyAll");
      const patched = JSON.parse(String(fetchSpy.mock.calls[2][1]?.body));
      expect(patched).toEqual({ body: { contentType: "HTML", content: history.replace('<body class="mail">', '<body class="mail"><p>New answer</p>') } });
      for (const [input] of fetchSpy.mock.calls) expect(new URL(String(input)).hostname).toBe("graph.microsoft.com");
    } finally { fetchSpy.mockRestore(); }
  });

  it("supports native creation with no body and preserves draft identity on a failed insert", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      if (init?.method === "POST") return Response.json({ id: "retained-draft" });
      return Response.json({ id: "retained-draft", isDraft: false });
    });
    try {
      const tool = registeredTools().outlook_mail_write;
      expect((await tool.execute("native-only", { action: "reply_draft", messageId: "original" })).details).toMatchObject({ ok: true, item: { id: "retained-draft" } });
      expect(fetchSpy).toHaveBeenCalledTimes(1);
      const failed = (await tool.execute("insert-failed", { action: "reply_draft", messageId: "original", bodyHtml: "<p>Reply</p>" })).details;
      expect(failed).toMatchObject({ ok: false, draftCreated: true, item: { id: "retained-draft" } });
      expect(fetchSpy.mock.calls.filter(([, init]) => init?.method === "PATCH")).toHaveLength(0);
    } finally { fetchSpy.mockRestore(); }
  });

  it("prepends to an existing draft without new creation or attachment writes", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => Response.json(init?.method === "PATCH" ? { id: "existing" } : { id: "existing", isDraft: true, body: { contentType: "html", content: '<div>Quote<img src="https://images.example.invalid/logo"></div>' } }));
    try {
      const result = await registeredTools().outlook_mail_write.execute("prepend", { action: "prepend_reply_draft", messageId: "existing", bodyHtml: "<p>Answer</p>" });
      expect(result.details).toMatchObject({ ok: true, item: { id: "existing" } });
      expect(fetchSpy).toHaveBeenCalledTimes(2);
      expect(JSON.parse(String(fetchSpy.mock.calls[1][1]?.body)).body.content).toBe('<p>Answer</p><div>Quote<img src="https://images.example.invalid/logo"></div>');
    } finally { fetchSpy.mockRestore(); }
  });

  it("marks a terminal 1,000-result mail search as truncated with unknown completeness", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = new URL(String(input));
      const page = Number(url.searchParams.get("$skiptoken") ?? "0");
      return Response.json({
        value: Array.from({ length: 50 }, (_, index) => ({ id: `message-${page * 50 + index}` })),
        ...(page < 19 ? { "@odata.nextLink": `https://graph.microsoft.com/v1.0/me/messages?$skiptoken=${page + 1}` } : {}),
      });
    });
    try {
      const tool = registeredTools().outlook_mail_read;
      let continuation: string | undefined;
      let result: any;
      let count = 0;
      do {
        result = (await tool.execute("search", { action: "search_messages", search: "match", limit: 50, ...(continuation ? { continuation } : {}) })).details;
        count += result.items.length;
        continuation = result.continuation;
      } while (continuation);
      expect(count).toBe(1000);
      expect(result).toMatchObject({ ok: true, truncated: true, providerResultLimit: 1000, completeness: "unknown", warning: "provider_search_result_limit_reached" });
      expect(result).not.toHaveProperty("continuation");
      expect(fetchSpy).toHaveBeenCalledTimes(20);
    } finally {
      fetchSpy.mockRestore();
    }
  });
});
