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

function registeredTools() {
  const factories: Array<(context: any) => any> = [];
  entry.register({
    pluginConfig: {
      enabled: true,
      policy: graphPolicyFixture(),
    },
    registerTool: (factory: any) => factories.push(factory),
    on: vi.fn(),
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  return Object.fromEntries(factories.map((factory) => {
    const tool = factory({ agentId: "main" });
    return [tool.name, tool];
  }));
}

describe("execution contracts", () => {
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

  it.each([
    ["reply_draft", "createReply", { action: "reply_draft", messageId: "original", bodyText: "Reply" }],
    ["reply_all_draft", "createReplyAll", { action: "reply_all_draft", messageId: "original", bodyHtml: "<p>Reply all</p>" }],
    ["forward_draft", "createForward", { action: "forward_draft", messageId: "original", bodyText: "Forward", to: ["person@example.invalid"] }],
  ])("%s creates a complete draft in one Graph mutation", async (action, endpoint, params) => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      expect(init?.method).toBe("POST");
      return Response.json({ id: `${action}-draft` }, { status: 201 });
    });
    try {
      const response = await registeredTools().outlook_mail_write.execute("draft", params);
      expect(response.details).toEqual({ ok: true, action, item: { id: `${action}-draft` } });
      expect(fetchSpy).toHaveBeenCalledTimes(1);
      expect(String(fetchSpy.mock.calls[0][0])).toContain(`/messages/original/${endpoint}`);
      expect(JSON.parse(String(fetchSpy.mock.calls[0][1]?.body))).toHaveProperty("message.body");
      expect(fetchSpy.mock.calls.some(([, init]) => init?.method === "PATCH")).toBe(false);
    } finally {
      fetchSpy.mockRestore();
    }
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
