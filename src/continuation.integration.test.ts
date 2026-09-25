import { describe, expect, it, vi } from "vitest";
import { getToolPluginMetadata } from "openclaw/plugin-sdk/tool-plugin";

const mocks = vi.hoisted(() => ({
  readCredential: vi.fn(async () => ({ clientId: "synthetic", refreshToken: "synthetic", tenant: "common", scopes: ["Files.Read", "Calendars.Read", "Mail.Read", "Tasks.Read"] })),
}));
vi.mock("./credential.js", () => ({
  readCredential: mocks.readCredential,
  selectScope: (_credential: unknown, allowed: string[]) => allowed[0],
  exchangeRefreshToken: vi.fn(async () => "synthetic-access-token"),
  tokenForAuthorizedOperation: vi.fn(async () => { await mocks.readCredential(); return "synthetic-access-token"; }),
}));

import entry, { READ_ACTION_FIELDS } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

describe("opaque continuation tool integration", () => {
  it("round-trips opaque handles across Calendar, Mail, OneDrive, and To Do", async () => {
    const factories: Array<(context: any) => any> = [];
    const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
    entry.register({
      pluginConfig: { enabled: true, policy: graphPolicyFixture() },
      registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger,
    } as any);
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = new URL(String(input));
      const continued = url.searchParams.has("$skiptoken");
      const value = [{ id: continued ? "two" : "one", name: "item", displayName: "item", title: "item", subject: "item", file: { mimeType: "text/plain" } }];
      return new Response(JSON.stringify({ value, ...(!continued ? { "@odata.nextLink": `https://graph.microsoft.com${url.pathname}?$skiptoken=next` } : {}) }), { status: 200 });
    });
    try {
      const cases = [
        { tool: factories[1]({ agentId: "fixture-reader" }), params: { rootLabel: "synthetic_documents", relativePath: "SYNTHETIC_FOLDER", limit: 1 } },
        { tool: factories[9]({ agentId: "main" }), params: { action: "list_events", startDateTime: "2026-09-01T00:00:00Z", endDateTime: "2026-09-02T00:00:00Z", limit: 1 } },
        { tool: factories[11]({ agentId: "main" }), params: { action: "list_messages", folder: "inbox", limit: 1 } },
        { tool: factories[13]({ agentId: "main" }), params: { action: "list_lists", limit: 1 } },
      ];
      for (const { tool, params } of cases) {
        const first = await tool.execute("first", params);
        expect(first.details.continuation).toMatch(/^mgc1_[A-Za-z0-9_-]{43}$/);
        expect(JSON.stringify(first.details)).not.toContain("graph.microsoft.com");
        const second = await tool.execute("second", { ...params, continuation: first.details.continuation });
        expect(second.details).toMatchObject({ ok: true, truncated: false });
        expect(second.details).not.toHaveProperty("continuation");
      }
      const calendarCalls = fetchSpy.mock.calls.filter(([input]) => new URL(String(input)).pathname.endsWith("/me/calendarView"));
      expect(calendarCalls).toHaveLength(2);
      for (const [, init] of calendarCalls) expect(new Headers(init?.headers).get("Prefer")).toContain('outlook.timezone="UTC"');
      expect(logger.info.mock.calls.flat().join(" ")).not.toMatch(/mgc1_|graph\.microsoft\.com|\$skiptoken/);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("round-trips exact filename fallback state under the same pinned root and criteria", async () => {
    const factories: Array<(context: any) => any> = [];
    const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
    entry.register({
      pluginConfig: { enabled: true, policy: graphPolicyFixture() },
      registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger,
    } as any);
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) return new Response(JSON.stringify({ value: [] }), { status: 200 });
      if (url.pathname.endsWith("/items/synthetic-root/children")) return url.searchParams.has("$skiptoken")
        ? new Response(JSON.stringify({ value: [{ id: "nested-folder", name: "SYNTHETIC_NESTED_FOLDER", folder: {}, parentReference: { id: "synthetic-root", driveId: "synthetic-drive" } }] }), { status: 200 })
        : new Response(JSON.stringify({ value: [{ id: "root-file", name: "SYNTHETIC_RECORD.pdf", file: { mimeType: "application/pdf" }, parentReference: { id: "synthetic-root", driveId: "synthetic-drive" } }], "@odata.nextLink": `${url.origin}${url.pathname}?$skiptoken=opaque%2Bpage` }), { status: 200 });
      return new Response(JSON.stringify({ value: [{ id: "nested-file", name: "synthetic_record.PDF", file: { mimeType: "application/pdf" }, parentReference: { id: "nested-folder", driveId: "synthetic-drive" } }] }), { status: 200 });
    });
    try {
      const tool = factories[0]({ agentId: "fixture-reader" });
      const params = { rootLabel: "synthetic_documents", query: "SYNTHETIC_RECORD.pdf", mode: "filename_exact", exhaustive: true, limit: 1 };
      const first = await tool.execute("first", params);
      expect(first.details).toMatchObject({ ok: true, items: [expect.objectContaining({ id: "root-file" })], truncated: true, scan_complete: false, match_satisfied: true });
      expect(first.details).not.toHaveProperty("fallback");
      expect(first.details.continuation).toMatch(/^mgc1_[A-Za-z0-9_-]{43}$/);
      expect(JSON.stringify(first.details)).not.toContain("graph.microsoft.com");

      const second = await tool.execute("second", { ...params, continuation: first.details.continuation });
      expect(second.details).toMatchObject({ ok: true, items: [expect.objectContaining({ id: "nested-file" })], truncated: false, scan_complete: true, match_satisfied: true });
      expect(fetchSpy.mock.calls.some(([input]) => String(input).includes("/search("))).toBe(false);
      const credentialCalls = mocks.readCredential.mock.calls.length;
      const networkCalls = fetchSpy.mock.calls.length;
      const modified = await tool.execute("modified", { ...params, query: "Other.pdf", continuation: first.details.continuation });
      expect(modified.details).toEqual({ ok: false, error: "invalid_continuation" });
      const changedMode = await tool.execute("changed-mode", { ...params, mode: "filename_contains", continuation: first.details.continuation });
      expect(changedMode.details).toEqual({ ok: false, error: "invalid_continuation" });
      const changedExhaustive = await tool.execute("changed-exhaustive", { ...params, exhaustive: false, continuation: first.details.continuation });
      expect(changedExhaustive.details).toEqual({ ok: false, error: "invalid_continuation" });
      expect(mocks.readCredential).toHaveBeenCalledTimes(credentialCalls);
      expect(fetchSpy).toHaveBeenCalledTimes(networkCalls);
      expect(logger.info.mock.calls.flat().join(" ")).not.toMatch(/mgc1_|graph\.microsoft\.com|\$skiptoken/);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("binds provider-search continuation to agent, action, root, and query without leaking its URL", async () => {
    const factories: Array<(context: any) => any> = [];
    const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
    entry.register({
      pluginConfig: { enabled: true, policy: graphPolicyFixture() },
      registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger,
    } as any);
    const searchPath = "/v1.0/drives/synthetic-drive/items/synthetic-root/search(q='quarterly')";
    const next = `https://graph.microsoft.com${searchPath}?$skiptoken=A%2BB%252F&$top=1`;
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = new URL(String(input));
      if (url.pathname.includes("/search(")) return url.searchParams.has("$skiptoken")
        ? new Response(JSON.stringify({ value: [{ id: "second", name: "quarterly second", file: { mimeType: "text/plain" } }] }), { status: 200 })
        : new Response(JSON.stringify({ value: [{ id: "first", name: "quarterly first", file: { mimeType: "text/plain" } }], "@odata.nextLink": next }), { status: 200 });
      const itemId = decodeURIComponent(url.pathname.split("/").at(-1)!);
      return new Response(JSON.stringify(itemId === "synthetic-root"
        ? { id: itemId }
        : { id: itemId, parentReference: { id: "synthetic-root", driveId: "synthetic-drive" } }), { status: 200 });
    });
    try {
      const params = { rootLabel: "synthetic_documents", query: "quarterly", limit: 1 };
      const tool = factories[0]({ agentId: "fixture-reader" });
      const first = await tool.execute("first", params);
      expect(first.details).toMatchObject({ ok: true, items: [expect.objectContaining({ id: "first" })], truncated: true });
      expect(first.details.continuation).toMatch(/^mgc1_[A-Za-z0-9_-]{43}$/);
      expect(JSON.stringify(first.details)).not.toContain("graph.microsoft.com");

      const credentialCalls = mocks.readCredential.mock.calls.length;
      const networkCalls = fetchSpy.mock.calls.length;
      const attacks = [
        [factories[0]({ agentId: "other-agent" }), { ...params, continuation: first.details.continuation }],
        [factories[1]({ agentId: "fixture-reader" }), { rootLabel: "synthetic_documents", relativePath: "", limit: 1, continuation: first.details.continuation }],
        [tool, { ...params, rootLabel: "other_root", continuation: first.details.continuation }],
        [tool, { ...params, query: "different", continuation: first.details.continuation }],
      ] as const;
      for (const [attackedTool, attackedParams] of attacks) expect((await attackedTool.execute("attack", attackedParams)).details).toEqual({ ok: false, error: "invalid_continuation" });
      expect(mocks.readCredential).toHaveBeenCalledTimes(credentialCalls);
      expect(fetchSpy).toHaveBeenCalledTimes(networkCalls);

      const second = await tool.execute("second", { ...params, continuation: first.details.continuation });
      expect(second.details).toMatchObject({ ok: true, items: [expect.objectContaining({ id: "second" })], truncated: false });
      expect(fetchSpy.mock.calls.some(([input]) => String(input) === next)).toBe(true);
      expect(JSON.stringify(second.details)).not.toContain("graph.microsoft.com");
      expect(logger.info.mock.calls.flat().join(" ")).not.toMatch(/mgc1_|graph\.microsoft\.com|\$skiptoken/);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("blocks a modified calendarView range before another credential or network access", async () => {
    const factories: Array<(context: any) => any> = [];
    entry.register({ pluginConfig: { enabled: true, policy: graphPolicyFixture() }, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response(JSON.stringify({ value: [{ id: "one" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=next" }), { status: 200 }));
    try {
      const tool = factories[9]({ agentId: "main" });
      const original = { action: "list_events", startDateTime: "2026-09-01T00:00:00Z", endDateTime: "2026-09-02T00:00:00Z", limit: 1 };
      const first = await tool.execute("first", original);
      const credentialCalls = mocks.readCredential.mock.calls.length;
      const networkCalls = fetchSpy.mock.calls.length;
      const attacked = await tool.execute("attack", { ...original, startDateTime: "1900-01-01T00:00:00Z", endDateTime: "2100-01-01T00:00:00Z", continuation: first.details.continuation });
      expect(attacked.details).toEqual({ ok: false, error: "invalid_continuation" });
      expect(mocks.readCredential).toHaveBeenCalledTimes(credentialCalls);
      expect(fetchSpy).toHaveBeenCalledTimes(networkCalls);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("rejects every schema-accepted field irrelevant to each read action before credentials or network", async () => {
    const factories: Array<(context: any) => any> = [];
    entry.register({ pluginConfig: { enabled: true }, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
    const values: Record<string, unknown> = {
      calendarId: "calendar-id", startDateTime: "2026-09-07T09:00:00Z", endDateTime: "2026-09-07T10:00:00Z", eventId: "event-id",
      schedules: ["person@example.invalid"], timeZone: "UTC", limit: 1, search: "needle", searchFields: ["subject"], eventType: "singleInstance",
      showAs: "busy", sensitivity: "normal", importance: "high", categories: ["Blue"], isAllDay: false, isCancelled: false, hasAttachments: false,
      isOnlineMeeting: false, organizer: "owner@example.invalid", attendee: "guest@example.invalid", includeBody: true, bodyContentType: "text",
      availabilityViewInterval: 30, continuation: `mgc1_${"a".repeat(43)}`,
      folder: "inbox", folderId: "folder-id", parentFolderId: "parent-id", recursive: true, includeHidden: true, messageId: "message-id", attachmentId: "attachment-id",
      searchKql: "subject:needle", receivedAfter: "2026-09-01T00:00:00Z", receivedBefore: "2026-09-02T00:00:00Z", sentAfter: "2026-09-01T00:00:00Z",
      sentBefore: "2026-09-02T00:00:00Z", createdAfter: "2026-09-01T00:00:00Z", modifiedAfter: "2026-09-01T00:00:00Z", isRead: false,
      isDraft: false, inferenceClassification: "focused", orderBy: "receivedDateTime", orderDirection: "asc", includeUniqueBody: true, includeHeaders: true,
      listId: "list-id", taskId: "task-id", status: "notStarted", isReminderOn: false,
    };
    const cases = [
      { service: "calendar" as const, factory: 9 },
      { service: "mail" as const, factory: 11 },
      { service: "todo" as const, factory: 13 },
    ];
    const credentialCalls = mocks.readCredential.mock.calls.length;
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      for (const { service, factory } of cases) {
        const tool = factories[factory]({ agentId: "main" });
        const serviceFields = new Set(Object.values(READ_ACTION_FIELDS[service]).flat());
        for (const [action, allowed] of Object.entries(READ_ACTION_FIELDS[service])) {
          for (const field of serviceFields) {
            if (field === "action" || allowed.includes(field)) continue;
            const response = await tool.execute(`${service}:${action}:${field}`, { action, [field]: values[field] });
            const error = service === "calendar" && field === "calendarId" ? "invalid_calendar_target" : "invalid_read_parameter";
            expect(response.details, `${service}:${action}:${field}`).toEqual({ ok: false, error });
          }
        }
      }
      expect(mocks.readCredential).toHaveBeenCalledTimes(credentialCalls);
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("assigns every shared-schema field to a consuming action and keeps OneDrive read schemas narrow", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const shared = [
      ["calendar", "outlook_calendar_read"],
      ["mail", "outlook_mail_read"],
      ["todo", "microsoft_todo_read"],
    ] as const;
    for (const [service, toolName] of shared) {
      const schema = metadata.tools.find((tool) => tool.name === toolName)!.parameters as any;
      const schemaFields = Object.keys(schema.properties).sort();
      const routedFields = [...new Set(Object.values(READ_ACTION_FIELDS[service]).flat())].sort();
      expect(routedFields, service).toEqual(schemaFields);
    }

    const oneDriveSchemas = {
      onedrive_search: ["rootLabel", "agentsInstructionAck", "query", "mode", "exhaustive", "limit", "continuation"],
      onedrive_list: ["rootLabel", "relativePath", "agentsInstructionAck", "limit", "continuation"],
      onedrive_read: ["rootLabel", "relativePath", "agentsInstructionAck", "mode"],
      onedrive_download: ["rootLabel", "relativePath", "agentsInstructionAck"],
    };
    for (const [toolName, fields] of Object.entries(oneDriveSchemas)) {
      const schema = metadata.tools.find((tool) => tool.name === toolName)!.parameters as any;
      expect(Object.keys(schema.properties), toolName).toEqual(fields);
    }
    const searchTool = metadata.tools.find((tool) => tool.name === "onedrive_search")!;
    const searchSchema = searchTool.parameters as any;
    expect(searchSchema.properties.mode.anyOf.map((entry: any) => entry.const)).toEqual(["provider", "filename_exact", "filename_stem", "filename_contains"]);
    expect(searchSchema.properties.exhaustive).toMatchObject({ type: "boolean", default: false, description: expect.stringContaining("Filename modes only") });
    expect(searchTool.description).toContain("scan completion");
  });

  it("rejects ignored criterion modifiers before credentials", async () => {
    const factories: Array<(context: any) => any> = [];
    entry.register({ pluginConfig: { enabled: true }, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
    const credentialCalls = mocks.readCredential.mock.calls.length;
    const cases = [
      [factories[0]({ agentId: "fixture-reader" }), { rootLabel: "synthetic_documents", query: "../README.md", mode: "filename_exact" }, "invalid_search"],
      [factories[0]({ agentId: "fixture-reader" }), { rootLabel: "synthetic_documents", query: "README", mode: "provider", exhaustive: true }, "invalid_search"],
      [factories[9]({ agentId: "main" }), { action: "search_events", startDateTime: "2026-09-07T09:00:00Z", endDateTime: "2026-09-07T10:00:00Z", searchFields: ["subject"], showAs: "busy" }, "invalid_search"],
      [factories[11]({ agentId: "main" }), { action: "list_messages", orderDirection: "asc" }, "invalid_order"],
      [factories[13]({ agentId: "main" }), { action: "search_tasks", listId: "list-id", searchFields: ["title"], status: "notStarted" }, "invalid_search"],
    ] as const;
    for (const [tool, params, error] of cases) expect((await tool.execute("dependency", params)).details).toEqual({ ok: false, error });
    expect(mocks.readCredential).toHaveBeenCalledTimes(credentialCalls);
  });
});
