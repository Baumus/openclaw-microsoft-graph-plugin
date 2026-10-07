import { describe, expect, it, vi } from "vitest";

const credentials = vi.hoisted(() => ({ token: vi.fn(async () => "synthetic-token") }));
vi.mock("./credential.js", async () => {
  const actual = await vi.importActual<typeof import("./credential.js")>("./credential.js");
  return { ...actual, tokenForAuthorizedOperation: credentials.token };
});

import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

const params = { action: "delete_exact", calendarId: "synthetic-calendar", subject: "Exact", eventDate: "2026-10-07", timeZone: "UTC" };
function setup(policy = graphPolicyFixture()) {
  const factories: Array<(context: any) => any> = [];
  const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
  entry.register({ pluginConfig: { enabled: true, policy }, registerTool: (factory: any) => factories.push(factory), on: (name: string, handler: any) => { hooks[name] = handler; }, logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
  const context = { agentId: "main", sessionId: "calendar-exact-test" };
  const tool = factories.map((factory) => factory(context)).find((candidate) => candidate.name === "outlook_calendar_write");
  return { tool, before: (id: string, input: Record<string, unknown> = params) => hooks.before_tool_call({ toolName: tool.name, toolCallId: id, params: input }, context) };
}

function graphFixture(options: { secondPage?: boolean; malformed?: boolean; currentSubject?: string; deletedBeforeExecute?: boolean } = {}) {
  const calls: string[] = [];
  const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input, init) => {
    const url = new URL(String(input));
    const path = url.pathname.replace(/^\/v1\.0/, "");
    calls.push(`${init?.method ?? "GET"} ${path}`);
    const event = (id: string, subject = "Exact") => ({ id, subject, start: { dateTime: "2026-10-07T09:00:00", timeZone: "UTC" } });
    if (init?.method === "DELETE") return new Response(null, { status: 204 });
    if (path.endsWith("/calendarView")) return new Response(JSON.stringify(options.malformed ? {} : url.searchParams.has("$skiptoken")
      ? { value: [event("event-2")] }
      : { value: [event("event-1")], ...(options.secondPage ? { "@odata.nextLink": `https://graph.microsoft.com/v1.0${path}?$skiptoken=page2` } : {}) }), { status: 200 });
    if (path.endsWith("/events/event-1")) return options.deletedBeforeExecute ? new Response(JSON.stringify({ error: { code: "ErrorItemNotFound" } }), { status: 404 }) : new Response(JSON.stringify(event("event-1", options.currentSubject)), { status: 200 });
    throw new Error(`unexpected path ${path}`);
  });
  return { calls, fetchSpy };
}

describe("calendar exact-subject/date delete", () => {
  it.each([
    ["duplicate on later page", { secondPage: true }, "exact_search_ambiguous"],
    ["malformed page", { malformed: true }, "invalid_provider_response"],
  ] as const)("fails closed on %s", async (_name, options, error) => {
    const { before } = setup();
    const fixture = graphFixture(options);
    try {
      expect(await before(`blocked-${_name}`)).toEqual({ block: true, blockReason: error });
      expect(fixture.calls.some((call) => call.startsWith("DELETE"))).toBe(false);
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it("binds the exact calendar/event identity and rejects param drift", async () => {
    const { before, tool } = setup();
    const fixture = graphFixture();
    try {
      const gate = await before("bound");
      expect(gate.params.eventId).toBe("event-1");
      expect(gate.requireApproval).toMatchObject({ severity: "critical", allowedDecisions: ["allow-once", "deny"] });
      expect(gate.requireApproval.description).toContain('calendar ID "synthetic-calendar", event ID "event-1"');
      await gate.requireApproval.onResolution("allow-once");
      expect((await tool.execute("bound", { ...gate.params, eventId: "event-2" })).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
      expect(fixture.calls.some((call) => call.startsWith("DELETE"))).toBe(false);
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it("rechecks stale subject and avoids DELETE", async () => {
    const { before, tool } = setup();
    const fixture = graphFixture({ currentSubject: "Changed" });
    try {
      const gate = await before("stale");
      await gate.requireApproval.onResolution("allow-once");
      expect((await tool.execute("stale", gate.params)).details).toMatchObject({ ok: false, error: "exact_target_changed", mutationApplied: false, retrySafety: "safe_after_correction" });
      expect(fixture.calls.some((call) => call.startsWith("DELETE"))).toBe(false);
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it("does not DELETE when the approved event disappears before execution", async () => {
    const { before, tool } = setup();
    const fixture = graphFixture({ deletedBeforeExecute: true });
    try {
      const gate = await before("deleted");
      await gate.requireApproval.onResolution("allow-once");
      expect((await tool.execute("deleted", gate.params)).details).toMatchObject({ ok: false });
      expect(fixture.calls.filter((call) => call.startsWith("DELETE"))).toHaveLength(0);
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it("deletes only the approved, revalidated event and returns its identity", async () => {
    const { before, tool } = setup();
    const fixture = graphFixture();
    try {
      const gate = await before("success");
      await gate.requireApproval.onResolution("allow-once");
      expect((await tool.execute("success", gate.params)).details).toMatchObject({ ok: true, deleted: true, calendarId: "synthetic-calendar", eventId: "event-1", subject: "Exact", eventDate: "2026-10-07" });
      expect(fixture.calls.filter((call) => call.startsWith("DELETE"))).toEqual(["DELETE /me/calendars/synthetic-calendar/events/event-1"]);
    } finally { fixture.fetchSpy.mockRestore(); }
  });

  it("requires read authority before credentials and network", async () => {
    const policy = graphPolicyFixture();
    policy.services.calendar.agents.main.operations = ["delete"];
    const { before } = setup(policy);
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    credentials.token.mockClear();
    try {
      expect(await before("denied")).toEqual({ block: true, blockReason: "access_denied" });
      expect(credentials.token).not.toHaveBeenCalled();
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally { fetchSpy.mockRestore(); }
  });
});
