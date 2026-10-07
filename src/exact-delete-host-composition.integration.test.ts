import { readFile, readdir } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { describe, expect, it, vi } from "vitest";

vi.mock("./credential.js", async () => {
  const actual = await vi.importActual<typeof import("./credential.js")>("./credential.js");
  return { ...actual, tokenForAuthorizedOperation: vi.fn(async () => "synthetic-token") };
});
import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

async function hostRunner(): Promise<(registry: any) => any> {
  const dir = process.env.OPENCLAW_TEST_HOST_DIST || dirname(fileURLToPath(import.meta.resolve("openclaw")));
  for (const name of (await readdir(dir)).filter((x) => /^hooks-[A-Za-z0-9_-]+\.mjs$/.test(x)).sort()) {
    if (!(await readFile(join(dir, name), "utf8")).includes("createHookRunner as t")) continue;
    const module = await import(pathToFileURL(join(dir, name)).href);
    if (typeof module.t === "function") return module.t;
  }
  throw new Error("installed_openclaw_hook_runner_not_found");
}
function candidate(pluginConfig: Record<string, unknown> = {}) {
  let hook: any;
  const factories: any[] = [];
  const policy = graphPolicyFixture();
  policy.services.calendar.agents.main.resources!.push("calendar-1");
  entry.register({
    pluginConfig: { enabled: true, policy, ...pluginConfig },
    registerTool: (factory: any) => factories.push(factory),
    on: (name: string, handler: any, options: any) => { if (name === "before_tool_call") hook = { handler, priority: options?.priority, timeoutMs: options?.timeoutMs }; },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  const context = { agentId: "main", sessionId: "host-exact-delete" };
  return { hook, context, tool: (name: string) => factories.map((factory) => factory(context)).find((tool) => tool.name === name) };
}

describe("installed OpenClaw host resolves exact-delete approval parameters", () => {
  it("aborts a slow exact scan before its registered host-hook deadline", async () => {
    const instance = candidate({ readOperationTimeoutMs: 20 });
    expect(instance.hook.timeoutMs).toBeGreaterThan(20);
    let aborted = false;
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation((_input, init) => new Promise((_resolve, reject) => {
      init?.signal?.addEventListener("abort", () => { aborted = true; reject(init.signal?.reason); }, { once: true });
    }));
    try {
      const runner = (await hostRunner())({
        hooks: [], plugins: [], trustedToolPolicies: [],
        typedHooks: [{ pluginId: "microsoft-graph", hookName: "before_tool_call", priority: instance.hook.priority, timeoutMs: instance.hook.timeoutMs, source: "candidate", handler: instance.hook.handler }],
      });
      const gate = await runner.runBeforeToolCall({ toolName: "microsoft_todo_write", toolCallId: "slow-scan", params: { action: "delete_task_exact", title: "Synthetic exact target" } }, instance.context);
      expect(gate).toMatchObject({ block: true });
      expect(gate.requireApproval).toBeUndefined();
      expect(aborted).toBe(true);
    } finally { fetchSpy.mockRestore(); }
  });

  it.each([
    ["microsoft_todo_write", { action: "delete_task_exact", title: "Synthetic exact target" }, { listId: "list-1", taskId: "task-1" }],
    ["outlook_calendar_write", { action: "delete_exact", calendarId: "calendar-1", subject: "Synthetic exact target", eventDate: "2026-10-07", timeZone: "UTC" }, { eventId: "event-1" }],
  ] as const)("binds resolved IDs for %s through the real host hook reducer", async (toolName, params, resolved) => {
    const instance = candidate();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input, init) => {
      const path = new URL(String(input)).pathname.replace(/^\/v1\.0/, "");
      const json = (value: unknown) => new Response(JSON.stringify(value), { status: 200 });
      if (path === "/me/todo/lists") return json({ value: [{ id: "list-1", isOwner: true, isShared: false }] });
      if (path === "/me/todo/lists/list-1/tasks") return json({ value: [{ id: "task-1", title: "Synthetic exact target" }] });
      if (path === "/me/calendars/calendar-1/calendarView") return json({ value: [{ id: "event-1", subject: "Synthetic exact target", start: { dateTime: "2026-10-07T09:00:00", timeZone: "UTC" } }] });
      throw new Error("unexpected provider path: " + path);
    });
    try {
      const runner = (await hostRunner())({
        hooks: [], plugins: [], trustedToolPolicies: [],
        typedHooks: [{ pluginId: "microsoft-graph", hookName: "before_tool_call", priority: instance.hook.priority, source: "candidate", handler: instance.hook.handler }],
      });
      const id = "exact-" + toolName;
      const gate = await runner.runBeforeToolCall({ toolName, toolCallId: id, params }, instance.context);
      expect(gate.params).toMatchObject({ ...params, ...resolved });
      expect(gate.requireApproval).toMatchObject({ severity: "critical", allowedDecisions: ["allow-once", "deny"] });
      for (const value of Object.values(resolved)) expect(gate.requireApproval.description).toContain(value);
      await gate.requireApproval.onResolution("allow-once");
      const [key] = Object.keys(resolved);
      expect((await instance.tool(toolName).execute(id, { ...gate.params, [key]: "other-id" })).details)
        .toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
      expect(fetchSpy).not.toHaveBeenCalledWith(expect.anything(), expect.objectContaining({ method: "DELETE" }));
    } finally { fetchSpy.mockRestore(); }
  });
});
