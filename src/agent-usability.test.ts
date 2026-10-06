import { describe, expect, it, vi } from "vitest";
import { getToolPluginMetadata } from "openclaw/plugin-sdk/tool-plugin";

const credential = vi.hoisted(() => ({ tokenForAuthorizedOperation: vi.fn(async () => "synthetic-token") }));
vi.mock("./credential.js", async () => ({ ...await vi.importActual<typeof import("./credential.js")>("./credential.js"), ...credential }));

import entry, { beforeMicrosoftGraphToolCall, callerCapabilities, lifecycleResult } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function runtime(policy = graphPolicyFixture()) {
  const factories: Array<(context: any) => any> = [];
  const hooks: Record<string, (...args: any[]) => any> = {};
  entry.register({ pluginConfig: { enabled: true, policy }, registerTool: (factory: any) => factories.push(factory), on: (name: string, handler: any) => { hooks[name] = handler; }, logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
  const context = { agentId: "main", sessionId: "agent-usability" };
  const tools = Object.fromEntries(factories.map((factory) => { const tool = factory(context); return [tool.name, tool]; }));
  return { tools, hooks, context };
}

describe("agent-facing Microsoft Graph contracts", () => {
  it("advertises concrete guidance for every runtime tool and action-specific schema hints", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const { tools, context } = runtime();
    expect(metadata.tools).toHaveLength(26);
    for (const definition of metadata.tools) {
      const concrete = tools[definition.name];
      expect(concrete.description, definition.name).toBeTruthy();
      expect(concrete.description.length, definition.name).toBeGreaterThan(90);
      expect(concrete.description, definition.name).not.toBe(`Microsoft Graph ${definition.name} operation.`);
      expect(definition.description).toBe(concrete.description);
      expect(concrete.name).toBe(definition.name);
      expect(context.agentId).toBe("main");
    }
    for (const name of ["outlook_calendar_read", "outlook_calendar_write", "outlook_mail_read", "outlook_mail_write", "microsoft_todo_read", "microsoft_todo_write"]) {
      const action = (metadata.tools.find((tool) => tool.name === name)!.parameters as any).properties.action;
      expect(action.description, name).toMatch(/(?:list|create|update|delete|send|search)/);
      expect(action.description, name).toMatch(/Id|ID/);
    }
    for (const definition of metadata.tools) {
      const schema = definition.parameters as any;
      const isWrite = [
        "onedrive_upload", "onedrive_update", "onedrive_metadata_update", "onedrive_create_folder", "onedrive_delete",
        "onedrive_root_folder_create", "onedrive_root_folder_delete_exact",
        "outlook_calendar_write", "outlook_calendar_event_create", "outlook_calendar_event_delete_exact",
        "outlook_mail_write",
        "microsoft_todo_write", "microsoft_todo_default_task_create", "microsoft_todo_task_delete_exact",
      ].includes(definition.name);
      expect(Boolean(schema.properties?.timeoutMs), definition.name).toBe(isWrite);
      if (isWrite) expect(schema.properties.timeoutMs.maximum).toBe(600_000);
    }
  });

  it("strips timeout metadata before semantic approval binding and provider execution", async () => {
    const { tools, hooks, context } = runtime();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) =>
      init?.method === "POST"
        ? new Response(null, { status: 202 })
        : new Response(JSON.stringify({ id: "draft-1", internetMessageId: "<draft-1@example.test>", isDraft: true }), {
          status: 200, headers: { "content-type": "application/json" },
        }));
    try {
      const params = { action: "send_draft", messageId: "draft-1", timeoutMs: 180_000 };
      const gate = await hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "budgeted-send", params }, context);
      expect(gate.params).toEqual({ action: "send_draft", messageId: "draft-1" });
      expect(gate.requireApproval).toMatchObject({ severity: "critical", timeoutMs: 120_000, allowedDecisions: ["allow-once", "deny"] });
      gate.requireApproval.onResolution("allow-once");
      const response = (await tools.outlook_mail_write.execute("budgeted-send", params)).details;
      expect(response).toMatchObject({ ok: true, sent: true, deliveryStatus: "unknown", mutationApplied: true, retrySafety: "readback_before_retry" });
      expect(fetchSpy.mock.calls.some((call) => String(call[0]).includes("/messages/draft-1/send"))).toBe(true);
      expect(JSON.stringify(fetchSpy.mock.calls)).not.toContain("timeoutMs");
    } finally { fetchSpy.mockRestore(); }
  });

  it("rejects invalid transport, policy-denied delete, deny, and expiry before provider writes", async () => {
    const policy = graphPolicyFixture();
    const { hooks, tools, context } = runtime(policy);
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      const invalid = await hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "invalid-budget", params: { action: "send_draft", messageId: "draft-1", timeoutMs: 600_001 } }, context);
      expect(invalid).toEqual({ block: true, blockReason: "invalid_tool_parameters" });
      expect(invalid.requireApproval).toBeUndefined();

      const deniedDelete = await hooks.before_tool_call({ toolName: "onedrive_delete", toolCallId: "delete-denied", params: { rootLabel: "synthetic_documents", relativePath: "file.txt", timeoutMs: 180_000 } }, context);
      expect(deniedDelete).toEqual({ block: true, blockReason: "access_denied" });
      expect(deniedDelete.requireApproval).toBeUndefined();

      for (const decision of ["deny", "timeout"] as const) {
        const callId = `send-${decision}`;
        const params = { action: "send_draft", messageId: "draft-1", timeoutMs: 180_000 };
        const gate = await hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: callId, params }, context);
        gate.requireApproval.onResolution(decision);
        expect((await tools.outlook_mail_write.execute(callId, params)).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed", mutationApplied: false });
      }
      expect(fetchSpy).not.toHaveBeenCalled();
      expect(credential.tokenForAuthorizedOperation).not.toHaveBeenCalled();
    } finally { fetchSpy.mockRestore(); }
  });

  it("leaves unrelated host tools untouched", async () => {
    expect(await beforeMicrosoftGraphToolCall({ enabled: true, policy: graphPolicyFixture() }, { toolName: "unrelated_tool", params: { arbitrary: true } }, { agentId: "main", sessionId: "unrelated" })).toBeUndefined();
  });

  it("reports uncertainty, partial results, empty results, and send acceptance without delivery certainty", () => {
    expect(lifecycleResult("outlook_mail_write", { ok: false, error: "request_timeout" })).toMatchObject({ phase: "failed", mutationApplied: "unknown", retrySafety: "readback_before_retry" });
    for (const error of ["workspace_file_unavailable", "workspace_file_changed", "workspace_context_unavailable", "exactly_one_source_required"]) {
      const result = lifecycleResult("onedrive_upload", { ok: false, error }) as Record<string, unknown>;
      expect(result).toMatchObject({ phase: "failed", mutationApplied: false, retrySafety: "safe_after_correction" });
      expect(String(result.nextAction).length).toBeGreaterThan(40);
    }

    expect(lifecycleResult("outlook_mail_read", { ok: true, items: [], truncated: true })).toMatchObject({ phase: "partial", noResults: false, code: "partial_results" });
    expect(lifecycleResult("outlook_mail_read", { ok: true, items: [], truncated: false })).toMatchObject({ phase: "complete", noResults: true, code: "ok" });
    expect(lifecycleResult("outlook_calendar_write", { ok: false, action: "multiwrite", outcome: "partial", operations: [{ applied: true }, { applied: false, status: 400 }] })).toMatchObject({ phase: "partial", mutationApplied: true, retrySafety: "readback_before_retry" });
    expect(lifecycleResult("microsoft_todo_default_task_create", { ok: true, action: "create_task", item: { id: "task-1" } })).toMatchObject({ phase: "complete", mutationApplied: true, retrySafety: "do_not_repeat" });
  });

  it("exposes only effective caller grants and prerequisites without credentials or foreign grants", async () => {
    const policy = graphPolicyFixture();
    policy.services.mail.agents.foreign = { operations: ["send"], resources: ["private-mailbox"] };
    const result = callerCapabilities({ enabled: true, policy }, "main");
    expect(result).toMatchObject({ ok: true, roots: [{ rootLabel: "synthetic_documents", actions: ["write"] }], services: { mail: { actions: expect.arrayContaining(["send"]), resources: ["me"] } } });
    const serialized = JSON.stringify(result);
    expect(serialized).not.toMatch(/private-mailbox|foreign|synthetic-drive|synthetic-root|refreshToken|accessToken/);
    expect((await runtime(policy).tools.microsoft_graph_capabilities.execute("readiness", {})).details).toMatchObject({ ok: true, roots: expect.any(Array) });
  });

  it("omits non-me Mail and To Do grants that their handlers cannot execute", () => {
    const policy = graphPolicyFixture();
    policy.services.mail.agents.main = { operations: ["send"], resources: ["other-mailbox"] };
    policy.services.todo.agents.main = { operations: ["create"], resources: ["other-list"] };
    const result = callerCapabilities({ enabled: true, policy }, "main") as any;
    expect(result.services.mail).toMatchObject({ actions: [], resources: [], limitation: expect.stringContaining("me grant") });
    expect(result.services.todo).toMatchObject({ actions: [], resources: [], limitation: expect.stringContaining("me grant") });
    expect(JSON.stringify(result)).not.toMatch(/other-mailbox|other-list/);
  });

  it("projects only executable me resources from mixed Mail and To Do grants without disclosing foreign resources", () => {
    const policy = graphPolicyFixture();
    policy.services.mail.agents.main = { operations: ["read", "send"], resources: ["me", "other-mailbox"] };
    policy.services.todo.agents.main = { operations: ["read", "create"], resources: ["other-list", "me"] };
    policy.services.mail.agents.foreign = { operations: ["delete"], resources: ["foreign-mailbox"] };
    const result = callerCapabilities({ enabled: true, policy }, "main") as any;
    expect(result.services.mail).toMatchObject({ actions: ["read", "send"], resources: ["me"], limitation: expect.stringContaining("other resource grants are unsupported") });
    expect(result.services.todo).toMatchObject({ actions: ["read", "create"], resources: ["me"], limitation: expect.stringContaining("other resource grants are unsupported") });
    expect(JSON.stringify(result)).not.toMatch(/other-mailbox|other-list|foreign-mailbox|foreign/);
  });
});
