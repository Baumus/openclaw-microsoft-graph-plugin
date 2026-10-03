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
    expect(metadata.tools).toHaveLength(17);
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
      const isWrite = ["onedrive_upload", "onedrive_update", "onedrive_metadata_update", "onedrive_create_folder", "onedrive_delete", "outlook_calendar_write", "outlook_mail_write", "microsoft_todo_write"].includes(definition.name);
      expect(Boolean(schema.properties?.timeoutMs), definition.name).toBe(isWrite);
      if (isWrite) expect(schema.properties.timeoutMs.maximum).toBe(600_000);
    }
  });

  it("strips timeout metadata before semantic approval binding and provider execution", async () => {
    const { tools, hooks, context } = runtime();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response(null, { status: 202 }));
    try {
      const params = { action: "send_draft", messageId: "draft-1", timeoutMs: 180_000 };
      const gate = await hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "budgeted-send", params }, context);
      expect(gate.params).toEqual({ action: "send_draft", messageId: "draft-1" });
      expect(gate.requireApproval).toMatchObject({ severity: "critical", timeoutMs: 120_000, allowedDecisions: ["allow-once", "deny"] });
      gate.requireApproval.onResolution("allow-once");
      const response = (await tools.outlook_mail_write.execute("budgeted-send", params)).details;
      expect(response).toMatchObject({ ok: true, sent: true, deliveryStatus: "unknown", mutationApplied: true, retrySafety: "readback_before_retry" });
      expect(String(fetchSpy.mock.calls[0][0])).toContain("/messages/draft-1/send");
      expect(JSON.stringify(fetchSpy.mock.calls[0])).not.toContain("timeoutMs");
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
    expect(lifecycleResult("outlook_mail_read", { ok: true, items: [], truncated: true })).toMatchObject({ phase: "partial", noResults: false, code: "partial_results" });
    expect(lifecycleResult("outlook_mail_read", { ok: true, items: [], truncated: false })).toMatchObject({ phase: "complete", noResults: true, code: "ok" });
    expect(lifecycleResult("outlook_calendar_write", { ok: false, action: "multiwrite", outcome: "partial", operations: [{ applied: true }, { applied: false, status: 400 }] })).toMatchObject({ phase: "partial", mutationApplied: true, retrySafety: "readback_before_retry" });
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
});
