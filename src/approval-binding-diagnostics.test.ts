import { describe, expect, it, vi } from "vitest";
import { concrete, NativeApprovalSnapshotStore, beforeMicrosoftGraphToolCall, approvalDiagnosticCategory } from "./index.js";
import { normalizedCriteria } from "./continuation.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

describe("approval binding diagnostics", () => {
  const params = { rootLabel: "synthetic_documents", relativePath: "private-name.md", sourceWorkspacePath: "private-source.md" };
  const root = { label: "private-label", drive_id: "private-drive", item_id: "private-root" };
  it.each(["agent", "session", "tool", "params", "root", "missing", "match"])("reports only comparison flags: %s", (changed) => {
    const store = new NativeApprovalSnapshotStore();
    store.record("private-call", { agentId: "private-agent", sessionId: "private-session", toolName: "onedrive_upload",
      params: normalizedCriteria(params), oneDriveRoot: root });
    const report = vi.fn();
    const matched = store.consume(changed === "missing" ? "different-call" : "private-call",
      changed === "agent" ? "other" : "private-agent", changed === "session" ? "other" : "private-session",
      changed === "tool" ? "onedrive_update" : "onedrive_upload", changed === "params" ? {} : params,
      changed === "root" ? undefined : root, undefined, report);
    expect(matched).toBe(changed === "missing" ? undefined : changed === "match");
    expect(report).toHaveBeenCalledOnce();
    expect(Object.values(report.mock.calls[0][0]).every(value => typeof value === "boolean")).toBe(true);
    expect(JSON.stringify(report.mock.calls)).not.toMatch(/private-/);
    if (changed !== "missing") expect(store.consume("private-call", "private-agent", "private-session", "onedrive_upload", params, root)).toBeUndefined();
  });
  it("correlates missing-snapshot and preflight without raw identities or values", async () => {
    const logger = { info: vi.fn(), warn: vi.fn() };
    const config = { enabled: true, policy: graphPolicyFixture() };
    const call = "private-call-id";
    // A denied hook still provides a safe preflight observation.
    await beforeMicrosoftGraphToolCall(config, { toolName: "onedrive_upload", toolCallId: call, params },
      { agentId: "private-agent", sessionId: "private-session" }, undefined, undefined, undefined, undefined, undefined, logger);
    const action = vi.fn();
    const tool = concrete("onedrive_upload", {}, "private-agent", "private-session", logger, action);
    expect((await tool.execute(call, params)).details).toMatchObject({ error: "approval_context_invalid_or_changed", mutationApplied: false });
    expect(action).not.toHaveBeenCalled();
    const logs = logger.info.mock.calls.map(([text]) => JSON.parse(text)).filter(entry => entry.event === "microsoft_graph_approval_binding");
    expect(logs.map(entry => entry.stage)).toEqual(["preflight", "approval_wait", "consume", "execute_approval_error"]);
    expect(logs[0].call).toBe(logs[1].call);
    expect(logs[0].process).toBe(logs[1].process);
    expect(logs[2].snapshot_present).toBe(false);
    expect(JSON.stringify(logs)).not.toMatch(/private-|synthetic_documents/);
    expect(Object.keys(logs[2]).sort()).toEqual(["call", "event", "process", "snapshot_present", "stage", "tool"]);
  });
  it("maps arbitrary valid-looking tokens and instruction text to unknown", () => {
    for (const value of ["private_token_123", "allow-once-secret", "onedrive_agents_instructions_required:{private-content}", "__proto__", null]) {
      expect(approvalDiagnosticCategory(value)).toBe("unknown");
    }
    expect(approvalDiagnosticCategory("approval_context_invalid_or_changed")).toBe("approval_binding");
  });
  it("does not let diagnostic logger failure change a hook denial", async () => {
    const logger = { info: () => { throw new Error("synthetic logger failure"); }, warn: vi.fn() };
    const result = await beforeMicrosoftGraphToolCall({ enabled: false }, { toolName: "onedrive_upload", toolCallId: "x", params },
      {}, undefined, undefined, undefined, undefined, undefined, logger);
    expect(result).toEqual({ block: true, blockReason: "connector_disabled" });
  });
});
