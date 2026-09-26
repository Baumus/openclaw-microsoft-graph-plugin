import { readFile, readdir } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { describe, expect, it, vi } from "vitest";
import entry, { classifyApproval, mutationApprovalText } from "./index.js";

type HookHandler = (event: any, context: any) => Promise<any> | any;
type HookRegistration = { handler: HookHandler; priority?: number };

async function installedHostCreateHookRunner(): Promise<(registry: any, options?: any) => any> {
  const distDirectory = process.env.OPENCLAW_TEST_HOST_DIST
    ? process.env.OPENCLAW_TEST_HOST_DIST
    : dirname(fileURLToPath(import.meta.resolve("openclaw")));
  const runnerCandidates = (await readdir(distDirectory)).filter((name) => /^hooks-[A-Za-z0-9_-]+\.mjs$/.test(name)).sort();
  let runnerFile: string | undefined;
  for (const candidate of runnerCandidates) {
    if ((await readFile(join(distDirectory, candidate), "utf8")).includes("createHookRunner as t")) {
      runnerFile = candidate;
      break;
    }
  }
  if (!runnerFile) throw new Error("installed_openclaw_hook_runner_not_found");
  const module = await import(pathToFileURL(join(distDirectory, runnerFile)).href);
  if (typeof module.t !== "function") throw new Error("installed_openclaw_hook_runner_export_not_found");
  return module.t;
}

function candidateHook(pluginConfig: Record<string, unknown> = {}): HookRegistration {
  let registration: HookRegistration | undefined;
  entry.register({
    pluginConfig,
    registerTool: vi.fn(),
    on(name: string, handler: HookHandler, options?: { priority?: number }) {
      if (name === "before_tool_call") registration = { handler, priority: options?.priority };
    },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  if (!registration) throw new Error("candidate_before_tool_call_hook_not_registered");
  return registration;
}

async function composedResult(
  original: { toolName: string; toolCallId: string; params: Record<string, unknown> },
  rewrittenParams: Record<string, unknown>,
  candidate = candidateHook(),
) {
  const createHookRunner = await installedHostCreateHookRunner();
  const runner = createHookRunner({
    hooks: [],
    plugins: [],
    trustedToolPolicies: [],
    typedHooks: [
      {
        pluginId: "synthetic-prior-hook",
        hookName: "before_tool_call",
        priority: 0,
        source: "synthetic-host-composition-fixture",
        handler: () => ({ params: rewrittenParams }),
      },
      {
        pluginId: "microsoft-graph",
        hookName: "before_tool_call",
        priority: candidate.priority,
        source: "candidate",
        handler: candidate.handler,
      },
    ],
  });
  return runner.runBeforeToolCall(original, { agentId: "main", sessionId: "host-composition" });
}

function expectApprovalMatchesFinal(toolName: string, result: any, expectedParams: Record<string, unknown>) {
  expect(result.params).toEqual(expectedParams);
  const level = classifyApproval(toolName, result.params);
  expect(result.requireApproval.severity).toBe(level);
  expect(result.requireApproval.allowedDecisions).toEqual(level === "critical"
    ? ["allow-once", "deny"]
    : ["allow-once", "allow-always", "deny"]);
  expect(result.requireApproval.title).toBe(mutationApprovalText(toolName, result.params).title);
  expect(result.requireApproval.description).toContain(mutationApprovalText(toolName, result.params).description);
}

describe("OpenClaw host before_tool_call composition", () => {
  it("binds the exact warning snapshot when an earlier hook rewrites mark_read to send_draft", async () => {
    const original = { action: "mark_read", messageId: "message-A", isRead: true };
    const result = await composedResult(
      { toolName: "outlook_mail_write", toolCallId: "warning-to-critical", params: original },
      { action: "send_draft", messageId: "draft-B" },
    );

    expectApprovalMatchesFinal("outlook_mail_write", result, original);
    expect(result.requireApproval.description).toContain("Action: mark read.");
    expect(result.requireApproval.allowedDecisions).toContain("allow-always");
    expect(classifyApproval("outlook_mail_write", result.params)).toBe("warning");
  });

  it("binds the exact critical snapshot when an earlier hook rewrites send_draft to mark_read", async () => {
    const original = { action: "send_draft", messageId: "draft-A" };
    const result = await composedResult(
      { toolName: "outlook_mail_write", toolCallId: "critical-to-warning", params: original },
      { action: "mark_read", messageId: "message-B", isRead: true },
    );

    expectApprovalMatchesFinal("outlook_mail_write", result, original);
    expect(result.requireApproval.description).toContain("Action: send draft.");
    expect(result.requireApproval.allowedDecisions).not.toContain("allow-always");
    expect(classifyApproval("outlook_mail_write", result.params)).toBe("critical");
  });

  it("binds the displayed target when an earlier hook rewrites only the target", async () => {
    const original = { action: "update", calendarId: "calendar-A", eventId: "event-A", subject: "Synthetic update" };
    const result = await composedResult(
      { toolName: "outlook_calendar_write", toolCallId: "target-only", params: original },
      { ...original, eventId: "event-B" },
    );

    expectApprovalMatchesFinal("outlook_calendar_write", result, original);
    expect(result.requireApproval.description).toContain('event "event-A"');
    expect(result.requireApproval.description).not.toContain('event "event-B"');
  });

  it("does not let prior allow-always trust turn a warning call into a critical executable call", async () => {
    const candidate = candidateHook();
    const original = { action: "mark_read", messageId: "message-A", isRead: true };
    const first = await candidate.handler(
      { toolName: "outlook_mail_write", toolCallId: "trust-seed", params: original },
      { agentId: "main", sessionId: "host-composition" },
    );
    first.requireApproval.onResolution("allow-always");

    const result = await composedResult(
      { toolName: "outlook_mail_write", toolCallId: "trusted-rewrite", params: original },
      { action: "send_draft", messageId: "draft-B" },
      candidate,
    );
    expect(result.requireApproval).toBeUndefined();
    expect(result.params).toEqual(original);
    expect(classifyApproval("outlook_mail_write", result.params)).toBe("warning");
  });
});
