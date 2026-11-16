import { mkdtemp, mkdir, writeFile, rm } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { expect, it, vi } from "vitest";
import plugin, { beforeMicrosoftGraphToolCall, NativeApprovalSnapshotStore } from "./index.js";
import { workspaceStagingStore } from "./stage-workspace-file.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

// Opt-in local diagnostic: real installed host code, only synthetic local state.
const host = process.env.MS_GRAPH_DIAGNOSTIC_HOST_DIST;
it.skipIf(!host)("qualifies exact markdown input through installed host approval callbacks", async () => {
  const load = (file: string) => import(/* @vite-ignore */ join(host!, file));
  const hooks = await load("hook-runner-global-Da5_MbMg.mjs");
  const core = await load("agent-tools.before-tool-call-BlgyX1yZ.mjs");
  const mode = await load("embedded-mode--Mk71zY-.mjs");
  const directory = await mkdtemp(join(tmpdir(), "mg-host-contract-"));
  const previous = process.env.OPENCLAW_STATE_DIR;
  const workspaceDir = join(directory, "workspace");
  const stateDir = join(directory, "state");
  await mkdir(join(workspaceDir, "reports"), { recursive: true });
  await mkdir(stateDir);
  await writeFile(join(workspaceDir, "reports", "briefing.md"), "# Synthetic briefing\n");
  process.env.OPENCLAW_STATE_DIR = stateDir;
  const policy = graphPolicyFixture();
  delete policy.services.onedrive.allowed_roots[0].agents_instructions;
  policy.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
  const typedHooks: any[] = [];
  const factories: any[] = [];
  let generation = "generation-A";
  const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
  const lookup = vi.fn(async () => ({ sessionId: generation }));
  const config = { enabled: true, policy, warningApprovalsRequired: true };
  plugin.register({
    pluginConfig: config,
    runtime: { state: { resolveStateDir: () => stateDir }, agent: { session: {
      getSessionEntryAsync: lookup,
    } } },
    on: (hookName: string, handler: any, options: any) => typedHooks.push({ pluginId: "microsoft-graph", hookName, handler, priority: options?.priority }),
    registerTool: (factory: any) => factories.push(factory), logger,
  } as never);
  hooks.a({ plugins: [{ id: "microsoft-graph" }], typedHooks, hooks: [], trustedToolPolicies: [] });
  const context = { agentId: "main", sessionId: "generation-A", sessionKey: "agent:main:main", workspaceDir };
  const tool = factories.map(factory => factory(context)).find(tool => tool.name === "onedrive_upload");
  const params = { rootLabel: "synthetic_documents", relativePath: "reports/briefing.md",
    sourceWorkspacePath: "reports/briefing.md", contentType: "text/markdown", timeoutMs: 180000 };
  let decision = "allow-once";
  mode.n(true);
  core.C({ request: async () => ({ decision }) });
  try {
    for (const senderIsOwner of [true, false]) {
      const toolCallId = "markdown-" + senderIsOwner;
      const outcome = await core._({ toolName: "onedrive_upload", params, toolCallId,
        ctx: { ...context, requester: { senderIsOwner, channel: "webchat" } } });
      expect(outcome.blocked).toBe(false);
      expect(outcome.params).toEqual(params);
      expect((await tool.execute(toolCallId, outcome.params)).details.error).toBe("credential_vault_unavailable");
      expect((await tool.execute(toolCallId, outcome.params)).details.error).toBe("approval_context_invalid_or_changed");
    }
    // Authentic trace: native admission and later dispatch preflight the same
    // invocation under a policy that already permits warning mutations.
    config.warningApprovalsRequired = false;
    const twice = { toolName: "onedrive_upload", params, toolCallId: "native-then-dispatch", ctx: context };
    expect((await core._(twice)).blocked).toBe(false);
    expect((await core._(twice)).blocked).toBe(false);
    expect((await tool.execute(twice.toolCallId, params)).details.error).toBe("credential_vault_unavailable");
    expect((await tool.execute(twice.toolCallId, params)).details.error).toBe("approval_context_invalid_or_changed");
    config.warningApprovalsRequired = true;
    // Native report mode defers approval; execution must await its callback.
    const deferred = await core._({ toolName: "onedrive_upload", params, toolCallId: "deferred",
      approvalMode: "defer", ctx: context });
    expect(deferred.blocked).toBe(false);
    expect(deferred.deferredApproval).toBeDefined();
    const resolved = await core.b({ deferredApproval: deferred.deferredApproval });
    expect(resolved.blocked).toBe(false);
    expect((await tool.execute("deferred", resolved.params)).details.error).toBe("credential_vault_unavailable");
    const injected = await core._({ toolName: "onedrive_upload", params, toolCallId: "injected", ctx: context });
    expect(injected.blocked).toBe(false);
    expect((await tool.execute("injected", { ...params, contentType: "text/plain" })).details.error).toBe("approval_context_invalid_or_changed");
    const changedCall = await core._({ toolName: "onedrive_upload", params, toolCallId: "bound-call", ctx: context });
    expect(changedCall.blocked).toBe(false);
    expect((await tool.execute("other-call", params)).details.error).toBe("approval_context_invalid_or_changed");
    const rootFailure = await beforeMicrosoftGraphToolCall({ enabled: true, policy },
      { toolName: "onedrive_upload", toolCallId: "root-mismatch", params }, context,
      { policyValidator: () => { const changed = structuredClone(policy); changed.services.onedrive.allowed_roots[0].item_id = "synthetic-changed-root"; return changed; } },
      undefined, undefined, undefined, lookup, logger);
    expect(rootFailure?.blockReason).toBe("approval_context_invalid_or_changed");
    const leaseSpy = vi.spyOn(workspaceStagingStore, "bind").mockImplementationOnce(() => { throw new Error("private_lease_token"); });
    const leaseFailure = await core._({ toolName: "onedrive_upload", params, toolCallId: "lease-failure", ctx: context });
    expect(leaseFailure.blocked).toBe(true);
    leaseSpy.mockRestore();
    const snapshotSpy = vi.spyOn(NativeApprovalSnapshotStore.prototype, "record").mockImplementationOnce(() => { throw new Error("private_snapshot_token"); });
    const snapshotFailure = await core._({ toolName: "onedrive_upload", params, toolCallId: "snapshot-failure", ctx: context });
    expect((await tool.execute("snapshot-failure", snapshotFailure.params)).details.error).toBe("approval_context_invalid_or_changed");
    snapshotSpy.mockRestore();
    const pending = await core._({ toolName: "onedrive_upload", params, toolCallId: "abort-wait", approvalMode: "defer", ctx: context });
    let release!: () => void;
    lookup.mockImplementationOnce(async () => { await new Promise<void>(resolve => { release = resolve; }); return { sessionId: generation }; });
    const pendingResolution = core.b({ deferredApproval: pending.deferredApproval });
    while (!release) await new Promise(resolve => setTimeout(resolve, 0));
    const aborted = new AbortController(); aborted.abort();
    expect((await tool.execute("abort-wait", params, aborted.signal)).details.error).toBe("approval_context_invalid_or_changed");
    release(); await pendingResolution;
    const odd = await beforeMicrosoftGraphToolCall({ enabled: true, policy },
      { toolName: "onedrive_upload", toolCallId: "odd-decision", params }, context, undefined, undefined, undefined, undefined, lookup, logger);
    await (odd as any).requireApproval.onResolution("private_decision_token");
    expect((await tool.execute("odd-decision", params)).details.error).toBe("approval_context_invalid_or_changed");
    decision = "deny";
    const denied = await core._({ toolName: "onedrive_upload", params, toolCallId: "denied", ctx: context });
    expect(denied.blocked).toBe(true);
    expect((await tool.execute("denied", params)).details.error).toBe("approval_context_invalid_or_changed");
    decision = "allow-once";
    const changed = await core._({ toolName: "onedrive_upload", params, toolCallId: "reset", ctx: context });
    expect(changed.blocked).toBe(false);
    // Use real Core noninteractive-surface rejection, not the synthetic approval broker.
    core.S(); mode.n(false);
    const unavailable = await core._({ toolName: "onedrive_upload", params, toolCallId: "no-surface",
      ctx: { ...context, trigger: "internal", config: {} } });
    expect(unavailable.blocked).toBe(true);
    expect(unavailable.reason).toBe("Plugin approval unavailable: internal runs have no approval-capable initiating surface.");
    expect((await tool.execute("no-surface", params)).details.error).toBe("approval_context_invalid_or_changed");
    generation = "generation-B";
    expect((await tool.execute("reset", params)).details.error).toBe("approval_context_invalid_or_changed");
    const diagnostics = logger.info.mock.calls.map(([line]) => JSON.parse(line)).filter(row => row.event === "microsoft_graph_approval_binding");
    for (const stage of ["preflight", "hook_session", "approval_route", "hook_result", "resolution", "bind_session", "recorded", "consume", "execute_session", "artifact_reused", "lease_bind_error", "snapshot_bind_error", "instruction_preflight_error", "approval_wait_error", "execute_approval_error"]) {
      expect(diagnostics.some(row => row.stage === stage), stage).toBe(true);
    }
    const safe = JSON.stringify(diagnostics);
    for (const value of [params.rootLabel, params.relativePath, params.sourceWorkspacePath, context.sessionKey, context.sessionId, "markdown-true", "text/markdown", "private_lease_token", "private_snapshot_token", "private_decision_token"]) expect(safe).not.toContain(value);
    expect(diagnostics.some(row => row.stage === "resolution" && row.decision === "unknown")).toBe(true);
    for (const stage of ["lease_bind_error", "snapshot_bind_error"]) expect(diagnostics.some(row => row.stage === stage && row.category === "unknown")).toBe(true);
    expect(diagnostics.some(row => row.stage === "instruction_preflight_error" && row.category === "approval_binding")).toBe(true);
    expect(diagnostics.some(row => row.stage === "approval_wait_error" && row.aborted)).toBe(true);
    const consumed = diagnostics.filter(row => row.stage === "consume");
    expect(consumed.some(row => row.snapshot_present === false)).toBe(true);
    expect(consumed.some(row => row.params_match === false)).toBe(true);
    expect(diagnostics.some(row => row.stage === "execute_session" && row.current === false)).toBe(true);
  } finally {
    core.S(); mode.n(false); hooks.o();
    if (previous === undefined) delete process.env.OPENCLAW_STATE_DIR;
    else process.env.OPENCLAW_STATE_DIR = previous;
    await rm(directory, { recursive: true, force: true });
  }
}, 30000);
