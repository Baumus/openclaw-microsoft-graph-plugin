import { mkdtemp, mkdir, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

let directory: string | undefined;
const originalStateDir = process.env.OPENCLAW_STATE_DIR;
afterEach(async () => {
  if (originalStateDir === undefined) delete process.env.OPENCLAW_STATE_DIR;
  else process.env.OPENCLAW_STATE_DIR = originalStateDir;
  if (directory) await rm(directory, { recursive: true, force: true });
  directory = undefined;
});

async function candidate(warningApprovalsRequired = true) {
  directory = await mkdtemp(join(tmpdir(), "mg-session-key-"));
  const stateDir = join(directory, "state");
  const workspaceDir = join(directory, "workspace");
  await mkdir(stateDir);
  await mkdir(join(workspaceDir, "reports"), { recursive: true });
  await writeFile(join(workspaceDir, "reports", "onepager.pdf"), "synthetic private content");
  process.env.OPENCLAW_STATE_DIR = stateDir;
  const policy = graphPolicyFixture();
  delete policy.services.onedrive.allowed_roots[0].agents_instructions;
  policy.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
  let storedId: string | undefined = "generation-A";
  const getSessionEntry = vi.fn(({ agentId, sessionKey, readConsistency }: { agentId: string; sessionKey: string; readConsistency: string }) => {
    expect({ agentId, sessionKey, readConsistency }).toEqual({ agentId: "main", sessionKey: "agent:main:main", readConsistency: "latest" });
    return storedId ? { sessionId: storedId } : undefined;
  });
  const hooks: Record<string, (event: unknown, context: unknown) => Promise<any>> = {};
  const factories: Array<(context: any) => any> = [];
  entry.register({
    pluginConfig: { enabled: true, policy, warningApprovalsRequired },
    runtime: { state: { resolveStateDir: () => stateDir }, agent: { session: { getSessionEntry } } },
    on: (name: string, handler: (event: unknown, context: unknown) => Promise<any>) => { hooks[name] = handler; },
    registerTool: (factory: (context: any) => any) => factories.push(factory),
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as never);
  const toolContext = { agentId: "main", sessionKey: "agent:main:main", sessionId: "generation-A", workspaceDir };
  const tool = factories.map((factory) => factory(toolContext)).find((candidate) => candidate.name === "onedrive_upload");
  const event = { toolName: "onedrive_upload", toolCallId: `session-key-${Math.random()}`, params: {
    rootLabel: "synthetic_documents", relativePath: "onepager.pdf", sourceWorkspacePath: "reports/onepager.pdf",
  } };
  const hookContext = { agentId: "main", sessionKey: "agent:main:main", workspaceDir };
  const staged = async () => {
    const root = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging");
    const runs = await readdir(root).catch(() => []);
    return (await Promise.all(runs.map((run) => readdir(join(root, run))))).flat().filter((name) => name !== ".workspace-staging-owner");
  };
  return { hook: hooks.before_tool_call, tool, event, hookContext, getSessionEntry, setStoredId: (id: string | undefined) => { storedId = id; }, staged };
}

describe("host sessionKey-only approval context", () => {
  it("binds a one-call workspace upload to the persisted generation and reaches the credential gate", async () => {
    const { hook, tool, event, hookContext, getSessionEntry, staged } = await candidate();
    const approval = await hook(event, hookContext);
    expect(approval.requireApproval.severity).toBe("warning");
    expect(await staged()).toHaveLength(1);
    await approval.requireApproval.onResolution("allow-once");
    // OpenClaw 2026.9.8 shallow-merges hook overrides into the original call.
    const hostParams = { ...event.params, ...approval.params };
    expect(hostParams).toHaveProperty("sourceWorkspacePath", "reports/onepager.pdf");
    expect(hostParams).toHaveProperty("sourceMediaUri");
    expect((await tool.execute(event.toolCallId, hostParams)).details).toMatchObject({ ok: false, error: "credential_vault_unavailable" });
    expect(await staged()).toEqual([]);
    expect(getSessionEntry).toHaveBeenCalledTimes(3);
  });

  it("binds the host-merged params when warning approval is policy-disabled", async () => {
    const { hook, tool, event, hookContext, staged } = await candidate(false);
    const decision = await hook(event, hookContext);
    expect(decision.requireApproval).toBeUndefined();
    const hostParams = { ...event.params, ...decision.params };
    expect(hostParams).toHaveProperty("sourceWorkspacePath", "reports/onepager.pdf");
    expect(hostParams).toHaveProperty("sourceMediaUri");
    expect((await tool.execute(event.toolCallId, hostParams)).details).toMatchObject({ ok: false, error: "credential_vault_unavailable" });
    expect(await staged()).toEqual([]);
  });

  it("rejects a later rewrite of the original workspace path", async () => {
    const { hook, tool, event, hookContext, staged } = await candidate(false);
    const decision = await hook(event, hookContext);
    const altered = { ...event.params, ...decision.params, sourceWorkspacePath: "reports/other.pdf" };
    expect((await tool.execute(event.toolCallId, altered)).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed", mutationApplied: false });
    expect(await staged()).toEqual([]);
  });

  it("fails closed when the stored session disappears before preflight", async () => {
    const { hook, event, hookContext, setStoredId, staged } = await candidate();
    setStoredId(undefined);
    expect(await hook(event, hookContext)).toEqual({ block: true, blockReason: "trusted_session_identity_required" });
    expect(await staged()).toEqual([]);
  });

  it("rejects a reset before approval resolution and cleans the owned staging copy", async () => {
    const { hook, event, hookContext, setStoredId, staged } = await candidate();
    const approval = await hook(event, hookContext);
    setStoredId("generation-B");
    await expect(approval.requireApproval.onResolution("allow-once")).rejects.toThrow("approval_context_invalid_or_changed");
    expect(await staged()).toEqual([]);
  });

  it("rejects a reset after approval but before execution and cleans the owned staging copy", async () => {
    const { hook, tool, event, hookContext, setStoredId, staged } = await candidate();
    const approval = await hook(event, hookContext);
    await approval.requireApproval.onResolution("allow-once");
    setStoredId("generation-B");
    expect((await tool.execute(event.toolCallId, approval.params)).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed", mutationApplied: false });
    expect(await staged()).toEqual([]);
  });

  it("rejects disagreement between hook sessionId and persisted sessionKey generation", async () => {
    const { hook, event, hookContext, staged } = await candidate();
    expect(await hook(event, { ...hookContext, sessionId: "generation-B" })).toEqual({ block: true, blockReason: "trusted_session_identity_required" });
    expect(await staged()).toEqual([]);
  });
});
