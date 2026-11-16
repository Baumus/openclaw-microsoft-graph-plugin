import { mkdtemp, mkdir, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import entry from "./index.js";
import { workspaceStagingStore } from "./stage-workspace-file.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

let directory: string | undefined;
const originalStateDir = process.env.OPENCLAW_STATE_DIR;
afterEach(async () => {
  vi.restoreAllMocks();
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
  let lookupFails = false;
  const getSessionEntryAsync = vi.fn(async ({ agentId, sessionKey, readConsistency }: { agentId: string; sessionKey: string; readConsistency: string }) => {
    expect({ agentId, sessionKey, readConsistency }).toEqual({ agentId: "main", sessionKey: "agent:main:main", readConsistency: "latest" });
    if (lookupFails) throw new Error("session store unavailable");
    return storedId ? { sessionId: storedId } : undefined;
  });
  const deprecatedGetSessionEntry = vi.fn(() => { throw new Error("deprecated session getter called"); });
  const hooks: Record<string, (event: unknown, context: unknown) => Promise<any>> = {};
  const factories: Array<(context: any) => any> = [];
  entry.register({
    pluginConfig: { enabled: true, policy, warningApprovalsRequired },
    runtime: { state: { resolveStateDir: () => stateDir }, agent: { session: { getSessionEntryAsync, getSessionEntry: deprecatedGetSessionEntry } } },
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
  return { hook: hooks.before_tool_call, tool, event, hookContext, getSessionEntryAsync, deprecatedGetSessionEntry, setStoredId: (id: string | undefined) => { storedId = id; }, setLookupFails: (value: boolean) => { lookupFails = value; }, policy, staged };
}

describe("host sessionKey-only approval context", () => {
  it("reuses the bound staged bytes for repeated host preflight of the same call", async () => {
    const { hook, tool, event, hookContext, staged } = await candidate(false);
    event.params = { ...event.params, contentType: "text/markdown", timeoutMs: 180000 } as any;
    const first = await hook(event, hookContext);
    expect(first.block).not.toBe(true);
    const files = await staged();
    const second = await hook(event, hookContext);
    expect(second.block).not.toBe(true);
    expect(second.params).toEqual(event.params);
    expect(await staged()).toEqual(files);
    expect((await tool.execute(event.toolCallId, event.params)).details.error).toBe("credential_vault_unavailable");
    expect(await staged()).toEqual([]);
    expect((await tool.execute(event.toolCallId, event.params)).details.error).toBe("approval_context_invalid_or_changed");
  });

  it.each(["destination", "root", "session"])("rejects changed %s on repeated preflight", async (field) => {
    const c = await candidate(false);
    await c.hook(c.event, c.hookContext);
    const event = structuredClone(c.event);
    if (field === "destination") event.params.relativePath = "different.pdf";
    if (field === "root") c.policy.services.onedrive.allowed_roots[0].item_id = "changed-root";
    if (field === "session") c.setStoredId("generation-B");
    const result = await c.hook(event, c.hookContext);
    // A changed stored generation is resolved by the hook but mismatches the
    // recorded snapshot; no caller can borrow the old artifact.
    expect(result.block).toBe(true);
    expect((await c.tool.execute(c.event.toolCallId, c.event.params)).details.error).toBe("approval_context_invalid_or_changed");
    expect(await c.staged()).toEqual([]);
  });

  it("cannot resurrect a consumed snapshot while duplicate preflight awaits currentness", async () => {
    const c = await candidate(false);
    await c.hook(c.event, c.hookContext);
    let release!: () => void;
    c.getSessionEntryAsync.mockImplementationOnce(async () => ({ sessionId: "generation-A" }));
    c.getSessionEntryAsync.mockImplementationOnce(async () => {
      await new Promise<void>(resolve => { release = resolve; });
      return { sessionId: "generation-A" };
    });
    const duplicate = c.hook(c.event, c.hookContext);
    await vi.waitFor(() => expect(release).toBeTypeOf("function"));
    const cleanup = vi.spyOn(workspaceStagingStore, "cleanup");
    expect((await c.tool.execute(c.event.toolCallId, c.event.params)).details.error).toBe("credential_vault_unavailable");
    const executionCleanupCalls = cleanup.mock.calls.length;
    release();
    expect((await duplicate).block).toBe(true);
    expect(cleanup.mock.calls).toHaveLength(executionCleanupCalls);
    expect((await c.tool.execute(c.event.toolCallId, c.event.params)).details.error).toBe("approval_context_invalid_or_changed");
    expect(await c.staged()).toEqual([]);
  });

  it.each([true, false])("preserves Codex relay input with warning approval %s", async (required) => {
    const { hook, tool, event, hookContext, staged } = await candidate(required);
    const original = JSON.stringify(event.params);
    const decision = await hook(event, hookContext);
    expect(Object.keys(decision.params).filter(key => !(key in event.params)).sort())
      .toEqual([]);
    expect(JSON.stringify(decision.params)).toBe(original);
    if (required) await decision.requireApproval.onResolution("allow-once");
    expect(JSON.stringify(decision.params)).toBe(original);
    expect(await staged()).toHaveLength(1);
    expect((await tool.execute(event.toolCallId, event.params)).details)
      .toMatchObject({ ok: false, error: "credential_vault_unavailable" });
    expect(await staged()).toEqual([]);
  });

  it("binds a one-call workspace upload to the persisted generation and reaches the credential gate", async () => {
    const { hook, tool, event, hookContext, getSessionEntryAsync, deprecatedGetSessionEntry, staged } = await candidate();
    const approval = await hook(event, hookContext);
    expect(approval.requireApproval.severity).toBe("warning");
    expect(await staged()).toHaveLength(1);
    await approval.requireApproval.onResolution("allow-once");
    // OpenClaw 2026.9.8 shallow-merges hook overrides into the original call.
    const hostParams = { ...event.params, ...approval.params };
    expect(hostParams).toHaveProperty("sourceWorkspacePath", "reports/onepager.pdf");
    expect(hostParams).toEqual(event.params);
    expect((await tool.execute(event.toolCallId, hostParams)).details).toMatchObject({ ok: false, error: "credential_vault_unavailable" });
    expect(await staged()).toEqual([]);
    expect(getSessionEntryAsync).toHaveBeenCalledTimes(3);
    expect(deprecatedGetSessionEntry).not.toHaveBeenCalled();
  });

  it("binds the host-merged params when warning approval is policy-disabled", async () => {
    const { hook, tool, event, hookContext, staged } = await candidate(false);
    const decision = await hook(event, hookContext);
    expect(decision.requireApproval).toBeUndefined();
    const hostParams = { ...event.params, ...decision.params };
    expect(hostParams).toHaveProperty("sourceWorkspacePath", "reports/onepager.pdf");
    expect(hostParams).toEqual(event.params);
    expect((await tool.execute(event.toolCallId, hostParams)).details).toMatchObject({ ok: false, error: "credential_vault_unavailable" });
    expect(await staged()).toEqual([]);
  });

  it.each(["sourceMediaUri", "sourceSha256", "sourceByteSize", "contentType"])("rejects injected %s before execution", async (field) => {
    const { hook, tool, event, hookContext, staged } = await candidate(false);
    await hook(event, hookContext);
    const altered = { ...event.params, [field]: field === "sourceByteSize" ? 1 : "synthetic-injection" };
    expect((await tool.execute(event.toolCallId, altered)).details)
      .toMatchObject({ ok: false, error: "approval_context_invalid_or_changed", mutationApplied: false });
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

  it("fails closed when the awaited session lookup rejects", async () => {
    const { hook, event, hookContext, setLookupFails, deprecatedGetSessionEntry, staged } = await candidate();
    setLookupFails(true);
    expect(await hook(event, hookContext)).toEqual({ block: true, blockReason: "trusted_session_identity_required" });
    expect(await staged()).toEqual([]);
    expect(deprecatedGetSessionEntry).not.toHaveBeenCalled();
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
