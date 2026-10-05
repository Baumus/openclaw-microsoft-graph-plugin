import { createHash } from "node:crypto";
import { mkdtemp, mkdir, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import entry, { concrete, NativeApprovalSnapshotStore, openProtectedMediaUploadSource } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
import { workspaceStagingStore } from "./stage-workspace-file.js";

let directory: string | undefined;
const originalStateDir = process.env.OPENCLAW_STATE_DIR;
async function stagedFiles(stateDir: string) {
  const directory = join(stateDir, "media", "inbound", "baumus-msgraph-workspace-staging");
  const runs = await readdir(directory).catch(() => []);
  return (await Promise.all(runs.map((run) => readdir(join(directory, run))))).flat().filter((file) => file !== ".workspace-staging-owner");
}
afterEach(async () => {
  vi.restoreAllMocks();
  if (originalStateDir === undefined) delete process.env.OPENCLAW_STATE_DIR;
  else process.env.OPENCLAW_STATE_DIR = originalStateDir;
  if (directory) await rm(directory, { recursive: true, force: true });
  directory = undefined;
});

describe("workspace-file OneDrive approval preflight", () => {
  async function approvalWithFailingBoundCleanup(toolCallId: string) {
    directory = await mkdtemp(join(tmpdir(), "mg-bound-cleanup-"));
    const stateDir = join(directory, "state");
    const workspaceDir = join(directory, "workspace");
    await mkdir(join(workspaceDir, "reports"), { recursive: true });
    await mkdir(stateDir);
    process.env.OPENCLAW_STATE_DIR = stateDir;
    await writeFile(join(workspaceDir, "reports", "onepager.pdf"), "synthetic");
    const policy = graphPolicyFixture();
    delete policy.services.onedrive.allowed_roots[0].agents_instructions;
    policy.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
    const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
    const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
    entry.register({ pluginConfig: { enabled: true, policy }, runtime: { state: { resolveStateDir: () => stateDir } },
      registerTool: vi.fn(), on: (name: string, handler: (...args: any[]) => Promise<any>) => { hooks[name] = handler; }, logger } as never);
    const originalBind = workspaceStagingStore.bind.bind(workspaceStagingStore);
    const cleanup = vi.fn();
    vi.spyOn(workspaceStagingStore, "bind").mockImplementationOnce((id, name, lease, sessionId, onExpire) => {
      cleanup.mockRejectedValueOnce(new Error("private_workspace_cleanup_path"))
        .mockImplementation(() => lease.cleanup());
      originalBind(id, name, { cleanup }, sessionId, onExpire);
    });
    const context = { agentId: "main", sessionId: toolCallId, workspaceDir };
    const approval = await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId, params: {
      rootLabel: "synthetic_documents", relativePath: "onepager.pdf", sourceWorkspacePath: "reports/onepager.pdf",
    } }, context) as { params: Record<string, unknown>; requireApproval: { onResolution(decision: string): Promise<void> } };
    expect(approval.requireApproval).toBeDefined();
    return { approval, cleanup, context, logger, policy, stateDir };
  }

  it("contains denied approval cleanup failures, warns without private details, and retries the lease", async () => {
    const { approval, cleanup, context, logger, policy, stateDir } = await approvalWithFailingBoundCleanup("cleanup-deny");
    const action = vi.fn(async () => ({ ok: true }));
    await expect(approval.requireApproval.onResolution("deny")).resolves.toBeUndefined();
    expect(logger.warn).toHaveBeenCalledWith("workspace_staging_cleanup_deferred");
    expect(JSON.stringify(logger.warn.mock.calls)).not.toContain("private_workspace_cleanup_path");
    expect(await stagedFiles(stateDir)).toHaveLength(1);
    const tool = concrete("onedrive_upload", {}, context.agentId, context.sessionId, logger as never, action, { enabled: true, policy });
    expect((await tool.execute("cleanup-deny", approval.params)).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
    expect(action).not.toHaveBeenCalled();
    await workspaceStagingStore.retryFailedCleanups();
    expect(cleanup).toHaveBeenCalledTimes(2);
    expect(await stagedFiles(stateDir)).toEqual([]);
  });

  it("sanitizes a failed allow-once snapshot and its bound cleanup without approving execution", async () => {
    const { approval, cleanup, context, logger, policy, stateDir } = await approvalWithFailingBoundCleanup("cleanup-allow-failure");
    vi.spyOn(NativeApprovalSnapshotStore.prototype, "record").mockImplementationOnce(() => { throw new Error("private_snapshot_failure"); });
    await expect(approval.requireApproval.onResolution("allow-once")).rejects.toThrow("approval_context_invalid_or_changed");
    expect(logger.warn).toHaveBeenCalledWith("workspace_staging_cleanup_deferred");
    const action = vi.fn(async () => ({ ok: true }));
    const tool = concrete("onedrive_upload", {}, context.agentId, context.sessionId, logger as never, action, { enabled: true, policy });
    expect((await tool.execute("cleanup-allow-failure", approval.params)).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
    expect(action).not.toHaveBeenCalled();
    expect(JSON.stringify(logger.warn.mock.calls)).not.toMatch(/private_snapshot_failure|private_workspace_cleanup_path/);
    await workspaceStagingStore.retryFailedCleanups();
    expect(cleanup).toHaveBeenCalledTimes(2);
    expect(await stagedFiles(stateDir)).toEqual([]);
  });

  it.each(["applied", "failed"])("reports a sanitized deferred cleanup after a synthetic %s write", async (execution) => {
    const { approval, cleanup, context, logger, policy, stateDir } = await approvalWithFailingBoundCleanup(`cleanup-${execution}`);
    await approval.requireApproval.onResolution("allow-once");
    const action = vi.fn(async () => {
      if (execution === "failed") throw new Error("request_timeout");
      return { ok: true, item: { id: "synthetic-upload" } };
    });
    const tool = concrete("onedrive_upload", {}, context.agentId, context.sessionId, logger as never, action, { enabled: true, policy });
    const response = await tool.execute(`cleanup-${execution}`, approval.params);
    expect(action).toHaveBeenCalledTimes(1);
    expect(response.details).toMatchObject(execution === "applied"
      ? { ok: true, outcome: "applied_with_warning", cleanupWarning: "workspace_staging_cleanup_deferred", phase: "partial", mutationApplied: true, retrySafety: "do_not_repeat" }
      : { ok: false, error: "request_timeout", cleanupWarning: "workspace_staging_cleanup_deferred", phase: "failed", mutationApplied: "unknown", retrySafety: "readback_before_retry" });
    if (execution === "applied") expect((response.details as { nextAction: string }).nextAction).toContain("Do not repeat a completed write");
    expect(logger.warn).toHaveBeenCalledWith("workspace_staging_cleanup_deferred");
    expect(JSON.stringify([response, logger.info.mock.calls, logger.warn.mock.calls])).not.toContain("private_workspace_cleanup_path");
    expect(await stagedFiles(stateDir)).toHaveLength(1);
    await workspaceStagingStore.retryFailedCleanups();
    expect(cleanup).toHaveBeenCalledTimes(2);
    expect(await stagedFiles(stateDir)).toEqual([]);
  });

  it.each(["pre-binding", "before-binding"])("queues %s cleanup failures and returns a fixed block reason", async (failurePoint) => {
    directory = await mkdtemp(join(tmpdir(), "mg-stage-cleanup-"));
    const stateDir = join(directory, "state");
    const workspaceDir = join(directory, "workspace");
    await mkdir(join(workspaceDir, "reports"), { recursive: true });
    await mkdir(stateDir);
    process.env.OPENCLAW_STATE_DIR = stateDir;
    await writeFile(join(workspaceDir, "reports", "onepager.pdf"), "synthetic");
    const policy = graphPolicyFixture();
    delete policy.services.onedrive.allowed_roots[0].agents_instructions;
    policy.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
    const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
    entry.register({ pluginConfig: { enabled: true, policy }, runtime: { state: { resolveStateDir: () => stateDir } },
      registerTool: vi.fn(), on: (name: string, handler: (...args: any[]) => Promise<any>) => { hooks[name] = handler; },
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as never);
    const original = workspaceStagingStore.reserveShared.bind(workspaceStagingStore);
    const release = vi.fn();
    vi.spyOn(workspaceStagingStore, "reserveShared").mockImplementationOnce(async (...args) => {
      const reservation = await original(...args);
      release.mockRejectedValueOnce(new Error("workspace_private_cleanup_details")).mockImplementation(reservation.release);
      return { publish: reservation.publish, release };
    });
    const retry = vi.spyOn(workspaceStagingStore, "retryCleanup");
    if (failurePoint === "before-binding") vi.spyOn(workspaceStagingStore, "bind").mockImplementationOnce(() => { throw new Error("workspace_private_binding_details"); });
    const params = { rootLabel: "synthetic_documents", relativePath: "onepager.pdf", sourceWorkspacePath: "reports/onepager.pdf",
      ...(failurePoint === "pre-binding" ? { sourceSha256: "0".repeat(64), sourceByteSize: 9 } : {}) };
    expect(await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: `cleanup-${failurePoint}`, params },
      { agentId: "main", sessionId: "cleanup", workspaceDir })).toEqual({ block: true, blockReason: "workspace_file_unavailable" });
    expect(retry).toHaveBeenCalledTimes(1);
    expect(release).toHaveBeenCalledTimes(1);
    await workspaceStagingStore.retryFailedCleanups();
    expect(release).toHaveBeenCalledTimes(2);
    expect(await stagedFiles(stateDir)).toEqual([]);
  });
  it("shares the owned-artifact ledger across separately imported hook and tool bundles", async () => {
    directory = await mkdtemp(join(tmpdir(), "mg-cross-realm-stage-"));
    const stateDir = join(directory, "state");
    const workspaceDir = join(directory, "workspace");
    await mkdir(join(workspaceDir, "reports"), { recursive: true });
    await mkdir(stateDir, { recursive: true });
    process.env.OPENCLAW_STATE_DIR = stateDir;
    await writeFile(join(workspaceDir, "reports", "onepager.pdf"), "synthetic");
    const policy = graphPolicyFixture();
    delete policy.services.onedrive.allowed_roots[0].agents_instructions;
    policy.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
    const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
    const factories: Array<(context: any) => any> = [];
    const firstSpecifier = `./index.js?stage-hook-realm=1`;
    const secondSpecifier = `./index.js?stage-tool-realm=2`;
    const first = await import(firstSpecifier);
    const second = await import(secondSpecifier);
    const api = {
      pluginConfig: { enabled: true, policy },
      runtime: { state: { resolveStateDir: () => stateDir } },
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    };
    first.default.register({ ...api, on: (name: string, handler: (...args: any[]) => Promise<any>) => { hooks[name] = handler; }, registerTool: vi.fn() } as never);
    second.default.register({ ...api, on: vi.fn(), registerTool: (factory: (context: any) => any) => factories.push(factory) } as never);
    const context = { agentId: "main", sessionId: "cross-realm-stage", workspaceDir };
    const tool = factories.map((factory) => factory(context)).find((candidate) => candidate.name === "onedrive_upload");
    const approval = await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: "cross-realm-stage", params: {
      rootLabel: "synthetic_documents", relativePath: "onepager.pdf", sourceWorkspacePath: "reports/onepager.pdf",
    } }, context);
    await approval.requireApproval.onResolution("allow-once");
    expect((await stagedFiles(stateDir)).length).toBe(1);
    expect((await tool.execute("cross-realm-stage", { ...approval.params, relativePath: "modified.pdf" })).details)
      .toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
    expect(await stagedFiles(stateDir)).toEqual([]);

    const abandoned = await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: "cross-realm-abandoned", params: {
      rootLabel: "synthetic_documents", relativePath: "onepager.pdf", sourceWorkspacePath: "reports/onepager.pdf",
    } }, context);
    expect(abandoned.requireApproval).toBeDefined();
    expect((await stagedFiles(stateDir)).length).toBe(1);
    hooks.session_end({}, context);
    await vi.waitFor(async () => expect(await stagedFiles(stateDir)).toEqual([]));
    await abandoned.requireApproval.onResolution("allow-once");
    expect((await tool.execute("cross-realm-abandoned", abandoned.params)).details)
      .toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
  });

  it("stages a workspace PDF in the same upload call before approval, then binds its fingerprint", async () => {
    directory = await mkdtemp(join(tmpdir(), "mg-one-call-"));
    const stateDir = join(directory, "state");
    const workspaceDir = join(directory, "workspace");
    await mkdir(join(workspaceDir, "reports"), { recursive: true });
    await mkdir(stateDir, { recursive: true });
    process.env.OPENCLAW_STATE_DIR = stateDir;
    const bytes = Buffer.from("%PDF-1.7\nsynthetic onepager");
    await writeFile(join(workspaceDir, "reports", "onepager.pdf"), bytes);
    const policy = graphPolicyFixture();
    delete policy.services.onedrive.allowed_roots[0].agents_instructions;
    policy.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
    const factories = new Map<string, (context: unknown) => { name: string }>();
    const hooks: Record<string, (...args: unknown[]) => Promise<unknown>> = {};
    entry.register({
      pluginConfig: { enabled: true, policy },
      runtime: { state: { resolveStateDir: () => stateDir } },
      registerTool: (factory: (context: unknown) => { name: string }, options: { name: string }) => factories.set(options.name, factory),
      on: (name: string, handler: (...args: unknown[]) => Promise<unknown>) => { hooks[name] = handler; },
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    } as never);
    const context = { agentId: "main", sessionId: "one-call", workspaceDir };
    factories.get("onedrive_upload")!(context);
    const params = { rootLabel: "synthetic_documents", relativePath: "onepager.pdf", sourceWorkspacePath: "reports/onepager.pdf" };
    const denied = await hooks.before_tool_call(
      { toolName: "onedrive_upload", toolCallId: "denied-upload", params: { rootLabel: "synthetic_documents", relativePath: "onepager.pdf", sourceWorkspacePath: "reports/onepager.pdf" } },
      { agentId: "fixture-reader", sessionId: "one-call", workspaceDir },
    );
    expect(denied).toEqual({ block: true, blockReason: "access_denied" });
    expect(await readdir(join(stateDir, "media", "inbound")).catch(() => [])).toEqual([]);
    const ambiguous = await hooks.before_tool_call(
      { toolName: "onedrive_upload", toolCallId: "ambiguous-upload", params: { rootLabel: "synthetic_documents", relativePath: "onepager.pdf", sourceWorkspacePath: "reports/onepager.pdf", sourceMediaUri: "media://inbound/other.pdf" } },
      context,
    );
    expect(ambiguous).toEqual({ block: true, blockReason: "exactly_one_source_required" });
    const result = await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: "one-call-upload", params }, context) as {
      params: Record<string, unknown>;
      requireApproval: { onResolution: (decision: string) => Promise<void> };
    };
    expect(result).toMatchObject({ requireApproval: expect.anything() });
    expect(result.params).not.toHaveProperty("sourceWorkspacePath");
    expect(result.params).toMatchObject({
      rootLabel: "synthetic_documents",
      relativePath: "onepager.pdf",
      contentType: "application/pdf",
      sourceSha256: createHash("sha256").update(bytes).digest("hex"),
      sourceByteSize: bytes.byteLength,
    });
    expect(result.params.sourceMediaUri).toMatch(/^media:\/\/inbound\//);
    const source = await openProtectedMediaUploadSource(result.params.sourceMediaUri as string, workspaceDir);
    try {
      expect(await source.readChunk(0, bytes.length)).toEqual(bytes);
    } finally {
      await source.close();
    }
    await result.requireApproval.onResolution("deny");
    expect(await stagedFiles(stateDir)).toEqual([]);
  });
});
