import { createHash } from "node:crypto";
import { mkdtemp, mkdir, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import entry, { openProtectedMediaUploadSource } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

let directory: string | undefined;
const originalStateDir = process.env.OPENCLAW_STATE_DIR;
afterEach(async () => {
  if (originalStateDir === undefined) delete process.env.OPENCLAW_STATE_DIR;
  else process.env.OPENCLAW_STATE_DIR = originalStateDir;
  if (directory) await rm(directory, { recursive: true, force: true });
  directory = undefined;
});

describe("workspace-file OneDrive approval preflight", () => {
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
    expect((await readdir(join(stateDir, "media", "inbound"))).length).toBe(1);
    expect((await tool.execute("cross-realm-stage", { ...approval.params, relativePath: "modified.pdf" })).details)
      .toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
    expect(await readdir(join(stateDir, "media", "inbound"))).toEqual([]);

    const abandoned = await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: "cross-realm-abandoned", params: {
      rootLabel: "synthetic_documents", relativePath: "onepager.pdf", sourceWorkspacePath: "reports/onepager.pdf",
    } }, context);
    expect(abandoned.requireApproval).toBeDefined();
    expect((await readdir(join(stateDir, "media", "inbound"))).length).toBe(1);
    hooks.session_end({}, context);
    await vi.waitFor(async () => expect(await readdir(join(stateDir, "media", "inbound"))).toEqual([]));
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
    expect(await readdir(join(stateDir, "media", "inbound"))).toEqual([]);
  });
});
