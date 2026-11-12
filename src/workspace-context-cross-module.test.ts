import { createHash } from "node:crypto";
import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

let directory: string | undefined;
const previousStateDir = process.env.OPENCLAW_STATE_DIR;
afterEach(async () => {
  if (previousStateDir === undefined) delete process.env.OPENCLAW_STATE_DIR;
  else process.env.OPENCLAW_STATE_DIR = previousStateDir;
  if (directory) await rm(directory, { recursive: true, force: true });
  directory = undefined;
});

describe("trusted OneDrive workspace context across native module imports", () => {
  async function setup() {
    directory = await mkdtemp(join(tmpdir(), "mg-workspace-context-"));
    const stateDir = join(directory, "state");
    const workspaceDir = join(directory, "workspace");
    await mkdir(stateDir);
    await mkdir(workspaceDir);
    process.env.OPENCLAW_STATE_DIR = stateDir;
    const bytes = Buffer.from("synthetic cross-module workspace fixture\n");
    await writeFile(join(workspaceDir, "fixture.txt"), bytes);
    const firstSpecifier = `./index.js?workspace-context-hook=${directory}`;
    const secondSpecifier = `./index.js?workspace-context-factory=${directory}`;
    const hookModule = await import(firstSpecifier);
    const factoryModule = await import(secondSpecifier);
    expect(hookModule).not.toBe(factoryModule);
    const policy = graphPolicyFixture();
    delete policy.services.onedrive.allowed_roots[0].agents_instructions;
    policy.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
    policy.services.onedrive.allowed_roots[0].agents.other = structuredClone(policy.services.onedrive.allowed_roots[0].agents.main);
    const config = { enabled: true, policy };
    const hooks: Record<string, (...args: any[]) => any> = {};
    const factories = new Map<string, (context: any) => any>();
    const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
    const api = { pluginConfig: config, logger, runtime: { state: { resolveStateDir: () => stateDir } } };
    hookModule.default.register({ ...api, registerTool: vi.fn(), on: (name: string, handler: any) => { hooks[name] = handler; } });
    factoryModule.default.register({ ...api, on: vi.fn(), registerTool: (factory: any, options: any) => factories.set(options.name, factory) });
    const context = { agentId: "main", sessionId: directory!, workspaceDir };
    factories.get("onedrive_upload")!(context);
    const hookContext = { agentId: context.agentId, sessionId: context.sessionId };
    const params = { rootLabel: "synthetic_documents", relativePath: "fixture.txt", sourceWorkspacePath: "fixture.txt" };
    return { hookModule, factoryModule, hooks, factories, context, hookContext, params, bytes, config, logger, stateDir };
  }

  it("stages exact bytes when the hook has no workspace and only another module's factory has bound it", async () => {
    const { hooks, hookContext, params, bytes, context, factoryModule, logger, config } = await setup();
    const gate = await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: "cross-workspace", params }, hookContext);
    expect(gate.block).toBeUndefined();
    expect(gate.params.sourceSha256).toBe(createHash("sha256").update(bytes).digest("hex"));
    expect(gate.params.sourceByteSize).toBe(bytes.length);
    expect(gate.params.contentType).toBe("text/plain");
    await gate.requireApproval.onResolution("allow-once");
    const action = vi.fn(async (p: any) => {
      const source = await factoryModule.openProtectedMediaUploadSource(p.sourceMediaUri, context.workspaceDir);
      try { expect(await source.readChunk(0, bytes.length)).toEqual(bytes); }
      finally { await source.close(); }
      return { ok: true };
    });
    const tool = factoryModule.concrete("onedrive_upload", {}, context.agentId, context.sessionId, logger, action, config);
    expect((await tool.execute("cross-workspace", gate.params)).details.ok).toBe(true);
    expect(action).toHaveBeenCalledTimes(1);
    expect((await tool.execute("cross-workspace", gate.params)).details.error).toBe("approval_context_invalid_or_changed");
    await expect(factoryModule.openProtectedMediaUploadSource(gate.params.sourceMediaUri, context.workspaceDir)).rejects.toThrow();
    hooks.session_end({}, hookContext);
  });

  it("does not borrow another caller/session workspace or accept a caller-selected workspace", async () => {
    const { hooks, hookContext, params } = await setup();
    for (const context of [{ ...hookContext, sessionId: "unbound-session" }, { ...hookContext, agentId: "other" }]) {
      expect(await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: "unbound", params }, context))
        .toEqual({ block: true, blockReason: "workspace_context_unavailable" });
    }
    expect(await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: "forged", params: { ...params, workspaceDir: "/etc" } }, hookContext))
      .toEqual({ block: true, blockReason: "invalid_tool_parameters" });
    hooks.session_end({}, hookContext);
  });

  it("clears another module's workspace binding on session end", async () => {
    const { hooks, hookContext, params } = await setup();
    hooks.session_end({}, hookContext);
    expect(await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: "ended", params }, hookContext))
      .toEqual({ block: true, blockReason: "workspace_context_unavailable" });
  });

  it("retains the 64-context bound across module imports", async () => {
    const { hooks, factories, context, hookContext, params } = await setup();
    for (let i = 0; i < 64; i++) factories.get("onedrive_upload")!({ ...context, sessionId: `${context.sessionId}-${i}` });
    expect(await hooks.before_tool_call({ toolName: "onedrive_upload", toolCallId: "evicted", params }, hookContext))
      .toEqual({ block: true, blockReason: "workspace_context_unavailable" });
    for (let i = 0; i < 64; i++) hooks.session_end({}, { sessionId: `${context.sessionId}-${i}` });
  });
});
