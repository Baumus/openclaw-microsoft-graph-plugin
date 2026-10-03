import { createHash, randomBytes } from "node:crypto";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({ tokenForAuthorizedOperation: vi.fn() }));
vi.mock("./credential.js", async (importOriginal) => ({
  ...await importOriginal<typeof import("./credential.js")>(),
  tokenForAuthorizedOperation: mocks.tokenForAuthorizedOperation,
}));

import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function key(): string { return randomBytes(32).toString("base64url"); }
function vaultPolicy() {
  const current = graphPolicyFixture();
  return { version: 2 as const, rules: current.rules, services: current.services };
}
function registeredTools(agentId: string) {
  const factories: Array<(context: any) => any> = [];
  entry.register({
    pluginConfig: { enabled: true, credentialVaultKey: key(), policy: vaultPolicy() },
    runtime: { state: { resolveStateDir: () => "/synthetic-state-never-read-when-denied" } },
    registerTool: (factory: any) => factories.push(factory), on: vi.fn(),
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  return Object.fromEntries(factories.map((factory) => {
    const tool = factory({ agentId });
    return [tool.name, tool];
  }));
}

function registeredRuntime(agentId: string, workspaceDir?: string) {
  const factories: Array<(context: any) => any> = [];
  const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
  entry.register({
    pluginConfig: { enabled: true, credentialVaultKey: key(), policy: vaultPolicy() },
    runtime: { state: { resolveStateDir: () => "/synthetic-state-never-read-when-denied" } },
    registerTool: (factory: any) => factories.push(factory), on: (name: string, handler: any) => { hooks[name] = handler; },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  return { hooks, tools: Object.fromEntries(factories.map((factory) => { const tool = factory({ agentId, sessionId: "denied-session", workspaceDir }); return [tool.name, tool]; })) };
}

beforeEach(() => mocks.tokenForAuthorizedOperation.mockReset());

describe("vault authorization ordering", () => {
  it("denies before plugin-side key selection, vault access, decryption, or network", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const response = await registeredTools("unauthorized-agent").outlook_calendar_read.execute("denied", { action: "list_calendars" });
    expect(response.details).toMatchObject({ ok: false, error: "access_denied" });
    expect(mocks.tokenForAuthorizedOperation).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
    fetchSpy.mockRestore();
  });

  it("reaches vault credential access only after the exact operation is authorized", async () => {
    mocks.tokenForAuthorizedOperation.mockRejectedValueOnce(new Error("credential_vault_unavailable"));
    const response = await registeredTools("main").outlook_calendar_read.execute("authorized", { action: "list_calendars" });
    expect(response.details).toMatchObject({ ok: false, error: "credential_vault_unavailable" });
    expect(mocks.tokenForAuthorizedOperation).toHaveBeenCalledOnce();
    expect(mocks.tokenForAuthorizedOperation.mock.calls[0][0]).toMatchObject({ policy: { version: 2 } });
  });

  it.each([
    ["onedrive_upload", { rootLabel: "synthetic_documents", relativePath: "new.txt", sourceMediaUri: "media://inbound/new.txt" }],
    ["onedrive_update", { rootLabel: "synthetic_documents", relativePath: "existing.txt", sourceMediaUri: "media://inbound/existing.txt" }],
    ["onedrive_metadata_update", { rootLabel: "synthetic_documents", relativePath: "existing.txt", name: "renamed.txt" }],
    ["onedrive_create_folder", { rootLabel: "synthetic_documents", parentRelativePath: "", name: "new-folder" }],
    ["onedrive_delete", { rootLabel: "synthetic_documents", relativePath: "existing.txt" }],
  ])("denies %s before instruction credential access or Graph", async (toolName, params) => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const { hooks } = registeredRuntime("unauthorized-agent");
    await expect(hooks.before_tool_call({ toolName, params }, { agentId: "unauthorized-agent", sessionId: "denied-session", requester: { senderIsOwner: true, channel: "synthetic-channel" } }))
      .resolves.toEqual({ block: true, blockReason: "access_denied" });
    expect(mocks.tokenForAuthorizedOperation).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
    fetchSpy.mockRestore();
  });

  it("requires separate read authority for instruction discovery after mutation authorization", async () => {
    const workspaceDir = await mkdtemp(join(tmpdir(), "microsoft-graph-vault-ordering-"));
    await mkdir(join(workspaceDir, "media", "inbound"), { recursive: true });
    const bytes = Buffer.from("n");
    await writeFile(join(workspaceDir, "media", "inbound", "new.txt"), bytes);
    const { hooks } = registeredRuntime("main", workspaceDir);
    const event = {
      toolName: "onedrive_upload",
      toolCallId: "instruction-read-authority",
      params: { rootLabel: "synthetic_documents", relativePath: "new.txt", sourceMediaUri: "media://inbound/new.txt", sourceSha256: createHash("sha256").update(bytes).digest("hex"), sourceByteSize: bytes.byteLength },
    };
    const context = { agentId: "main", sessionId: "denied-session", requester: { senderIsOwner: true, channel: "synthetic-channel" } };
    await expect(hooks.before_tool_call(event, context)).resolves.toEqual({ block: true, blockReason: "access_denied" });
    expect(mocks.tokenForAuthorizedOperation).not.toHaveBeenCalled();
    await rm(workspaceDir, { recursive: true, force: true });
  });
});
