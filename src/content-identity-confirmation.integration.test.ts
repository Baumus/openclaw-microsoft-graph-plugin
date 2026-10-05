import { createHash } from "node:crypto";
import { link, mkdir, mkdtemp, readdir, rename, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("./credential.js", async () => {
  const actual = await vi.importActual<typeof import("./credential.js")>("./credential.js");
  const readCredential = vi.fn(async () => ({ clientId: "synthetic", refreshToken: "synthetic", tenant: "common", scopes: ["Files.ReadWrite"] }));
  const exchangeRefreshToken = vi.fn(async () => "synthetic-token");
  return {
    ...actual,
    readCredential,
    exchangeRefreshToken,
    tokenForAuthorizedOperation: vi.fn(async () => { await readCredential(); return exchangeRefreshToken(); }),
  };
});

import { exchangeRefreshToken, readCredential } from "./credential.js";
import entry, { beforeMicrosoftGraphToolCall } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
import { OneDriveAgentsSessionCache } from "./onedrive-agents-instructions.js";

let workspaceDir = "";
const originalStateDir = process.env.OPENCLAW_STATE_DIR;

function digest(bytes: Uint8Array) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function consumeBody(body: unknown): Promise<void> {
  if (body && typeof (body as any)[Symbol.asyncIterator] === "function") {
    for await (const _chunk of body as AsyncIterable<Uint8Array>) { /* consume upload stream */ }
  }
}

function runtime(warningApprovalsRequired: boolean, managedRoot = false) {
  const policy = graphPolicyFixture();
  if (!managedRoot) delete policy.services.onedrive.allowed_roots[0].agents_instructions;
  policy.services.onedrive.allowed_roots[0].agents.main.permissions.read = true;
  const hooks: Record<string, (...args: any[]) => Promise<any> | any> = {};
  const factories: Array<(context: any) => any> = [];
  entry.register({
    pluginConfig: { enabled: true, warningApprovalsRequired, policy },
    runtime: { state: { resolveStateDir: () => workspaceDir } },
    registerTool: (factory: any) => factories.push(factory),
    on: (name: string, handler: any) => { hooks[name] = handler; },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  const toolContext = { agentId: "main", sessionId: "approval-session", workspaceDir };
  const tools = Object.fromEntries(factories.map((factory) => {
    const tool = factory(toolContext);
    return [tool.name, tool];
  }));
  return { hooks, tools, context: toolContext };
}

function graphSuccess(bytes: Uint8Array, update = false) {
  return vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
    if (update && init?.method !== "PUT") return Response.json({ eTag: "before", file: {} });
    await consumeBody(init?.body);
    return Response.json({ id: "item-1", name: "SYNTHETIC_RECORD.pdf", size: bytes.byteLength, file: { mimeType: "application/pdf" } }, { status: update ? 200 : 201 });
  });
}

beforeEach(async () => {
  workspaceDir = await mkdtemp(join(tmpdir(), "msgraph-content-identity-"));
  process.env.OPENCLAW_STATE_DIR = workspaceDir;
  await mkdir(join(workspaceDir, "media", "inbound"), { recursive: true });
  vi.mocked(readCredential).mockClear();
  vi.mocked(exchangeRefreshToken).mockClear();
});

afterEach(async () => {
  vi.restoreAllMocks();
  if (originalStateDir === undefined) delete process.env.OPENCLAW_STATE_DIR;
  else process.env.OPENCLAW_STATE_DIR = originalStateDir;
  await rm(workspaceDir, { recursive: true, force: true });
});

describe.each([
  ["onedrive_upload", false],
  ["onedrive_update", true],
] as const)("content identity preconditions for %s", (toolName, update) => {
  it("reclaims owned workspace copies on denial, cancellation, timeout, fingerprint mismatch and both execution outcomes", async () => {
    const bytes = Buffer.from("%PDF-1.7\nsynthetic staging");
    await mkdir(join(workspaceDir, "reports"));
    await writeFile(join(workspaceDir, "reports", "onepager.pdf"), bytes);
    const { hooks, tools, context } = runtime(true);
    const inbound = join(workspaceDir, "media", "inbound");
    const params = { rootLabel: "synthetic_documents", relativePath: "SYNTHETIC_RECORD.pdf", sourceWorkspacePath: "reports/onepager.pdf" };
    const preflight = (callId: string, input: Record<string, unknown> = params) => hooks.before_tool_call({ toolName, toolCallId: callId, params: input }, context);
    const mediaFiles = async () => {
      const entries = await readdir(inbound);
      const namespace = "baumus-msgraph-workspace-staging";
      if (!entries.includes(namespace)) return entries;
      const runs = await readdir(join(inbound, namespace));
      const staged = (await Promise.all(runs.map(async (run) => readdir(join(inbound, namespace, run))))).flat().filter((file) => file !== ".workspace-staging-owner");
      return [...entries.filter((entry) => entry !== namespace), ...staged];
    };
    const originalFiles = await mediaFiles();

    const mismatch = await preflight("stage-mismatch", { ...params, sourceSha256: "0".repeat(64), sourceByteSize: bytes.length });
    expect(mismatch).toEqual({ block: true, blockReason: "invalid_source_fingerprint" });
    expect(await mediaFiles()).toEqual(originalFiles);

    for (const decision of ["deny", "cancelled", "timeout"]) {
      const approval = await preflight(`stage-${decision}`);
      expect((await mediaFiles()).length).toBe(originalFiles.length + 1);
      await approval.requireApproval.onResolution(decision);
      expect(await mediaFiles()).toEqual(originalFiles);
    }

    const missingId = await hooks.before_tool_call({ toolName, params }, context);
    expect(missingId).toEqual({ block: true, blockReason: "approval_context_tool_call_id_required" });
    expect(await mediaFiles()).toEqual(originalFiles);

    const changed = await preflight("stage-changed");
    await changed.requireApproval.onResolution("allow-once");
    expect((await tools[toolName].execute("stage-changed", { ...changed.params, relativePath: "changed.pdf" })).details)
      .toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
    expect(await mediaFiles()).toEqual(originalFiles);

    const failed = await preflight("stage-failed");
    await failed.requireApproval.onResolution("allow-once");
    vi.spyOn(globalThis, "fetch").mockRejectedValueOnce(new Error("synthetic_provider_failure"));
    expect((await tools[toolName].execute("stage-failed", failed.params)).details).toMatchObject({ ok: false });
    expect(await mediaFiles()).toEqual(originalFiles);
    vi.restoreAllMocks();

    const approved = await preflight("stage-success");
    await approved.requireApproval.onResolution("allow-once");
    graphSuccess(bytes, update);
    expect((await tools[toolName].execute("stage-success", approved.params)).details).toMatchObject({ ok: true });
    expect(await mediaFiles()).toEqual(originalFiles);

    await writeFile(join(inbound, "existing.pdf"), bytes);
    const existing = await preflight("existing-media", {
      rootLabel: params.rootLabel, relativePath: params.relativePath, sourceMediaUri: "media://inbound/existing.pdf",
    });
    await existing.requireApproval.onResolution("deny");
    expect(await mediaFiles()).toEqual([...originalFiles, "existing.pdf"]);
  });

  it.each([
    {
      name: "missing",
      prepare: async (_sourcePath: string, _bytes: Buffer) => undefined,
      expected: "invalid_source_media_uri",
    },
    {
      name: "replaced",
      prepare: async (sourcePath: string, bytes: Buffer) => {
        await writeFile(sourcePath, bytes);
        await rename(sourcePath, `${sourcePath}.approved`);
        await writeFile(sourcePath, Buffer.alloc(bytes.byteLength, 0x72));
      },
      expected: "invalid_source_fingerprint",
    },
    {
      name: "symlinked",
      prepare: async (sourcePath: string, bytes: Buffer) => {
        const targetPath = `${sourcePath}.target`;
        await writeFile(targetPath, bytes);
        await symlink(targetPath, sourcePath);
      },
      expected: "invalid_source_media_uri",
    },
    {
      name: "hardlinked",
      prepare: async (sourcePath: string, bytes: Buffer) => {
        const targetPath = `${sourcePath}.target`;
        await writeFile(targetPath, bytes);
        await link(targetPath, sourcePath);
      },
      expected: "invalid_source_media_uri",
    },
    {
      name: "SHA-256-mismatched",
      prepare: async (sourcePath: string, bytes: Buffer) => writeFile(sourcePath, bytes),
      expected: "invalid_source_fingerprint",
      claimedSha256: "0".repeat(64),
    },
    {
      name: "byte-size-mismatched",
      prepare: async (sourcePath: string, bytes: Buffer) => writeFile(sourcePath, bytes),
      expected: "invalid_source_fingerprint",
      claimedByteSizeDelta: 1,
    },
    {
      name: "uppercase-SHA-256",
      prepare: async (sourcePath: string, bytes: Buffer) => writeFile(sourcePath, bytes),
      expected: "invalid_source_fingerprint",
      uppercaseSha256: true,
    },
  ])("rejects a cold managed-root $name artifact before instruction or provider boundaries", async ({ prepare, expected, claimedSha256, claimedByteSizeDelta = 0, uppercaseSha256 = false }) => {
    const bytes = Buffer.from(`cold-managed-${toolName}`);
    const sourcePath = join(workspaceDir, "media", "inbound", "cold.pdf");
    await prepare(sourcePath, bytes);
    const { context } = runtime(true, true);
    const credentialReader = vi.fn(async () => ({ clientId: "synthetic", refreshToken: "synthetic", tenant: "common", scopes: ["Files.Read"] }));
    const tokenExchange = vi.fn(async () => "synthetic-token");
    const candidateReader = vi.fn(async () => null);
    const graph = vi.spyOn(globalThis, "fetch").mockRejectedValue(new Error("graph_must_not_be_called"));
    const sourceSha256 = claimedSha256 ?? (uppercaseSha256 ? digest(bytes).toUpperCase() : digest(bytes));
    const params = {
      rootLabel: "synthetic_documents",
      relativePath: "SYNTHETIC_FOLDER/COLD.pdf",
      sourceMediaUri: "media://inbound/cold.pdf",
      sourceSha256,
      sourceByteSize: bytes.byteLength + claimedByteSizeDelta,
      contentType: "application/pdf",
    };

    expect(await beforeMicrosoftGraphToolCall(
      { enabled: true, warningApprovalsRequired: true, policy: graphPolicyFixture() },
      { toolName, params },
      context,
      { credentialReader, tokenExchange, candidateReader, cache: new OneDriveAgentsSessionCache() },
    )).toEqual({ block: true, blockReason: expected });
    expect(credentialReader).not.toHaveBeenCalled();
    expect(tokenExchange).not.toHaveBeenCalled();
    expect(candidateReader).not.toHaveBeenCalled();
    expect(graph).not.toHaveBeenCalled();
  });

  it("derives and binds the fingerprint without agent-supplied fields", async () => {
    const bytes = Buffer.from(`native-fingerprint-${toolName}`);
    await writeFile(join(workspaceDir, "media", "inbound", "auto.pdf"), bytes);
    const { hooks, tools, context } = runtime(true);
    const params = {
      rootLabel: "synthetic_documents",
      relativePath: "SYNTHETIC_FOLDER/AUTO.pdf",
      sourceMediaUri: "media://inbound/auto.pdf",
      contentType: "application/pdf",
    };
    const approval = await hooks.before_tool_call({ toolName, toolCallId: "auto", params }, context);
    expect(approval.params).toEqual({ ...params, sourceSha256: digest(bytes), sourceByteSize: bytes.byteLength });
    expect(approval.requireApproval.description).toContain(digest(bytes));
    expect(readCredential).not.toHaveBeenCalled();
    approval.requireApproval.onResolution("allow-once");
    const graph = graphSuccess(bytes, update);
    expect((await tools[toolName].execute("auto", approval.params)).details).toMatchObject({ ok: true, source_sha256: digest(bytes), source_byte_size: bytes.byteLength });
    expect(graph).toHaveBeenCalled();
  });

  it("validates protected media and fingerprints before credentials when approvals are disabled", async () => {
    const bytes = Buffer.from(`synthetic-record-${toolName}`);
    await writeFile(join(workspaceDir, "media", "inbound", "wrong.pdf"), Buffer.alloc(bytes.byteLength, 0x7a));
    const { hooks, tools, context } = runtime(false);
    const params = {
      rootLabel: "synthetic_documents",
      relativePath: "SYNTHETIC_FOLDER/SYNTHETIC_RECORD.pdf",
      sourceMediaUri: "media://inbound/wrong.pdf",
      sourceSha256: digest(bytes),
      sourceByteSize: bytes.byteLength,
      contentType: "application/pdf",
      chatConfirmed: true,
      chatConfirmationToken: `mgw1_${"A".repeat(43)}`,
    };

    expect(await hooks.before_tool_call({ toolName, params }, context)).toEqual({ block: true, blockReason: "invalid_source_fingerprint" });
    expect((await tools[toolName].execute("wrong", params)).details).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed" });
    expect(readCredential).not.toHaveBeenCalled();
    expect(exchangeRefreshToken).not.toHaveBeenCalled();

    const invalid = await hooks.before_tool_call({ ...{ toolName }, params: { ...params, sourceMediaUri: "file:///tmp/private" } }, context);
    expect(invalid).toEqual({ block: true, blockReason: "invalid_source_media_uri" });
    expect(readCredential).not.toHaveBeenCalled();
  });

  it("fails closed when the approved pathname is replaced before execution", async () => {
    const approvedBytes = Buffer.from(`approved-race-${toolName}`);
    const replacementBytes = Buffer.alloc(approvedBytes.byteLength, 0x72);
    const sourcePath = join(workspaceDir, "media", "inbound", "race.pdf");
    await writeFile(sourcePath, approvedBytes);
    const { hooks, tools, context } = runtime(true);
    const params = {
      rootLabel: "synthetic_documents",
      relativePath: "SYNTHETIC_FOLDER/SYNTHETIC_RACE.pdf",
      sourceMediaUri: "media://inbound/race.pdf",
      sourceSha256: digest(approvedBytes),
      sourceByteSize: approvedBytes.byteLength,
      contentType: "application/pdf",
    };

    const approval = await hooks.before_tool_call({ toolName, toolCallId: "replaced", params }, context);
    expect(approval.requireApproval.description).toContain(digest(approvedBytes));
    approval.requireApproval.onResolution("allow-once");
    await rename(sourcePath, `${sourcePath}.approved`);
    await writeFile(sourcePath, replacementBytes);

    expect((await tools[toolName].execute("replaced", params)).details).toMatchObject({ ok: false, error: "invalid_source_fingerprint" });
    expect(readCredential).not.toHaveBeenCalled();
    expect(exchangeRefreshToken).not.toHaveBeenCalled();
  });

  it("treats legacy chat fields as inert while preserving authorized execution", async () => {
    const bytes = Buffer.from(`approved-${toolName}`);
    await writeFile(join(workspaceDir, "media", "inbound", "approved.pdf"), bytes);
    const { hooks, tools, context } = runtime(false);
    const params = {
      rootLabel: "synthetic_documents",
      relativePath: "SYNTHETIC_FOLDER/SYNTHETIC_APPROVED.pdf",
      sourceMediaUri: "media://inbound/approved.pdf",
      sourceSha256: digest(bytes),
      sourceByteSize: bytes.byteLength,
      contentType: "application/pdf",
      chatConfirmed: false,
      chatConfirmationToken: `mgw1_${"A".repeat(43)}`,
    };
    expect(await hooks.before_tool_call({ toolName, toolCallId: "write", params }, context)).toEqual({ params });
    const fetchSpy = graphSuccess(bytes, update);
    expect((await tools[toolName].execute("write", params)).details).toMatchObject({ ok: true, source_sha256: digest(bytes), source_byte_size: bytes.byteLength });
    expect(readCredential).toHaveBeenCalledTimes(1);
    expect(fetchSpy).toHaveBeenCalledTimes(update ? 2 : 1);
  });

  it("never lets legacy chat fields bypass required native approval", async () => {
    const bytes = Buffer.from("n");
    await writeFile(join(workspaceDir, "media", "inbound", "native.pdf"), bytes);
    const { hooks, context } = runtime(true);
    const params = {
      rootLabel: "synthetic_documents",
      relativePath: "SYNTHETIC_FOLDER/SYNTHETIC_NATIVE.pdf",
      sourceMediaUri: "media://inbound/native.pdf",
      sourceSha256: digest(bytes),
      sourceByteSize: bytes.byteLength,
      chatConfirmed: true,
      chatConfirmationToken: `mgw1_${"A".repeat(43)}`,
    };
    const result = await hooks.before_tool_call({ toolName, toolCallId: `legacy-native-approval-${toolName}`, params }, context);
    expect(result.requireApproval).toMatchObject({ severity: "warning", allowedDecisions: ["allow-once", "allow-always", "deny"] });
    expect(readCredential).not.toHaveBeenCalled();
  });
});
