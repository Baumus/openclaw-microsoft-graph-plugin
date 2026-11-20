import { mkdtemp, mkdir, writeFile, rm, symlink, link, truncate, readdir } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { afterEach, expect, it, vi } from "vitest";
vi.mock("./credential.js", () => ({
  tokenForAuthorizedOperation: vi.fn(async () => { throw new Error("synthetic_credential_gate"); }),
  readCredential: vi.fn(), selectScope: vi.fn(), exchangeRefreshToken: vi.fn(),
}));
import { tokenForAuthorizedOperation, readCredential } from "./credential.js";
import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
let dir: string;
const previous = process.env.OPENCLAW_STATE_DIR;
afterEach(async () => {
  vi.clearAllMocks(); vi.restoreAllMocks();
  if (previous === undefined) delete process.env.OPENCLAW_STATE_DIR; else process.env.OPENCLAW_STATE_DIR = previous;
  if (dir) await rm(dir, { recursive: true, force: true });
});
async function stagedFiles(stateDir: string): Promise<string[]> {
  const root = join(stateDir, "media");
  const entries = await readdir(root, { recursive: true, withFileTypes: true });
  return entries.filter(entry => entry.isFile() && !entry.name.startsWith(".")).map(entry => join(entry.parentPath, entry.name));
}
async function setup(warnings = true) {
  dir = await mkdtemp(join(tmpdir(), "mg-mail-workspace-"));
  const workspaceDir = join(dir, "workspace"), stateDir = join(dir, "state");
  await mkdir(workspaceDir); await mkdir(stateDir);
  process.env.OPENCLAW_STATE_DIR = stateDir;
  await writeFile(join(workspaceDir, "file.pdf"), "synthetic attachment");
  const hooks: Record<string, any> = {}, factories: any[] = [];
  const config = { enabled: true, warningApprovalsRequired: warnings, policy: graphPolicyFixture() };
  entry.register({ pluginConfig: config, runtime: { state: { resolveStateDir: () => stateDir } },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    on: (n: string, h: any) => hooks[n] = h, registerTool: (f: any) => factories.push(f),
  } as never);
  const ctx = { agentId: "main", sessionId: dir, workspaceDir };
  const tool = factories.map(f => f(ctx)).find(t => t.name === "outlook_mail_write");
  const original = { action: "add_attachment", messageId: "synthetic-draft", attachmentName: "file.pdf", sourceWorkspacePath: "file.pdf" };
  const fetch = vi.spyOn(globalThis, "fetch").mockRejectedValue(new Error("unexpected_network"));
  return { ctx, tool, original, hooks, workspaceDir, stateDir, fetch };
}
it.each([true, false])("preserves actual Host merge and reaches credential gate (approval=%s)", async warnings => {
  const x = await setup(warnings);
  const result = await x.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "mail-ok", params: x.original }, x.ctx);
  expect(result.block).not.toBe(true);
  expect(result.params.sourceWorkspacePath).toBe("file.pdf");
  expect(result.params).toEqual(x.original);
  if (result.requireApproval) await result.requireApproval.onResolution("allow-once");
  await x.tool.execute("mail-ok", { ...x.original, ...result.params });
  expect(tokenForAuthorizedOperation).toHaveBeenCalled();
  expect(x.fetch).not.toHaveBeenCalled();
});
it.each(["missing", "dual", "absolute", "traversal", "symlink", "hardlink", "caller"])("rejects %s before credentials/network", async kind => {
  const x = await setup();
  const params: any = { ...x.original };
  if (kind === "missing") delete params.sourceWorkspacePath;
  if (kind === "dual") params.attachmentMediaUri = "media://inbound/file.pdf";
  if (kind === "absolute") params.sourceWorkspacePath = join(x.workspaceDir, "file.pdf");
  if (kind === "traversal") params.sourceWorkspacePath = "../file.pdf";
  if (kind === "symlink") { await symlink("file.pdf", join(x.workspaceDir, "link.pdf")); params.sourceWorkspacePath = "link.pdf"; }
  if (kind === "hardlink") { await link(join(x.workspaceDir, "file.pdf"), join(x.workspaceDir, "link.pdf")); params.sourceWorkspacePath = "link.pdf"; }
  const result = await x.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "mail-bad", params }, kind === "caller" ? { ...x.ctx, agentId: "fixture-reader" } : x.ctx);
  expect(result.block).toBe(true);
  expect(tokenForAuthorizedOperation).not.toHaveBeenCalled(); expect(readCredential).not.toHaveBeenCalled(); expect(x.fetch).not.toHaveBeenCalled();
});
it("transfers approved snapshot bytes after original source edits and cleans the private copy", async () => {
  const x = await setup();
  const result = await x.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "mail-transfer", params: x.original }, x.ctx);
  await writeFile(join(x.workspaceDir, "file.pdf"), "new original bytes");
  await result.requireApproval.onResolution("allow-once");
  vi.mocked(tokenForAuthorizedOperation).mockResolvedValueOnce("synthetic-access-token");
  x.fetch.mockImplementationOnce(async (_url, init) => {
    const body = JSON.parse(String(init?.body));
    expect(Buffer.from(body.contentBytes, "base64").toString()).toBe("synthetic attachment");
    return new Response(JSON.stringify({ id: "synthetic-attachment", name: "file.pdf", size: 20 }));
  });
  const response = await x.tool.execute("mail-transfer", { ...x.original, ...result.params });
  expect(response.details.ok).toBe(true);
  expect(x.fetch).toHaveBeenCalledTimes(1);
  expect(await stagedFiles(x.stateDir)).toEqual([]);
});

it.each(["oversize", "no-workspace", "other-session"])("fails closed for %s before credentials", async kind => {
  const x = await setup();
  if (kind === "oversize") await truncate(join(x.workspaceDir, "file.pdf"), 64 * 1024 * 1024 + 1);
  const ctx = kind === "no-workspace" ? { agentId: "main", sessionId: "unbound" }
    : kind === "other-session" ? { ...x.ctx, sessionId: "unbound", workspaceDir: undefined } : x.ctx;
  const result = await x.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "mail-extra-bad", params: x.original }, ctx);
  expect(result.block).toBe(true);
  expect(tokenForAuthorizedOperation).not.toHaveBeenCalled(); expect(x.fetch).not.toHaveBeenCalled();
});

it("rejects staged mutation between staging and approval", async () => {
  const x = await setup();
  const result = await x.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "mail-before-approval", params: x.original }, x.ctx);
  await writeFile((await stagedFiles(x.stateDir))[0], "changed");
  await result.requireApproval.onResolution("allow-once");
  expect((await x.tool.execute("mail-before-approval", { ...x.original, ...result.params })).details.ok).toBe(false);
  expect(tokenForAuthorizedOperation).not.toHaveBeenCalled(); expect(x.fetch).not.toHaveBeenCalled();
});

it.each(["source-param", "staged-bytes"])("rejects changed %s after approval before credentials", async mutation => {
  const x = await setup();
  const result = await x.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "mail-changed", params: x.original }, x.ctx);
  expect(result.requireApproval).toBeDefined();
  await result.requireApproval.onResolution("allow-once");
  const merged = { ...x.original, ...result.params };
  if (mutation === "source-param") merged.sourceWorkspacePath = "other.pdf";
  else await writeFile((await stagedFiles(x.stateDir))[0], "changed staged content");
  const response = await x.tool.execute("mail-changed", merged);
  expect(response.details.ok).toBe(false);
  expect(tokenForAuthorizedOperation).not.toHaveBeenCalled(); expect(readCredential).not.toHaveBeenCalled(); expect(x.fetch).not.toHaveBeenCalled();
});

it.each([true, false])("keeps native attachment input byte-stable and executes original params (approval=%s)", async warnings => {
  const x = await setup(warnings);
  const first = await x.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: `native-mail-${warnings}`, params: x.original }, x.ctx);
  expect(first.block).not.toBe(true);
  expect(first.params).toEqual(x.original);
  if (first.requireApproval) await first.requireApproval.onResolution("allow-once");
  await rm(join(x.workspaceDir, "file.pdf"));
  const second = await x.hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: `native-mail-${warnings}`, params: x.original }, x.ctx);
  expect(second.block).not.toBe(true);
  expect(second.params).toEqual(x.original);
  if (second.requireApproval) await second.requireApproval.onResolution("allow-once");
  vi.mocked(tokenForAuthorizedOperation).mockResolvedValueOnce("synthetic-access-token");
  x.fetch.mockImplementationOnce(async (_url, init) => {
    expect(Buffer.from(JSON.parse(String(init?.body)).contentBytes, "base64").toString()).toBe("synthetic attachment");
    return Response.json({ id: "synthetic-attachment", name: "file.pdf", size: 20 });
  });
  expect((await x.tool.execute(`native-mail-${warnings}`, x.original)).details.ok).toBe(true);
  expect(x.fetch).toHaveBeenCalledTimes(1);
});
