import { mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import { readdirSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { accountScope, type AccountScope } from "./account-scope.js";
import { downloadOneDriveFile, downloadOutlookFileAttachment, publishPrivateMediaBytes } from "./attachment-download.js";
import { graphRequest } from "./graph.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

let stateDir: string;
let oldStateDir: string | undefined;
beforeEach(async () => {
  stateDir = await mkdtemp(join(tmpdir(), "graph-media-revocation-"));
  oldStateDir = process.env.OPENCLAW_STATE_DIR;
  process.env.OPENCLAW_STATE_DIR = stateDir;
});
afterEach(async () => {
  if (oldStateDir === undefined) delete process.env.OPENCLAW_STATE_DIR;
  else process.env.OPENCLAW_STATE_DIR = oldStateDir;
  await rm(stateDir, { recursive: true, force: true });
});

async function inboundNames() {
  try { return await readdir(join(stateDir, "media", "inbound")); }
  catch (error) { if ((error as NodeJS.ErrnoException).code === "ENOENT") return []; throw error; }
}
function tempBytes() {
  const dir = join(stateDir, "media", "inbound");
  try { return readdirSync(dir).filter(name => name.endsWith(".tmp")).reduce((n, name) => n + statSync(join(dir, name)).size, 0); }
  catch (error) { if ((error as NodeJS.ErrnoException).code === "ENOENT") return 0; throw error; }
}
function account() {
  let generation = 1;
  let revokeAtTempBytes = Number.POSITIVE_INFINITY;
  const assertCurrent = () => {
    if (tempBytes() >= revokeAtTempBytes) generation = 2;
    if (generation !== 1) throw new Error("graph_account_context_changed");
  };
  const scope: AccountScope = { identity: "synthetic-account:1", stateDir, policy: {}, assertCurrent, assertFresh: async () => assertCurrent() };
  return { scope, revoke: () => { generation = 2; }, revokeAtTemp: (bytes: number) => { revokeAtTempBytes = bytes; } };
}

const bytes = Buffer.from("synthetic private media bytes");
const root = graphPolicyFixture().services.onedrive.allowed_roots[0];
type Mode = "valid" | "after_response" | "midstream" | "commit";

function binaryBody(mode: Mode, reads: { count: number }) {
  if (mode !== "midstream") return new Response(bytes);
  const part = bytes.subarray(0, 8);
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      reads.count++;
      controller.enqueue(reads.count === 1 ? part : bytes.subarray(8));
      if (reads.count > 1) controller.close();
    },
  }, { highWaterMark: 0 });
  return new Response(stream);
}

describe.each(["onedrive", "outlook"] as const)("real SDK %s media revocation", provider => {
  async function run(mode: Mode) {
    const a = account();
    const reads = { count: 0 };
    if (mode === "midstream") a.revokeAtTemp(8);
    if (mode === "commit") a.revokeAtTemp(bytes.byteLength);
    const fetchFn = (async (input: RequestInfo | URL) => {
      const url = String(input);
      if (provider === "onedrive") {
        if (url.endsWith("/content")) { const response = binaryBody(mode, reads); if (mode === "after_response") a.revoke(); return response; }
        return Response.json({ id: "item-1", name: "private.txt", size: bytes.byteLength, file: { mimeType: "text/plain" } });
      }
      if (url.endsWith("/$value")) { const response = binaryBody(mode, reads); if (mode === "after_response") a.revoke(); return response; }
      return Response.json({ "@odata.type": "#microsoft.graph.fileAttachment", id: "attachment-1", name: "private.txt", contentType: "text/plain", size: bytes.byteLength });
    }) as typeof fetch;
    const action = provider === "onedrive"
      ? () => downloadOneDriveFile({ root, relativePath: "private.txt", token: "synthetic-token", fetchFn })
      : () => downloadOutlookFileAttachment({ owner: { kind: "message", messageId: "message-1" }, attachmentId: "attachment-1", token: "synthetic-token", maxBytes: 1024, readIdleTimeoutMs: 1000, fetchFn });
    return { result: accountScope.run(a.scope, async () => action()), reads };
  }
  for (const mode of ["after_response", "midstream", "commit"] as const) {
    it(`rolls back ${mode} revocation with no final or temp media`, async () => {
      const { result, reads } = await run(mode);
      await expect(result).rejects.toThrow("graph_account_context_changed");
      expect(await inboundNames()).toEqual([]);
      if (mode === "midstream") expect(reads.count).toBe(1);
    });
  }
  it("still publishes an authorized download", async () => {
    const { result } = await run("valid");
    const saved = await result;
    expect(saved.artifact.uri).toMatch(/^media:\/\/inbound\//);
    expect(await inboundNames()).toEqual([saved.artifact.id]);
    expect(await readFile(join(stateDir, "media", "inbound", saved.artifact.id))).toEqual(bytes);
  });
});

describe("real SDK To Do media revocation", () => {
  async function run(mode: "valid" | "after_response" | "commit") {
    const a = account();
    if (mode === "commit") a.revokeAtTemp(bytes.byteLength);
    const fetchFn = (async () => {
      const response = Response.json({ id: "attachment-1", name: "todo.txt", contentType: "text/plain", size: bytes.byteLength, contentBytes: bytes.toString("base64") });
      if (mode === "after_response") a.revoke();
      return response;
    }) as typeof fetch;
    return accountScope.run(a.scope, async () => {
      const item = await graphRequest("synthetic-token", "/me/todo/lists/list-1/tasks/task-1/attachments/attachment-1", {}, fetchFn);
      return publishPrivateMediaBytes(Buffer.from(item.contentBytes, "base64"), item.name, item.contentType, 1024);
    });
  }
  for (const mode of ["after_response", "commit"] as const) {
    it(`rolls back ${mode} revocation with no final or temp media`, async () => {
      await expect(run(mode)).rejects.toThrow("graph_account_context_changed");
      expect(await inboundNames()).toEqual([]);
    });
  }
  it("still publishes an authorized attachment", async () => {
    const saved = await run("valid");
    expect(saved.artifact.uri).toMatch(/^media:\/\/inbound\//);
    expect(await inboundNames()).toEqual([saved.artifact.id]);
    expect(await readFile(join(stateDir, "media", "inbound", saved.artifact.id))).toEqual(bytes);
  });
});
