import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("./credential.js", async () => ({
  ...await vi.importActual<typeof import("./credential.js")>("./credential.js"),
  tokenForAuthorizedOperation: vi.fn(async () => "synthetic-token"),
}));
import entry, { approvalInventory, attachmentWritePlan, classifyApproval } from "./index.js";
import { tokenForAuthorizedOperation } from "./credential.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
import { DIRECT_ATTACHMENT_MAX_BYTES, OUTLOOK_ATTACHMENT_MAX_BYTES } from "./graph.js";

const params = { action: "add_message_attachment", messageId: "draft+opaque=", sourceMessageId: "source%opaque=" };
const mime = Buffer.from("From: sender@example.test\r\nTo: recipient@example.test\r\nSubject: Synthetic original\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=synthetic\r\n\r\n--synthetic\r\nContent-Type: text/plain\r\n\r\nOriginal body\r\n--synthetic--\r\n");
function registered(policy = graphPolicyFixture(), warningApprovalsRequired = true, agentId = "main") {
  const factories: Array<(context: any) => any> = [];
  const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
  const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
  entry.register({ pluginConfig: { enabled: true, policy, warningApprovalsRequired, requestTimeoutMs: 1000 },
    registerTool: (factory: any) => factories.push(factory), on: (name: string, handler: any) => { hooks[name] = handler; }, logger } as any);
  const context = { agentId, sessionId: "synthetic-message-attachment" };
  const tool = factories.map((factory) => factory(context)).find((tool) => tool.name === "outlook_mail_write");
  const gate = (p = params, id = "call") => hooks.before_tool_call({ toolName: tool.name, toolCallId: id, params: p }, context);
  return { tool, gate, logger, async run(p: Record<string, unknown> = params) {
    const decision = await gate(p as typeof params);
    if (decision?.block) return { details: { ok: false, error: decision.blockReason } };
    await decision?.requireApproval?.onResolution("allow-once");
    return tool.execute("call", decision?.params ?? p);
  } };
}
function transport(options: { draft?: unknown; changedDraft?: boolean; bytes?: Buffer; sourceStatus?: number; postThrows?: boolean; echo?: boolean } = {}) {
  let reads = 0;
  return vi.spyOn(globalThis, "fetch").mockImplementation(async (input, init) => {
    const url = String(input);
    expect(new Headers(init?.headers).get("authorization")).toBe("Bearer synthetic-token");
    if (url.endsWith("/me/messages/draft%2Bopaque%3D?$select=id,isDraft")) {
      reads += 1;
      return Response.json({ id: params.messageId, isDraft: options.changedDraft && reads > 1 ? false : options.draft ?? true });
    }
    if (url.endsWith("/me/messages/source%25opaque%3D/$value")) {
      expect(new Headers(init?.headers).get("accept")).toBe("message/rfc822");
      return new Response(new Uint8Array(options.bytes ?? mime), { status: options.sourceStatus ?? 200 });
    }
    if (url.endsWith("/me/messages/draft%2Bopaque%3D/attachments")) {
      expect(init?.method).toBe("POST");
      if (options.postThrows) throw new TypeError("synthetic network failure");
      const body = JSON.parse(String(init?.body));
      expect(body).toMatchObject({ "@odata.type": "#microsoft.graph.fileAttachment", contentType: "message/rfc822", contentBytes: (options.bytes ?? mime).toString("base64") });
      return Response.json({ ...options.echo ? body : {}, id: "attachment-synthetic", size: 9999 }, { status: 201 });
    }
    throw new Error("unexpected synthetic route");
  });
}
beforeEach(() => { vi.mocked(tokenForAuthorizedOperation).mockClear(); });
afterEach(() => { vi.restoreAllMocks(); });

describe("attach original messages to existing Outlook drafts", () => {
  it("preserves exact original MIME bytes, encodes both IDs, minimizes echoed bytes and never sends", async () => {
    const fetchSpy = transport({ echo: true });
    const { run, logger } = registered();
    const response = await run({ ...params, attachmentName: "original.eml" });
    expect(response.details).toMatchObject({ ok: true, action: "add_message_attachment", attachment: { id: "attachment-synthetic", name: "original.eml", contentType: "message/rfc822", size: mime.length } });
    expect(fetchSpy).toHaveBeenCalledTimes(4);
    expect(JSON.stringify(response)).not.toContain("contentBytes");
    expect(JSON.stringify(response)).not.toContain(mime.toString("base64"));
    expect(JSON.stringify(logger.info.mock.calls)).not.toContain("Original body");
    expect(tokenForAuthorizedOperation).toHaveBeenCalledWith(expect.objectContaining({ allowedScopes: ["Mail.ReadWrite"] }));
  });
  it("discovers the action and retains warning classification", async () => {
    const { tool } = registered();
    expect(tool.parameters.properties.action.anyOf.map((a: any) => a.const)).toContain(params.action);
    expect(classifyApproval(tool.name, params)).toBe("warning");
    expect((await approvalInventory(graphPolicyFixture())).approvals).toContainEqual({ tool: tool.name, action: params.action, level: "warning" });
  });
  it.each(["read", "update"])("requires independent mail %s policy before approval, credentials or Graph", async (missing) => {
    const policy = graphPolicyFixture();
    policy.services.mail!.agents.main.operations = policy.services.mail!.agents.main.operations.filter((operation) => operation !== missing);
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    expect((await registered(policy).run()).details).toMatchObject({ ok: false, error: "access_denied" });
    expect(tokenForAuthorizedOperation).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
  });
  it.each([
    { sourceMessageId: undefined }, { sourceMessageId: "../source" }, { sourceMessageId: params.messageId },
    { messageId: "../draft" }, { attachmentName: "../mail.eml" }, { attachmentName: "original.txt" },
    { attachmentName: "original\r\n.eml" }, { attachmentMediaUri: "media://inbound/other" }, { bodyText: "ignored" },
  ])("rejects malformed or incompatible inputs before credentials and network: %j", async (changes) => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    expect((await registered().run({ ...params, ...changes })).details.ok).toBe(false);
    expect(tokenForAuthorizedOperation).not.toHaveBeenCalled();
    expect(fetchSpy).not.toHaveBeenCalled();
  });
  it.each([false, undefined, "true"])("fails closed unless the target is explicitly a draft: %s", async (draft) => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(Response.json({ id: params.messageId, isDraft: draft }));
    expect((await registered().run()).details).toMatchObject({ ok: false, error: "invalid_mail_draft_target" });
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });
  it("rejects target identity mismatch and a draft sent during the MIME read", async () => {
    const fetchSpy = transport({ changedDraft: true });
    expect((await registered().run()).details).toMatchObject({ ok: false, error: "invalid_mail_draft_target" });
    expect(fetchSpy).toHaveBeenCalledTimes(3);
    fetchSpy.mockResolvedValue(Response.json({ id: "different-target", isDraft: true }));
    expect((await registered().run()).details).toMatchObject({ ok: false, error: "invalid_mail_draft_target" });
    expect(fetchSpy).toHaveBeenCalledTimes(4);
  });
  it("rejects empty MIME without writing", async () => {
    const bytes = Buffer.alloc(0);
    const fetchSpy = transport({ bytes });
    const { run, logger } = registered();
    expect((await run()).details).toMatchObject({ ok: false, error: "invalid_provider_response", diagnosticStep: "source_mime_read", writeAttempted: false, mutationApplied: false, retrySafety: "safe_after_correction" });
    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(logger.info).toHaveBeenCalledWith(expect.stringContaining('"diagnostic_step":"source_mime_read"'));
  });
  it("supports the maximum direct MIME size", async () => {
    const fetchSpy = transport({ bytes: Buffer.alloc(DIRECT_ATTACHMENT_MAX_BYTES - 1, 65), echo: true });
    expect((await registered().run()).details).toMatchObject({ ok: true, attachment: { size: DIRECT_ATTACHMENT_MAX_BYTES - 1 } });
    expect(fetchSpy).toHaveBeenCalledTimes(4);
  });
  it("never writes after a missing source", async () => {
    const fetchSpy = transport({ sourceStatus: 404 });
    expect((await registered().run()).details.ok).toBe(false);
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });
  it("requires an approval snapshot and cannot reuse approval for another source or target", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const { gate, tool } = registered();
    expect((await tool.execute("direct", params)).details).toMatchObject({ error: "approval_context_invalid_or_changed" });
    for (const field of ["sourceMessageId", "messageId"]) {
      const decision = await gate();
      expect(decision.requireApproval.severity).toBe("warning");
      await decision.requireApproval.onResolution("allow-once");
      expect((await tool.execute("call", { ...decision.params, [field]: "different" })).details).toMatchObject({ error: "approval_context_invalid_or_changed" });
    }
    expect(fetchSpy).not.toHaveBeenCalled();
    expect(tokenForAuthorizedOperation).not.toHaveBeenCalled();
  });
  it("supports explicitly policy-disabled warning prompts without bypassing the snapshot", async () => {
    transport();
    expect((await registered(graphPolicyFixture(), false).run()).details.ok).toBe(true);
  });
  it("fails before Graph if Mail.ReadWrite consent is missing", async () => {
    vi.mocked(tokenForAuthorizedOperation).mockRejectedValueOnce(new Error("credential_scope_missing"));
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    expect((await registered().run()).details).toMatchObject({ ok: false, error: "credential_scope_missing", mutationApplied: false });
    expect(fetchSpy).not.toHaveBeenCalled();
  });
  it("requires readback after an accepted POST with an invalid receipt", async () => {
    const fetchSpy = transport();
    const original = fetchSpy.getMockImplementation()!;
    fetchSpy.mockImplementation(async (input, init) => init?.method === "POST" ? Response.json({}, { status: 201 }) : original(input, init));
    expect((await registered().run()).details).toMatchObject({ ok: false, error: "invalid_provider_response", mutationApplied: "unknown", retrySafety: "readback_before_retry" });
    expect(fetchSpy.mock.calls.filter(([, init]) => init?.method === "POST")).toHaveLength(1);
  });
  it("reports an uncertain POST without blindly retrying", async () => {
    const fetchSpy = transport({ postThrows: true });
    const response = await registered().run();
    expect(response.details).toMatchObject({ ok: false, mutationApplied: "unknown" });
    expect(fetchSpy.mock.calls.filter(([, init]) => init?.method === "POST")).toHaveLength(1);
  });
  it("accepts the provider maximum and rejects the next byte in the shared attachment planner", () => {
    // One lazy zero-filled backing allocation; no base64 encoding or 150 MB transport copy.
    const backing = Buffer.alloc(OUTLOOK_ATTACHMENT_MAX_BYTES + 1);
    const fields = { attachmentName: "original.eml", attachmentContentType: "message/rfc822" };
    expect(attachmentWritePlan(fields, "#microsoft.graph.fileAttachment", OUTLOOK_ATTACHMENT_MAX_BYTES, backing.subarray(0, OUTLOOK_ATTACHMENT_MAX_BYTES)))
      .toMatchObject({ mode: "session", size: OUTLOOK_ATTACHMENT_MAX_BYTES });
    expect(() => attachmentWritePlan(fields, "#microsoft.graph.fileAttachment", OUTLOOK_ATTACHMENT_MAX_BYTES, backing)).toThrow("file_too_large");
  });
  it("rejects a MIME Content-Length above the provider maximum before upload", async () => {
    const fetchSpy = transport();
    const original = fetchSpy.getMockImplementation()!;
    fetchSpy.mockImplementation(async (input, init) => String(input).endsWith("/$value")
      ? new Response("synthetic", { headers: { "content-length": String(OUTLOOK_ATTACHMENT_MAX_BYTES + 1) } }) : original(input, init));
    expect((await registered().run()).details).toMatchObject({ ok: false, error: "provider_response_too_large" });
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });
  it.each([DIRECT_ATTACHMENT_MAX_BYTES, 10 * 320 * 1024 + 37, 11 * 1024 * 1024 + 37])("uploads %i original MIME bytes through a safe session without credential leakage", async (size) => {
    const bytes = Buffer.alloc(size, 65);
    mime.copy(bytes);
    const fetchSpy = transport({ bytes });
    const original = fetchSpy.getMockImplementation()!;
    const chunks: Buffer[] = [];
    const uploadUrl = "https://outlook.office.com/api/v2.0/AttachmentSessions('synthetic')?token=synthetic-secret";
    let offset = 0;
    fetchSpy.mockImplementation(async (input, init) => {
      const url = String(input);
      if (url.endsWith("/attachments/createUploadSession")) {
        expect(JSON.parse(String(init?.body))).toEqual({ AttachmentItem: { attachmentType: "file", name: "original-message.eml", contentType: "message/rfc822", size } });
        expect(new Headers(init?.headers).get("authorization")).toBe("Bearer synthetic-token");
        return Response.json({ uploadUrl, nextExpectedRanges: ["0-"] }, { status: 201 });
      }
      if (url === uploadUrl) {
        expect(init?.method).toBe("PUT");
        expect(new Headers(init?.headers).has("authorization")).toBe(false);
        expect(init?.redirect).toBe("error");
        const chunk = Buffer.from(init?.body as Uint8Array);
        expect(new Headers(init?.headers).get("content-range")).toBe(`bytes ${offset}-${offset + chunk.length - 1}/${size}`);
        chunks.push(chunk);
        offset += chunk.length;
        return offset === size ? new Response(null, { status: 201 }) : Response.json({ nextExpectedRanges: [`${offset}-`] });
      }
      return original(input, init);
    });
    const { run, logger } = registered();
    const response = await run();
    expect(response.details).toMatchObject({ ok: true, upload_mode: "session", bytes: size, attachment: { name: "original-message.eml", contentType: "message/rfc822", size } });
    expect(Buffer.concat(chunks).equals(bytes)).toBe(true);
    expect(JSON.stringify(response)).not.toContain("synthetic-secret");
    expect(JSON.stringify(response)).not.toContain("contentBytes");
    expect(JSON.stringify(logger.info.mock.calls)).not.toContain("synthetic-token");
    expect(JSON.stringify(logger.info.mock.calls)).not.toContain("synthetic-secret");
  });
  it.each(["unsafe-session", "uncertain-chunk"])("fails closed on %s without retrying attachment writes", async (failure) => {
    const fetchSpy = transport({ bytes: Buffer.alloc(DIRECT_ATTACHMENT_MAX_BYTES, 65) });
    const original = fetchSpy.getMockImplementation()!;
    const uploadUrl = "https://outlook.office.com/api/v2.0/AttachmentSessions('synthetic')";
    fetchSpy.mockImplementation(async (input, init) => {
      if (String(input).endsWith("/attachments/createUploadSession")) return Response.json({ uploadUrl: failure === "unsafe-session" ? "https://attacker.example.test/upload" : uploadUrl, nextExpectedRanges: ["0-"] });
      if (String(input) === uploadUrl) {
        expect(new Headers(init?.headers).has("authorization")).toBe(false);
        if (init?.method === "DELETE") return new Response(null, { status: 204 });
        throw new TypeError("synthetic interrupted chunk");
      }
      return original(input, init);
    });
    const response = await registered().run();
    expect(response.details).toMatchObject({ ok: false, mutationApplied: "unknown", retrySafety: "readback_before_retry" });
    expect(response.details).toMatchObject({ diagnosticStep: failure === "unsafe-session" ? "upload_session_validation" : "upload_chunk", writeAttempted: true });
    expect(fetchSpy.mock.calls.filter(([input]) => String(input).includes("attacker.example.test"))).toHaveLength(0);
    expect(fetchSpy.mock.calls.filter(([, init]) => init?.method === "POST")).toHaveLength(1);
    expect(fetchSpy.mock.calls.filter(([, init]) => init?.method === "PUT")).toHaveLength(failure === "unsafe-session" ? 0 : 1);
  });
  it("does not confuse the per-request timeout with the whole transfer deadline", async () => {
    const fetchSpy = transport();
    const original = fetchSpy.getMockImplementation()!;
    fetchSpy.mockImplementation(async (input, init) => {
      if (!init?.method || init.method === "GET") await new Promise((resolve) => setTimeout(resolve, 400));
      return original(input, init);
    });
    // Three independently bounded 400 ms reads exceed the configured 1,000 ms request timeout overall.
    expect((await registered().run()).details.ok).toBe(true);
  });

});
