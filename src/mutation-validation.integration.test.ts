import { describe, expect, it, vi } from "vitest";

const credential = vi.hoisted(() => ({
  readCredential: vi.fn(async () => { throw new Error("credential_must_not_be_read"); }),
  exchangeRefreshToken: vi.fn(async () => { throw new Error("oauth_must_not_run"); }),
  tokenForAuthorizedOperation: vi.fn(async () => { throw new Error("credential_must_not_be_read"); }),
}));

vi.mock("./credential.js", async () => {
  const actual = await vi.importActual<typeof import("./credential.js")>("./credential.js");
  return { ...actual, ...credential };
});

import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function registeredTools() {
  const factories: Array<(context: any) => any> = [];
  const hooks: Record<string, (...args: any[]) => Promise<any> | any> = {};
  const policy = graphPolicyFixture();
  const root = policy.services.onedrive.allowed_roots[0];
  delete root.agents_instructions;
  root.permissions.delete = true;
  root.agents.main.permissions = { read: true, write: true, delete: true };
  entry.register({
    pluginConfig: { enabled: true, policy },
    registerTool: (factory: any) => factories.push(factory),
    on: (name: string, handler: any) => { hooks[name] = handler; },
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  const context = { agentId: "main", sessionId: "mutation-validation" };
  return Object.fromEntries(factories.map((factory) => {
    const tool = factory(context);
    return [tool.name, { ...tool, async execute(toolCallId: string, params: Record<string, unknown>, signal?: AbortSignal) {
      const gate = await hooks.before_tool_call({ toolName: tool.name, toolCallId, params }, context);
      if (gate?.block) return { details: { ok: false, error: gate.blockReason } };
      gate?.requireApproval?.onResolution("allow-once");
      return tool.execute(toolCallId, gate?.params ?? params, signal);
    } }];
  }));
}

const malformedMutations: Array<{ tool: string; action: string; params: Record<string, unknown>; error: string }> = [
  { tool: "onedrive_upload", action: "upload", params: { rootLabel: "synthetic_documents", relativePath: "a.bin", sourceMediaUri: "/tmp/a.bin" }, error: "invalid_source_media_uri" },
  { tool: "onedrive_upload", action: "upload_unsupported_type", params: { rootLabel: "synthetic_documents", relativePath: "a.bin", sourceMediaUri: "media/inbound/a.bin", contentType: "application/x-unsupported" }, error: "invalid_source_media_uri" },
  { tool: "onedrive_update", action: "update", params: { rootLabel: "synthetic_documents", relativePath: "a.bin", sourceMediaUri: "../a.bin" }, error: "invalid_source_media_uri" },
  { tool: "onedrive_metadata_update", action: "metadata_update", params: { rootLabel: "synthetic_documents", relativePath: "a.bin" }, error: "invalid_drive_metadata" },
  { tool: "onedrive_metadata_update", action: "metadata_update", params: { rootLabel: "synthetic_documents", relativePath: "a.bin", fileSystemInfo: { lastModifiedDateTime: "2026-02-30T12:00:00Z" } }, error: "invalid_datetime" },
  { tool: "onedrive_create_folder", action: "create_folder", params: { rootLabel: "synthetic_documents", parentRelativePath: "", name: "bad/name" }, error: "invalid_drive_name" },
  { tool: "onedrive_delete", action: "delete", params: { rootLabel: "synthetic_documents", relativePath: "../outside" }, error: "invalid_relative_path" },

  { tool: "outlook_calendar_write", action: "create", params: { action: "create", subject: "Missing end", startDateTime: "2026-09-08T09:00:00" }, error: "invalid_event_payload" },
  { tool: "outlook_calendar_write", action: "update", params: { action: "update", eventId: "event" }, error: "invalid_event_payload" },
  { tool: "outlook_calendar_write", action: "multiwrite", params: { action: "multiwrite", operations: [
    { operationId: "duplicate", kind: "update", eventId: "event-1", subject: "One" },
    { operationId: "duplicate", kind: "update", eventId: "event-2", subject: "Two" },
  ] }, error: "invalid_operation_id" },
  { tool: "outlook_calendar_write", action: "respond", params: { action: "respond", eventId: "event" }, error: "invalid_response" },
  { tool: "outlook_calendar_write", action: "attach", params: { action: "attach", eventId: "event" }, error: "invalid_attachment" },
  { tool: "outlook_calendar_write", action: "delete", params: { action: "delete" }, error: "invalid_resource_id" },

  { tool: "outlook_mail_write", action: "create_draft", params: { action: "create_draft", subject: "Missing body" }, error: "invalid_mail_payload" },
  { tool: "outlook_mail_write", action: "update_draft", params: { action: "update_draft", messageId: "message" }, error: "invalid_mail_payload" },
  { tool: "outlook_mail_write", action: "update_properties", params: { action: "update_properties", messageId: "message" }, error: "invalid_mail_payload" },
  { tool: "outlook_mail_write", action: "reply_draft", params: { action: "reply_draft", messageId: "message" }, error: "invalid_mail_payload" },
  { tool: "outlook_mail_write", action: "reply_all_draft", params: { action: "reply_all_draft", messageId: "message" }, error: "invalid_mail_payload" },
  { tool: "outlook_mail_write", action: "forward_draft", params: { action: "forward_draft", messageId: "message", bodyText: "Forward" }, error: "invalid_mail_payload" },
  { tool: "outlook_mail_write", action: "copy", params: { action: "copy", messageId: "message", destination: "inbox", destinationFolderId: "folder" }, error: "invalid_destination" },
  { tool: "outlook_mail_write", action: "add_attachment", params: { action: "add_attachment", messageId: "message" }, error: "invalid_attachment" },
  { tool: "outlook_mail_write", action: "move", params: { action: "move", messageId: "message" }, error: "invalid_destination" },
  { tool: "outlook_mail_write", action: "mark_read", params: { action: "mark_read", messageId: "message" }, error: "invalid_read_state" },
  { tool: "outlook_mail_write", action: "send_draft", params: { action: "send_draft", messageId: "message", bodyText: "must not be ignored" }, error: "invalid_write_parameter" },
  { tool: "outlook_mail_write", action: "delete", params: { action: "delete", messageId: "message", categories: ["must-not-be-ignored"] }, error: "invalid_write_parameter" },

  { tool: "microsoft_todo_write", action: "create_list", params: { action: "create_list" }, error: "invalid_title" },
  { tool: "microsoft_todo_write", action: "update_list", params: { action: "update_list", listId: "list" }, error: "invalid_title" },
  { tool: "microsoft_todo_write", action: "delete_list", params: { action: "delete_list", listId: "list", title: "ignored" }, error: "invalid_write_parameter" },
  { tool: "microsoft_todo_write", action: "create_task", params: { action: "create_task", listId: "list" }, error: "invalid_title" },
  { tool: "microsoft_todo_write", action: "update_task", params: { action: "update_task", listId: "list", taskId: "task" }, error: "invalid_task" },
  { tool: "microsoft_todo_write", action: "delete_task", params: { action: "delete_task", listId: "list", taskId: "task", title: "ignored" }, error: "invalid_write_parameter" },
  { tool: "microsoft_todo_write", action: "add_checklist", params: { action: "add_checklist", listId: "list", taskId: "task", title: "Done", checklistCheckedDateTime: "2026-09-07T12:00:00Z" }, error: "invalid_write_parameter" },
  { tool: "microsoft_todo_write", action: "update_checklist", params: { action: "update_checklist", listId: "list", taskId: "task", checklistItemId: "check" }, error: "invalid_checklist" },
  { tool: "microsoft_todo_write", action: "update_checklist", params: { action: "update_checklist", listId: "list", taskId: "task", checklistItemId: "check", checklistCheckedDateTime: "2026-02-30T12:00:00Z" }, error: "invalid_datetime" },
  { tool: "microsoft_todo_write", action: "delete_checklist", params: { action: "delete_checklist", listId: "list", taskId: "task", checklistItemId: "check", title: "ignored" }, error: "invalid_write_parameter" },
  { tool: "microsoft_todo_write", action: "add_linked_resource", params: { action: "add_linked_resource", listId: "list", taskId: "task", linkedResourceWebUrl: "https://example.invalid" }, error: "invalid_linked_resource" },
  { tool: "microsoft_todo_write", action: "update_linked_resource", params: { action: "update_linked_resource", listId: "list", taskId: "task", linkedResourceId: "link" }, error: "invalid_linked_resource" },
  { tool: "microsoft_todo_write", action: "delete_linked_resource", params: { action: "delete_linked_resource", listId: "list", taskId: "task", linkedResourceId: "link", title: "ignored" }, error: "invalid_write_parameter" },
  { tool: "microsoft_todo_write", action: "add_attachment", params: { action: "add_attachment", listId: "list", taskId: "task" }, error: "invalid_attachment" },
  { tool: "microsoft_todo_write", action: "delete_attachment", params: { action: "delete_attachment", listId: "list", taskId: "task", attachmentId: "attachment", title: "ignored" }, error: "invalid_write_parameter" },
];

describe("pre-service mutation validation", () => {
  it.each(malformedMutations)("$tool:$action rejects before credentials, OAuth, or Graph", async ({ tool, params, error }) => {
    credential.readCredential.mockClear();
    credential.exchangeRefreshToken.mockClear();
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      const result = await registeredTools()[tool].execute("malformed", params);
      expect(result.details).toMatchObject({ ok: false, error });
      expect(credential.readCredential).not.toHaveBeenCalled();
      expect(credential.exchangeRefreshToken).not.toHaveBeenCalled();
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      fetchSpy.mockRestore();
    }
  });
});
