import { createHash, randomBytes } from "node:crypto";
import { readFileSync } from "node:fs";
import { link, mkdir, mkdtemp, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { getToolPluginMetadata } from "openclaw/plugin-sdk/tool-plugin";
import entry, { approvalInventory, assertOwnedTodoList, assertWriteActionFields, attachmentWritePlan, boundedCollectionPage, calendarApprovalCriteria, calendarCollectionPath, calendarEventPath, calendarEventPayload, calendarPageIsTruncated, calendarViewPath, calendarViewQuery, calendarWindowDateTime, classifyApproval, collectFilteredCollection, dateTimeTimeZone, deadlineSignal, enforceOneDriveInstructionPreflight, eventGetFields, eventMatchesSearch, eventReadFields, fileAttachmentPayload, listMailFolders, mailFolderCollectionPath, mailListQuery, mailMessageActionPath, mailMessageCollectionPath, mailMessagePayload, mailReplyForwardPlan, NativeApprovalSnapshotStore, nextGraphPath, normalizedWarningApprovalAction, oneDriveAgentsInstructions, oneDriveInstructionDirectories, oneDriveWriteApprovalCriteria, readProtectedMediaSource, readOperationTimeout, schedulePage, taskPayload, todoTaskChildPath, todoTaskMatches, todoTaskPath, validateAttachmentContent, validateProtectedMediaUri, WarningApprovalTrustStore, WRITE_ACTION_FIELDS } from "./index.js";
import { beforeMicrosoftGraphToolCall, enforceOneDriveInstructionExecution } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
import { ONEDRIVE_AGENTS_MAX_DEPTH, OneDriveAgentsSessionCache, oneDriveAgentsSessionCache } from "./onedrive-agents-instructions.js";

describe("microsoft-graph plugin contract", () => {
  it("registers CLI metadata without accessing the restricted runtime", () => {
    const registerCli = vi.fn();
    const runtime = new Proxy(Object.create(null), {
      get() {
        throw new Error('Plugin "microsoft-graph" runtime is intentionally unavailable during "cli-metadata" registration.');
      },
    });

    expect(() => entry.register({
      registrationMode: "cli-metadata",
      pluginConfig: {},
      runtime,
      registerCli,
      registerGatewayMethod: () => { throw new Error("gateway registration is unavailable during CLI metadata collection"); },
      registerTool: () => { throw new Error("tool registration is unavailable during CLI metadata collection"); },
      on: () => { throw new Error("hook registration is unavailable during CLI metadata collection"); },
    } as any)).not.toThrow();
    expect(registerCli).toHaveBeenCalledTimes(1);
    expect(registerCli.mock.calls[0]?.[1]).toMatchObject({
      commands: ["microsoft-graph"],
      descriptors: [{
        name: "microsoft-graph",
        description: "Microsoft Graph operator commands",
        hasSubcommands: true,
      }],
    });
  });

  it("defers vault-key use until after per-operation authorization", () => {
    const current = graphPolicyFixture();
    const policy = { version: 2 as const, rules: current.rules, services: current.services };
    const registerTool = vi.fn();
    const registerCli = vi.fn();
    const registerGatewayMethod = vi.fn();
    const api = {
      pluginConfig: { enabled: true, credentialVaultKey: "invalid", policy },
      registerTool,
      registerCli,
      registerGatewayMethod,
      runtime: { state: { resolveStateDir: vi.fn(() => "/synthetic/state") } },
      on: vi.fn(),
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    };

    expect(() => entry.register(api as any)).not.toThrow();
    expect(registerTool).toHaveBeenCalled();
    expect(registerCli).toHaveBeenCalledTimes(1);
    expect(registerGatewayMethod).toHaveBeenCalledTimes(8);
    expect(registerGatewayMethod.mock.calls.map(([method, _handler, options]) => [method, options.scope])).toEqual([
      ["microsoft-graph.configuration.validate", "operator.admin"],
      ["microsoft-graph.configuration.resolveFolder", "operator.admin"],
      ["microsoft-graph.credentials.status", "operator.read"],
      ["microsoft-graph.credentials.migrate-from-pass", "operator.admin"],
      ["microsoft-graph.credentials.restore-pass", "operator.admin"],
      ["microsoft-graph.credentials.recover-refresh", "operator.admin"],
      ["microsoft-graph.credentials.device-start", "operator.admin"],
      ["microsoft-graph.credentials.device-status", "operator.admin"],
    ]);
  });

  it("declares one coherent optional tool surface", () => {
    const metadata = getToolPluginMetadata(entry)!;
    expect(metadata.tools.map((tool) => tool.name)).toEqual([
      "onedrive_search", "onedrive_list", "onedrive_read", "onedrive_download", "onedrive_upload", "onedrive_update", "onedrive_metadata_update", "onedrive_create_folder", "onedrive_delete",
      "outlook_calendar_read", "outlook_calendar_write", "outlook_mail_read", "outlook_mail_write", "microsoft_todo_read", "microsoft_todo_write",
      "onedrive_agents_instructions",
    ]);
    expect(metadata.tools.every((tool) => tool.optional)).toBe(true);
    for (const tool of metadata.tools) expect(JSON.stringify(tool.parameters)).not.toMatch(/agent_?id|access_?token|refresh_?token|secret/i);
  });

  it("keeps manifest aligned with runtime metadata", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const manifest = JSON.parse(readFileSync(new URL("../openclaw.plugin.json", import.meta.url), "utf8"));
    expect(manifest.cliCommands).toEqual([{
      name: "microsoft-graph",
      description: "Microsoft Graph operator commands",
      hasSubcommands: true,
    }]);
    expect(manifest.contracts.tools).toEqual(metadata.tools.map((tool) => tool.name));
    expect(manifest.configSchema).toEqual(metadata.configSchema);
    expect(manifest.configContracts.secretInputs.paths).toEqual([
      { path: "credentialVaultKey", expected: "string", ownerKind: "capability" },
    ]);
    expect(metadata.activation).toEqual({ onStartup: true });
    expect(manifest.toolMetadata.onedrive_agents_instructions).toEqual({ optional: true, replaySafe: true, sideEffecting: false });
  });

  it("publishes the complete structured default-deny policy schema", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const properties = (metadata.configSchema as any).properties;
    expect(properties).not.toHaveProperty("policyPath");
    expect(properties).not.toHaveProperty("readCredential");
    expect(properties).not.toHaveProperty("writeCredential");
    expect(properties).not.toHaveProperty("credentialBackend");
    expect(properties).not.toHaveProperty("credentialVaultKeys");
    expect(properties.credentialVaultKey).toMatchObject({ type: "object", additionalProperties: false, required: ["source", "provider", "id"] });
    expect(properties.credentialVaultKey).not.toHaveProperty("anyOf");
    expect(properties.warningApprovalsRequired).toMatchObject({ type: "boolean", default: true });
    expect(properties.policy).toMatchObject({ type: "object", additionalProperties: false });
    expect(properties.policy.properties.version).toEqual({ const: 2, type: "number" });
    expect(properties.policy.properties.rules.properties.default).toEqual({ const: "deny", type: "string" });
    expect(properties.policy.properties).not.toHaveProperty("account");
  });

  it("publishes the exact lean AGENTS.md preflight schema", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const tool = metadata.tools.find((candidate) => candidate.name === "onedrive_agents_instructions")!;
    expect(tool.optional).toBe(true);
    expect(Object.keys((tool.parameters as any).properties)).toEqual(["rootLabel", "relativeDirectory", "acknowledgement"]);
    expect((tool.parameters as any).required).toEqual(["rootLabel"]);
    expect((tool.parameters as any).additionalProperties).toBe(false);
    expect((tool.parameters as any).properties.relativeDirectory).toMatchObject({ type: "string", maxLength: 1024, default: "" });
    expect(JSON.stringify(tool.parameters)).not.toMatch(/agent_?id|session_?id|agent_managed|access_?token|secret/i);
  });

  it("binds write schemas to the native provider endpoint limits", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const configSchema = metadata.configSchema as any;
    expect(configSchema.properties.requestTimeoutMs).toMatchObject({ minimum: 1000, maximum: 30000, default: 5000 });
    expect(configSchema.properties.readOperationTimeoutMs).toMatchObject({ minimum: 5000, maximum: 300000, default: 30000 });
    expect(configSchema.properties.calendarMultiwriteTimeoutMs).toMatchObject({ minimum: 30000, maximum: 1800000, default: 300000 });
    expect(configSchema.properties.maxReadBytes).toMatchObject({ maximum: 250 * 1024 * 1024 * 1024, default: 250 * 1024 * 1024 * 1024 });
    expect(configSchema.properties.maxReadOutputBytes).toMatchObject({ maximum: 1024 * 1024, default: 262144 });
    expect(configSchema.properties.maxAttachmentDownloadBytes).toMatchObject({ maximum: 150 * 1024 * 1024, default: 150 * 1024 * 1024 });
    expect(configSchema.properties.attachmentDownloadTimeoutMs).toMatchObject({ minimum: 30000, maximum: 1800000, default: 600000 });
    for (const obsolete of ["maxAttachmentStageBytes", "attachmentStageTtlSeconds", "attachmentStageTimeoutMs"]) expect(configSchema.properties).not.toHaveProperty(obsolete);
    const uploadSchema = metadata.tools.find((tool) => tool.name === "onedrive_upload")?.parameters as any;
    const updateSchema = metadata.tools.find((tool) => tool.name === "onedrive_update")?.parameters as any;
    expect(uploadSchema.properties.sourceMediaUri).toMatchObject({ type: "string", maxLength: 4096, pattern: expect.stringContaining("media://inbound/") });
    expect(uploadSchema.properties.sourceSha256).toMatchObject({ type: "string", minLength: 64, maxLength: 64, pattern: "^[a-f0-9]{64}$" });
    expect(uploadSchema.properties.sourceByteSize).toMatchObject({ type: "integer", minimum: 0, maximum: 250 * 1024 * 1024 * 1024 });
    expect(uploadSchema.required).toEqual(expect.arrayContaining(["sourceMediaUri", "sourceSha256", "sourceByteSize"]));
    expect(updateSchema.required).toEqual(expect.arrayContaining(["sourceMediaUri", "sourceSha256", "sourceByteSize"]));
    for (const toolName of ["outlook_calendar_write", "outlook_mail_write"]) {
      const schema = metadata.tools.find((tool) => tool.name === toolName)?.parameters as any;
      expect(schema.properties.attachmentMediaUri).toMatchObject({ type: "string", maxLength: 4096 });
    }
    const todoSchema = metadata.tools.find((tool) => tool.name === "microsoft_todo_write")?.parameters as any;
    expect(todoSchema.properties.attachmentMediaUri).toMatchObject({ type: "string", maxLength: 4096 });
    for (const schema of [uploadSchema, updateSchema, todoSchema, ...["outlook_calendar_write", "outlook_mail_write"].map((name) => metadata.tools.find((tool) => tool.name === name)?.parameters as any)]) expect(JSON.stringify(schema)).not.toMatch(/contentBase64|sourceFilePath|attachmentContentBase64/);
    for (const toolName of ["onedrive_upload", "onedrive_update", "onedrive_metadata_update", "onedrive_create_folder", "onedrive_delete", "outlook_calendar_write", "outlook_mail_write", "microsoft_todo_write"]) {
      const schema = metadata.tools.find((tool) => tool.name === toolName)?.parameters as any;
      expect(schema.properties.chatConfirmed.description, toolName).toContain("ignored");
      expect(schema.properties.chatConfirmed.description, toolName).toContain("never authorize");
      expect(schema.properties.chatConfirmationToken.description, toolName).toContain("ignored");
    }
    expect(configSchema.properties).not.toHaveProperty("maxWriteBytes");
  });

  it("requires OneDrive fingerprints in the semantic write effect", () => {
    const fingerprint = "a".repeat(64);
    const root = { label: "synthetic_documents", drive_id: "synthetic-drive", item_id: "synthetic-root" };
    expect(oneDriveWriteApprovalCriteria("onedrive_upload", {
      rootLabel: "synthetic_documents",
      relativePath: " SYNTHETIC_FOLDER//a.pdf ",
      sourceMediaUri: "media://inbound/first.pdf",
      sourceSha256: fingerprint,
      sourceByteSize: 42,
    }, root)).toEqual({
      operation: "upload",
      authorizedRoot: { label: "synthetic_documents", driveId: "synthetic-drive", itemId: "synthetic-root" },
      relativePath: "SYNTHETIC_FOLDER/a.pdf",
      contentType: "application/octet-stream",
      sourceSha256: fingerprint,
      sourceByteSize: 42,
    });
    expect(() => oneDriveWriteApprovalCriteria("onedrive_upload", {
      rootLabel: "synthetic_documents",
      relativePath: "a.pdf",
      sourceMediaUri: "media://inbound/a.pdf",
    }, root)).toThrow("invalid_source_fingerprint");
    expect(() => oneDriveWriteApprovalCriteria("onedrive_upload", {
      rootLabel: "synthetic_documents",
      relativePath: "a.pdf",
      sourceMediaUri: "media://inbound/a.pdf",
      sourceSha256: fingerprint,
    }, root)).toThrow("invalid_source_fingerprint");
  });

  it("keeps allow-always trust process-local and exact-scope", () => {
    const scope = { agentId: "main", toolName: "outlook_mail_write", action: "mark_read" };
    const store = new WarningApprovalTrustStore();
    expect(store.has(scope)).toBe(false);
    store.grant(scope);
    expect(store.has(scope)).toBe(true);
    expect(store.has({ ...scope, agentId: "other" })).toBe(false);
    expect(store.has({ ...scope, toolName: "microsoft_todo_write" })).toBe(false);
    expect(store.has({ ...scope, action: "move" })).toBe(false);
    expect(new WarningApprovalTrustStore().has(scope)).toBe(false);
    const bounded = new WarningApprovalTrustStore(1);
    bounded.grant(scope);
    expect(() => bounded.grant({ ...scope, action: "move" })).toThrow("warning_approval_trust_capacity_exceeded");
    store.clear();
    expect(store.has(scope)).toBe(false);
    expect(normalizedWarningApprovalAction("onedrive_metadata_update", {})).toBe("metadata_update");
    expect(normalizedWarningApprovalAction("outlook_mail_write", { action: "mark_read" })).toBe("mark_read");
    expect(normalizedWarningApprovalAction("outlook_mail_write", { action: "INVALID" })).toBe("unknown");
  });

  it("keeps caller cancellation ahead of the plugin deadline", () => {
    const controller = new AbortController();
    const reason = new Error("caller_cancelled");
    controller.abort(reason);
    const bounded = deadlineSignal(controller.signal, 30_000, 5_000);
    expect(bounded.aborted).toBe(true);
    expect(bounded.reason).toBe(reason);
  });

  it("reads OneDrive upload sources only from protected media/inbound staging", async () => {
    const workspace = await mkdtemp(join(tmpdir(), "microsoft-graph-staging-"));
    const inbound = join(workspace, "media", "inbound");
    const staged = join(inbound, "batch", "SYNTHETIC_RECORD.xlsx");
    const outside = join(workspace, "outside.xlsx");
    try {
      await mkdir(join(inbound, "batch"), { recursive: true });
      await writeFile(staged, Buffer.from([0, 255, 1, 254]));
      await writeFile(outside, "outside");
      expect(await readProtectedMediaSource("media://inbound/batch/SYNTHETIC_RECORD.xlsx", workspace, 4)).toEqual(Buffer.from([0, 255, 1, 254]));
      await expect(readProtectedMediaSource(outside, workspace, 1024)).rejects.toThrow("invalid_source_media_uri");
      await expect(readProtectedMediaSource("media/inbound/batch/SYNTHETIC_RECORD.xlsx", workspace, 1024)).rejects.toThrow("invalid_source_media_uri");
      await expect(readProtectedMediaSource("media://inbound/batch/SYNTHETIC_RECORD.xlsx", workspace, 3)).rejects.toThrow("invalid_source_media_uri");
      await symlink(staged, join(inbound, "symlink.xlsx"));
      await expect(readProtectedMediaSource("media://inbound/symlink.xlsx", workspace, 1024)).rejects.toThrow("invalid_source_media_uri");
      await link(staged, join(inbound, "hardlink.xlsx"));
      await expect(readProtectedMediaSource("media://inbound/batch/SYNTHETIC_RECORD.xlsx", workspace, 1024)).rejects.toThrow("invalid_source_media_uri");
    } finally {
      await rm(workspace, { recursive: true, force: true });
    }
  });

  it("accepts only canonical protected media URIs", () => {
    expect(validateProtectedMediaUri("media://inbound/a.bin")).toBe("a.bin");
    for (const value of [undefined, "/workspace/media/inbound/a.bin", "media/inbound/a.bin", "media://inbound/../a.bin", "media://inbound/a.bin?x=1", "media://inbound/a\\b"]) expect(() => validateProtectedMediaUri(value)).toThrow("invalid_source_media_uri");
  });

  it("exposes optional calendarId on calendar reads and writes", () => {
    const metadata = getToolPluginMetadata(entry)!;
    for (const name of ["outlook_calendar_read", "outlook_calendar_write"]) {
      const schema = metadata.tools.find((tool) => tool.name === name)?.parameters as any;
      expect(schema.properties.calendarId).toMatchObject({ type: "string", minLength: 1, maxLength: 512, description: expect.stringContaining("list_calendars") });
      expect(schema.required ?? []).not.toContain("calendarId");
    }
  });

  it("routes calendar operations to the default or selected calendar", () => {
    expect(calendarCollectionPath()).toBe("/me/events");
    expect(calendarViewPath()).toBe("/me/calendarView");
    expect(calendarEventPath(undefined, "event-1")).toBe("/me/events/event-1");
    expect(calendarCollectionPath("fixture+calendar=")).toBe("/me/calendars/fixture%2Bcalendar%3D/events");
    expect(calendarViewPath("fixture+calendar=")).toBe("/me/calendars/fixture%2Bcalendar%3D/calendarView");
    expect(calendarEventPath("fixture+calendar=", "event+1=")).toBe("/me/calendars/fixture%2Bcalendar%3D/events/event%2B1%3D");
    expect(() => calendarCollectionPath("../calendar")).toThrow("invalid_resource_id");
  });

  it("exposes mailbox-wide search and exact custom mail-folder targeting", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const schema = metadata.tools.find((tool) => tool.name === "outlook_mail_read")?.parameters as any;
    expect(schema.properties.action.anyOf.map((entry: any) => entry.const)).toEqual(expect.arrayContaining(["list_folders", "search_messages"]));
    expect(schema.properties.folderId).toMatchObject({ type: "string", description: expect.stringContaining("list_folders") });
    expect(mailFolderCollectionPath()).toBe("/me/mailFolders");
    expect(mailFolderCollectionPath("custom+folder=")).toBe("/me/mailFolders/custom%2Bfolder%3D/childFolders");
    expect(mailMessageCollectionPath(undefined, undefined, true)).toBe("/me/messages");
    expect(mailMessageCollectionPath(undefined, "custom+folder=")).toBe("/me/mailFolders/custom%2Bfolder%3D/messages");
    expect(() => mailMessageCollectionPath("inbox", "custom-folder")).toThrow("invalid_mail_folder_target");
  });

  it("supports the stable Microsoft Graph v1.0 event settings including reminders and recurrence", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const schema = metadata.tools.find((tool) => tool.name === "outlook_calendar_write")?.parameters as any;
    expect(schema.properties.showAs.anyOf.map((entry: any) => entry.const)).toEqual(["free", "tentative", "busy", "oof", "workingElsewhere", "unknown"]);
    for (const field of ["isReminderOn", "reminderMinutesBeforeStart", "isAllDay", "categories", "importance", "sensitivity", "allowNewTimeProposals", "hideAttendees", "isOnlineMeeting", "onlineMeetingProvider", "responseRequested", "recurrence", "transactionId", "locations", "attendeeDetails"]) expect(schema.properties[field]).toBeTruthy();
    expect(calendarEventPayload({ subject: "Synthetic Event", bodyHtml: "<p>Synthetic body</p>", startDateTime: "2026-09-08T09:00:00", endDateTime: "2026-09-08T10:00:00", location: "Synthetic Location", showAs: "busy", isReminderOn: true, reminderMinutesBeforeStart: 15, categories: ["Synthetic Category"], importance: "high" })).toEqual({
      subject: "Synthetic Event",
      body: { contentType: "HTML", content: "<p>Synthetic body</p>" },
      start: { dateTime: "2026-09-08T09:00:00", timeZone: "UTC" },
      end: { dateTime: "2026-09-08T10:00:00", timeZone: "UTC" },
      location: { displayName: "Synthetic Location" },
      showAs: "busy",
      isReminderOn: true,
      reminderMinutesBeforeStart: 15,
      categories: ["Synthetic Category"],
      importance: "high",
    });
    expect(calendarEventPayload({ reminderMinutesBeforeStart: 30 })).toEqual({ reminderMinutesBeforeStart: 30, isReminderOn: true });
    expect(() => calendarEventPayload({ reminderMinutesBeforeStart: 15, isReminderOn: false })).toThrow("invalid_reminder");
    expect(() => calendarEventPayload({ recurrence: { pattern: { type: "weekly", interval: 1 }, range: { type: "noEnd", startDate: "2026-09-08" } } })).toThrow("invalid_recurrence");
  });

  it("rejects Graph-invalid event timing and recurrence combinations", () => {
    expect(() => calendarEventPayload({ startDateTime: "2026-09-08T10:00:00", endDateTime: "2026-09-08T09:00:00", timeZone: "UTC" })).toThrow("invalid_event_time_order");
    expect(() => calendarEventPayload({ startDateTime: "2026-02-30T09:00:00", endDateTime: "2026-03-01T10:00:00", timeZone: "UTC" })).toThrow("invalid_datetime");
    expect(() => calendarEventPayload({ startDateTime: "2026-09-08T24:00:00", endDateTime: "2026-09-09T01:00:00", timeZone: "UTC" })).toThrow("invalid_datetime");
    expect(() => calendarEventPayload({ startDateTime: "2024-02-29T09:00:00", endDateTime: "2024-02-29T10:00:00", timeZone: "UTC" })).not.toThrow();
    expect(() => calendarEventPayload({ startDateTime: "2026-09-08T09:00:00", endDateTime: "2026-09-08T10:00:00", startTimeZone: "Unresolvable Graph Zone A", endTimeZone: "Unresolvable Graph Zone B" })).toThrow("invalid_event_time_zone");
    expect(() => calendarEventPayload({ startDateTime: "2026-09-08T00:00:00", endDateTime: "2026-09-09T00:00:00", startTimeZone: "Pacific Standard Time", endTimeZone: "UTC", isAllDay: true })).toThrow("invalid_all_day_event");
    expect(() => calendarEventPayload({ startDateTime: "2026-09-08T00:01:00", endDateTime: "2026-09-09T00:00:00", timeZone: "UTC", isAllDay: true })).toThrow("invalid_all_day_event");
    expect(calendarEventPayload({ startDateTime: "2026-09-08T09:00:00", recurrence: { pattern: { type: "relativeMonthly", interval: 1, daysOfWeek: ["monday"] }, range: { type: "noEnd", startDate: "2026-09-08" } } }).recurrence).toEqual({ pattern: { type: "relativeMonthly", interval: 1, daysOfWeek: ["monday"], index: "first" }, range: { type: "noEnd", startDate: "2026-09-08" } });
    expect(() => calendarEventPayload({ startDateTime: "2026-09-08T09:00:00", recurrence: { pattern: { type: "daily", interval: 1, dayOfMonth: 8 }, range: { type: "noEnd", startDate: "2026-09-08" } } })).toThrow("invalid_recurrence");
    expect(() => calendarEventPayload({ startDateTime: "2026-09-08T09:00:00", recurrence: { pattern: { type: "daily", interval: 1 }, range: { type: "endDate", startDate: "2026-09-08", endDate: "2026-09-07" } } })).toThrow("invalid_recurrence");
    expect(() => calendarEventPayload({ startDateTime: "2026-02-28T09:00:00", recurrence: { pattern: { type: "daily", interval: 1 }, range: { type: "noEnd", startDate: "2026-02-30" } } })).toThrow("invalid_recurrence");
    expect(() => calendarEventPayload({ startDateTime: "2026-02-28T09:00:00", recurrence: { pattern: { type: "daily", interval: 1 }, range: { type: "endDate", startDate: "2026-02-28", endDate: "2026-02-30" } } })).toThrow("invalid_recurrence");
    expect(() => calendarEventPayload({ startDateTime: "2024-02-29T09:00:00", recurrence: { pattern: { type: "daily", interval: 1 }, range: { type: "noEnd", startDate: "2024-02-29" } } })).not.toThrow();
    expect(() => calendarEventPayload({ startDateTime: "2026-09-08T09:00:00", recurrence: { pattern: { type: "daily", interval: 1 }, range: { type: "noEnd", startDate: "2026-09-09" } } })).toThrow("invalid_recurrence");
    expect(calendarEventPayload({ startDateTime: "2026-09-08T00:00:00", endDateTime: "2026-09-09T00:00:00", isAllDay: true })).toMatchObject({ start: { timeZone: "UTC" }, end: { timeZone: "UTC" }, isAllDay: true });
    const crossMidnight = calendarEventPayload({ startDateTime: "2026-09-07T23:00:00Z", endDateTime: "2026-09-08T00:00:00Z", timeZone: "Europe/Berlin", recurrence: { pattern: { type: "daily", interval: 1 }, range: { type: "noEnd", startDate: "2026-09-08" } } });
    expect(crossMidnight).toMatchObject({ start: { dateTime: "2026-09-08T01:00:00" }, recurrence: { range: { startDate: "2026-09-08" } } });
    expect(() => calendarEventPayload({ startDateTime: "2026-09-07T23:00:00Z", endDateTime: "2026-09-08T00:00:00Z", timeZone: "Europe/Berlin", recurrence: { pattern: { type: "daily", interval: 1 }, range: { type: "noEnd", startDate: "2026-09-07" } } })).toThrow("invalid_recurrence");
    expect(calendarEventPayload({ startDateTime: "2026-09-07T22:00:00Z", endDateTime: "2026-09-08T22:00:00Z", timeZone: "Europe/Berlin", isAllDay: true })).toMatchObject({ start: { dateTime: "2026-09-08T00:00:00" }, end: { dateTime: "2026-09-09T00:00:00" }, isAllDay: true });
  });

  it("exposes bounded event search across stable event metadata", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const schema = metadata.tools.find((tool) => tool.name === "outlook_calendar_read")?.parameters as any;
    expect(schema.properties.action.anyOf.map((entry: any) => entry.const)).toContain("search_events");
    expect(schema.properties.eventType.anyOf.map((entry: any) => entry.const)).not.toContain("seriesMaster");
    expect(eventReadFields()).toContain("changeKey");
    expect(eventReadFields()).not.toContain("body");
    expect(eventReadFields(true)).toContain("body");
    expect(eventReadFields()).not.toContain("cancelledOccurrences");
    expect(eventGetFields()).toContain("cancelledOccurrences");
    expect(calendarPageIsTruncated(25, 50)).toBe(true);
    expect(calendarPageIsTruncated(50, 50)).toBe(false);
    expect(calendarPageIsTruncated(50, 50, "/me/calendarView?$skiptoken=next")).toBe(true);
    const item = { subject: "Synthetic planning record", bodyPreview: "Synthetic preview", location: { displayName: "Synthetic Location" }, organizer: { emailAddress: { address: "organizer@example.invalid" } }, attendees: [{ emailAddress: { address: "attendee@example.invalid" } }], categories: ["Synthetic Category"], showAs: "busy", isAllDay: false };
    expect(eventMatchesSearch(item, { search: "planning", searchFields: ["subject"], showAs: "busy", attendee: "attendee@example.invalid" })).toBe(true);
    expect(eventMatchesSearch(item, { search: "absent", searchFields: ["location"] })).toBe(false);
  });

  it("uses explicit offsets for local calendar windows and validates continuation scope", () => {
    expect(calendarWindowDateTime("2026-01-08T09:00:00", "Europe/Berlin")).toBe("2026-01-08T09:00:00+01:00");
    expect(calendarWindowDateTime("2026-07-08T09:00:00", "Europe/Berlin")).toBe("2026-07-08T09:00:00+02:00");
    expect(calendarWindowDateTime("2026-07-08T09:00:00Z", "Europe/Berlin")).toBe("2026-07-08T09:00:00Z");
    const query = calendarViewQuery("2026-07-08T09:00:00", "2026-07-08T10:00:00", "Europe/Berlin", 25);
    expect(query.get("startDateTime")).toBe("2026-07-08T09:00:00+02:00");
    expect(query.has("$select")).toBe(false);
    expect(nextGraphPath("https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=next", "/me/calendarView")).toBe("/me/calendarView?$skiptoken=next");
    expect(nextGraphPath("/me/calendarView?$skiptoken=next", "/me/calendarView")).toBe("/me/calendarView?$skiptoken=next");
    expect(() => nextGraphPath("https://evil.invalid/v1.0/me/calendarView?$skiptoken=next", "/me/calendarView")).toThrow("invalid_provider_response");
    for (const prefix of ["/me/calendars", "/me/calendarView", "/me/messages", "/me/messages/message/attachments", "/me/mailFolders", "/me/mailFolders/folder/childFolders", "/me/todo/lists", "/me/todo/lists/list/tasks", "/me/todo/lists/list/tasks/task/attachments"]) expect(nextGraphPath(`${prefix}?$skiptoken=next`, prefix)).toBe(`${prefix}?$skiptoken=next`);
    const attacks = [
      "/me/calendarViewExtra?$skiptoken=x",
      "/me/calendarView/../events?$skiptoken=x",
      "/me/calendarView/%2e%2e/events?$skiptoken=x",
      "/me/calendarView/%252e%252e/events?$skiptoken=x",
      "https://graph.microsoft.com.evil.invalid/v1.0/me/calendarView?$skiptoken=x",
      "https://user@graph.microsoft.com/v1.0/me/calendarView?$skiptoken=x",
      "https://graph.microsoft.com:444/v1.0/me/calendarView?$skiptoken=x",
      "https://graph.microsoft.com/v1.0/me/events?$skiptoken=x",
      "https://graph.microsoft.com/v1.0/me/calendars/other/calendarView?$skiptoken=x",
    ];
    for (const value of attacks) expect(() => nextGraphPath(value, "/me/calendarView"), value).toThrow("invalid_provider_response");
  });

  it("normalizes offset-bearing dateTimeTimeZone values across DST without emitting offsets", () => {
    expect(dateTimeTimeZone("2026-03-29T00:30:00Z", "W. Europe Standard Time")).toEqual({ dateTime: "2026-03-29T01:30:00", timeZone: "W. Europe Standard Time" });
    expect(dateTimeTimeZone("2026-03-29T01:30:00Z", "W. Europe Standard Time")).toEqual({ dateTime: "2026-03-29T03:30:00", timeZone: "W. Europe Standard Time" });
    expect(dateTimeTimeZone("2026-10-25T00:30:00Z", "Europe/Berlin")).toEqual({ dateTime: "2026-10-25T02:30:00", timeZone: "Europe/Berlin" });
    expect(dateTimeTimeZone("2026-10-25T01:30:00Z", "Europe/Berlin")).toEqual({ dateTime: "2026-10-25T02:30:00", timeZone: "Europe/Berlin" });
    expect(dateTimeTimeZone("2026-07-08T09:00:00+02:00", "UTC")).toEqual({ dateTime: "2026-07-08T07:00:00", timeZone: "UTC" });
    expect(() => dateTimeTimeZone("2026-07-08T09:00:00Z", "Unsupported Synthetic Zone")).toThrow("invalid_datetime_timezone");
    expect(dateTimeTimeZone("2026-07-08T09:00:00", "Pacific Standard Time")).toEqual({ dateTime: "2026-07-08T09:00:00", timeZone: "Pacific Standard Time" });
    expect(calendarEventPayload({ startDateTime: "2026-07-08T09:00:00+02:00", endDateTime: "2026-07-08T10:00:00+02:00", timeZone: "UTC" })).toMatchObject({ start: { dateTime: "2026-07-08T07:00:00", timeZone: "UTC" }, end: { dateTime: "2026-07-08T08:00:00", timeZone: "UTC" } });
    expect(() => calendarEventPayload({ startDateTime: "2026-07-08T00:00:00+02:00", endDateTime: "2026-07-09T00:00:00+02:00", timeZone: "UTC", isAllDay: true })).toThrow("invalid_all_day_event");
  });

  it("keeps client-filtered paging safe within a page and across provider pages", async () => {
    const prefix = "/me/calendarView";
    const samePage = await collectFilteredCollection(`${prefix}?page=1`, prefix, ["id", "subject"], 1, 10, (entry) => entry.subject === "match", async () => ({ value: [{ id: "one", subject: "match" }, { id: "two", subject: "match" }], "@odata.nextLink": `${prefix}?page=2` }));
    expect(samePage).toEqual({ items: [{ id: "one", subject: "match" }], scanned: 1, truncated: true });
    expect(samePage).not.toHaveProperty("continuation");

    const pages: Record<string, unknown> = {
      [`${prefix}?page=1`]: { value: [{ id: "skip", subject: "other" }], "@odata.nextLink": `https://graph.microsoft.com/v1.0${prefix}?page=2` },
      [`${prefix}?page=2`]: { value: [{ id: "match", subject: "match" }] },
    };
    const acrossPages = await collectFilteredCollection(`${prefix}?page=1`, prefix, ["id", "subject"], 2, 10, (entry) => entry.subject === "match", async (path) => pages[path]);
    expect(acrossPages).toEqual({ items: [{ id: "match", subject: "match" }], scanned: 2, truncated: false });
    expect(acrossPages).not.toHaveProperty("continuation");
  });

  it("rejects repeated and cyclic nextLinks in client-filtered reads", async () => {
    const prefix = "/me/calendarView";
    await expect(collectFilteredCollection(`${prefix}?page=1`, prefix, ["id"], 2, 10, () => false, async (path) => ({ value: [], "@odata.nextLink": path }))).rejects.toThrow("invalid_provider_response");
    const pages: Record<string, unknown> = {
      [`${prefix}?page=1`]: { value: [], "@odata.nextLink": `${prefix}?page=2` },
      [`${prefix}?page=2`]: { value: [], "@odata.nextLink": `${prefix}?page=1` },
    };
    await expect(collectFilteredCollection(`${prefix}?page=1`, prefix, ["id"], 2, 10, () => false, async (path) => pages[path])).rejects.toThrow("invalid_provider_response");
  });

  it("resumes bounded recursive mail-folder traversal without dropping queued work", async () => {
    const signal = deadlineSignal(undefined, 1_000, 100);
    const first = await listMailFolders("token", undefined, true, true, 1, signal, undefined, async () => ({
      value: [{ id: "parent", displayName: "Parent", childFolderCount: 1 }],
    }));
    expect(first).toMatchObject({ items: [{ id: "parent", path: "Parent", depth: 0 }], truncated: true, continuationState: { kind: "mail_folders" } });
    const second = await listMailFolders("token", undefined, true, true, 1, signal, first.continuationState, async (path) => {
      expect(path).toContain("/me/mailFolders/parent/childFolders");
      return { value: [{ id: "child", displayName: "Child", parentFolderId: "parent", childFolderCount: 0 }] };
    });
    expect(second).toEqual({ items: [expect.objectContaining({ id: "child", path: "Parent/Child", depth: 1 })], truncated: false });
  });

  it("uses the mail read-operation deadline across multiple individually bounded requests", async () => {
    const requestTimeoutMs = 100;
    const operation = deadlineSignal(undefined, readOperationTimeout({ readOperationTimeoutMs: 1_000 }, { action: "list_folders" }), requestTimeoutMs);
    let request = 0;
    const fetchFn = vi.fn(async () => {
      await new Promise((resolve) => setTimeout(resolve, 30));
      request += 1;
      return new Response(JSON.stringify(request === 1
        ? { value: [{ id: "one", displayName: "One", childFolderCount: 0 }], "@odata.nextLink": "/me/mailFolders?$skiptoken=next" }
        : { value: [{ id: "two", displayName: "Two", childFolderCount: 0 }] }), { status: 200 });
    });
    const result = await listMailFolders("token", undefined, false, true, 2, operation, undefined, async (path) => {
      const { graphRequest } = await import("./graph.js");
      return graphRequest("token", path, { signal: operation, requestTimeoutMs, readRetries: 0 }, fetchFn as typeof fetch);
    });
    expect(result).toMatchObject({ items: [{ id: "one" }, { id: "two" }], truncated: false });
    expect(fetchFn).toHaveBeenCalledTimes(2);
    expect(operation.aborted).toBe(false);
    expect(readOperationTimeout({ attachmentDownloadTimeoutMs: 900_000 }, { action: "download_attachment" })).toBe(900_000);
  });

  it("rejects repeated nextLinks in recursive mail-folder traversal", async () => {
    const signal = deadlineSignal(undefined, 1_000, 100);
    await expect(listMailFolders("token", undefined, true, true, 2, signal, undefined, async (path) => ({ value: [], "@odata.nextLink": path }))).rejects.toThrow("invalid_provider_response");
  });

  it("enforces the v1.0 getSchedule boundary and reports provider overflow honestly", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const schema = metadata.tools.find((tool) => tool.name === "outlook_calendar_read")?.parameters as any;
    expect(schema.properties.schedules.maxItems).toBe(20);
    const exact = schedulePage({ value: Array.from({ length: 20 }, (_, index) => ({ scheduleId: `person-${index}@example.invalid` })) });
    expect(exact.items).toHaveLength(20);
    expect(exact.truncated).toBe(false);
    const overflow = schedulePage({ value: Array.from({ length: 21 }, (_, index) => ({ scheduleId: `person-${index}@example.invalid` })) });
    expect(overflow.items).toHaveLength(20);
    expect(overflow.truncated).toBe(true);
  });

  it("converts protected-media bytes to Graph wire payloads with raw-byte bounds", () => {
    const raw = Buffer.alloc(1024, 7);
    expect(fileAttachmentPayload({ attachmentName: "near-limit.bin" }, raw, undefined, 1024)).toMatchObject({ name: "near-limit.bin", contentBytes: raw.toString("base64") });
    expect(() => fileAttachmentPayload({ attachmentName: "too-large.bin" }, Buffer.alloc(1025), undefined, 1024)).toThrow("file_too_large");
    expect(() => fileAttachmentPayload({ attachmentName: "bad.bin", attachmentContentType: "text/plain\r\nx: y" }, Buffer.from("a"))).toThrow("invalid_attachment");
    expect(attachmentWritePlan({ attachmentName: "large.bin" }, "#microsoft.graph.fileAttachment", 150 * 1024 * 1024, Buffer.alloc(3 * 1024 * 1024)).mode).toBe("session");
    expect(() => attachmentWritePlan({ attachmentName: "too-large.bin" }, "#microsoft.graph.taskFileAttachment", 25 * 1024 * 1024, Buffer.alloc(25 * 1024 * 1024 + 1))).toThrow("file_too_large");
    expect(() => validateAttachmentContent({ contentBytes: raw.toString("base64") }, 1024)).not.toThrow();
    expect(() => validateAttachmentContent({ contentBytes: Buffer.alloc(1025).toString("base64") }, 1024)).toThrow("file_too_large");
    expect(() => validateAttachmentContent({ contentBytes: "YR==" }, 1024)).toThrow("invalid_provider_response");
  });

  it("rejects direct OneDrive mutations without an approval snapshot before policy, credentials, or network", async () => {
    const factories: Array<(context: any) => any> = [];
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      entry.register({ pluginConfig: { enabled: true }, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
      for (const index of [4, 5]) {
        const tool = factories[index]({ agentId: "main" });
        const response = await tool.execute("x", { rootLabel: "synthetic_documents", relativePath: "a.bin", sourceMediaUri: "/tmp/a.bin" });
        expect(response.details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
      }
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("requires approval snapshots before direct OneDrive connector checks", async () => {
    const factories: Array<(context: any) => any> = [];
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      entry.register({ pluginConfig: { enabled: false }, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
      const tools = Object.fromEntries(factories.map((factory) => {
        const tool = factory({ agentId: "main" });
        return [tool.name, tool];
      }));
      expect((await tools.onedrive_upload.execute("upload", { rootLabel: "synthetic_documents", relativePath: "large.bin", sourceMediaUri: "media://inbound/large.bin" })).details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
      expect((await tools.onedrive_update.execute("update", { rootLabel: "synthetic_documents", relativePath: "large.bin", sourceMediaUri: "media://inbound/large.bin" })).details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("keeps provider nextLinks internal while reporting truncation honestly", () => {
    for (const prefix of ["/me/calendars", "/me/messages", "/me/todo/lists/list/tasks"]) {
      const nextLink = `https://graph.microsoft.com/v1.0${prefix}?$skiptoken=next`;
      expect(boundedCollectionPage({ value: [{ id: "one" }], "@odata.nextLink": nextLink }, ["id"], 1, prefix)).toEqual({ items: [{ id: "one" }], truncated: true, providerNextLink: `${prefix}?$skiptoken=next` });
      expect(() => boundedCollectionPage({ value: [{ id: "one" }, { id: "two" }], "@odata.nextLink": nextLink }, ["id"], 1, prefix)).toThrow("invalid_provider_response");
    }
  });

  it("applies documented recurrence defaults without relaxing required fields", () => {
    const weekly = calendarEventPayload({ startDateTime: "2026-09-08T09:00:00", recurrence: { pattern: { type: "weekly", interval: 1, daysOfWeek: ["tuesday"] }, range: { type: "noEnd", startDate: "2026-09-08" } } });
    expect((weekly.recurrence as any).pattern.firstDayOfWeek).toBe("sunday");
    expect(() => calendarEventPayload({ startDateTime: "2026-09-08T09:00:00", recurrence: { pattern: { type: "weekly", interval: 1 }, range: { type: "noEnd", startDate: "2026-09-08" } } })).toThrow("invalid_recurrence");
  });

  it("supports full bounded mail KQL/filter search and stable message properties", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const readSchema = metadata.tools.find((tool) => tool.name === "outlook_mail_read")?.parameters as any;
    const writeSchema = metadata.tools.find((tool) => tool.name === "outlook_mail_write")?.parameters as any;
    for (const field of ["searchKql", "receivedAfter", "receivedBefore", "sentAfter", "sentBefore", "isRead", "hasAttachments", "isDraft", "importance", "inferenceClassification", "categories", "orderBy", "includeBody", "includeUniqueBody", "includeHeaders"]) expect(readSchema.properties[field]).toBeTruthy();
    expect(readSchema.properties.action.anyOf.map((entry: any) => entry.const)).toEqual(expect.arrayContaining(["list_attachments", "download_attachment"]));
    for (const obsoleteAction of ["get_attachment", "read_staged_attachment"]) expect(readSchema.properties.action.anyOf.map((entry: any) => entry.const)).not.toContain(obsoleteAction);
    for (const obsoleteField of ["mode", "stagedHandle", "offset", "length", "contentBytes"]) expect(readSchema.properties).not.toHaveProperty(obsoleteField);
    const calendarReadSchema = metadata.tools.find((tool) => tool.name === "outlook_calendar_read")?.parameters as any;
    expect(calendarReadSchema.properties.action.anyOf.map((entry: any) => entry.const)).toEqual(expect.arrayContaining(["list_attachments", "download_attachment"]));
    for (const obsoleteAction of ["get_attachment", "read_staged_attachment"]) expect(calendarReadSchema.properties.action.anyOf.map((entry: any) => entry.const)).not.toContain(obsoleteAction);
    expect(calendarReadSchema.properties).not.toHaveProperty("contentBytes");
    expect(writeSchema.properties.action.anyOf.map((entry: any) => entry.const)).toEqual(expect.arrayContaining(["update_properties", "reply_all_draft", "copy", "add_attachment"]));
    for (const field of ["replyTo", "categories", "importance", "inferenceClassification", "isDeliveryReceiptRequested", "isReadReceiptRequested", "flagStatus", "flagStartDateTime", "flagDueDateTime", "internetMessageId", "internetMessageHeaders", "destinationFolderId", "attachmentMediaUri"]) expect(writeSchema.properties[field]).toBeTruthy();

    const kql = mailListQuery({ action: "search_messages", searchKql: "from:person@example.invalid AND hasAttachments:true", includeHeaders: true }, 25);
    expect(kql.query.get("$search")).toBe('"from:person@example.invalid AND hasAttachments:true"');
    expect(kql.fields).toContain("internetMessageHeaders");
    expect(kql.fields).toContain("changeKey");
    const filtered = mailListQuery({ action: "list_messages", receivedAfter: "2026-09-01T00:00:00Z", isRead: false, importance: "high", orderBy: "receivedDateTime" }, 10);
    expect(filtered.query.get("$filter")).toBe("receivedDateTime ge 2026-09-01T00:00:00Z and isRead eq false and importance eq 'high'");
    expect(filtered.query.get("$orderby")).toBe("receivedDateTime desc");
    expect(() => mailListQuery({ action: "search_messages", search: "synthetic-query", isRead: false }, 10)).toThrow("invalid_search_combination");
    expect(() => mailListQuery({ action: "list_messages", receivedAfter: "2026-02-30T12:00:00Z" }, 10)).toThrow("invalid_datetime");
  });

  it("maps stable writable mail properties without allowing sender impersonation", () => {
    expect(mailMessagePayload({ subject: "Draft", bodyText: "Text", replyTo: ["reply@example.invalid"], categories: ["Synthetic Category"], importance: "high", isReadReceiptRequested: true, flagStatus: "flagged", flagStartDateTime: "2026-09-08T09:00:00", flagDueDateTime: "2026-09-09T09:00:00", internetMessageId: "<draft@example.invalid>", internetMessageHeaders: [{ name: "X-Trace", value: "bounded" }] }, true, true)).toEqual({
      subject: "Draft", body: { contentType: "Text", content: "Text" }, replyTo: [{ emailAddress: { address: "reply@example.invalid" } }], categories: ["Synthetic Category"], importance: "high", isReadReceiptRequested: true,
      flag: { flagStatus: "flagged", startDateTime: { dateTime: "2026-09-08T09:00:00", timeZone: "UTC" }, dueDateTime: { dateTime: "2026-09-09T09:00:00", timeZone: "UTC" } },
      internetMessageHeaders: [{ name: "X-Trace", value: "bounded" }],
      internetMessageId: "<draft@example.invalid>",
    });
    expect(() => mailMessagePayload({ subject: "Not a draft" }, false)).toThrow("invalid_mail_payload");
    expect(() => mailMessagePayload({ isReadReceiptRequested: true }, false)).toThrow("invalid_mail_payload");
    expect(() => mailMessagePayload({ internetMessageHeaders: [{ name: "X-Late", value: "no" }] }, true)).toThrow("invalid_internet_headers");
    expect(() => mailMessagePayload({ flagDueDateTime: "2026-09-09T09:00:00" }, false)).toThrow("invalid_followup_flag");
    expect(mailMessagePayload({ bodyText: "Reply", bcc: ["hidden@example.invalid"], replyTo: ["reply@example.invalid"], categories: ["Blue category"], importance: "high", isDeliveryReceiptRequested: true }, true)).toEqual({
      body: { contentType: "Text", content: "Reply" },
      bccRecipients: [{ emailAddress: { address: "hidden@example.invalid" } }],
      replyTo: [{ emailAddress: { address: "reply@example.invalid" } }],
      categories: ["Blue category"],
      importance: "high",
      isDeliveryReceiptRequested: true,
    });
  });

  it("fully validates reply and forward patches before any Graph draft creation", () => {
    expect(mailReplyForwardPlan({ action: "reply_draft", bodyText: "Reply" })).toEqual({
      endpoint: "createReply",
      body: { message: { body: { contentType: "Text", content: "Reply" } } },
    });
    expect(mailReplyForwardPlan({ action: "reply_all_draft", bodyText: "Reply all" })).toEqual({ endpoint: "createReplyAll", body: { message: { body: { contentType: "Text", content: "Reply all" } } } });
    expect(mailReplyForwardPlan({ action: "forward_draft", bodyHtml: "<p>Forward</p>", to: ["person@example.invalid"] })).toEqual({ endpoint: "createForward", body: { message: { body: { contentType: "HTML", content: "<p>Forward</p>" }, toRecipients: [{ emailAddress: { address: "person@example.invalid" } }] } } });
    expect(() => mailReplyForwardPlan({ action: "reply_draft", bodyText: "Reply", categories: ["Blue category"] })).toThrow("invalid_write_parameter");
    expect(() => mailReplyForwardPlan({ action: "reply_draft", bodyText: "Reply", bodyHtml: "<p>Reply</p>" })).toThrow("invalid_body_format");
    expect(() => mailReplyForwardPlan({ action: "reply_draft", bodyText: "Reply", internetMessageHeaders: [{ name: "X-Late", value: "no" }] })).toThrow("invalid_write_parameter");
    expect(() => mailReplyForwardPlan({ action: "forward_draft", bodyText: "Forward", replyTo: ["reply@example.invalid"] })).toThrow("invalid_write_parameter");
    expect(() => mailReplyForwardPlan({ action: "forward_draft", bodyText: "Forward", to: ["person@example.invalid"], flagDueDateTime: "2026-09-09T09:00:00" })).toThrow("invalid_write_parameter");
    expect(mailMessageActionPath("message+id=", "createReplyAll")).toBe("/me/messages/message%2Bid%3D/createReplyAll");
    expect(mailMessageActionPath("message", "copy")).toBe("/me/messages/message/copy");
    expect(mailMessageActionPath("message", "attachments")).toBe("/me/messages/message/attachments");
  });

  it("rejects direct mail mutations without an approval snapshot before any network request", async () => {
    const factories: Array<(context: any) => any> = [];
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      entry.register({ pluginConfig: { enabled: true }, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
      const tool = factories[12]({ agentId: "main" });
      const response = await tool.execute("x", { action: "reply_draft", messageId: "message", bodyText: "Reply", bodyHtml: "<p>Reply</p>" });
      expect(response.details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("exposes and maps stable To Do settings, bounded search, and task relationships", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const readSchema = metadata.tools.find((tool) => tool.name === "microsoft_todo_read")?.parameters as any;
    const writeSchema = metadata.tools.find((tool) => tool.name === "microsoft_todo_write")?.parameters as any;
    expect(readSchema.properties.action.anyOf.map((entry: any) => entry.const)).toEqual(expect.arrayContaining(["search_lists", "search_tasks", "list_checklist", "list_linked_resources", "list_attachments", "get_attachment"]));
    expect(writeSchema.properties.action.anyOf.map((entry: any) => entry.const)).toEqual(expect.arrayContaining(["update_checklist", "delete_checklist", "add_linked_resource", "update_linked_resource", "delete_linked_resource", "add_attachment", "delete_attachment"]));
    for (const field of ["bodyHtml", "categories", "recurrence", "isReminderOn", "completedDateTime"]) expect(writeSchema.properties[field]).toBeTruthy();
    expect(taskPayload({ title: "Synthetic Task", bodyHtml: "<p>Synthetic details</p>", categories: ["Synthetic Category"], isReminderOn: false, completedDateTime: "2026-09-08T12:00:00" })).toEqual({ title: "Synthetic Task", body: { contentType: "html", content: "<p>Synthetic details</p>" }, categories: ["Synthetic Category"], isReminderOn: false, completedDateTime: { dateTime: "2026-09-08T12:00:00", timeZone: "UTC" } });
    expect(todoTaskMatches({ title: "Synthetic task record", body: { content: "Synthetic content" }, categories: ["Synthetic Category"], status: "inProgress", isReminderOn: false }, { search: "synthetic content", searchFields: ["body"], categories: ["Synthetic Category"], status: "inProgress", isReminderOn: false })).toBe(true);
    expect(todoTaskMatches({ title: "Synthetic task record", categories: ["Synthetic Category"] }, { search: "absent", searchFields: ["title"] })).toBe(false);
    expect(todoTaskPath("list+id=", "task+id=")).toBe("/me/todo/lists/list%2Bid%3D/tasks/task%2Bid%3D");
    expect(todoTaskChildPath("list", "task", "checklistItems", "check")).toBe("/me/todo/lists/list/tasks/task/checklistItems/check");
    expect(todoTaskChildPath("AAList=", "BBTask=", "checklistItems", "CCCheck=")).toBe("/me/todo/lists/AAList%3D/tasks/BBTask%3D/checklistItems/CCCheck%3D");
    expect(todoTaskChildPath("AAList=", "BBTask=", "checklistItems", "CCCheck=")).not.toContain("%253D");
    expect(todoTaskChildPath("list", "task", "linkedResources", "link")).toBe("/me/todo/lists/list/tasks/task/linkedResources/link");
    expect(todoTaskChildPath("list", "task", "attachments", "attachment")).toBe("/me/todo/lists/list/tasks/task/attachments/attachment");
  });

  it("covers every shared write-schema field with an explicit action allowlist", () => {
    const metadata = getToolPluginMetadata(entry)!;
    for (const [domain, toolName] of [["calendar", "outlook_calendar_write"], ["mail", "outlook_mail_write"], ["todo", "microsoft_todo_write"]] as const) {
      const schema = metadata.tools.find((tool) => tool.name === toolName)?.parameters as any;
      const schemaFields = Object.keys(schema.properties).sort();
      const matrixFields = [...new Set([...Object.values(WRITE_ACTION_FIELDS[domain]).flat(), "chatConfirmed", "chatConfirmationToken"])].sort();
      expect(matrixFields, domain).toEqual(schemaFields);
      for (const [action, allowed] of Object.entries(WRITE_ACTION_FIELDS[domain])) {
        const allowedFields = new Set(allowed);
        for (const field of matrixFields) {
          if (field === "action" || field === "chatConfirmed" || field === "chatConfirmationToken" || allowedFields.has(field)) continue;
          expect(() => assertWriteActionFields({ action, [field]: true }, WRITE_ACTION_FIELDS[domain]), `${domain}:${action}:${field}`).toThrow("invalid_write_parameter");
        }
      }
    }
  });

  it("exhaustively classifies only explicit read tools as approval-free", () => {
    const metadata = getToolPluginMetadata(entry)!;
    const readTools = new Set(["onedrive_search", "onedrive_list", "onedrive_read", "onedrive_download", "onedrive_agents_instructions", "outlook_calendar_read", "outlook_mail_read", "microsoft_todo_read"]);
    const oneDriveWrites = new Set(["onedrive_upload", "onedrive_update", "onedrive_metadata_update", "onedrive_create_folder", "onedrive_delete"]);
    for (const tool of metadata.tools) {
      const actions = (tool.parameters as any).properties?.action?.anyOf?.map((value: any) => value.const) ?? [undefined];
      for (const action of actions) {
        const level = classifyApproval(tool.name, action === undefined ? {} : { action });
        if (readTools.has(tool.name)) expect(level, `${tool.name}:${action ?? "call"}`).toBe("none");
        else {
          expect(oneDriveWrites.has(tool.name) || tool.name.endsWith("_write"), tool.name).toBe(true);
          expect(level, `${tool.name}:${action ?? "call"}`).not.toBe("none");
        }
      }
    }
  });

  it("fails closed for unknown or missing actions on current write tools", () => {
    for (const toolName of ["outlook_calendar_write", "outlook_mail_write", "microsoft_todo_write"]) {
      expect(classifyApproval(toolName, { action: "future_action" })).toBe("warning");
      expect(classifyApproval(toolName, {})).toBe("warning");
    }
    for (const toolName of ["outlook_calendar_read", "outlook_mail_read", "microsoft_todo_read"]) {
      expect(classifyApproval(toolName, { action: "future_action" })).toBe("warning");
      expect(classifyApproval(toolName, {})).toBe("warning");
    }
  });

  it("exports a sanitized policy and classifier-derived approval inventory", async () => {
    const inventory = await approvalInventory(graphPolicyFixture());
    expect(inventory.roots.length).toBeGreaterThan(0);
    expect(inventory.services.map((service) => service.service)).toEqual(["calendar", "mail", "todo"]);
    const calendarGrants = inventory.services.find((service) => service.service === "calendar")?.grants;
    const mainCalendar = calendarGrants?.find((grant) => grant.agentId === "main")?.resources.find((resource) => resource !== "me");
    expect(mainCalendar).toBeTruthy();
    expect(calendarGrants).toContainEqual(expect.objectContaining({ agentId: "secondary-agent", operations: ["read", "create", "update", "respond", "attach", "delete"], resources: ["me", mainCalendar] }));
    expect(inventory.approvals).toContainEqual({ tool: "onedrive_delete", condition: "delete", level: "critical" });
    expect(inventory.approvals).toContainEqual({ tool: "onedrive_agents_instructions", condition: "read", level: "none" });
    expect(inventory.approvals).toContainEqual({ tool: "outlook_calendar_write", action: "create", level: "warning" });
    expect(inventory.confirmationModes).toEqual({ none: "none", warning: "native-plugin-approval", critical: "native-plugin-approval" });
    for (const tool of ["outlook_calendar_write", "outlook_mail_write", "microsoft_todo_write"]) {
      expect(inventory.approvals).toContainEqual({ tool, condition: "unknown_or_missing_action", level: "warning" });
    }
    expect(JSON.stringify(inventory)).not.toMatch(/credentials|secret|token|password/i);
  });

  it("uses configurable native warning approval and process-scoped allow-always trust", async () => {
    const register = (pluginConfig: Record<string, unknown>) => {
      const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
      entry.register({ pluginConfig, registerTool: vi.fn(), on: (name: string, handler: any) => { hooks[name] = handler; }, logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
      return hooks;
    };
    const context = { agentId: "main", sessionId: "native-approval" };
    const warningEvent = { toolName: "outlook_mail_write", toolCallId: "warning-call", params: { action: "mark_read", messageId: "message-1", isRead: true } };

    for (const pluginConfig of [{}, { warningApprovalsRequired: true }]) {
      const request = await register(pluginConfig).before_tool_call(warningEvent, context);
      expect(request.requireApproval).toMatchObject({
        severity: "warning",
        allowedDecisions: ["allow-once", "allow-always", "deny"],
      });
      expect(request.requireApproval.description).not.toContain("message-1");
    }

    const disabled = await register({ warningApprovalsRequired: false }).before_tool_call(warningEvent, context);
    expect(disabled).toEqual({ params: warningEvent.params });
    const serviceScoped = register({ policy: { ...graphPolicyFixture(), rules: { default: "deny", warningApprovalsByService: { mail: false, todo: true } } } });
    expect(await serviceScoped.before_tool_call(warningEvent, context)).toEqual({ params: warningEvent.params });
    expect((await serviceScoped.before_tool_call({ toolName: "microsoft_todo_write", toolCallId: "scoped-todo", params: { action: "create_list", title: "List" } }, context)).requireApproval).toMatchObject({ severity: "warning" });
    expect((await serviceScoped.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "scoped-critical", params: { action: "send_draft", messageId: "message-1" } }, context)).requireApproval).toMatchObject({ severity: "critical", allowedDecisions: ["allow-once", "deny"] });

    const hooks = register({ warningApprovalsRequired: true });
    const first = await hooks.before_tool_call(warningEvent, context);
    first.requireApproval.onResolution("allow-once");
    expect((await hooks.before_tool_call(warningEvent, context)).requireApproval).toBeTruthy();
    first.requireApproval.onResolution("allow-always");
    expect(await hooks.before_tool_call(warningEvent, context)).toEqual({ params: warningEvent.params });
    expect((await hooks.before_tool_call({ ...warningEvent, toolCallId: "other-agent-call" }, { ...context, agentId: "other" })).requireApproval).toBeTruthy();
    expect((await hooks.before_tool_call({ ...warningEvent, toolCallId: "move-call", params: { action: "move", messageId: "message-1", destination: "archive" } }, context)).requireApproval).toBeTruthy();
    expect((await hooks.before_tool_call({ toolName: "microsoft_todo_write", toolCallId: "todo-call", params: { action: "create_list", title: "List" } }, context)).requireApproval).toBeTruthy();

    const legacy = await register({ warningApprovalsRequired: true }).before_tool_call({
      ...warningEvent,
      toolCallId: "legacy-call",
      params: { ...warningEvent.params, chatConfirmed: true, chatConfirmationToken: `mgw1_${"A".repeat(43)}` },
    }, context);
    expect(legacy.requireApproval).toMatchObject({ severity: "warning", allowedDecisions: ["allow-once", "allow-always", "deny"] });

    const critical = await hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "critical-call", params: { action: "send_draft", messageId: "message-1", chatConfirmed: true, chatConfirmationToken: `mgw1_${"A".repeat(43)}` } }, context);
    expect(critical.requireApproval).toMatchObject({ severity: "critical", allowedDecisions: ["allow-once", "deny"] });
    expect(critical.requireApproval.onResolution).toEqual(expect.any(Function));
  });

  it("consumes exact execution snapshots once and fails closed without evicting live bindings", () => {
    const snapshots = new NativeApprovalSnapshotStore(1);
    snapshots.record("call-1", { agentId: "main", sessionId: "session-1", toolName: "outlook_mail_write", params: JSON.stringify({ action: "mark_read" }) });
    expect(() => snapshots.record("call-2", { agentId: "main", sessionId: "session-1", toolName: "outlook_mail_write", params: JSON.stringify({ action: "send_draft" }) })).toThrow("approval_context_capacity_exceeded");
    expect(snapshots.consume("call-1", "main", "session-1", "outlook_mail_write", { action: "send_draft" })).toBe(false);
    expect(snapshots.consume("call-1", "main", "session-1", "outlook_mail_write", { action: "mark_read" })).toBeUndefined();
  });

  it.each([
    ["wrong call ID", "wrong-call", "main", "session-1", "outlook_mail_write", { action: "mark_read", messageId: "message-1", isRead: true }, undefined],
    ["missing call ID", "", "main", "session-1", "outlook_mail_write", { action: "mark_read", messageId: "message-1", isRead: true }, undefined],
    ["cross-agent", "bound-call", "other", "session-1", "outlook_mail_write", { action: "mark_read", messageId: "message-1", isRead: true }, false],
    ["cross-session", "bound-call", "main", "session-2", "outlook_mail_write", { action: "mark_read", messageId: "message-1", isRead: true }, false],
    ["cross-tool", "bound-call", "main", "session-1", "microsoft_todo_write", { action: "mark_read", messageId: "message-1", isRead: true }, false],
    ["cross-action", "bound-call", "main", "session-1", "outlook_mail_write", { action: "send_draft", messageId: "message-1" }, false],
    ["cross-params", "bound-call", "main", "session-1", "outlook_mail_write", { action: "mark_read", messageId: "message-2", isRead: true }, false],
  ] as const)("rejects an approval snapshot on %s", (_case, callId, agentId, sessionId, toolName, params, expected) => {
    const snapshots = new NativeApprovalSnapshotStore();
    snapshots.record("bound-call", {
      agentId: "main",
      sessionId: "session-1",
      toolName: "outlook_mail_write",
      params: JSON.stringify({ action: "mark_read", isRead: true, messageId: "message-1" }),
    });
    expect(snapshots.consume(callId, agentId, sessionId, toolName, params)).toBe(expected);
    if (callId === "bound-call") {
      expect(snapshots.consume("bound-call", "main", "session-1", "outlook_mail_write", { action: "mark_read", messageId: "message-1", isRead: true })).toBeUndefined();
    }
  });

  it.each([
    { action: "mark_read", messageId: "message-1", isRead: true },
    { action: "send_draft", messageId: "draft-1" },
  ])("blocks $action approval when the host omits toolCallId", async (params) => {
    expect(await beforeMicrosoftGraphToolCall(
      {},
      { toolName: "outlook_mail_write", params },
      { agentId: "main", sessionId: "missing-tool-call-id" },
    )).toEqual({ block: true, blockReason: "approval_context_tool_call_id_required" });
  });

  it("consumes an approved mutation snapshot exactly once at tool execution", async () => {
    const hooks: Record<string, (...args: any[]) => Promise<any> | any> = {};
    const factories: Array<(context: any) => any> = [];
    entry.register({
      pluginConfig: {},
      registerTool: (factory: any) => factories.push(factory),
      on: (name: string, handler: any) => { hooks[name] = handler; },
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    } as any);
    const context = { agentId: "main", sessionId: "single-use-snapshot" };
    const tools = Object.fromEntries(factories.map((factory) => {
      const tool = factory(context);
      return [tool.name, tool];
    }));
    const params = { action: "mark_read", messageId: "message-1", isRead: true };
    const approval = await hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId: "single-use-call", params }, context);
    expect(approval.requireApproval).toMatchObject({ severity: "warning" });
    expect((await tools.outlook_mail_write.execute("single-use-call", params)).details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
    approval.requireApproval.onResolution("allow-once");

    expect((await tools.outlook_mail_write.execute("", params)).details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
    expect((await tools.outlook_mail_write.execute("wrong-call", params)).details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
    expect((await tools.outlook_mail_write.execute("single-use-call", params)).details).toEqual({ ok: false, error: "connector_disabled" });
    expect((await tools.outlook_mail_write.execute("single-use-call", params)).details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
  });

  it.each(["deny", "timeout", "cancelled"] as const)("does not bind execution after native approval resolution %s", async (decision) => {
    const hooks: Record<string, (...args: any[]) => Promise<any> | any> = {};
    const factories: Array<(context: any) => any> = [];
    entry.register({
      pluginConfig: {},
      registerTool: (factory: any) => factories.push(factory),
      on: (name: string, handler: any) => { hooks[name] = handler; },
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    } as any);
    const context = { agentId: "main", sessionId: `approval-${decision}` };
    const tool = factories.map((factory) => factory(context)).find((candidate) => candidate.name === "outlook_mail_write");
    const params = { action: "send_draft", messageId: "draft-1" };
    const toolCallId = `approval-${decision}`;
    const approval = await hooks.before_tool_call({ toolName: "outlook_mail_write", toolCallId, params }, context);
    approval.requireApproval.onResolution(decision);

    expect((await tool.execute(toolCallId, params)).details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
  });

  it.each(["drive_id", "item_id"] as const)("binds approved OneDrive mutations to the canonical %s", async (identityField) => {
    const policy = graphPolicyFixture();
    const root = policy.services.onedrive.allowed_roots[0];
    delete root.agents_instructions;
    root.agents.main.permissions.read = true;
    const hooks: Record<string, (...args: any[]) => Promise<any> | any> = {};
    const factories: Array<(context: any) => any> = [];
    entry.register({
      pluginConfig: { enabled: true, policy },
      registerTool: (factory: any) => factories.push(factory),
      on: (name: string, handler: any) => { hooks[name] = handler; },
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    } as any);
    const context = { agentId: "main", sessionId: `root-identity-${identityField}` };
    const tools = Object.fromEntries(factories.map((factory) => {
      const tool = factory(context);
      return [tool.name, tool];
    }));
    const params = { rootLabel: root.label, relativePath: "folder/file.txt", description: "updated" };
    const toolCallId = `root-identity-${identityField}`;
    const approval = await hooks.before_tool_call({ toolName: "onedrive_metadata_update", toolCallId, params }, context);
    expect(approval.requireApproval).toMatchObject({ severity: "warning" });
    approval.requireApproval.onResolution("allow-once");

    root[identityField] = `${root[identityField]}-changed`;
    expect((await tools.onedrive_metadata_update.execute(toolCallId, params)).details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
  });

  it("renders privacy-minimized action, target, and risk details in native approvals", async () => {
    const policy = graphPolicyFixture();
    const root = policy.services.onedrive.allowed_roots[0];
    delete root.agents_instructions;
    root.permissions = { read: true, write: true, delete: true };
    root.agents.main.permissions = { read: true, write: true, delete: true };
    const config = { enabled: true, warningApprovalsRequired: true, policy };
    const workspaceDir = await mkdtemp(join(tmpdir(), "microsoft-graph-approval-copy-"));
    await mkdir(join(workspaceDir, "media", "inbound"), { recursive: true });
    const uploadBytes = Buffer.alloc(17, 0x63);
    const uploadSha256 = createHash("sha256").update(uploadBytes).digest("hex");
    await writeFile(join(workspaceDir, "media", "inbound", "file.txt"), uploadBytes);
    const context = { agentId: "main", sessionId: "approval-copy", workspaceDir };

    const upload: any = await beforeMicrosoftGraphToolCall(config, {
      toolName: "onedrive_upload",
      toolCallId: "approval-copy-upload",
      params: { rootLabel: root.label, relativePath: `${"long-folder/".repeat(80)}file.txt`, sourceMediaUri: "media://inbound/file.txt", sourceSha256: uploadSha256, sourceByteSize: uploadBytes.byteLength },
    }, context);
    expect(upload.requireApproval.description).toContain(`OneDrive root "${root.label}", path "long-folder/`);
    expect(upload.requireApproval.description).toContain(`content SHA-256 ${uploadSha256}, 17 bytes`);
    expect(upload.requireApproval.description.length).toBeLessThanOrEqual(512);

    const calendar: any = await beforeMicrosoftGraphToolCall(config, {
      toolName: "outlook_calendar_write",
      toolCallId: "approval-copy-calendar",
      params: { action: "update", calendarId: "calendar-1", eventId: "event-1", subject: "Updated" },
    }, context);
    expect(calendar.requireApproval.description).toContain('calendar "calendar-1", event "event-1"');

    const multiwrite: any = await beforeMicrosoftGraphToolCall(config, {
      toolName: "outlook_calendar_write",
      toolCallId: "approval-copy-multiwrite",
      params: { action: "multiwrite", operations: [
        { operationId: "create-1", kind: "create", calendarId: `calendar-${"a".repeat(200)}`, subject: "One", startDateTime: "2099-01-15T08:00:00", endDateTime: "2099-01-15T09:00:00" },
        { operationId: "update-1", kind: "update", calendarId: `calendar-${"b".repeat(200)}`, eventId: "event-1", subject: "Two" },
        { operationId: "update-2", kind: "update", calendarId: `calendar-${"c".repeat(200)}`, eventId: "event-2", subject: "Three" },
      ] },
    }, context);
    expect(multiwrite.requireApproval.description).toContain("Target: 3 calendar operations across 3 calendar(s)");
    expect(multiwrite.requireApproval.description.length).toBeLessThanOrEqual(512);

    const draft: any = await beforeMicrosoftGraphToolCall(config, {
      toolName: "outlook_mail_write",
      toolCallId: "approval-copy-draft",
      params: { action: "create_draft", subject: "Private", bodyText: "Private", to: ["one@example.invalid"], cc: ["two@example.invalid"] },
    }, context);
    expect(draft.requireApproval.description).toContain("recipient count 2");
    expect(draft.requireApproval.description).not.toContain("one@example.invalid");
    expect(draft.requireApproval.description).not.toContain("two@example.invalid");

    const send: any = await beforeMicrosoftGraphToolCall(config, {
      toolName: "outlook_mail_write",
      toolCallId: "approval-copy-send",
      params: { action: "send_draft", messageId: "private-message-id" },
    }, context);
    expect(send.requireApproval.description).toContain("recipients come from the draft; recipient count is unavailable in this call");
    expect(send.requireApproval.description).not.toContain("private-message-id");

    const todo: any = await beforeMicrosoftGraphToolCall(config, {
      toolName: "microsoft_todo_write",
      toolCallId: "approval-copy-todo",
      params: { action: "create_task", listId: "list-1", title: "PRIVATE TODO TITLE" },
    }, context);
    expect(todo.requireApproval.description).toContain('To Do list "list-1", task "new task"');
    expect(todo.requireApproval.description).not.toContain("PRIVATE TODO TITLE");
    await rm(workspaceDir, { recursive: true, force: true });
  });

  it("offers the required native decisions for every declared mutation action", async () => {
    const policy = graphPolicyFixture();
    const root = policy.services.onedrive.allowed_roots[0];
    delete root.agents_instructions;
    root.permissions = { read: true, write: true, delete: true };
    root.agents.main.permissions = { read: true, write: true, delete: true };
    const workspaceDir = await mkdtemp(join(tmpdir(), "microsoft-graph-all-approvals-"));
    await mkdir(join(workspaceDir, "media", "inbound"), { recursive: true });
    const newBytes = Buffer.alloc(7, 0x61);
    const existingBytes = Buffer.alloc(8, 0x62);
    await writeFile(join(workspaceDir, "media", "inbound", "new.txt"), newBytes);
    await writeFile(join(workspaceDir, "media", "inbound", "existing.txt"), existingBytes);
    const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
    entry.register({
      pluginConfig: { enabled: true, warningApprovalsRequired: true, policy },
      registerTool: vi.fn(),
      on: (name: string, handler: any) => { hooks[name] = handler; },
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    } as any);
    const context = { agentId: "main", sessionId: "all-mutation-approvals", workspaceDir };
    const metadata = getToolPluginMetadata(entry)!;
    const oneDriveParams: Record<string, Record<string, unknown>> = {
      onedrive_upload: { rootLabel: root.label, relativePath: "new.txt", sourceMediaUri: "media://inbound/new.txt", sourceSha256: createHash("sha256").update(newBytes).digest("hex"), sourceByteSize: newBytes.byteLength },
      onedrive_update: { rootLabel: root.label, relativePath: "existing.txt", sourceMediaUri: "media://inbound/existing.txt", sourceSha256: createHash("sha256").update(existingBytes).digest("hex"), sourceByteSize: existingBytes.byteLength },
      onedrive_metadata_update: { rootLabel: root.label, relativePath: "existing.txt", name: "renamed.txt" },
      onedrive_create_folder: { rootLabel: root.label, parentRelativePath: "", name: "folder" },
      onedrive_delete: { rootLabel: root.label, relativePath: "existing.txt" },
    };
    for (const tool of metadata.tools.filter((candidate) => candidate.name in oneDriveParams || candidate.name.endsWith("_write"))) {
      const actions = (tool.parameters as any).properties?.action?.anyOf?.map((value: any) => value.const) ?? [undefined];
      for (const action of actions) {
        const calendarParams = tool.name === "outlook_calendar_write" && action === "create"
          ? { action, subject: "Synthetic fixture", startDateTime: "2099-01-15T08:00:00", endDateTime: "2099-01-15T09:00:00", timeZone: "Europe/Berlin" }
          : tool.name === "outlook_calendar_write" && action === "update"
            ? { action, eventId: "event-1", subject: "Fixture" }
            : tool.name === "outlook_calendar_write" && action === "multiwrite"
              ? { action, operations: [{ operationId: "create-1", kind: "create", subject: "Synthetic fixture", startDateTime: "2099-01-15T08:00:00", endDateTime: "2099-01-15T09:00:00", timeZone: "Europe/Berlin" }] }
              : tool.name === "outlook_calendar_write" && action === "attach"
                ? { action, eventId: "event-1" }
                : undefined;
        const params = oneDriveParams[tool.name] ?? calendarParams ?? { action };
        const level = classifyApproval(tool.name, params);
        const result = await hooks.before_tool_call({ toolName: tool.name, toolCallId: `all-approvals-${tool.name}-${action ?? "call"}`, params }, context);
        expect(result.requireApproval, `${tool.name}:${action ?? "call"}`).toMatchObject(level === "critical"
          ? { severity: "critical", allowedDecisions: ["allow-once", "deny"] }
          : { severity: "warning", allowedDecisions: ["allow-once", "allow-always", "deny"] });
        expect(result.requireApproval.description, `${tool.name}:${action ?? "call"}`).toMatch(/Action: .+\. Target: .+\. Risk: .+\./);
        expect(result.requireApproval.description.length, `${tool.name}:${action ?? "call"}`).toBeLessThanOrEqual(512);
      }
    }
    await rm(workspaceDir, { recursive: true, force: true });
  });

  it("runs managed-root checks before warning execution and keeps critical approval call-bound", async () => {
    const policy = graphPolicyFixture();
    const root = policy.services.onedrive!.allowed_roots[0];
    root.permissions = { read: true, write: true, delete: true };
    root.agents = { main: { permissions: { read: true, write: true, delete: true } } };
    const config = { enabled: true, policy };
    const context = { agentId: "main", sessionId: "approval-order", requester: { senderIsOwner: true, channel: "synthetic-channel" } };
    const credentialReader = vi.fn(async () => ({ clientId: "client", refreshToken: "refresh", tenant: "tenant", scopes: ["Files.Read", "offline_access"] }));
    const tokenExchange = vi.fn(async () => "token");
    const candidateReader = vi.fn(async () => null);
    const dependencies = { credentialReader, tokenExchange, candidateReader, cache: new OneDriveAgentsSessionCache() };

    const warningParams = { rootLabel: root.label, relativePath: "folder/file.txt", description: "updated" };
    const warning = await beforeMicrosoftGraphToolCall(config, { toolName: "onedrive_metadata_update", toolCallId: "managed-warning", params: warningParams }, context, dependencies);
    expect(warning).toMatchObject({ requireApproval: { severity: "warning", allowedDecisions: ["allow-once", "allow-always", "deny"] } });
    expect(credentialReader).toHaveBeenCalledTimes(1);
    expect(tokenExchange).toHaveBeenCalledTimes(1);
    expect(candidateReader).toHaveBeenCalled();

    credentialReader.mockClear(); tokenExchange.mockClear(); candidateReader.mockClear();
    await expect(beforeMicrosoftGraphToolCall({ ...config, warningApprovalsRequired: false }, {
      toolName: "onedrive_metadata_update",
      toolCallId: "warning-disabled-call",
      params: { ...warningParams, chatConfirmed: true, chatConfirmationToken: `mgw1_${"A".repeat(43)}` },
    }, context, { ...dependencies, cache: new OneDriveAgentsSessionCache() })).resolves.toMatchObject({ params: warningParams });
    expect(credentialReader).toHaveBeenCalledTimes(1);
    expect(tokenExchange).toHaveBeenCalledTimes(1);
    expect(candidateReader).toHaveBeenCalled();

    credentialReader.mockClear(); tokenExchange.mockClear(); candidateReader.mockClear();
    const criticalParams = { rootLabel: root.label, relativePath: "folder/file.txt" };
    const critical = await beforeMicrosoftGraphToolCall(config, { toolName: "onedrive_delete", toolCallId: "managed-critical", params: criticalParams }, context, { ...dependencies, cache: new OneDriveAgentsSessionCache() });
    expect(critical).toMatchObject({ requireApproval: { severity: "critical", allowedDecisions: ["allow-once", "deny"] } });
    expect(credentialReader).not.toHaveBeenCalled();
    expect(tokenExchange).not.toHaveBeenCalled();
    expect(candidateReader).not.toHaveBeenCalled();

    await enforceOneDriveInstructionExecution(config, context, "onedrive_delete", criticalParams, undefined, { ...dependencies, cache: new OneDriveAgentsSessionCache() });
    expect(credentialReader).toHaveBeenCalledTimes(1);
    expect(tokenExchange).toHaveBeenCalledTimes(1);
    expect(candidateReader).toHaveBeenCalled();
  });

  it("validates calendar writes before native approval and reuses exact action trust", async () => {
    expect(() => calendarApprovalCriteria({
      action: "create",
      subject: "Synthetic timezone fixture",
      startDateTime: "2099-01-15T08:00:00",
      endDateTime: "2099-01-15T09:00:00",
      timeZone: "Not/A-Timezone",
    })).toThrow("invalid_datetime_timezone");

    const hooks: Record<string, (...args: any[]) => Promise<any>> = {};
    entry.register({ pluginConfig: {}, registerTool: vi.fn(), on: (name: string, handler: any) => { hooks[name] = handler; }, logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
    const approvalContext = { agentId: "main", sessionId: "session-timezone-alias", requester: { senderIsOwner: true, channel: "synthetic-channel" } };
    const intended = {
      action: "create",
      subject: "Synthetic timezone-alias test event",
      startDateTime: "2099-01-15T08:00:00",
      endDateTime: "2099-01-15T09:00:00",
      timeZone: "Europe/Berlin",
      sensitivity: "normal",
      showAs: "busy",
      isReminderOn: true,
      reminderMinutesBeforeStart: 15,
    };
    const invalid = await hooks.before_tool_call({ toolName: "outlook_calendar_write", params: { ...intended, timeZone: "Not/A-Timezone" } }, approvalContext);
    expect(invalid).toEqual({ block: true, blockReason: "invalid_datetime_timezone" });

    const challenge = await hooks.before_tool_call({ toolName: "outlook_calendar_write", toolCallId: "timezone-challenge", params: intended }, approvalContext);
    expect(challenge.requireApproval).toMatchObject({ severity: "warning" });
    challenge.requireApproval.onResolution("allow-always");
    const equivalentAlias = await hooks.before_tool_call({
      toolName: "outlook_calendar_write",
      toolCallId: "timezone-alias-call",
      params: { ...intended, timeZone: "W. Europe Standard Time" },
    }, approvalContext);
    expect(equivalentAlias).toMatchObject({ params: { ...intended, timeZone: "W. Europe Standard Time" } });
  });

  it("fails closed before secret access without runtime identity", async () => {
    const factories: Array<(context: any) => any> = [];
    entry.register({ pluginConfig: { enabled: true, credentialVaultKey: randomBytes(32).toString("base64url"), policy: graphPolicyFixture() }, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
    const tool = factories[9]({ agentId: undefined });
    const response = await tool.execute("x", { action: "list_calendars" });
    expect(response.details).toEqual({ ok: false, error: "trusted_agent_identity_required" });
  });

  it("requires a stable session identity and rejects unauthorized AGENTS.md discovery before credentials or Graph", async () => {
    const policyValidator = vi.fn(() => graphPolicyFixture());
    const credentialReader = vi.fn(async () => ({ clientId: "client", refreshToken: "refresh", tenant: "tenant", scopes: ["Files.Read", "offline_access"] }));
    const tokenExchange = vi.fn(async () => "token");
    const candidateReader = vi.fn(async () => new TextEncoder().encode("instructions"));
    const dependencies = { policyValidator, credentialReader, tokenExchange, candidateReader, cache: new OneDriveAgentsSessionCache() };

    await expect(oneDriveAgentsInstructions({ enabled: true, policy: graphPolicyFixture() }, { agentId: "fixture-reader" }, { rootLabel: "synthetic_documents" }, undefined, dependencies)).rejects.toThrow("trusted_session_identity_required");
    expect(policyValidator).not.toHaveBeenCalled();
    await expect(oneDriveAgentsInstructions({ enabled: true, policy: graphPolicyFixture() }, { agentId: "main", sessionId: "session" }, { rootLabel: "synthetic_documents" }, undefined, dependencies)).rejects.toThrow("access_denied");
    expect(credentialReader).not.toHaveBeenCalled();
    expect(tokenExchange).not.toHaveBeenCalled();
    expect(candidateReader).not.toHaveBeenCalled();
  });

  it("rejects traversal and depth before credential or network and performs zero credential/Graph reads on a cache hit", async () => {
    const credentialReader = vi.fn(async () => ({ clientId: "client", refreshToken: "refresh", tenant: "tenant", scopes: ["Files.Read", "offline_access"] }));
    const tokenExchange = vi.fn(async () => "token");
    const candidateReader = vi.fn(async (_root, path: string) => path === "AGENTS.md" ? new TextEncoder().encode("root") : null);
    const cache = new OneDriveAgentsSessionCache();
    const dependencies = { credentialReader, tokenExchange, candidateReader, cache };
    const context = { agentId: "fixture-reader", sessionId: "session" };
    const config = { enabled: true, policy: graphPolicyFixture() };

    await expect(oneDriveAgentsInstructions(config, context, { rootLabel: "synthetic_documents", relativeDirectory: "../escape" }, undefined, dependencies)).rejects.toThrow("invalid_relative_path");
    const tooDeep = Array.from({ length: ONEDRIVE_AGENTS_MAX_DEPTH + 1 }, (_, index) => `d${index}`).join("/");
    await expect(oneDriveAgentsInstructions(config, context, { rootLabel: "synthetic_documents", relativeDirectory: tooDeep }, undefined, dependencies)).rejects.toThrow("instruction_depth_exceeded");
    expect(credentialReader).not.toHaveBeenCalled();
    expect(candidateReader).not.toHaveBeenCalled();

    const first = await oneDriveAgentsInstructions(config, context, { rootLabel: "synthetic_documents", relativeDirectory: "a" }, undefined, dependencies);
    expect(first).toMatchObject({ managed: true, cacheHit: false, instructionsIncluded: true });
    expect(credentialReader).toHaveBeenCalledTimes(1);
    expect(candidateReader).toHaveBeenCalledTimes(2);
    credentialReader.mockClear();
    tokenExchange.mockClear();
    candidateReader.mockClear();
    const second = await oneDriveAgentsInstructions(config, context, { rootLabel: "synthetic_documents", relativeDirectory: "a", acknowledgement: first.acknowledgement }, undefined, dependencies);
    expect(second).toMatchObject({ managed: true, cacheHit: true, instructionsIncluded: false });
    expect(second).not.toHaveProperty("instructions");
    expect(credentialReader).not.toHaveBeenCalled();
    expect(tokenExchange).not.toHaveBeenCalled();
    expect(candidateReader).not.toHaveBeenCalled();
  });

  it("treats roots without explicit instruction trust as unmanaged without credential or Graph access", async () => {
    const policy = graphPolicyFixture();
    delete policy.services.onedrive.allowed_roots[0].agents_instructions;
    const credentialReader = vi.fn();
    const tokenExchange = vi.fn();
    const candidateReader = vi.fn();
    const result = await oneDriveAgentsInstructions(
      { enabled: true, policy },
      { agentId: "fixture-reader", sessionId: "session" },
      { rootLabel: "synthetic_documents", relativeDirectory: "a" },
      undefined,
      { credentialReader, tokenExchange, candidateReader, cache: new OneDriveAgentsSessionCache() },
    );
    expect(result).toMatchObject({ managed: false, instructionsIncluded: false, chain: [] });
    expect(credentialReader).not.toHaveBeenCalled();
    expect(tokenExchange).not.toHaveBeenCalled();
    expect(candidateReader).not.toHaveBeenCalled();
  });

  it("automatically blocks ordinary OneDrive tools until the exact instruction chain is acknowledged", async () => {
    const cache = new OneDriveAgentsSessionCache();
    const candidateReader = vi.fn(async (_root, path: string) => new TextEncoder().encode(path));
    const dependencies = {
      credentialReader: async () => ({ clientId: "client", refreshToken: "refresh", tenant: "tenant", scopes: ["Files.Read", "offline_access"] }),
      tokenExchange: async () => "token",
      candidateReader,
      cache,
    };
    const config = { enabled: true, policy: graphPolicyFixture() };
    const context = { agentId: "fixture-reader", sessionId: "gate-session" };
    const params = { rootLabel: "synthetic_documents", relativePath: "a/file.md", mode: "text" };
    const first = await enforceOneDriveInstructionPreflight(config, context, "onedrive_read", params, undefined, dependencies);
    expect(first).toMatchObject({ block: true, blockReason: expect.stringContaining("onedrive_agents_instructions_required:") });
    const payload = JSON.parse(first!.blockReason.slice("onedrive_agents_instructions_required:".length));
    expect(payload.instructions.map((entry: any) => entry.relativePath)).toEqual(["AGENTS.md", "a/AGENTS.md"]);
    expect(payload.acknowledgement).toMatch(/^[A-Za-z0-9_-]{32}$/);
    await expect(enforceOneDriveInstructionPreflight(config, context, "onedrive_read", { ...params, agentsInstructionAck: payload.acknowledgement }, undefined, dependencies)).resolves.toBeUndefined();
    await expect(enforceOneDriveInstructionPreflight(config, context, "onedrive_read", params, undefined, dependencies)).resolves.toBeUndefined();
    expect(candidateReader).toHaveBeenCalledTimes(2);

    const deeper = await enforceOneDriveInstructionPreflight(config, context, "onedrive_read", { ...params, relativePath: "a/b/file.md" }, undefined, dependencies);
    const deeperPayload = JSON.parse(deeper!.blockReason.slice("onedrive_agents_instructions_required:".length));
    expect(deeperPayload.instructions.map((entry: any) => entry.relativePath)).toEqual(["a/b/AGENTS.md"]);
  });

  it("covers every ordinary OneDrive route with an automatically derived directory scope", () => {
    expect(oneDriveInstructionDirectories("onedrive_search", { rootLabel: "r" })).toEqual([""]);
    expect(oneDriveInstructionDirectories("onedrive_list", { rootLabel: "r", relativePath: "a/b" })).toEqual(["a/b"]);
    for (const tool of ["onedrive_read", "onedrive_download", "onedrive_upload", "onedrive_update"]) {
      expect(oneDriveInstructionDirectories(tool, { rootLabel: "r", relativePath: "a/b.txt" })).toEqual(["a"]);
    }
    expect(oneDriveInstructionDirectories("onedrive_create_folder", { rootLabel: "r", parentRelativePath: "a" })).toEqual(["a"]);
    expect(oneDriveInstructionDirectories("onedrive_metadata_update", { rootLabel: "r", relativePath: "a/file", destinationRelativePath: "b" })).toEqual(["a", "a/file", "b"]);
    expect(oneDriveInstructionDirectories("onedrive_delete", { rootLabel: "r", relativePath: "a/folder" })).toEqual(["a", "a/folder"]);
    expect(oneDriveInstructionDirectories("outlook_mail_read", {})).toBeUndefined();
  });

  it("batches every metadata move scope behind one instruction receipt and one retry", async () => {
    const cache = new OneDriveAgentsSessionCache();
    const candidateReader = vi.fn(async (_root, path: string) => new TextEncoder().encode(path));
    const dependencies = {
      credentialReader: async () => ({ clientId: "client", refreshToken: "refresh", tenant: "tenant", scopes: ["Files.Read", "offline_access"] }),
      tokenExchange: async () => "token",
      candidateReader,
      cache,
    };
    const config = { enabled: true, policy: graphPolicyFixture() };
    const context = { agentId: "fixture-reader", sessionId: "metadata-batch-session" };
    const params = { rootLabel: "synthetic_documents", relativePath: "a/folder", destinationRelativePath: "b" };
    const first = await enforceOneDriveInstructionPreflight(config, context, "onedrive_metadata_update", params, undefined, dependencies);
    const payload = JSON.parse(first!.blockReason.slice("onedrive_agents_instructions_required:".length));
    expect(payload.relativeDirectories).toEqual(["a", "a/folder", "b"]);
    expect(payload.instructions.map((entry: any) => entry.relativePath)).toEqual([
      "AGENTS.md", "a/AGENTS.md", "a/folder/AGENTS.md", "b/AGENTS.md",
    ]);
    await expect(enforceOneDriveInstructionPreflight(
      config,
      context,
      "onedrive_metadata_update",
      { ...params, agentsInstructionAck: payload.acknowledgement },
      undefined,
      dependencies,
    )).resolves.toBeUndefined();
    expect(candidateReader).toHaveBeenCalledTimes(4);
  });

  it("keeps cancelled discovery bursts inside the global provider concurrency limit", async () => {
    let active = 0;
    let peak = 0;
    const candidateReader = vi.fn(async (_root, _path, _token, _maxBytes, signal?: AbortSignal) => {
      active += 1;
      peak = Math.max(peak, active);
      try {
        return await new Promise<Uint8Array>((_resolve, reject) => signal?.addEventListener("abort", () => reject(signal.reason), { once: true }));
      } finally {
        active -= 1;
      }
    });
    const dependencies = {
      credentialReader: async () => ({ clientId: "client", refreshToken: "refresh", tenant: "tenant", scopes: ["Files.Read", "offline_access"] }),
      tokenExchange: async () => "token",
      candidateReader,
      cache: new OneDriveAgentsSessionCache(),
    };
    const controllers = Array.from({ length: 8 }, () => new AbortController());
    const calls = controllers.map((controller, index) => oneDriveAgentsInstructions(
      { enabled: true, policy: graphPolicyFixture(), maxConcurrent: 2 },
      { agentId: "fixture-reader", sessionId: `burst-${index}` },
      { rootLabel: "synthetic_documents" },
      controller.signal,
      dependencies,
    ));
    const settled = Promise.allSettled(calls);
    await vi.waitFor(() => expect(peak).toBe(2));
    for (const controller of controllers) controller.abort(new DOMException("cancelled", "AbortError"));
    await settled;
    expect(peak).toBe(2);
    expect(active).toBe(0);
    expect(candidateReader).toHaveBeenCalledTimes(2);
  });

  it("applies the whole-operation timeout to instruction discovery", async () => {
    const policy = graphPolicyFixture();
    const candidateReader = vi.fn(async (_root, _path, _token, _maxBytes, signal?: AbortSignal) => new Promise<Uint8Array | null>((_resolve, reject) => {
      const abort = () => reject(signal?.reason);
      if (signal?.aborted) abort();
      else signal?.addEventListener("abort", abort, { once: true });
    }));
    await expect(oneDriveAgentsInstructions(
      { enabled: true, policy, readOperationTimeoutMs: 20, requestTimeoutMs: 20 },
      { agentId: "fixture-reader", sessionId: "timeout-session" },
      { rootLabel: "synthetic_documents" },
      undefined,
      {
        credentialReader: async () => ({ clientId: "client", refreshToken: "refresh", tenant: "tenant", scopes: ["Files.Read", "offline_access"] }),
        tokenExchange: async () => "token",
        candidateReader,
        cache: new OneDriveAgentsSessionCache(),
      },
    )).rejects.toMatchObject({ name: "TimeoutError" });
    expect(candidateReader).toHaveBeenCalledTimes(1);
  });

  it("registers the instruction tool and clears its cache on typed session_end", async () => {
    const factories = new Map<string, (context: any) => any>();
    const hooks: Record<string, (...args: any[]) => unknown> = {};
    entry.register({
      pluginConfig: { enabled: true },
      registerTool: (factory: any, options: { name?: string }) => factories.set(options.name ?? "", factory),
      on: (name: string, handler: any) => { hooks[name] = handler; },
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    } as any);
    expect(factories.has("onedrive_agents_instructions")).toBe(true);
    const preflightBlocked = await hooks.before_tool_call(
      { toolName: "onedrive_read", params: { rootLabel: "synthetic_documents", relativePath: "a/file.md", mode: "text" } },
      { agentId: "fixture-reader" },
    );
    expect(preflightBlocked).toEqual({ block: true, blockReason: "trusted_session_identity_required" });
    const missingSessionTool = factories.get("onedrive_agents_instructions")!({ agentId: "fixture-reader" });
    expect((await missingSessionTool.execute("call", { rootLabel: "synthetic_documents", relativeDirectory: "" })).details).toEqual({ ok: false, error: "trusted_session_identity_required" });

    oneDriveAgentsSessionCache.clear();
    await oneDriveAgentsSessionCache.discover({ agentId: "fixture-reader", sessionId: "ended", rootPin: "pin", rootLabel: "root", relativeDirectory: "", load: async () => new TextEncoder().encode("root") });
    expect(oneDriveAgentsSessionCache.stats().roots).toBe(1);
    hooks.session_end({ sessionId: "ended", messageCount: 1 }, { agentId: "fixture-reader", sessionId: "ended" });
    expect(oneDriveAgentsSessionCache.stats().roots).toBe(0);
  });

  it("rejects raw provider continuation URLs pre-service across all four domains", async () => {
    const factories: Array<(context: any) => any> = [];
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      entry.register({ pluginConfig: { enabled: false }, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
      const raw = "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=raw";
      const calls = [
        [factories[1]({ agentId: "fixture-reader" }), { rootLabel: "synthetic_documents", relativePath: "", continuation: raw }],
        [factories[9]({ agentId: "main" }), { action: "list_events", startDateTime: "2026-09-01T00:00:00Z", endDateTime: "2026-09-02T00:00:00Z", continuation: raw }],
        [factories[11]({ agentId: "main" }), { action: "list_messages", continuation: raw }],
        [factories[13]({ agentId: "main" }), { action: "list_lists", continuation: raw }],
      ] as const;
      for (const [tool, params] of calls) expect((await tool.execute("x", params)).details).toEqual({ ok: false, error: "invalid_continuation" });
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("rejects calendarId where the action is not calendar-scoped before credential access", async () => {
    const factories: Array<(context: any) => any> = [];
    entry.register({ pluginConfig: {}, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
    const tool = factories[9]({ agentId: "main" });
    const response = await tool.execute("x", { action: "list_calendars", calendarId: "fixture-calendar" });
    expect(response.details).toEqual({ ok: false, error: "invalid_calendar_target" });
  });

  it("rejects invalid calendar read dates and timezones before service or credential access", async () => {
    const factories: Array<(context: any) => any> = [];
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      entry.register({ pluginConfig: {}, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
      const tool = factories[9]({ agentId: "main" });
      const malformed = "2026-02-30T09:00:00Z";
      const valid = "2026-03-01T10:00:00Z";
      const bases = [
        { action: "list_events" },
        { action: "search_events", search: "meeting" },
        { action: "get_schedule", schedules: ["person@example.invalid"] },
      ];
      for (const base of bases) {
        expect((await tool.execute("x", { ...base, startDateTime: malformed, endDateTime: valid })).details).toEqual({ ok: false, error: "invalid_datetime" });
        expect((await tool.execute("x", { ...base, startDateTime: valid, endDateTime: malformed })).details).toEqual({ ok: false, error: "invalid_datetime" });
      }
      for (const base of bases.slice(0, 2)) {
        expect((await tool.execute("x", { ...base, startDateTime: valid, endDateTime: "2026-03-01T11:00:00Z", timeZone: "Unsupported Synthetic Zone" })).details).toEqual({ ok: false, error: "invalid_datetime_timezone" });
      }
      expect((await tool.execute("x", { action: "get_schedule", schedules: ["person@example.invalid"], startDateTime: "2026-03-01T10:00:00", endDateTime: "2026-03-01T11:00:00", timeZone: "Unsupported Synthetic Zone" })).details).toEqual({ ok: false, error: "invalid_datetime_timezone" });
      expect((await tool.execute("x", { action: "get_event", eventId: "event", timeZone: "Unsupported Synthetic Zone" })).details).toEqual({ ok: false, error: "invalid_datetime_timezone" });
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("rejects direct calendar mutations without an approval snapshot before credential access", async () => {
    const factories: Array<(context: any) => any> = [];
    entry.register({ pluginConfig: {}, registerTool: (factory: any) => factories.push(factory), on: vi.fn(), logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
    const tool = factories[10]({ agentId: "main" });
    const response = await tool.execute("x", { action: "respond", eventId: "event", response: "forward" });
    expect(response.details).toEqual({ ok: false, error: "approval_context_invalid_or_changed" });
  });

  it("blocks mutations to shared or non-owned To Do lists", () => {
    expect(() => assertOwnedTodoList({ isOwner: true, isShared: false })).not.toThrow();
    expect(() => assertOwnedTodoList({ isOwner: false, isShared: true })).toThrow("access_denied");
    expect(() => assertOwnedTodoList({ isOwner: true, isShared: true })).toThrow("access_denied");
  });
});
