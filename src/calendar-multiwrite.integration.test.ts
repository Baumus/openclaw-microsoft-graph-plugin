import { describe, expect, it, vi } from "vitest";

const credential = vi.hoisted(() => ({
  readCredential: vi.fn(async () => ({ clientId: "synthetic", refreshToken: "synthetic", tenant: "common", scopes: ["Calendars.ReadWrite"] })),
  exchangeRefreshToken: vi.fn(async () => "synthetic-token"),
  tokenForAuthorizedOperation: vi.fn(async () => "synthetic-token"),
}));

vi.mock("./credential.js", async () => {
  const actual = await vi.importActual<typeof import("./credential.js")>("./credential.js");
  return { ...actual, ...credential };
});

import entry, { planCalendarMultiwrite } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function calendarTool(config: Record<string, unknown> = {}) {
  const factories: Array<(context: any) => any> = [];
  entry.register({
    pluginConfig: { enabled: true, policy: graphPolicyFixture(), ...config },
    registerTool: (factory: any) => factories.push(factory),
    on: vi.fn(),
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as any);
  return factories.map((factory) => factory({ agentId: "main" })).find((tool) => tool.name === "outlook_calendar_write");
}

function createOperation(index: number) {
  return {
    operationId: `create-${index}`,
    kind: "create",
    ...(index % 2 ? { calendarId: "synthetic-calendar" } : {}),
    subject: `Synthetic ${index}`,
    startDateTime: `2026-10-${String((index % 20) + 1).padStart(2, "0")}T09:00:00`,
    endDateTime: `2026-10-${String((index % 20) + 1).padStart(2, "0")}T10:00:00`,
    timeZone: "Europe/Berlin",
  };
}

function normalizedProviderEvent(body: Record<string, any>) {
  const withFraction = (value: any) => value && typeof value.dateTime === "string"
    ? { ...value, dateTime: `${value.dateTime}.0000000` }
    : value;
  const providerBody = body.body && typeof body.body.content === "string"
    ? { ...body.body, contentType: String(body.body.contentType).toLowerCase(), content: `<html><head><style>p { margin: 0; }</style></head><body>${body.body.content}</body></html>` }
    : body.body;
  return {
    ...body,
    ...(body.start ? { start: withFraction(body.start) } : {}),
    ...(body.end ? { end: withFraction(body.end) } : {}),
    ...(providerBody ? { body: providerBody } : {}),
    ...(Array.isArray(body.categories) ? { categories: [...body.categories].reverse() } : {}),
  };
}

describe("calendar multiwrite", () => {
  it("caps one confirmed multiwrite at 100 operations in schema and runtime", () => {
    const tool = calendarTool();
    expect(tool.parameters.properties.operations.maxItems).toBe(100);
    expect(planCalendarMultiwrite(Array.from({ length: 100 }, (_, index) => createOperation(index)))).toHaveLength(100);
    expect(() => planCalendarMultiwrite(Array.from({ length: 101 }, (_, index) => createOperation(index)))).toThrow("invalid_multiwrite");
  });

  it("validates, chunks above 20, tolerates out-of-order subresponses, and preserves operation order", async () => {
    const batches: any[] = [];
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      const envelope = JSON.parse(String(init?.body));
      batches.push(envelope);
      return Response.json({
        responses: [...envelope.requests].reverse().map((request: any) => ({ status: 201, id: request.id, body: { id: `event-${batches.length}-${request.id}`, ...normalizedProviderEvent(request.body) } })),
      });
    });
    try {
      const operations = Array.from({ length: 21 }, (_, index) => createOperation(index));
      const response = await calendarTool().execute("multi", { action: "multiwrite", operations });
      expect(response.details).toMatchObject({ ok: true, action: "multiwrite", outcome: "succeeded", atomic: false, batchesIssued: 2, retryOperationIds: [] });
      expect(response.details.operations.map((entry: any) => entry.operationId)).toEqual(operations.map((entry) => entry.operationId));
      expect(response.details.operations.every((entry: any) => entry.ok && entry.verification.matched)).toBe(true);
      expect(Object.keys(response.details.dedupeTransactionIds)).toHaveLength(21);
      expect(batches.map((batch) => batch.requests.length)).toEqual([20, 1]);
      expect(fetchSpy).toHaveBeenCalledTimes(2);
      for (const batch of batches) for (const request of batch.requests) expect(request.body.transactionId).toMatch(/^openclaw-msgraph-v2-[a-f0-9]{64}$/);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("returns partial outcomes with HTTP status, retry IDs, and stable create dedupe IDs", async () => {
    const seenTransactions: string[] = [];
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      const requests = JSON.parse(String(init?.body)).requests;
      seenTransactions.push(requests[0].body.transactionId);
      return Response.json({ responses: [
        { id: "2", status: 409, body: { error: { code: "ErrorConflict", message: "provider detail must not escape" } } },
        { id: "1", status: 201, body: { id: "created", ...requests[0].body } },
      ] });
    });
    try {
      const operations = [createOperation(1), { operationId: "update-1", kind: "update", eventId: "event-1", subject: "Updated" }];
      const first = (await calendarTool().execute("partial", { action: "multiwrite", operations })).details;
      const second = (await calendarTool().execute("partial-retry", { action: "multiwrite", operations })).details;
      expect(first).toMatchObject({ ok: false, outcome: "partial", atomic: false, retryOperationIds: ["update-1"] });
      expect(first.operations[1]).toEqual({ operationId: "update-1", kind: "update", status: 409, ok: false, error: { code: "ErrorConflict" } });
      expect(JSON.stringify(first)).not.toContain("provider detail");
      expect(seenTransactions[0]).toBe(seenTransactions[1]);
      expect(first.dedupeTransactionIds["create-1"]).toBe(seenTransactions[0]);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("fails a malformed response chunk closed and does not issue later chunks", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(Response.json({ responses: [{ id: "1", status: 201, body: { id: "only-one" } }] }));
    try {
      const operations = Array.from({ length: 21 }, (_, index) => createOperation(index));
      const response = (await calendarTool().execute("malformed", { action: "multiwrite", operations })).details;
      expect(response).toMatchObject({ ok: false, outcome: "failed", atomic: false, batchesIssued: 1 });
      expect(response.operations.slice(0, 20).every((entry: any) => entry.error.code === "invalid_provider_response")).toBe(true);
      expect(response.operations[20]).toMatchObject({ status: 0, ok: false, error: { code: "not_attempted" } });
      expect(response.retryOperationIds).toEqual(operations.map((entry) => entry.operationId));
      expect(fetchSpy).toHaveBeenCalledTimes(1);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("preserves completed receipts when a later batch transport fails", async () => {
    let calls = 0;
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      calls += 1;
      if (calls === 2) throw new Error("synthetic transport failure");
      const requests = JSON.parse(String(init?.body)).requests;
      return Response.json({ responses: requests.map((request: any) => ({ id: request.id, status: 201, body: { id: `event-${request.id}`, ...request.body } })) });
    });
    try {
      const operations = Array.from({ length: 21 }, (_, index) => createOperation(index));
      const response = (await calendarTool().execute("transport", { action: "multiwrite", operations })).details;
      expect(response).toMatchObject({ ok: false, outcome: "partial", atomic: false, batchesIssued: 2, retryOperationIds: ["create-20"] });
      expect(response.operations.slice(0, 20).every((entry: any) => entry.ok)).toBe(true);
      expect(response.operations[20]).toMatchObject({ status: 0, ok: false, error: { code: "provider_unavailable" } });
      expect(fetchSpy).toHaveBeenCalledTimes(2);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("applies one whole-operation deadline across all chunks within the operation cap", async () => {
    let calls = 0;
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      calls += 1;
      const requests = JSON.parse(String(init?.body)).requests;
      if (calls === 1) return Response.json({ responses: requests.map((request: any) => ({ id: request.id, status: 201, body: { id: `event-${request.id}`, ...normalizedProviderEvent(request.body) } })) });
      return await new Promise<Response>((_resolve, reject) => {
        if (init?.signal?.aborted) reject(init.signal.reason);
        else init?.signal?.addEventListener("abort", () => reject(init.signal?.reason), { once: true });
      });
    });
    try {
      const operations = Array.from({ length: 21 }, (_, index) => createOperation(index));
      const response = (await calendarTool({ calendarMultiwriteTimeoutMs: 1000, requestTimeoutMs: 5000 }).execute("whole-timeout", { action: "multiwrite", operations })).details;
      expect(response).toMatchObject({ ok: false, outcome: "partial", atomic: false, batchesIssued: 2, retryOperationIds: ["create-20"] });
      expect(response.operations.slice(0, 20).every((entry: any) => entry.ok)).toBe(true);
      expect(response.operations[20]).toMatchObject({ status: 0, ok: false, error: { code: "request_timeout" } });
      expect(fetchSpy).toHaveBeenCalledTimes(2);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("reports a semantically differing 201 create response as applied but unverified", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      const body = JSON.parse(String(init?.body));
      return Response.json({ id: "event-1", ...body, subject: "provider mismatch", providerOnly: "hidden" }, { status: 201 });
    });
    try {
      const operation = createOperation(1);
      const response = (await calendarTool().execute("single", { action: "create", ...operation, kind: undefined, operationId: undefined })).details;
      expect(response).toMatchObject({ ok: true, applied: true, appliedButUnverified: true, action: "create", warning: "response_verification_failed", verification: { matched: false, mismatches: ["subject"] } });
      expect(response.event).not.toHaveProperty("providerOnly");
      expect(fetchSpy).toHaveBeenCalledTimes(1);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("accepts documented provider normalization in single and batched event receipts", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      const request = JSON.parse(String(init?.body));
      if (Array.isArray(request.requests)) return Response.json({ responses: request.requests.map((entry: any) => ({ id: entry.id, status: entry.method === "POST" ? 201 : 200, body: { id: `event-${entry.id}`, ...normalizedProviderEvent(entry.body) } })) });
      return Response.json({ id: "event-single", ...normalizedProviderEvent(request) }, { status: 201 });
    });
    try {
      const create = { ...createOperation(2), bodyHtml: "<p>Hello <strong>Recipient</strong></p>", categories: ["Blue", "Green"] };
      const single = (await calendarTool().execute("single-normalized", { action: "create", ...create, kind: undefined, operationId: undefined })).details;
      expect(single).toMatchObject({ ok: true, action: "create", verification: { matched: true, mismatches: [] } });

      const batch = (await calendarTool().execute("batch-normalized", { action: "multiwrite", operations: [create, { operationId: "update-normalized", kind: "update", eventId: "event-existing", bodyHtml: "<p>Updated</p>", categories: ["A", "B"] }] })).details;
      expect(batch).toMatchObject({ ok: true, outcome: "succeeded", atomic: false });
      expect(batch.operations.every((entry: any) => entry.verification.matched)).toBe(true);
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("reports meaningful HTML changes in a single create receipt as applied but unverified", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      const request = JSON.parse(String(init?.body));
      const provider = normalizedProviderEvent(request);
      provider.body.content = provider.body.content.replace("<strong>Recipient</strong>", "Recipient");
      return Response.json({ id: "event-html-mismatch", ...provider }, { status: 201 });
    });
    try {
      const create = { ...createOperation(3), bodyHtml: "<p>Hello <strong>Recipient</strong></p>" };
      const response = (await calendarTool().execute("single-html-mismatch", { action: "create", ...create, kind: undefined, operationId: undefined })).details;
      expect(response).toMatchObject({ ok: true, applied: true, appliedButUnverified: true, action: "create", warning: "response_verification_failed", verification: { matched: false, mismatches: ["body"] } });
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("keeps meaningful HTML changes visible without retrying an applied batch operation", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      const requests = JSON.parse(String(init?.body)).requests;
      return Response.json({ responses: requests.map((entry: any, index: number) => {
        const provider: any = normalizedProviderEvent(entry.body);
        if (Array.isArray(provider.attendees)) provider.attendees = [...provider.attendees].reverse().map((attendee: any) => ({ ...attendee, type: String(attendee.type).toUpperCase(), emailAddress: { ...attendee.emailAddress, address: String(attendee.emailAddress.address).toUpperCase() } }));
        if (index === 1) provider.body.content = provider.body.content.replace("<em>important</em>", "important");
        return { id: entry.id, status: entry.method === "POST" ? 201 : 200, body: { id: `event-${entry.id}`, ...provider } };
      }) });
    });
    try {
      const operations = [
        { ...createOperation(4), attendees: ["first@example.com", "second@example.com"] },
        { operationId: "update-html-mismatch", kind: "update", eventId: "event-existing", bodyHtml: "<p>This is <em>important</em></p>" },
      ];
      const response = (await calendarTool().execute("batch-html-mismatch", { action: "multiwrite", operations })).details;
      expect(response).toMatchObject({ ok: true, outcome: "applied_with_warning", retryOperationIds: [] });
      expect(response.operations[0]).toMatchObject({ ok: true, verification: { matched: true } });
      expect(response.operations[1]).toMatchObject({ ok: true, applied: true, appliedButUnverified: true, warning: "response_verification_failed", verification: { matched: false, mismatches: ["body"] } });
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("keeps escaped HTML text visible as an unverified applied single receipt", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      const request = JSON.parse(String(init?.body));
      const provider = normalizedProviderEvent(request);
      provider.body.content = provider.body.content.replace("<strong>Recipient</strong>", "&lt;strong&gt;Recipient&lt;/strong&gt;");
      return Response.json({ id: "event-escaped-html-mismatch", ...provider }, { status: 201 });
    });
    try {
      const create = { ...createOperation(5), bodyHtml: "<p>Hello <strong>Recipient</strong></p>" };
      const response = (await calendarTool().execute("single-escaped-html-mismatch", { action: "create", ...create, kind: undefined, operationId: undefined })).details;
      expect(response).toMatchObject({ ok: true, applied: true, appliedButUnverified: true, action: "create", warning: "response_verification_failed", verification: { matched: false, mismatches: ["body"] } });
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("keeps escaped HTML text visible as an unverified applied batch receipt", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      const requests = JSON.parse(String(init?.body)).requests;
      return Response.json({ responses: requests.map((entry: any) => {
        const provider: any = normalizedProviderEvent(entry.body);
        provider.body.content = provider.body.content.replace("<em>important</em>", "&lt;em&gt;important&lt;/em&gt;");
        return { id: entry.id, status: entry.method === "POST" ? 201 : 200, body: { id: `event-${entry.id}`, ...provider } };
      }) });
    });
    try {
      const operation = { operationId: "update-escaped-html-mismatch", kind: "update", eventId: "event-existing", bodyHtml: "<p>This is <em>important</em></p>" };
      const response = (await calendarTool().execute("batch-escaped-html-mismatch", { action: "multiwrite", operations: [operation] })).details;
      expect(response).toMatchObject({ ok: true, outcome: "applied_with_warning", retryOperationIds: [] });
      expect(response.operations[0]).toMatchObject({ ok: true, applied: true, appliedButUnverified: true, warning: "response_verification_failed", verification: { matched: false, mismatches: ["body"] } });
    } finally {
      fetchSpy.mockRestore();
    }
  });

  it("does not retry a 2xx batch mutation whose event body cannot be verified", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(Response.json({ responses: [{ id: "1", status: 201, body: { subject: "omitted event id" } }] }));
    try {
      const operation = createOperation(6);
      const response = (await calendarTool().execute("batch-invalid-body", { action: "multiwrite", operations: [operation] })).details;
      expect(response).toMatchObject({ ok: true, outcome: "applied_with_warning", retryOperationIds: [] });
      expect(response.operations[0]).toMatchObject({ operationId: "create-6", status: 201, ok: true, applied: true, appliedButUnverified: true, warning: "invalid_provider_response", verification: { matched: false, responseValid: false } });
      expect(JSON.stringify(response)).not.toContain("omitted event id");
    } finally {
      fetchSpy.mockRestore();
    }
  });
});
