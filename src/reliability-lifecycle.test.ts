import { beforeEach, expect, it, vi } from "vitest";
const health = vi.hoisted(() => vi.fn());
vi.mock("./service-health.js", async () => ({ ...await vi.importActual("./service-health.js"), automaticServiceHealth: health }));
import { concrete } from "./index.js";
import { graphRequest } from "./graph.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn() };
const config = { enabled: true, policy: graphPolicyFixture() } as any;
beforeEach(() => { vi.clearAllMocks(); health.mockResolvedValue({ status: "not_authorized", evidence: "observed_request_failure" }); });
it("automatically enriches provider failures, preserving legacy contract", async () => {
  const tool = concrete("outlook_mail_read", {}, "main", "synthetic", logger, async () => { throw Object.assign(new Error("provider_access_denied"), { providerStatus: 403 }); }, config);
  const result = (await tool.execute("health-failure", { action: "list_messages" })).details as any;
  expect(result).toMatchObject({ error: "provider_access_denied", code: "provider_access_denied", failureCategory: "authorization", httpStatus: 403, serviceHealth: { status: "not_authorized" }, mutationApplied: false });
  expect(health).toHaveBeenCalledTimes(1);
});
it("never probes on success or local policy denial", async () => {
  await concrete("outlook_mail_read", {}, "main", "synthetic", logger, async () => ({ ok: true }), config).execute("success", { action: "list_messages" });
  await concrete("outlook_mail_read", {}, "main", "synthetic", logger, async () => { throw new Error("access_denied"); }, config).execute("denied", { action: "list_messages" });
  expect(health).not.toHaveBeenCalled();
});
it("preserves partial batch successes and decorates failing operations only", async () => {
  const tool = concrete("outlook_calendar_read", {}, "main", "synthetic", logger, async () => ({ ok: false, operations: [{ ok: true, status: 201, applied: true }, { ok: false, status: 503 }] }), config);
  const result = (await tool.execute("batch", { action: "list_events" })).details as any;
  expect(result.operations[0]).not.toHaveProperty("serviceHealth");
  expect(result.operations[1]).toMatchObject({ failureCategory: "transient_provider", serviceHealth: { status: "not_authorized" } });
});

it("delivers allowlisted attachment HTTP400 diagnosis through the connector boundary without private provider text", async () => {
  const fetchFn = vi.fn(async () => Response.json({ error: { code: "Request_BadRequest", message: "private@example.com", innerError: { requestId: "private-id" } } }, { status: 400 }));
  const tool = concrete("outlook_mail_read", {}, "main", "synthetic", logger, async () => graphRequest("synthetic", "/me/messages/synthetic/attachments?$top=5&$select=id", {}, fetchFn), config);
  const result = (await tool.execute("attachment-diagnosis", { action: "list_attachments", messageId: "synthetic", limit: 5 })).details as any;
  expect(result).toMatchObject({ ok: false, error: "provider_error_400", providerCode: "Request_BadRequest", httpStatus: 400, attempts: 1, mutationApplied: false });
  expect(JSON.stringify(result)).not.toMatch(/private|synthetic/);
  expect(JSON.stringify(logger.info.mock.calls)).not.toMatch(/private|synthetic/);
  expect(fetchFn).toHaveBeenCalledTimes(1);
});
