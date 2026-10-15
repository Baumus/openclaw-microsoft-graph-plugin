import { afterEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({ tokenForAuthorizedOperation: vi.fn().mockResolvedValue("synthetic-token") }));
vi.mock("./credential.js", async (original) => ({ ...await original<typeof import("./credential.js")>(), tokenForAuthorizedOperation: mocks.tokenForAuthorizedOperation }));
import { registerConfigurationUiMethods } from "./config-ui-rpc.js";

const policy = { version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: { main: { operations: ["read"] } } }, mail: { agents: {} }, todo: { agents: {} } } };
function calendarHandler() {
  const handlers = new Map<string, (arg: { params: unknown; respond: (ok: boolean, payload?: unknown, error?: unknown) => void }) => Promise<void>>();
  const api = { registerGatewayMethod: vi.fn((name, fn) => { handlers.set(name, fn); }) };
  registerConfigurationUiMethods(api, { policy, credentialVaultKey: "synthetic" }, () => "/synthetic-state");
  expect(api.registerGatewayMethod).toHaveBeenCalledWith("microsoft-graph.configuration.calendars", expect.any(Function), { scope: "operator.admin" });
  return handlers.get("microsoft-graph.configuration.calendars")!;
}
afterEach(() => { vi.restoreAllMocks(); mocks.tokenForAuthorizedOperation.mockClear(); });
describe("admin calendar-name discovery", () => {
  it("returns only IDs and names across bounded Graph pages", async () => {
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => {
      const url = String(input);
      if (url.includes("/me/calendar?")) return new Response(JSON.stringify({ id: "default-id", name: "Primary", event: "private" }), { status: 200 });
      if (url.includes("$skiptoken=page2")) return new Response(JSON.stringify({ value: [{ id: "other-id", name: "Other", events: ["private"] }] }), { status: 200 });
      return new Response(JSON.stringify({ value: [{ id: "default-id", name: "Primary" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendars?$skiptoken=page2" }), { status: 200 });
    });
    const respond = vi.fn(); await calendarHandler()({ params: {}, respond });
    expect(respond).toHaveBeenCalledWith(true, { defaultCalendarId: "default-id", calendars: [{ id: "default-id", name: "Primary" }, { id: "other-id", name: "Other" }] });
    expect(fetch).toHaveBeenCalledTimes(3);
    expect(mocks.tokenForAuthorizedOperation).toHaveBeenCalledWith(expect.objectContaining({ allowedScopes: ["Calendars.Read"] }));
  });
  it("rejects params before credential access and rejects an off-collection nextLink", async () => {
    const handler = calendarHandler(); const respond = vi.fn();
    await handler({ params: { agentId: "main" }, respond });
    expect(respond).toHaveBeenCalledWith(false, undefined, expect.objectContaining({ code: "INVALID_REQUEST" }));
    expect(mocks.tokenForAuthorizedOperation).not.toHaveBeenCalled();
    vi.spyOn(globalThis, "fetch").mockImplementation(async (input) => new Response(JSON.stringify(String(input).includes("/me/calendar?")
      ? { id: "default-id", name: "Primary" }
      : { value: [], "@odata.nextLink": "https://evil.invalid/v1.0/me/calendars?$skiptoken=secret" }), { status: 200 }));
    respond.mockClear(); await handler({ params: {}, respond });
    expect(respond).toHaveBeenCalledWith(false, undefined, { code: "INVALID_REQUEST", message: "Calendar list could not be loaded" });
  });
});
