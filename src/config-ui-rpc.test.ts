import { describe, expect, it, vi } from "vitest";
import { registerConfigurationUiMethods } from "./config-ui-rpc.js";

const policy = { version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: { main: { operations: ["read"], resources: ["me"] } } }, mail: { agents: {} }, todo: { agents: {} } } };
describe("configuration UI validation RPC", () => {
  it("registers admin-only, validates policy, and returns scopes only", () => {
    const handlers = new Map<string, (arg: { params: unknown; respond: (ok: boolean, payload?: unknown, error?: unknown) => void }) => void>();
    const api = { registerGatewayMethod: vi.fn((name, fn, _options) => { handlers.set(name, fn); }) };
    registerConfigurationUiMethods(api);
    expect(api.registerGatewayMethod).toHaveBeenCalledWith("microsoft-graph.configuration.validate", expect.any(Function), { scope: "operator.admin" });
    const handler = handlers.get("microsoft-graph.configuration.validate")!;
    const respond = vi.fn(); handler({ params: { policy }, respond });
    expect(respond).toHaveBeenCalledWith(true, { valid: true, requiredScopes: ["Calendars.Read"] });
    respond.mockClear(); handler({ params: { policy, credentialVaultKey: "secret" }, respond });
    expect(respond).toHaveBeenCalledWith(false, undefined, { code: "INVALID_REQUEST", message: "Invalid Microsoft Graph policy" });
    respond.mockClear(); handler({ params: { policy: { ...policy, rules: { default: "allow" } } }, respond });
    expect(respond).toHaveBeenCalledWith(false, undefined, { code: "INVALID_REQUEST", message: "Invalid Microsoft Graph policy" });
    expect(api.registerGatewayMethod).toHaveBeenCalledWith("microsoft-graph.configuration.resolveFolder", expect.any(Function), { scope: "operator.admin" });
    const folder = handlers.get("microsoft-graph.configuration.resolveFolder")!;
    const denied = vi.fn();
    folder({ params: { path: "/Safe/../Escape" }, respond: denied });
    expect(denied).toHaveBeenCalledWith(false, undefined, { code: "INVALID_REQUEST", message: "OneDrive folder could not be verified" });
  });
});
