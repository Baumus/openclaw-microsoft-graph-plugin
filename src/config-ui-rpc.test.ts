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
    expect(respond).toHaveBeenCalledWith(true, { valid: true, requiredScopes: ["Calendars.Read"], policyHash: expect.stringMatching(/^[a-f0-9]{64}$/) });
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

describe("running Graph generation status", () => {
  it("captures generation version and canonical policy identity with no provider or secret access", () => {
    const handlers = new Map<string, (arg: { params: unknown; respond: ReturnType<typeof vi.fn> }) => void>();
    const api = { version: "3.12.1", registerGatewayMethod: vi.fn((name, fn) => { handlers.set(name, fn); }) };
    const config = { policy: structuredClone(policy), get credentialVaultKey(): never { throw new Error("credential access forbidden"); } };
    const fetch = vi.spyOn(globalThis, "fetch").mockRejectedValue(new Error("provider access forbidden"));
    const stateDir = vi.fn(() => { throw new Error("vault access forbidden"); });
    registerConfigurationUiMethods(api, config, stateDir);
    expect(api.registerGatewayMethod).toHaveBeenCalledWith("microsoft-graph.configuration.applicationStatus", expect.any(Function), { scope: "operator.admin" });
    const respond = vi.fn(); handlers.get("microsoft-graph.configuration.applicationStatus")!({ params: {}, respond });
    const captured = respond.mock.calls[0][1];
    expect(captured).toEqual({ activeVersion: "3.12.1", activePolicyHash: expect.stringMatching(/^[a-f0-9]{64}$/) });
    api.version = "future"; config.policy.services.calendar.agents.main.operations = [];
    respond.mockClear(); handlers.get("microsoft-graph.configuration.applicationStatus")!({ params: {}, respond });
    expect(respond).toHaveBeenCalledWith(true, captured);
    const reverseObjects = (value: unknown): unknown => Array.isArray(value) ? value.map(reverseObjects) : value && typeof value === "object" ? Object.fromEntries(Object.entries(value).reverse().map(([k,v])=>[k,reverseObjects(v)])) : value;
    respond.mockClear(); handlers.get("microsoft-graph.configuration.validate")!({ params: { policy: reverseObjects(policy) }, respond });
    expect(respond.mock.calls[0][1].policyHash).toBe(captured.activePolicyHash);
    for (const params of [null, [], { unexpected: true }]) { respond.mockClear(); handlers.get("microsoft-graph.configuration.applicationStatus")!({ params, respond }); expect(respond).toHaveBeenCalledWith(false, undefined, expect.objectContaining({ code: "INVALID_REQUEST" })); }
    expect(fetch).not.toHaveBeenCalled(); expect(stateDir).not.toHaveBeenCalled(); fetch.mockRestore();
  });
  it("returns explicit unknown identity for missing or invalid policy", () => {
    for (const policy of [undefined, { rules: { default: "allow" } }]) {
      const registerGatewayMethod=vi.fn(); registerConfigurationUiMethods({ registerGatewayMethod }, { policy });
      const handler=registerGatewayMethod.mock.calls.find(c=>c[0]==="microsoft-graph.configuration.applicationStatus")![1]; const respond=vi.fn(); handler({ params: {}, respond });
      expect(respond).toHaveBeenCalledWith(true, { activePolicyHash: null, activeVersion: null });
    }
  });
});
