// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest";
import plugin from "./control-ui.js";

type Page = { mount(container: HTMLElement, context: { host: unknown; signal: AbortSignal }): { dispose(): void } };
const clientId = "12345678-1234-1234-1234-123456789012";
const tenant = "87654321-4321-4321-4321-210987654321";
const policy = { version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: {} }, mail: { agents: { main: { operations: ["read"] } } }, todo: { agents: {} } } };
function mount(initialIds?: { clientId: string; tenant: string }) {
  if (initialIds) localStorage.setItem("microsoft-graph.app-ids.v1", JSON.stringify(initialIds));
  let page: Page | undefined;
  const calls: Array<{ method: string; params: unknown }> = [];
  const responses: Record<string, unknown> = {
    "config.get": { hash: "same", configRevisionHash: "same", appliedConfigHash: "same", config: { plugins: { entries: { "microsoft-graph": { enabled: true, config: { credentialVaultKey: "secret-ref", policy } } } } } },
    "microsoft-graph.credentials.status": { ok: true, value: { policyVersion: 2, credential: { result: "missing" } } },
    "microsoft-graph.configuration.validate": { valid: true, requiredScopes: ["Mail.Read"] },
    "microsoft-graph.updateStatus": { updateAvailable: false },
    "microsoft-graph.credentials.device-status": { ok: true, value: { state: "failed", error: "device_authorization_declined" } },
    "microsoft-graph.credentials.device-start": { ok: true, value: { sessionId: "11111111-1111-1111-1111-111111111111", userCode: "ABCD-EFGH", verificationUri: "https://microsoft.com/devicelogin", expiresAt: new Date(Date.now() + 600_000).toISOString(), scopes: ["Mail.Read"] } },
  };
  const host = {
    locale: "de", connection: { connected: true, canAdmin: true }, agents: { rows: [{ id: "main", name: "Main" }] },
    subscribe: () => () => undefined,
    request: async (method: string, params: unknown) => { calls.push({ method, params }); if (!(method in responses)) throw new Error(`Unexpected RPC ${method}`); return responses[method]; },
    ui: { registerPage: (value: Page) => { page = value; return () => undefined; }, registerNavigation: () => () => undefined },
  };
  const stop = plugin.activate(host as never);
  const container = document.createElement("div");
  const controller = new AbortController();
  const view = page!.mount(container, { host, signal: controller.signal });
  const settle = async () => { for (let i = 0; i < 50; i++) await Promise.resolve(); };
  return { container, calls, responses, settle, dispose: () => { view.dispose(); controller.abort(); if (typeof stop === "function") stop(); } };
}
function clickByText(container: HTMLElement, text: string) {
  const element = [...container.querySelectorAll("button")].find(item => item.textContent === text);
  expect(element, `Button ${text}`).toBeTruthy(); element!.click();
}
describe("guided Microsoft sign-in", () => {
  beforeEach(() => { localStorage.clear(); });
  it("uses saved identifiers without claiming they were verified, then shows the device-code steps", async () => {
    const view = mount({ clientId, tenant }); await view.settle();
    expect(view.container.textContent).toContain("Anmeldung und Freigabe prüfen wir beim Verbinden");
    expect(view.container.querySelector("input[placeholder^='0000']")).toBeNull();
    clickByText(view.container, "Mit Microsoft verbinden"); await view.settle();
    expect(view.calls.find(call => call.method === "microsoft-graph.credentials.device-start")?.params).toEqual({ clientId, tenant });
    expect(view.container.textContent).toContain("Microsoft-Anmeldung abschließen");
    expect(view.container.textContent).toContain("ABCD-EFGH");
    expect(view.container.querySelector("a[href='https://microsoft.com/devicelogin']")?.getAttribute("target")).toBe("_blank");
    view.dispose();
  });
  it("shows a user-facing failure path and copies bounded diagnostics for an administrator", async () => {
    vi.useFakeTimers();
    try {
      const view = mount({ clientId, tenant }); await view.settle();
      clickByText(view.container, "Mit Microsoft verbinden"); await view.settle();
      await vi.advanceTimersByTimeAsync(4000); await view.settle();
      expect(view.container.textContent).toContain("Anmeldung nicht abgeschlossen");
      expect(view.container.textContent).toContain("Wir konnten die Anmeldung nicht als abgeschlossen erkennen");
      expect(view.container.querySelector("details")?.textContent).toContain("Anmeldeprotokolle öffnen");
      const writeText = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue();
      clickByText(view.container, "Fehlerangaben für Admin kopieren"); await view.settle();
      expect(writeText).toHaveBeenCalledWith(expect.stringContaining("Fehlercode: device_authorization_declined"));
      expect(writeText.mock.calls[0]?.[0]).not.toContain("ABCD-EFGH");
      view.dispose();
    } finally { vi.useRealTimers(); }
  });
  it("keeps admin setup separate and saves only valid identifiers after a successful start", async () => {
    const view = mount(); await view.settle();
    expect(view.container.textContent).toContain("Bitte deinen Admin um Anwendungs-ID und Verzeichnis-ID");
    expect(view.container.querySelector("details")?.textContent).toContain("Einrichtung für Admin öffnen");
    const fields = [...view.container.querySelectorAll(".mg-auth input")];
    (fields[0] as HTMLInputElement).value = clientId; (fields[1] as HTMLInputElement).value = tenant;
    clickByText(view.container, "Anmeldung starten"); await view.settle();
    expect(JSON.parse(localStorage.getItem("microsoft-graph.app-ids.v1") ?? "null")).toEqual({ clientId, tenant });
    view.dispose();
  });
});
