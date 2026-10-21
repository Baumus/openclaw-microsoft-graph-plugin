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
  let notify = () => undefined;
  const host = {
    locale: "de", connection: { connected: true, canAdmin: true }, agents: { rows: [{ id: "main", name: "Main" }] },
    subscribe: (listener: () => undefined) => { notify = listener; return () => undefined; },
    request: async (method: string, params: unknown) => { calls.push({ method, params }); if (!(method in responses)) throw new Error(`Unexpected RPC ${method}`); if (responses[method] instanceof Error) throw responses[method]; return responses[method]; },
    ui: { registerPage: (value: Page) => { page = value; return () => undefined; }, registerNavigation: () => () => undefined },
  };
  const stop = plugin.activate(host as never);
  const container = document.createElement("div");
  document.body.append(container);
  const controller = new AbortController();
  const view = page!.mount(container, { host, signal: controller.signal });
  const settle = async () => { for (let i = 0; i < 50; i++) await Promise.resolve(); };
  return { container, calls, responses, host, notify: () => notify(), settle, dispose: () => { view.dispose(); container.remove(); controller.abort(); if (typeof stop === "function") stop(); } };
}
function clickByText(container: HTMLElement, text: string) {
  const element = [...container.querySelectorAll("button")].find(item => item.textContent === text);
  expect(element, `Button ${text}`).toBeTruthy(); element!.click();
}
describe("guided Microsoft sign-in", () => {
  beforeEach(() => { localStorage.clear(); });
  it("guides renewal entirely in the UI with authoritative IDs even without browser storage", async () => {
    const view = mount();
    view.responses["microsoft-graph.credentials.status"] = { ok: true, value: { policyVersion: 2, credential: { result: "valid", missingScopes: ["Contacts.ReadWrite"] } } };
    await view.settle();
    expect(view.container.textContent).toContain("Microsoft-Zugang erneut autorisieren");
    expect(view.container.textContent).not.toContain("openclaw microsoft-graph");
    expect(view.container.querySelector(".mg-auth input")).toBeNull();
    expect(view.container.textContent).not.toContain("Verbindung hergestellt");
    view.responses["microsoft-graph.credentials.reauthorize-start"] = view.responses["microsoft-graph.credentials.device-start"];
    clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
    expect(view.calls.find(call => call.method === "microsoft-graph.credentials.reauthorize-start")?.params).toEqual({});
    expect(view.calls.some(call => call.method === "microsoft-graph.credentials.device-start")).toBe(false);
    expect(view.container.textContent).toContain("ABCD-EFGH");
    view.dispose();
  });
  it("offers an explicit UI renewal path without trusting stale browser identifiers", async () => {
    const view = mount({ clientId, tenant });
    view.responses["microsoft-graph.credentials.status"] = { ok: true, value: { policyVersion: 2, credential: { result: "valid", missingScopes: [] } } };
    await view.settle();
    expect(view.container.textContent).toContain("Microsoft-Berechtigungen erneuern");
    expect(view.container.textContent).not.toContain("openclaw microsoft-graph");
    const disclosure = view.container.querySelector<HTMLDetailsElement>(".mg-reauthorize-help")!;
    disclosure.open = true; disclosure.querySelector("summary")!.focus();
    view.notify(); await view.settle();
    expect(view.container.querySelector<HTMLDetailsElement>(".mg-reauthorize-help")!.open).toBe(true);
    expect(document.activeElement?.tagName).toBe("SUMMARY");
    view.responses["microsoft-graph.credentials.reauthorize-start"] = view.responses["microsoft-graph.credentials.device-start"];
    clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
    expect(view.calls.find(call => call.method === "microsoft-graph.credentials.reauthorize-start")?.params).toEqual({});
    view.dispose();
  });
  it("does not retain a connected status after a failed scope refresh and offers retry", async () => {
    const view = mount({ clientId, tenant });
    view.responses["microsoft-graph.credentials.status"] = { ok: true, value: { policyVersion: 2, credential: { result: "valid", missingScopes: [] } } };
    await view.settle();
    expect(view.container.textContent).toContain("Verbindung hergestellt");
    view.responses["microsoft-graph.credentials.status"] = new Error("gateway offline");
    clickByText(view.container, "Berechtigungen erneut prüfen"); await view.settle();
    expect(view.container.textContent).not.toContain("Verbindung hergestellt");
    expect(view.container.textContent).toContain("Der Zugangsstatus ist nicht verfügbar");
    view.responses["microsoft-graph.credentials.status"] = { ok: true, value: { policyVersion: 2, credential: { result: "valid", missingScopes: ["Contacts.ReadWrite"] } } };
    clickByText(view.container, "Berechtigungen erneut prüfen"); await view.settle();
    expect(view.container.textContent).toContain("Microsoft-Zugang erneut autorisieren");
    view.dispose();
  });
  it("shows Contacts as an optional default-deny service without changing legacy policy until selected", async () => {
    const view = mount(); await view.settle();
    clickByText(view.container, "2  Dienste");
    const card = [...view.container.querySelectorAll("fieldset")].find(fieldset => fieldset.querySelector("legend")?.textContent === "Kontakte");
    expect(card).toBeTruthy();
    const toggle = card!.querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    expect(toggle.checked).toBe(false);
    expect((policy.services as Record<string, unknown>).contacts).toBeUndefined();
    toggle.checked = true; toggle.dispatchEvent(new Event("change", { bubbles: true }));
    clickByText(view.container, "4  Prüfen"); await view.settle();
    expect(view.calls.find(call => call.method === "microsoft-graph.configuration.validate" && (call.params as any)?.policy?.services?.contacts)?.params).toMatchObject({ policy: { services: { contacts: { agents: { main: { operations: ["read", "create", "update", "delete"] } } } } } });
    view.dispose();
  });
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
  it("preserves expanded help, unfinished text, focus and caret on background updates", async () => {
    const view = mount(); await view.settle();
    const help = view.container.querySelector<HTMLDetailsElement>(".mg-auth details")!;
    help.open = true;
    const client = view.container.querySelector<HTMLInputElement>(".mg-auth input")!;
    client.value = "12345678-1234"; client.dispatchEvent(new Event("input", { bubbles: true }));
    client.focus(); client.setSelectionRange(3, 7);
    for (let i = 0; i < 3; i++) { view.notify(); await view.settle(); }
    const replacement = view.container.querySelector<HTMLInputElement>(".mg-auth input")!;
    expect(replacement.value).toBe("12345678-1234");
    expect(document.activeElement).toBe(replacement);
    expect([replacement.selectionStart, replacement.selectionEnd]).toEqual([3, 7]);
    expect(view.container.querySelector<HTMLDetailsElement>(".mg-auth details")!.open).toBe(true);
    expect(view.container.textContent).toContain("Anwendungs-ID (Client-ID)");
    expect(view.container.textContent).toContain("Verzeichnis-ID (Mandant)");
    expect(view.container.textContent).toContain("Nicht die Objekt-ID");
    clickByText(view.container, "Anmeldung starten"); await view.settle();
    expect(view.calls.some(call => call.method === "microsoft-graph.credentials.device-start")).toBe(false);
    view.dispose();
  });
  it("polls UI renewal to verified completion and handles cancellation without creating a new credential", async () => {
    vi.useFakeTimers();
    try {
      const view = mount();
      view.responses["microsoft-graph.credentials.status"] = { ok: true, value: { policyVersion: 2, credential: { result: "valid", missingScopes: ["Mail.Read"] } } };
      await view.settle();
      view.responses["microsoft-graph.credentials.reauthorize-start"] = view.responses["microsoft-graph.credentials.device-start"];
      clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
      expect(view.container.textContent).toContain("Microsoft-Anmeldung abschließen");
      view.responses["microsoft-graph.credentials.device-status"] = { ok: true, value: { state: "pending" } };
      const copy = [...view.container.querySelectorAll("button")].find(button => button.textContent === "Code kopieren")!;
      copy.focus();
      await vi.advanceTimersByTimeAsync(4000); await view.settle();
      expect(document.activeElement?.textContent).toBe("Code kopieren");
      view.responses["microsoft-graph.credentials.device-cancel"] = { ok: true, value: { state: "failed", error: "device_authorization_cancelled" } };
      clickByText(view.container, "Anmeldung abbrechen"); await view.settle();
      expect(view.container.textContent).not.toContain("ABCD-EFGH");
      expect(view.container.textContent).toContain("Die Anmeldung wurde abgebrochen");
      clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
      view.responses["microsoft-graph.credentials.status"] = { ok: true, value: { policyVersion: 2, credential: { result: "valid", missingScopes: [] } } };
      view.responses["microsoft-graph.credentials.device-status"] = { ok: true, value: { state: "created", scopes: ["Mail.Read"] } };
      await vi.advanceTimersByTimeAsync(4000); await view.settle();
      expect(view.container.textContent).toContain("Verbindung hergestellt");
      expect(view.container.textContent).not.toContain("ABCD-EFGH");
      view.dispose();
    } finally { vi.useRealTimers(); }
  });
  it("never offers credential actions to a non-admin, including after authority loss", async () => {
    const view = mount({ clientId, tenant }); await view.settle();
    view.host.connection.canAdmin = false; view.notify(); await view.settle();
    expect(view.container.textContent).toContain("Administratorrechte");
    expect(view.container.querySelector("button")).toBeNull();
    expect(view.calls.some(call => /credentials.(device-start|reauthorize-start)/.test(call.method))).toBe(false);
    view.dispose();
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
