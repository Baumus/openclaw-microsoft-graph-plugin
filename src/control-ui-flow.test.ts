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
function deferredReply() {
  let resolve!: (value: unknown) => void;
  let reject!: (reason: Error) => void;
  const promise = new Promise<unknown>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}
function renewalView() {
  const view = mount({ clientId, tenant });
  view.responses["microsoft-graph.credentials.status"] = { ok: true, value: { policyVersion: 2, credential: { result: "valid", missingScopes: ["Mail.Read"] } } };
  view.responses["microsoft-graph.credentials.reauthorize-start"] = view.responses["microsoft-graph.credentials.device-start"];
  view.responses["microsoft-graph.credentials.device-cancel"] = { ok: true, value: { state: "failed", error: "device_authorization_cancelled" } };
  return view;
}
const secondSession = {
  ok: true, value: { sessionId: "22222222-2222-2222-2222-222222222222", userCode: "JKLM-NOPQ", verificationUri: "https://microsoft.com/devicelogin", expiresAt: new Date(Date.now() + 600_000).toISOString(), scopes: ["Mail.Read"] },
};
describe("guided Microsoft sign-in", () => {
  beforeEach(() => { localStorage.clear(); });
  it.each(["failed", "rejected", "created"])("R-001 ignores delayed A status (%s) after cancellation and B start", async outcome => {
    vi.useFakeTimers();
    const view = renewalView();
    try {
      await view.settle();
      clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
      const old = deferredReply();
      view.responses["microsoft-graph.credentials.device-status"] = old.promise;
      await vi.advanceTimersByTimeAsync(4000); await view.settle();
      expect(view.calls.at(-1)).toMatchObject({ method: "microsoft-graph.credentials.device-status", params: { sessionId: "11111111-1111-1111-1111-111111111111" } });
      clickByText(view.container, "Anmeldung abbrechen"); await view.settle();
      view.responses["microsoft-graph.credentials.reauthorize-start"] = secondSession;
      clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
      const credentialReads = view.calls.filter(call => call.method === "microsoft-graph.credentials.status").length;
      if (outcome === "rejected") old.reject(new Error("synthetic delayed disconnect"));
      else old.resolve({ ok: true, value: outcome === "created" ? { state: "created", scopes: ["Mail.Read"] } : { state: "failed", error: "device_authorization_declined" } });
      await view.settle();
      expect(view.container.textContent).toContain("JKLM-NOPQ");
      expect(view.container.textContent).not.toContain("ABCD-EFGH");
      expect(view.container.textContent).not.toContain("Statusprüfung unterbrochen");
      expect(view.container.textContent).not.toContain("Anmeldung nicht abgeschlossen");
      expect(view.calls.filter(call => call.method === "microsoft-graph.credentials.status")).toHaveLength(credentialReads);
      view.responses["microsoft-graph.credentials.device-status"] = { ok: true, value: { state: "pending" } };
      await vi.advanceTimersByTimeAsync(4000); await view.settle();
      expect(view.calls.at(-1)).toMatchObject({ method: "microsoft-graph.credentials.device-status", params: { sessionId: secondSession.value.sessionId } });
      expect(view.container.textContent).toContain("JKLM-NOPQ");
      clickByText(view.container, "Anmeldung abbrechen"); await view.settle();
      expect(view.calls.at(-1)).toMatchObject({ method: "microsoft-graph.credentials.device-cancel", params: { sessionId: secondSession.value.sessionId } });
    } finally { view.dispose(); vi.useRealTimers(); }
  });
  it.each(["start", "status", "cancel"])("R-001 ignores stale %s success/rejection across authority loss and a newer session", async phase => {
    vi.useFakeTimers();
    try {
      for (const rejected of [false, true]) {
        const view = renewalView(); await view.settle();
        const old = deferredReply();
        if (phase === "start") view.responses["microsoft-graph.credentials.reauthorize-start"] = old.promise;
        clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
        if (phase === "status") {
          view.responses["microsoft-graph.credentials.device-status"] = old.promise;
          await vi.advanceTimersByTimeAsync(4000); await view.settle();
        }
        if (phase === "cancel") {
          view.responses["microsoft-graph.credentials.device-cancel"] = old.promise;
          clickByText(view.container, "Anmeldung abbrechen"); await view.settle();
        }
        view.host.connection.canAdmin = false; view.notify(); await view.settle();
        expect(view.container.textContent).not.toContain("ABCD-EFGH");
        view.host.connection.canAdmin = true; view.notify(); await view.settle();
        if (phase !== "start") {
          view.responses["microsoft-graph.credentials.device-cancel"] = { ok: true, value: { state: "failed", error: "device_authorization_cancelled" } };
          clickByText(view.container, "Anmeldung abbrechen"); await view.settle();
        }
        view.responses["microsoft-graph.credentials.reauthorize-start"] = secondSession;
        clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
        if (rejected) old.reject(new Error("synthetic stale response"));
        else old.resolve(phase === "start" ? view.responses["microsoft-graph.credentials.device-start"] : { ok: true, value: { state: "failed", error: "device_authorization_declined" } });
        await view.settle();
        expect(view.container.textContent).toContain("JKLM-NOPQ");
        expect(view.container.textContent).not.toContain("Anmeldung nicht abgeschlossen");
        expect(view.container.textContent).not.toContain("Statusprüfung unterbrochen");
        view.responses["microsoft-graph.credentials.device-status"] = { ok: true, value: { state: "pending" } };
        await vi.advanceTimersByTimeAsync(4000); await view.settle();
        expect(view.calls.at(-1)).toMatchObject({ method: "microsoft-graph.credentials.device-status", params: { sessionId: secondSession.value.sessionId } });
        clickByText(view.container, "Anmeldung abbrechen"); await view.settle();
        view.dispose();
      }
    } finally { vi.useRealTimers(); }
  });
  it.each(["start", "status", "cancel"])("R-001 ignores deferred %s results and errors after unmount", async phase => {
    vi.useFakeTimers();
    try {
      for (const rejected of [false, true]) {
        const view = renewalView(); await view.settle();
        const old = deferredReply();
        if (phase === "start") view.responses["microsoft-graph.credentials.reauthorize-start"] = old.promise;
        clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
        if (phase === "status") {
          view.responses["microsoft-graph.credentials.device-status"] = old.promise;
          await vi.advanceTimersByTimeAsync(4000); await view.settle();
        }
        if (phase === "cancel") {
          view.responses["microsoft-graph.credentials.device-cancel"] = old.promise;
          clickByText(view.container, "Anmeldung abbrechen"); await view.settle();
        }
        view.dispose();
        const calls = view.calls.length;
        if (rejected) old.reject(new Error("synthetic unmounted response"));
        else old.resolve(phase === "start" ? secondSession : { ok: true, value: { state: "created", scopes: ["Mail.Read"] } });
        await view.settle(); await vi.advanceTimersByTimeAsync(8000);
        expect(view.container.childElementCount).toBe(0);
        expect(view.calls).toHaveLength(calls);
      }
    } finally { vi.useRealTimers(); }
  });
  it("R-002 omits stale browser tenant attribution from vault-renewal diagnostics", async () => {
    vi.useFakeTimers();
    const view = renewalView();
    try {
      await view.settle();
      clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
      expect(view.calls.find(call => call.method === "microsoft-graph.credentials.reauthorize-start")?.params).toEqual({});
      await vi.advanceTimersByTimeAsync(4000); await view.settle();
      const writeText = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue();
      clickByText(view.container, "Fehlerangaben für Admin kopieren"); await view.settle();
      const diagnostic = writeText.mock.calls.at(-1)![0];
      expect(diagnostic).toContain("Fehlercode: device_authorization_declined");
      expect(diagnostic).not.toContain("Organisation:");
      expect(diagnostic).not.toContain(tenant);
      expect(diagnostic).not.toContain(clientId);
      expect(diagnostic).not.toContain("ABCD-EFGH");
    } finally { view.dispose(); vi.useRealTimers(); }
  });
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
