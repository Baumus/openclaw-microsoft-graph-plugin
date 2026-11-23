// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest";
import plugin from "./control-ui.js";
import { createHash } from "node:crypto";
import { canonicalPolicy } from "./config-policy-identity.js";

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
      if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
      clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
      const old = deferredReply();
      view.responses["microsoft-graph.credentials.device-status"] = old.promise;
      await vi.advanceTimersByTimeAsync(4000); await view.settle();
      expect(view.calls.at(-1)).toMatchObject({ method: "microsoft-graph.credentials.device-status", params: { sessionId: "11111111-1111-1111-1111-111111111111" } });
      clickByText(view.container, "Anmeldung abbrechen"); await view.settle();
      view.responses["microsoft-graph.credentials.reauthorize-start"] = secondSession;
      if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
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
        if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
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
        if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
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
        if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
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
      if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
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
    clickByText(view.container, "Konten");
    expect(view.container.textContent).toContain("Microsoft-Zugang erneut autorisieren");
    expect(view.container.textContent).not.toContain("openclaw microsoft-graph");
    expect(view.container.querySelector(".mg-auth input")).toBeNull();
    expect(view.container.textContent).not.toContain("Verbindung hergestellt");
    view.responses["microsoft-graph.credentials.reauthorize-start"] = view.responses["microsoft-graph.credentials.device-start"];
    if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
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
    clickByText(view.container, "Konten");
    expect(view.container.textContent).toContain("Microsoft-Berechtigungen erneuern");
    expect(view.container.textContent).not.toContain("openclaw microsoft-graph");
    const disclosure = view.container.querySelector<HTMLDetailsElement>(".mg-reauthorize-help")!;
    disclosure.open = true; disclosure.querySelector("summary")!.focus();
    view.notify(); await view.settle();
    expect(view.container.querySelector<HTMLDetailsElement>(".mg-reauthorize-help")!.open).toBe(true);
    expect(document.activeElement?.tagName).toBe("SUMMARY");
    view.responses["microsoft-graph.credentials.reauthorize-start"] = view.responses["microsoft-graph.credentials.device-start"];
    if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
      clickByText(view.container, "Microsoft-Berechtigungen erneuern"); await view.settle();
    expect(view.calls.find(call => call.method === "microsoft-graph.credentials.reauthorize-start")?.params).toEqual({});
    view.dispose();
  });
  it("does not retain a connected status after a failed scope refresh and offers retry", async () => {
    const view = mount({ clientId, tenant });
    view.responses["microsoft-graph.credentials.status"] = { ok: true, value: { policyVersion: 2, credential: { result: "valid", missingScopes: [] } } };
    await view.settle();
    clickByText(view.container, "Konten");
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
    const view = mount(); await view.settle(); if (view.container.querySelector(".mg-steps")) clickByText(view.container, "Konten");
    clickByText(view.container, "Agent-Zugriff");
    const card = [...view.container.querySelectorAll("fieldset")].find(fieldset => fieldset.querySelector("legend")?.textContent === "Kontakte");
    expect(card).toBeTruthy();
    const toggle = card!.querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    expect(toggle.checked).toBe(false);
    expect((policy.services as Record<string, unknown>).contacts).toBeUndefined();
    toggle.checked = true; toggle.dispatchEvent(new Event("change", { bubbles: true }));
    clickByText(view.container, "Änderungen prüfen"); await view.settle();
    expect(view.calls.find(call => call.method === "microsoft-graph.configuration.validate" && (call.params as any)?.policy?.services?.contacts)?.params).toMatchObject({ policy: { services: { contacts: { agents: { main: { operations: ["read", "create", "update", "delete"] } } } } } });
    view.dispose();
  });
  it("uses saved identifiers without claiming they were verified, then shows the device-code steps", async () => {
    const view = mount({ clientId, tenant }); await view.settle(); if (view.container.querySelector(".mg-steps")) clickByText(view.container, "Konten");
    expect(view.container.textContent).toContain("Anmeldung und Freigabe prüfen wir beim Verbinden");
    expect(view.container.querySelector("input[placeholder^='0000']")).toBeNull();
    clickByText(view.container, "Mit Microsoft verbinden"); await view.settle();
    expect(view.calls.find(call => call.method === "microsoft-graph.credentials.device-start")?.params).toEqual({ clientId, tenant });
    expect(view.container.textContent).toContain("Microsoft-Anmeldung abschließen");
    expect(view.container.textContent).toContain("ABCD-EFGH");
    expect(view.container.querySelector("a[href='https://microsoft.com/devicelogin']")?.getAttribute("target")).toBe("_blank");
    view.dispose();
  });
  it.each([
    ["device_authorization_declined", "Die Anmeldung wurde bei Microsoft abgelehnt", "device_authorization_declined", true],
    ["device_authorization_transport_failed", "Verbindung zum Microsoft-Anmeldedienst", "device_authorization_transport_failed", false],
    ["device_authorization_invalid_response", "war nicht auswertbar", "device_authorization_invalid_response", false],
    ["device_authorization_token_rejected", "Microsoft hat den Token-Abruf abgelehnt", "device_authorization_token_rejected", true],
    ["device_authorization_invalid_token_fields", "fehlen gültige Felder", "device_authorization_invalid_token_fields", false],
    ["credential_vault_write_failed", "lokalen verschlüsselten Speicher", "credential_vault_write_failed", false],
    ["synthetic-private-raw-error", "Die genaue Ursache ist nicht bekannt", "nicht verfügbar", false],
  ])("shows and copies only bounded diagnostics for %s", async (error, message, code, hasEntraLink) => {
    vi.useFakeTimers();
    try {
      const view = mount({ clientId, tenant }); await view.settle(); if (view.container.querySelector(".mg-steps")) clickByText(view.container, "Konten");
      view.responses["microsoft-graph.credentials.device-status"] = { ok: true, value: { state: "failed", error } };
      clickByText(view.container, "Mit Microsoft verbinden"); await view.settle();
      await vi.advanceTimersByTimeAsync(4000); await view.settle();
      expect(view.container.textContent).toContain("Anmeldung nicht abgeschlossen");
      expect(view.container.textContent).toContain("Wir konnten die Anmeldung nicht als abgeschlossen erkennen");
      expect(view.container.textContent).toContain(message);
      const details = view.container.querySelector(".mg-auth details");
      expect(details?.textContent).toContain(`Fehlercode: ${code}`);
      expect(Boolean(details?.querySelector('a[href="https://entra.microsoft.com/"]'))).toBe(hasEntraLink);
      expect(view.container.textContent).not.toContain("synthetic-private-raw-error");
      const writeText = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue();
      clickByText(view.container, "Fehlerangaben für Admin kopieren"); await view.settle();
      expect(writeText).toHaveBeenCalledWith(expect.stringContaining(`Fehlercode: ${code}`));
      expect(writeText.mock.calls[0]?.[0]).not.toContain("ABCD-EFGH");
      expect(writeText.mock.calls[0]?.[0]).not.toContain(tenant);
      expect(writeText.mock.calls[0]?.[0]).not.toContain("synthetic-private-raw-error");
      view.dispose();
    } finally { vi.useRealTimers(); }
  });
  it("preserves expanded help, unfinished text, focus and caret on background updates", async () => {
    const view = mount(); await view.settle(); if (view.container.querySelector(".mg-steps")) clickByText(view.container, "Konten");
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
      if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
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
      if (!view.container.querySelector("#microsoft-connect")) clickByText(view.container, "Konten");
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
    const view = mount({ clientId, tenant }); await view.settle(); if (view.container.querySelector(".mg-steps")) clickByText(view.container, "Konten");
    view.host.connection.canAdmin = false; view.notify(); await view.settle();
    expect(view.container.textContent).toContain("Administratorrechte");
    expect(view.container.querySelector("button")).toBeNull();
    expect(view.calls.some(call => /credentials.(device-start|reauthorize-start)/.test(call.method))).toBe(false);
    view.dispose();
  });
  it("keeps admin setup separate and saves only valid identifiers after a successful start", async () => {
    const view = mount(); await view.settle(); if (view.container.querySelector(".mg-steps")) clickByText(view.container, "Konten");
    expect(view.container.textContent).toContain("Bitte deinen Admin um Anwendungs-ID und Verzeichnis-ID");
    expect(view.container.querySelector("details")?.textContent).toContain("Einrichtung für Admin öffnen");
    const fields = [...view.container.querySelectorAll(".mg-auth input")];
    (fields[0] as HTMLInputElement).value = clientId; (fields[1] as HTMLInputElement).value = tenant;
    clickByText(view.container, "Anmeldung starten"); await view.settle();
    expect(JSON.parse(localStorage.getItem("microsoft-graph.app-ids.v1") ?? "null")).toEqual({ clientId, tenant });
    view.dispose();
  });
});


it.each(["lost response", "pending activation", "manual reload", "already active at retry"])("guides a new named connection through scoped sign-in with %s", async (mode) => {
  vi.useFakeTimers();
  const view = mount();
  try {
    const snapshot = view.responses["config.get"] as { config: { agents?: unknown } };
    snapshot.config.agents = { entries: { main: { tools: { alsoAllow: ["microsoft_graph_capabilities", "outlook_contacts_read", "outlook_contacts_write"] } } } };
    const request = view.host.request;
    const hash = (value: unknown) => createHash("sha256").update(canonicalPolicy(value)).digest("hex");
    const entry = (view.responses["config.get"] as { config: { plugins: { entries: Record<string, { config: { policy: typeof policy; accounts?: Record<string, typeof policy>; accountLabels?: Record<string, string> } }> } } }).config.plugins.entries["microsoft-graph"]!.config;
    let connected = false;
    let active = false;
    let rejectPatch!: (error: Error) => void;
    const pendingPatch = new Promise<never>((_, reject) => { rejectPatch = reject; });
    view.host.request = async (method, raw) => {
      const params = raw as { accountId?: string; raw?: string; policy?: typeof policy };
      if (method === "config.patch") {
        view.calls.push({ method, params });
        const change = JSON.parse(params.raw!).plugins.entries["microsoft-graph"].config;
        if (change.accountLabels) {
          entry.accounts = change.accounts;
          entry.accountLabels = change.accountLabels;
          if (mode === "pending activation") return pendingPatch;
          throw new Error("connection lost after write");
        }
        Object.assign(entry.accounts!.buro!.services, change.accounts.buro.services);
        return { ok: true };
      }
      if (method === "plugins.reload") {
        const attempts = view.calls.filter(call => call.method === method).length;
        view.calls.push({ method, params });
        if (mode === "manual reload" && attempts === 0) throw new Error("plugin busy");
        active = true;
        return { ok: true };
      }
      if (method === "microsoft-graph.configuration.validate") return { valid: true, policyHash: hash(params.policy), requiredScopes: ["Contacts.ReadWrite"] };
      if (method === "microsoft-graph.configuration.applicationStatus") {
        if (params.accountId && !active) throw new Error("account not active");
        return { activePolicyHash: hash(params.accountId ? entry.accounts?.[params.accountId] : policy) };
      }
      if (method === "microsoft-graph.credentials.status" && params.accountId)
        return { ok: true, value: { credential: { result: !active ? "unavailable" : connected ? "valid" : "missing" } } };
      if (method === "microsoft-graph.credentials.device-status") {
        expect(params.accountId).toBe("buro");
        connected = true;
        return { ok: true, value: { state: "created" } };
      }
      return request(method, raw);
    };
    await view.settle();
    clickByText(view.container, "Konten");
    const form = view.container.querySelector("form")!;
    const name = form.querySelector("input")!;
    name.value = "Büro"; name.dispatchEvent(new Event("input", { bubbles: true }));
    form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
    await view.settle();
    if (mode === "pending activation") {
      expect(view.container.textContent).toContain("Verbindung wird gespeichert …");
      form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
      await vi.advanceTimersByTimeAsync(1000);
      await view.settle();
    }
    expect(view.calls.filter(call => call.method === "config.patch")).toHaveLength(1);
    expect(entry.accountLabels).toEqual({ buro: "Büro" });
    expect(view.container.querySelector<HTMLSelectElement>(".mg-account-select")?.selectedOptions[0]?.textContent).toBe("Büro");
    expect(view.container.textContent).toContain("Verbindung gespeichert · Aktivierung ausstehend");
    expect(view.container.textContent).not.toContain("Anmeldung starten");
    if (mode === "pending activation") rejectPatch(new Error("saved but activation timed out"));
    await view.settle();
    expect(view.container.textContent).toContain("Verbindung gespeichert · Aktivierung ausstehend");
    expect(view.container.textContent).not.toContain("Plugin erneut laden");
    expect(view.calls.some(call => call.method === "plugins.reload")).toBe(false);
    if (mode === "manual reload" || mode === "already active at retry") {
      await vi.advanceTimersByTimeAsync(70_000);
      expect(view.container.textContent).toContain("Du musst die Verbindung nicht noch einmal anlegen.");
      active = mode === "already active at retry";
      clickByText(view.container, "Plugin erneut laden"); await view.settle();
      const reload = { method: "plugins.reload", params: { plugins: [{ pluginId: "microsoft-graph" }], waitForDrain: true } };
      if (mode === "manual reload") {
        expect(view.calls.filter(call => call.method === "plugins.reload")).toEqual([reload]);
        expect(view.container.textContent).toContain("Verbindung gespeichert · Aktivierung ausstehend");
        expect(view.container.textContent).not.toContain("Anmeldung starten");
        await vi.advanceTimersByTimeAsync(70_000);
        clickByText(view.container, "Plugin erneut laden"); await view.settle();
      }
      expect(view.calls.filter(call => call.method === "plugins.reload")).toEqual(mode === "manual reload" ? [reload, reload] : []);
    } else {
      active = true;
      await vi.advanceTimersByTimeAsync(5000);
      await view.settle();
    }
    expect(view.container.textContent).toContain("Verbindung aktiv");
    expect(view.calls.filter(call => call.method === "config.patch")).toHaveLength(1);
    clickByText(view.container, "Zugriff festlegen");
    const contacts = [...view.container.querySelectorAll("fieldset")].find(item => item.querySelector("legend")?.textContent === "Kontakte")!;
    const toggle = contacts.querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    toggle.checked = true; toggle.dispatchEvent(new Event("change", { bubbles: true }));
    clickByText(view.container, "Änderungen prüfen"); await view.settle();
    const confirm = view.container.querySelector<HTMLInputElement>(".mg-body input[type=checkbox]")!;
    confirm.checked = true; confirm.dispatchEvent(new Event("change", { bubbles: true }));
    clickByText(view.container, "Speichern und anwenden"); await view.settle();
    expect(view.calls.filter(call => call.method === "config.patch")).toHaveLength(2);
    expect(view.container.textContent).toContain("Zu den Konten");
    clickByText(view.container, "Zu den Konten");
    const client = view.container.querySelector<HTMLInputElement>('input[placeholder^="0000"]')!;
    const directory = view.container.querySelector<HTMLInputElement>('input[placeholder="example.onmicrosoft.com"]')!;
    client.value = clientId; directory.value = tenant;
    clickByText(view.container, "Anmeldung starten"); await view.settle();
    expect(view.calls.find(call => call.method.endsWith(".device-start"))?.params).toEqual({ clientId, tenant, accountId: "buro" });
    expect(view.container.textContent).toContain("ABCD-EFGH");
    await vi.advanceTimersByTimeAsync(15_000);
    expect(view.container.textContent).toContain("Verbindung hergestellt");
    expect(entry.policy).toEqual(policy);
    expect(view.calls.filter(call => call.method === "config.patch")).toHaveLength(2);
  } finally { view.dispose(); vi.useRealTimers(); }
});

it("removes only the confirmed connection and never reuses its ID", async () => {
  vi.useFakeTimers();
  const view = mount();
  const confirm = vi.fn((_message: string) => false);
  vi.stubGlobal("confirm", confirm);
  try {
    const snap = view.responses["config.get"] as {
      hash: string; configRevisionHash: string; appliedConfigHash: string;
      config: { plugins: { entries: Record<string, { config: Record<string, unknown> }> } };
    };
    const entry = snap.config.plugins.entries["microsoft-graph"]!.config;
    const accounts: Record<string, unknown> = { office: structuredClone(policy), private: structuredClone(policy) };
    const labels: Record<string, string> = { office: "Office", private: "Private" };
    entry.accounts = accounts;
    entry.accountLabels = labels;
    const request = view.host.request;
    const reload = deferredReply();
    view.host.request = async (method, params) => {
      if (method !== "config.patch") return request(method, params);
      view.calls.push({ method, params });
      const raw = JSON.parse((params as { raw: string }).raw).plugins.entries["microsoft-graph"].config;
      if (raw.retiredAccountIds) {
        expect((params as { replacePaths: string[] }).replacePaths).toContain("plugins.entries.microsoft-graph.config.accounts.office.services.mail.agents.main.operations");
        expect(raw).toEqual({ accounts: { office: null }, accountLabels: { office: null }, retiredAccountIds: { office: true } });
        delete accounts.office;
        delete labels.office;
        entry.retiredAccountIds = raw.retiredAccountIds;
        snap.configRevisionHash = "removal-saved";
        return reload.promise;
      }
      expect(Object.keys(raw.accounts)).toEqual(["office-2"]);
      Object.assign(accounts, raw.accounts);
      Object.assign(labels, raw.accountLabels);
      return { ok: true };
    };
    await view.settle();
    clickByText(view.container, "Konten");
    expect(view.container.textContent).not.toContain("Verbindung entfernen");
    const picker = view.container.querySelector<HTMLSelectElement>(".mg-account-select")!;
    picker.value = "1";
    picker.dispatchEvent(new Event("change", { bubbles: true }));
    await view.settle();
    clickByText(view.container, "Verbindung entfernen");
    expect(confirm.mock.calls[0]?.[0]).toContain("„Office“");
    expect(view.calls.filter(call => call.method === "config.patch")).toHaveLength(0);
    confirm.mockReturnValue(true);
    clickByText(view.container, "Verbindung entfernen");
    await view.settle();
    expect(view.container.textContent).toContain("Verbindung wird entfernt");
    await vi.advanceTimersByTimeAsync(1000);
    await view.settle();
    expect(view.container.textContent).toContain("Entfernung gespeichert");
    expect(accounts).toEqual({ private: policy });
    expect(entry.policy).toEqual(policy);
    expect(view.container.querySelector<HTMLSelectElement>(".mg-account-select")?.selectedOptions[0]?.textContent).toBe("Standardverbindung");
    reload.reject(new Error("reload still pending"));
    const form = view.container.querySelector("form")!;
    const name = form.querySelector("input")!;
    name.value = "Office";
    name.dispatchEvent(new Event("input", { bubbles: true }));
    form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
    await view.settle();
    expect(labels).toEqual({ private: "Private", "office-2": "Office" });
    expect(view.calls.filter(call => call.method === "config.patch")).toHaveLength(2);
    expect(view.calls.some(call => call.method === "plugins.reload")).toBe(false);
  } finally {
    vi.unstubAllGlobals();
    view.dispose();
    vi.useRealTimers();
  }
});
