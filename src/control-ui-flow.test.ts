// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest";
import { AccountPage } from "./account-control-ui.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";
function mount(connected = false, canAdmin = false, automatic = false, assistanceUnavailable = false, missingStatus = false) {
  const policy = graphPolicyFixture(); const calls: Array<{ method: string; params: any }> = [];
  const account = { profileId: "A", generation: 1, revision: 1, connected, shared: false, unattended: false, policy };
  const replies: Record<string, any> = {
    "automatic-migration": { automatic },
    "status": { ...(automatic ? { installation: { mode: "single-user", profileId: "A", revision: 1 }, automatic: { profileId: "A", invalidated: false } } : {}), profileId: "A", account, ceiling: policy, workers: [], warning: "Andere zugelassene Benutzer können freigegebene Aktionen auslösen. Tokens bleiben persönlich.", credential: { result: connected ? "valid" : "missing" } },
    "session-status": { binding: null, available: connected ? ["A", "shared-B"] : [] },
    "device-start": { sessionId: "sign-in-A", userCode: "ABCD-EFGH", verificationUri: "https://microsoft.com/devicelogin", expiresAt: new Date(Date.now() + 600000).toISOString() },
    "device-status": { state: "pending" }, "device-cancel": { state: "failed" }, "configure": account, "bind": {}, "disconnect": { disconnected: true },
  };
  const host = { connection: { connected: true, canAdmin }, agents: { rows: [{ id: "main", name: "Main" }], selectedId: "main" }, sessions: { rows: [{ key: "agent:main:session-A", sessionId: "session-A", agentId: "main" }], selectedKey: "agent:main:session-A" }, subscribe: () => () => {}, request: async (method: string, params: any) => {
    calls.push({ method, params }); const name = method.replace("microsoft-graph.accounts.", "");
    if (name === "configure") { account.revision++; account.policy = params.policy; account.shared = params.shared; account.unattended = params.unattended; }
    if (name === "disconnect") { account.connected = false; account.generation++; replies["status"].credential = { result: "missing" }; }
    if (!(name in replies)) throw new Error(`Unexpected method ${method}`); return replies[name];
  } };
  if (assistanceUnavailable) delete replies["automatic-migration"];
  if (missingStatus) delete replies["status"];
  const container = document.createElement("div"); document.body.append(container); const controller = new AbortController(); const view = new AccountPage(container, host as any, controller.signal);
  const settle = async () => { for (let i = 0; i < 50; i++) await Promise.resolve(); };
  const click = (label: string) => { const b = [...container.querySelectorAll("button")].find(b => b.textContent === label); expect(b).toBeDefined(); b!.click(); };
  return { container, calls, replies, account, settle, click, dispose: () => { view.dispose(); controller.abort(); container.remove(); } };
}
beforeEach(() => { vi.useRealTimers(); vi.stubGlobal("confirm", () => true); localStorage.clear(); });
describe("personal Control UI flow", () => {
  it("allows a non-admin authenticated user to manage only their own account without loading global config", async () => {
    const r = mount(); await r.settle(); expect(r.container.textContent).toContain("Mein Microsoft-Konto"); expect(r.container.textContent).toContain("nicht verbunden");
    const inputs = r.container.querySelectorAll<HTMLInputElement>("input"); inputs[0].value = "12345678-1234-1234-1234-123456789012"; inputs[0].dispatchEvent(new Event("input")); inputs[1].value = "common"; inputs[1].dispatchEvent(new Event("input"));
    r.click("Bei Microsoft anmelden"); await r.settle(); expect(r.container.textContent).toContain("ABCD-EFGH"); expect(r.container.querySelector("a")?.getAttribute("rel")).toBe("noopener noreferrer");
    expect(r.calls.every(c => c.method.startsWith("microsoft-graph.accounts."))).toBe(true); expect(r.calls.find(c => c.method.endsWith("device-start"))?.params).toEqual({ clientId: "12345678-1234-1234-1234-123456789012", tenant: "common" });
    expect(localStorage.length).toBe(0); r.dispose();
  });
  it("requires an explicit account choice and explicit sharing, preserving expected revisions", async () => {
    const r = mount(true); await r.settle(); const label = [...r.container.querySelectorAll("label")].find(l => l.textContent?.includes("Gemeinsame Nutzung"))!; const input = label.querySelector("input")!; input.checked = true; input.dispatchEvent(new Event("change"));
    r.click("Kontorechte speichern"); await r.settle(); expect(r.calls.find(c => c.method.endsWith("configure"))?.params).toMatchObject({ expectedRevision: 1, shared: true, unattended: false });
    const select = r.container.querySelector("select")!; select.value = "shared-B"; r.click("Dieses Konto für die Session verwenden"); await r.settle();
    expect(r.calls.find(c => c.method.endsWith("bind"))?.params).toEqual({ agentId: "main", sessionKey: "agent:main:session-A", sessionId: "session-A", profileId: "shared-B", expectedRevision: null }); r.dispose();
  });
  it("displays host-startup singleton migration without triggering migration on page load", async () => {
    const confirm = vi.spyOn(window, "confirm");
    const r = mount(true, true, true); await r.settle();
    expect(r.calls[0].method).toBe("microsoft-graph.accounts.status");
    expect(r.calls.some(c => c.method.endsWith("automatic-migration"))).toBe(false);
    expect(r.container.textContent).toContain("Automatische Einbenutzer-Übernahme aktiv");
    expect(r.container.textContent).toContain("keine Session- oder Worker-Konfiguration nötig");
    expect(r.container.textContent).toContain("Installationsmodus: Ein Benutzer");
    expect(r.container.textContent).toContain("keine wiederholten Benutzerverzeichnis-Abfragen");
    expect(r.container.textContent).toContain("vor dem Zugang weiterer Personen");
    expect(r.calls.some(c => /bind|configure|device-start/.test(c.method))).toBe(false);
    expect(confirm).not.toHaveBeenCalled(); expect(localStorage.length).toBe(0); r.dispose();
  });
  it("shows explicit setup fallback after assistance failure and never treats it as success", async () => {
    const r = mount(false, true, false, true); await r.settle();
    expect(r.calls.some(c => c.method.endsWith("automatic-migration"))).toBe(false);
    expect(r.container.textContent).toContain("Bei Microsoft anmelden");
    expect(r.container.textContent).not.toContain("Automatische Einbenutzer-Übernahme aktiv"); r.dispose();
  });
  it("renders fixed worker account without an editable selection; reconnect invalidates own generation", async () => {
    const r = mount(true); r.replies["session-status"] = { workerProfileId: "A", binding: null, available: ["A"] }; await r.settle();
    expect(r.container.textContent).toContain("Festes Worker-Konto: A"); expect(r.container.querySelector("select")).toBeNull();
    r.click("Trennen / neu verbinden"); await r.settle(); expect(r.calls.find(c => c.method.endsWith("disconnect"))?.params).toEqual({ expectedRevision: 1 }); expect(r.account.generation).toBe(2); expect(r.container.textContent).toContain("nicht verbunden"); r.dispose();
  });
});


describe("personal setup identity guidance", () => {
  it("does not offer personal setup or migration writes for the shared owner identity", async () => {
    const r = mount(true, true); r.replies.status.profileId = "gateway-owner"; await r.settle();
    expect(r.container.textContent).toContain("Persönliche Anmeldung erforderlich");
    expect(r.container.textContent).toContain("Einrichtung nicht abgeschlossen");
    expect(r.container.textContent).toContain("Dashboard-Token");
    expect(r.container.textContent).toContain("Einstellungen → Profil → Identität");
    expect(r.container.querySelector("input")).toBeNull();
    expect(r.calls.some(c => /provision|migrate-legacy|device-start|configure|disconnect/.test(c.method))).toBe(false); r.dispose();
  });
  it("shows missing identity as an incomplete setup with recovery rather than an endless loading screen", async () => {
    const r = mount(false, false, false, false, true); await r.settle();
    expect(r.container.textContent).toContain("Einrichtung nicht abgeschlossen");
    expect(r.container.textContent).toContain("Administratorrechte allein");
    expect([...r.container.querySelectorAll("button")].some(b => b.textContent === "Einrichtung erneut prüfen")).toBe(true); r.dispose();
  });
  it("distinguishes a confirmed personal connection from an unverified first agent request", async () => {
    const r = mount(true); await r.settle();
    expect(r.container.textContent).toContain("Bestätigtes persönliches OpenClaw-Profil: A");
    expect(r.container.textContent).toContain("Persönliche Verbindung eingerichtet");
    expect(r.container.textContent).toContain("Ein erfolgreicher Lesezugriff"); r.dispose();
  });
});
