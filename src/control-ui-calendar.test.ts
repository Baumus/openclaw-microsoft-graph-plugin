// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";
import plugin from "./control-ui.js";

type Page = { mount(container: HTMLElement, context: { host: unknown; signal: AbortSignal }): { dispose(): void } };
const initial = { version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: { main: { operations: ["read"], resources: undefined as string[] | undefined }, other: { operations: ["read"], resources: ["me"] } } }, mail: { agents: {} }, todo: { agents: {} } } };
function mount(policy = initial, calendarReply: unknown = { defaultCalendarId: "default-id", calendars: [{ id: "default-id", name: "Main" }, { id: "team-id", name: "Team" }, { id: "family-id", name: "Family" }] }) {
  let page: Page | undefined;
  const calls: Array<{ method: string; params: unknown }> = [];
  const host = {
    locale: "de", connection: { connected: true, canAdmin: true }, agents: { rows: [{ id: "main", name: "Main" }, { id: "other", name: "Other" }] },
    subscribe: () => () => undefined,
    request: async (method: string, params: unknown) => {
      calls.push({ method, params });
      if (method === "config.get") return { hash: "same", configRevisionHash: "same", appliedConfigHash: "same", config: { plugins: { entries: { "microsoft-graph": { enabled: true, config: { credentialVaultKey: "synthetic", policy } } } } } };
      if (method === "microsoft-graph.credentials.status") return { ok: true, value: { policyVersion: 2, credential: { result: "valid" } } };
      if (method === "microsoft-graph.configuration.validate") return { valid: true, requiredScopes: ["Calendars.Read"] };
      if (method === "microsoft-graph.configuration.applicationStatus") return { activePolicyHash: null };
      if (method === "microsoft-graph.updateStatus") return { updateAvailable: false };
      if (method === "microsoft-graph.configuration.calendars") { if (calendarReply instanceof Error) throw calendarReply; return calendarReply; }
      throw new Error(`Unexpected ${method}`);
    },
    ui: { registerPage: (value: Page) => { page = value; return () => undefined; }, registerNavigation: () => () => undefined },
  };
  const stop = plugin.activate(host as never);
  const container = document.createElement("div"); const controller = new AbortController();
  const view = page!.mount(container, { host, signal: controller.signal });
  const settle = async () => { for (let i = 0; i < 80; i++) await Promise.resolve(); };
  return { container, calls, settle, dispose: () => { view.dispose(); controller.abort(); if (typeof stop === "function") stop(); } };
}
function click(container: HTMLElement, text: string) {
  const button = [...container.querySelectorAll("button")].find(item => item.textContent?.includes(text));
  expect(button).toBeTruthy(); button!.click();
}
function toggle(container: HTMLElement, text: string) {
  const label = [...container.querySelectorAll("label")].find(item => item.textContent?.includes(text));
  expect(label).toBeTruthy(); const input = label!.querySelector("input")!; input.checked = !input.checked; input.dispatchEvent(new Event("change", { bubbles: true }));
}
describe("agent calendar access editor", () => {
  it("shows the effective default-only fallback, selects another calendar for only the chosen agent, and reviews the change", async () => {
    const view = mount(); await view.settle(); click(view.container, "Dienste"); await view.settle();
    expect(view.container.textContent).toContain("Eingeschränkt · nur Standardkalender");
    expect(view.container.textContent).toContain("Team");
    toggle(view.container, "Team");
    click(view.container, "Prüfen"); await view.settle();
    expect(view.container.textContent).toContain("Main: freigegeben · Team");
    const candidate = view.calls.filter(call => call.method === "microsoft-graph.configuration.validate").at(-1)!.params as { policy: typeof initial };
    expect(candidate.policy.services.calendar.agents.main.resources).toEqual(["me", "team-id"]);
    expect(candidate.policy.services.calendar.agents.other.resources).toEqual(["me"]);
    expect(candidate.policy.services.calendar.agents.main.operations).toEqual(["read"]);
    view.dispose();
  });
  it("keeps unresolvable configured IDs visible and preserves them when another calendar is selected", async () => {
    const policy = structuredClone(initial); policy.services.calendar.agents.main.resources = ["me", "old-id"];
    const view = mount(policy); await view.settle(); click(view.container, "Dienste"); await view.settle();
    expect(view.container.textContent).toContain("Kalender-ID: old-id · nicht verifizierbar");
    toggle(view.container, "Team"); click(view.container, "Prüfen"); await view.settle();
    const candidate = view.calls.filter(call => call.method === "microsoft-graph.configuration.validate").at(-1)!.params as { policy: typeof initial };
    expect(candidate.policy.services.calendar.agents.main.resources).toEqual(["me", "old-id", "team-id"]);
    view.dispose();
  });
  it("shows an honest fallback when discovery fails and retains the effective scope", async () => {
    const view = mount(initial, new Error("offline")); await view.settle(); click(view.container, "Dienste"); await view.settle();
    expect(view.container.textContent).toContain("Eingeschränkt · nur Standardkalender");
    expect(view.container.textContent).toContain("Kalenderliste nicht ladbar");
    expect(view.container.textContent).not.toContain("von 3 Kalendern");
    view.dispose();
  });
});
