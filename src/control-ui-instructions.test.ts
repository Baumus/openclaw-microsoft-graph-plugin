// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from "vitest";
import plugin from "./control-ui.js";

type Root = { label: string; path: string; drive_id: string; item_id: string; include_descendants: true; agents_instructions?: "trusted"; permissions: { read: boolean; write: boolean; delete: boolean }; agents: Record<string, { permissions: { read: boolean; write?: boolean } }> };
const root = (label: string, trusted = false): Root => ({ label, path: `/Synthetic/${label}`, drive_id: "synthetic-drive", item_id: label, include_descendants: true, ...(trusted ? { agents_instructions: "trusted" as const } : {}), permissions: { read: true, write: true, delete: false }, agents: { reader: { permissions: { read: true } }, editor: { permissions: { read: true, write: true } } } });
const initial = () => ({ version: 2, rules: { default: "deny", warningApprovalsByService: { mail: false } }, services: { onedrive: { allowed_roots: [root("Documents"), root("Projects", true)] }, calendar: { agents: { reader: { operations: ["read"], resources: ["me", "synthetic-calendar"] } } }, mail: { agents: { editor: { operations: ["read", "draft"] } } }, contacts: { agents: {} }, todo: { agents: {} } } });
type Policy = ReturnType<typeof initial>;
const cleanup: Array<() => void> = [];
afterEach(() => { for (const stop of cleanup.splice(0)) stop(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });
function mount(locale = "de", roots?: Root[]) {
  let saved = initial();
  if (roots) saved.services.onedrive.allowed_roots = roots;
  const baseline = structuredClone(saved);
  let page!: { mount(container: HTMLElement, context: unknown): { dispose(): void } };
  let listener = () => {};
  const calls: Array<{ method: string; params: Record<string, unknown> }> = [];
  const host = { locale, connection: { connected: true, canAdmin: true }, agents: { rows: [{ id: "reader", name: "Reader" }, { id: "editor", name: "Editor" }] },
    subscribe: (fn: () => void) => { listener = fn; return () => {}; },
    request: async (method: string, params: Record<string, unknown>) => {
      calls.push({ method, params: structuredClone(params) });
      if (method === "config.get") return { hash: "synthetic-revision", configRevisionHash: "same", appliedConfigHash: "same", config: { plugins: { entries: { "microsoft-graph": { enabled: true, config: { credentialVaultKey: "synthetic-ref", policy: structuredClone(saved) } } } } } };
      if (method === "microsoft-graph.credentials.status") return { ok: true, value: { policyVersion: 2, credential: { result: "valid" } } };
      if (method === "microsoft-graph.configuration.validate") return { valid: true, requiredScopes: [] };
      if (method === "microsoft-graph.configuration.applicationStatus") return { activePolicyHash: null };
      if (method === "microsoft-graph.updateStatus") return { updateAvailable: false };
      if (method === "microsoft-graph.configuration.resolveFolder") return { path: "/Synthetic/New", drive_id: "synthetic-drive", item_id: "new-item" };
      if (method === "config.patch") { const patch = JSON.parse(params.raw as string).plugins.entries["microsoft-graph"].config.policy; saved.services.onedrive.allowed_roots = patch.services.onedrive.allowed_roots; return {}; }
      throw new Error(`Unexpected RPC ${method}`);
    }, ui: { registerPage: (value: typeof page) => { page = value; return () => {}; }, registerNavigation: () => () => {} } };
  const stop = plugin.activate(host as never);
  const container = document.createElement("div"); document.body.append(container);
  const controller = new AbortController();
  const view = page.mount(container, { host, signal: controller.signal });
  cleanup.push(() => { view.dispose(); controller.abort(); if (typeof stop === "function") stop(); container.remove(); });
  const settle = async () => { for (let i = 0; i < 100; i++) await Promise.resolve(); };
  const click = (text: string) => { const node = [...container.querySelectorAll("button")].find(button => button.textContent?.includes(text)); expect(node, text).toBeTruthy(); node!.click(); };
  const toggle = (index: number) => { const input = container.querySelectorAll<HTMLInputElement>(".mg-instructions-control input")[index]!; input.focus(); input.click(); };
  const candidate = async () => { click("Prüfen"); await settle(); return (calls.filter(call => call.method === "microsoft-graph.configuration.validate").at(-1)!.params as { policy: Policy }).policy; };
  return { container, calls, baseline, settle, click, toggle, candidate, notify: () => listener(), saved: () => structuredClone(saved) };
}
describe("OneDrive root instruction opt-in", () => {
  it("renders one root-wide checkbox per card, defaults off, and preserves preexisting trusted values", async () => {
    const h = mount(); await h.settle();
    expect([...h.container.querySelectorAll<HTMLInputElement>(".mg-instructions-control input")].map(input => input.checked)).toEqual([false, true]);
    expect(h.container.querySelectorAll(".mg-root-card .mg-instructions-control")).toHaveLength(2);
    expect(h.container.querySelectorAll(".mg-instructions-control label")[0]!.textContent).toBe("AGENTS.md als Arbeitsanweisungen verwenden");
    expect(await h.candidate()).toEqual(h.baseline);
    expect(h.calls.some(call => call.method === "config.patch")).toBe(false);
  });
  it("leaves newly added roots unchecked without writing the optional property", async () => {
    const h = mount(); await h.settle();
    const path = h.container.querySelector<HTMLInputElement>(".mg-add-folder input")!; path.value = "/Synthetic/New"; path.dispatchEvent(new Event("input", { bubbles: true }));
    h.click("Ordner prüfen und hinzufügen"); await h.settle();
    expect(h.container.querySelectorAll<HTMLInputElement>(".mg-instructions-control input")[2]!.checked).toBe(false);
    expect((await h.candidate()).services.onedrive.allowed_roots[2]).not.toHaveProperty("agents_instructions");
  });
  it("roundtrips on/off without dirty state or unrelated grant changes and retains focus on the correct root", async () => {
    const h = mount(); await h.settle(); const confirm = vi.fn(); vi.stubGlobal("confirm", confirm);
    h.toggle(1); h.toggle(1); h.toggle(0); h.toggle(0);
    expect(document.activeElement).toBe(h.container.querySelector(".mg-instructions-control input"));
    expect(h.container.textContent).toContain("Keine ungespeicherten Änderungen");
    expect(await h.candidate()).toEqual(h.baseline);
    expect(confirm).not.toHaveBeenCalled();
  });
  it("restores each checkbox and info button for the collision labels foo and help-foo", async () => {
    const h = mount("de", [root("foo"), root("help-foo")]); await h.settle();
    for (const index of [0, 1]) {
      for (const selector of [".mg-instructions-control input", ".mg-instructions-info"]) {
        const control = h.container.querySelectorAll<HTMLElement>(selector)[index]!;
        control.focus(); h.notify(); await h.settle();
        const replacement = h.container.querySelectorAll<HTMLElement>(selector)[index]!;
        expect(replacement).not.toBe(control);
        expect(document.activeElement).toBe(replacement);
      }
    }
    h.toggle(1);
    expect(document.activeElement).toBe(h.container.querySelectorAll(".mg-instructions-control input")[1]);
  });
  it.each([0, 1])("saves root %i through the existing review flow with exact property and grant preservation", async index => {
    const h = mount(); await h.settle(); h.toggle(index);
    expect(h.calls.some(call => call.method === "config.patch")).toBe(false);
    const expected = structuredClone(h.baseline);
    if (index === 0) expected.services.onedrive.allowed_roots[0]!.agents_instructions = "trusted";
    else delete expected.services.onedrive.allowed_roots[1]!.agents_instructions;
    expect(await h.candidate()).toEqual(expected);
    const confirmation = h.container.querySelector<HTMLInputElement>(".mg-body input[type=checkbox]")!; confirmation.click();
    h.click("Speichern und anwenden"); await h.settle();
    const writes = h.calls.filter(call => call.method === "config.patch"); expect(writes).toHaveLength(1);
    const params = writes[0]!.params;
    expect(JSON.parse(params.raw as string).plugins.entries["microsoft-graph"].config.policy).toEqual({ services: { onedrive: { allowed_roots: expected.services.onedrive.allowed_roots } } });
    expect(params.replacePaths).toContain("plugins.entries.microsoft-graph.config.policy.services.onedrive.allowed_roots");
    expect(h.saved()).toEqual(expected);
    if (index === 1) expect(h.saved().services.onedrive.allowed_roots[1]).not.toHaveProperty("agents_instructions");
    expect(h.calls.some(call => call.method === "plugins.reload")).toBe(false);
  });
  it("associates each info button with unique native light-dismiss help and preserves the approved boundary wording", async () => {
    const h = mount(); await h.settle();
    const buttons = [...h.container.querySelectorAll(".mg-instructions-info")];
    const ids = buttons.map(button => button.getAttribute("aria-controls")); expect(new Set(ids).size).toBe(2);
    for (const button of buttons) {
      expect(button.getAttribute("aria-label")).toBe("Informationen zu AGENTS.md-Arbeitsanweisungen");
      expect(button.getAttribute("aria-expanded")).toBe("false");
      const panel = h.container.querySelector(`#${button.getAttribute("aria-controls")}`)!;
      expect(panel.getAttribute("popover")).toBe("auto");
      expect(panel.textContent).toContain("auch zukünftige Änderungen");
      expect(panel.textContent).toContain("garantiert keine vollständige Befolgung");
      expect(panel.textContent).toContain("bis zu einer Stunde");
    }
  });
  it("provides English fallback without altering the German label", async () => {
    const h = mount("en"); await h.settle();
    expect(h.container.querySelector(".mg-instructions-control label")!.textContent).toBe("Use AGENTS.md as working instructions");
    expect(h.container.querySelector(".mg-instructions-help")!.textContent).toContain("Acknowledgement does not guarantee full compliance.");
  });
});
