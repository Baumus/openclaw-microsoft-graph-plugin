// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from "vitest";
import plugin from "./control-ui.js";
import { requiredOneDriveTools } from "./control-ui-tool-exposure.js";
import { authorizeRoot, validatePolicy } from "./policy.js";

type Root = { label: string; path: string; drive_id: string; item_id: string; include_descendants: true; agents_instructions?: "trusted"; permissions: { read: boolean; write: boolean; delete: boolean }; agents: Record<string, { permissions: { read: boolean; write?: boolean } }> };
const root = (label: string, trusted = false): Root => ({ label: label.toLowerCase(), path: `/Synthetic/${label}`, drive_id: "synthetic-drive", item_id: label, include_descendants: true, ...(trusted ? { agents_instructions: "trusted" as const } : {}), permissions: { read: true, write: true, delete: false }, agents: { reader: { permissions: { read: true } }, editor: { permissions: { read: true, write: true } } } });
const initial = () => ({ version: 2, rules: { default: "deny", warningApprovalsByService: { mail: false } }, services: { onedrive: { allowed_roots: [root("Documents"), root("Projects", true)] }, calendar: { agents: { reader: { operations: ["read"], resources: ["me", "synthetic-calendar"] } } }, mail: { agents: { editor: { operations: ["read", "draft"] } } }, contacts: { agents: {} }, todo: { agents: {} } } });
type Policy = ReturnType<typeof initial>;
const cleanup: Array<() => void> = [];
afterEach(() => { for (const stop of cleanup.splice(0)) stop(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });
function mount(locale = "de", roots?: Root[]) {
  let saved = initial();
  if (roots) saved.services.onedrive.allowed_roots = roots;
  const baseline = structuredClone(saved);
  const selected = requiredOneDriveTools(saved.services.onedrive.allowed_roots);
  const agentEntries = Object.fromEntries(Object.entries(selected).map(([id, tools]) => [id, { tools: { alsoAllow: tools } }]));
  let page!: { mount(container: HTMLElement, context: unknown): { dispose(): void } };
  let listener = () => {};
  const calls: Array<{ method: string; params: Record<string, unknown> }> = [];
  const host = { locale, connection: { connected: true, canAdmin: true }, agents: { rows: [{ id: "reader", name: "Reader" }, { id: "editor", name: "Editor" }] },
    subscribe: (fn: () => void) => { listener = fn; return () => {}; },
    request: async (method: string, params: Record<string, unknown>) => {
      calls.push({ method, params: structuredClone(params) });
      if (method === "config.get") return { hash: "synthetic-revision", configRevisionHash: "same", appliedConfigHash: "same", config: { agents: { entries: agentEntries }, plugins: { entries: { "microsoft-graph": { enabled: true, config: { credentialVaultKey: "synthetic-ref", policy: structuredClone(saved) } } } } } };
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
  const candidate = async () => { click(locale === "en" ? "Review" : "Prüfen"); await settle(); return (calls.filter(call => call.method === "microsoft-graph.configuration.validate").at(-1)!.params as { policy: Policy }).policy; };
  return { container, calls, baseline, settle, click, toggle, candidate, notify: () => listener(), saved: () => structuredClone(saved) };
}

describe("OneDrive per-agent access removal", () => {
  const menu = (h: ReturnType<typeof mount>, index = 0) => h.container.querySelectorAll<HTMLButtonElement>(".mg-agent-menu-trigger")[index]!;
  const remove = (h: ReturnType<typeof mount>, index = 0) => h.container.querySelectorAll<HTMLButtonElement>(".mg-agent-menu [role=menuitem]")[index]!;
  it("removes only the selected explicit assignment and preserves root caps, other roots/services and grants", async () => {
    const h = mount(); await h.settle();
    const confirm = vi.fn(() => true); vi.stubGlobal("confirm", confirm);
    remove(h).click();
    expect(confirm).toHaveBeenCalledWith(expect.stringContaining("Reader (reader) auf /Synthetic/Documents"));
    expect(confirm).toHaveBeenCalledWith(expect.stringContaining("Andere Agenten bleiben unverändert"));
    expect(confirm).toHaveBeenCalledWith(expect.stringContaining("übergeordnete Ordner bleiben wirksam"));
    const expected = structuredClone(h.baseline); delete expected.services.onedrive.allowed_roots[0]!.agents.reader;
    expect(await h.candidate()).toEqual(expected);
    expect(h.calls.some(call => call.method === "config.patch")).toBe(false);
    expect(authorizeRoot(validatePolicy(expected), "editor", "documents", "write")).toBeTruthy();
    expect(() => authorizeRoot(validatePolicy(expected), "reader", "documents", "read")).toThrow("access_denied");
  });
  it("preserves restrictive root caps and effective parent grants through review and save", async () => {
    const parent = root("parent"); parent.path = "/Synthetic";
    const child = root("child", true); child.permissions.write = false;
    const h = mount("de", [child, parent]); await h.settle(); vi.stubGlobal("confirm", vi.fn(() => true));
    remove(h).click();
    const expected = structuredClone(h.baseline); delete expected.services.onedrive.allowed_roots[0]!.agents.reader;
    const candidate = await h.candidate(); expect(candidate).toEqual(expected);
    const checked = validatePolicy(candidate);
    expect(() => authorizeRoot(checked, "editor", "child", "write")).toThrow("access_denied");
    expect(authorizeRoot(checked, "reader", "parent", "read").path).toBe("/Synthetic");
    h.container.querySelector<HTMLInputElement>(".mg-body input[type=checkbox]")!.click();
    h.click("Speichern und anwenden"); await h.settle();
    expect(h.saved()).toEqual(expected);
    expect(h.calls.filter(call => call.method === "config.patch")).toHaveLength(1);
    expect(h.calls.some(call => call.method === "plugins.reload")).toBe(false);
  });
  it("cancel leaves the whole draft unchanged and restores the trigger focus", async () => {
    const h = mount(); await h.settle(); vi.stubGlobal("confirm", vi.fn(() => false));
    remove(h).click(); expect(document.activeElement).toBe(menu(h));
    expect(await h.candidate()).toEqual(h.baseline);
  });
  it("last assignment removal retains an empty deny-by-default root and the add-agent workflow", async () => {
    const only = root("Only", true); delete only.agents.editor;
    const h = mount("de", [only]); await h.settle(); vi.stubGlobal("confirm", vi.fn(() => true));
    remove(h).click();
    expect(h.container.textContent).toContain("Noch kein Agent berechtigt.");
    expect(document.activeElement).toBe(h.container.querySelector(".mg-add-agent"));
    h.click("+ Agent hinzufügen"); expect(h.container.querySelector(".mg-access-editor")).toBeTruthy();
    h.click("Abbrechen");
    const expected = structuredClone(h.baseline); expected.services.onedrive.allowed_roots[0]!.agents = {};
    const candidate = await h.candidate(); expect(candidate).toEqual(expected);
    expect(() => authorizeRoot(validatePolicy(candidate), "reader", "only", "read")).toThrow("access_denied");
    expect(() => authorizeRoot(validatePolicy(candidate), "editor", "only", "write")).toThrow("access_denied");
  });
  it("exposes bounded native light-dismiss menus with keyboard dismissal and no policy mutation", async () => {
    const h = mount(); await h.settle();
    const trigger = menu(h), action = remove(h), panel = action.parentElement!;
    expect(trigger.textContent).toBe("⋯");
    expect(trigger.getAttribute("aria-haspopup")).toBe("menu");
    expect(trigger.getAttribute("aria-controls")).toBe(panel.id);
    expect(panel.getAttribute("popover")).toBe("auto");
    expect(action.getAttribute("role")).toBe("menuitem");
    expect(new Set([...h.container.querySelectorAll(".mg-agent-menu")].map(node => node.id)).size).toBe(4);
    const show = vi.fn(), hide = vi.fn(); Object.assign(panel, { showPopover: show, hidePopover: hide });
    trigger.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    expect(show).toHaveBeenCalledOnce(); expect(document.activeElement).toBe(action);
    action.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(hide).toHaveBeenCalledOnce(); expect(document.activeElement).toBe(trigger);
    action.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true }));
    expect(hide).toHaveBeenCalledTimes(2);
    expect(await h.candidate()).toEqual(h.baseline);
  });
  it("keeps whole-folder removal distinct and localized, with inherited-access wording in English", async () => {
    const h = mount("en"); await h.settle();
    expect(remove(h).textContent).toBe("Remove access…");
    expect(h.container.textContent).toContain("Remove folder sharing");
    const confirm = vi.fn(() => false); vi.stubGlobal("confirm", confirm); remove(h).click();
    expect(confirm).toHaveBeenCalledWith(expect.stringContaining("Access through parent folders remains effective"));
    expect(await h.candidate()).toEqual(h.baseline);
  });
});

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
