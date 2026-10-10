// @vitest-environment happy-dom
import { createHash } from "node:crypto";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import plugin from "./control-ui.js";
import { canonicalPolicy } from "./config-policy-identity.js";
const initial = { version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: {} }, mail: { agents: { main: { operations: ["read"] } } }, todo: { agents: {} } } };
const clone = <T>(v: T): T => structuredClone(v);
const hash = (v: unknown) => createHash("sha256").update(canonicalPolicy(v)).digest("hex");
function deferred<T>() { let resolve!: (v: T) => void; let reject!: (v: unknown) => void; const promise = new Promise<T>((r,j) => { resolve=r; reject=j; }); return { promise, resolve, reject }; }
type Mode = "active" | "pending" | "persist-reject" | "reject" | "hang" | "persist-hang" | "unavailable";
function harness(mode: Mode = "active") {
  let saved = clone(initial); let active = hash(saved); let revision = "r1";
  let page: { mount(container: HTMLElement, context: unknown): { dispose(): void } };
  let listener = () => {};
  const patch = deferred<unknown>(); const reload = deferred<unknown>();
  const calls: { method: string; params: Record<string, unknown> }[] = [];
  let getOverride: (() => Promise<unknown>) | undefined;
  const snapshot = () => ({ hash: revision, configRevisionHash: "unrelated-new", appliedConfigHash: "unrelated-old", config: { plugins: { entries: { "microsoft-graph": { enabled: true, config: { credentialVaultKey: "synthetic-secret-ref", policy: clone(saved) } } } } } });
  const host = { locale: "de", connection: { connected: true, canAdmin: true }, agents: { rows: [{ id: "main", name: "Main" }] },
    subscribe: (fn: () => void) => { listener = fn; return () => { listener = () => {}; }; },
    request: async (method: string, params: Record<string, unknown>) => {
      calls.push({ method, params });
      if (method === "config.get") return getOverride ? getOverride() : snapshot();
      if (method === "microsoft-graph.configuration.applicationStatus") return { activePolicyHash: active, activeVersion: "3.12.1" };
      if (method === "microsoft-graph.configuration.validate") return { valid: true, requiredScopes: ["Mail.Read"], policyHash: hash(params.policy) };
      if (method === "microsoft-graph.credentials.status") return { ok: true, value: { policyVersion: 2, credential: { result: "valid" } } };
      if (method === "microsoft-graph.updateStatus") return { updateAvailable: false };
      if (method === "config.patch") {
        const raw = JSON.parse(params.raw as string); const changes = raw.plugins.entries["microsoft-graph"].config.policy;
        if (mode !== "reject" && mode !== "hang" && mode !== "unavailable") { saved = { ...saved, rules: { ...saved.rules, ...changes.rules } }; revision = "r2"; }
        if (mode === "active") active = hash(saved);
        if (mode === "reject") throw { code: "INVALID_REQUEST" };
        if (mode === "persist-reject" || mode === "unavailable") throw { code: "UNAVAILABLE" };
        if (mode === "hang" || mode === "persist-hang") return patch.promise;
        return {};
      }
      if (method === "plugins.reload") return reload.promise;
      throw new Error(`Unexpected RPC ${method}`);
    }, ui: { registerPage: (value: typeof page) => { page=value; return () => {}; }, registerNavigation: () => () => {} } };
  const stop = plugin.activate(host as never); const container = document.createElement("div"); document.body.append(container);
  let mounted: { dispose(): void }; let controller: AbortController;
  const mount = () => { controller = new AbortController(); mounted = page.mount(container, { host, signal: controller.signal }); };
  mount();
  const settle = async () => { for (let i=0;i<100;i++) await Promise.resolve(); };
  const click = (text: string) => { const node = [...container.querySelectorAll("button")].find(b => b.textContent === text); expect(node, text + "\n" + container.textContent).toBeTruthy(); node!.click(); };
  const edit = async () => { click("3  Freigaben"); const input = container.querySelector<HTMLInputElement>(".mg-body input[type=checkbox]")!; input.checked=false; input.dispatchEvent(new Event("change", { bubbles: true })); click("4  Prüfen"); await settle(); const confirm = container.querySelector<HTMLInputElement>(".mg-body input[type=checkbox]")!; confirm.checked=true; confirm.dispatchEvent(new Event("change", { bubbles: true })); };
  return { host, calls, container, settle, click, edit, patch, reload, snapshot, setGet: (fn?: () => Promise<unknown>) => { getOverride=fn; }, notify: () => listener(), setActive: () => { active=hash(saved); }, changeSaved: () => { saved.services.mail.agents.main.operations.push("draft"); revision="external"; }, remount: () => { mounted.dispose(); controller.abort(); mount(); }, dispose: () => { mounted.dispose(); controller.abort(); if (typeof stop === "function") stop(); container.remove(); } };
}
const cleanup: (() => void)[] = [];
function view(mode?: Mode) { const h=harness(mode); cleanup.push(h.dispose); return h; }
beforeEach(() => { localStorage.clear(); vi.useFakeTimers(); });
afterEach(() => { for (const fn of cleanup.splice(0)) fn(); vi.useRealTimers(); });
describe("policy save and application lifecycle", () => {
  it("proves Graph activation despite unrelated global revision mismatch and preserves connection", async () => { const h=view(); await h.settle(); expect(h.container.textContent).toContain("Verbindung hergestellt"); expect(h.container.textContent).not.toContain("Aktivierung noch nicht bestätigt"); expect(h.calls.filter(c=>c.method==="microsoft-graph.updateStatus")).toHaveLength(1); });
  it("locks editing immediately and issues one write for duplicate clicks, then confirms saved and active", async () => { const h=view(); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); h.click("Speichert …"); expect([...h.container.querySelectorAll(".mg-steps button")].every(b=>(b as HTMLButtonElement).disabled)).toBe(true); await h.settle(); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.container.textContent).toContain("Gespeichert und aktiv"); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false); });
  it("recognizes persistence after a rejected response and applies without resaving", async () => { const h=view("persist-reject"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); expect(h.container.textContent).toContain("Gespeichert"); expect(h.container.textContent).toContain("Verbindung hergestellt"); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.calls.filter(c=>c.method==="plugins.reload")[0]?.params).toEqual({ plugins: [{ pluginId: "microsoft-graph" }], waitForDrain: true }); h.setActive(); h.reload.resolve({ ok:true }); await h.settle(); expect(h.container.textContent).toContain("Gespeichert und aktiv"); });
  it("retains drafts after known precommit rejection and permits a corrected retry", async () => { const h=view("reject"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); expect(h.container.textContent).toContain("Nicht gespeichert. Dein Entwurf ist noch da."); expect(h.container.textContent).toContain("Ungespeicherte Änderungen"); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false); });
  it("does not claim saving succeeded after UNAVAILABLE with unchanged persistence", async () => { const h=view("unavailable"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); expect(h.container.textContent).toContain("Speichern noch nicht bestätigt"); expect(h.container.textContent).not.toContain("Deine Regeln sind gespeichert"); expect(h.container.textContent).not.toContain("Gespeichert."); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false); });
  it("bounds unknown late outcomes and never repeats a write", async () => { const h=view("hang"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); await vi.advanceTimersByTimeAsync(70000); await h.settle(); expect(h.container.textContent).toContain("Speichern noch nicht bestätigt"); const reads=h.calls.filter(c=>c.method==="config.get").length; await vi.advanceTimersByTimeAsync(120000); expect(h.calls.filter(c=>c.method==="config.get")).toHaveLength(reads); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false); });
  it("recognizes persisted timeout and remounts without losing the target or writing twice", async () => { const h=view("persist-hang"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); await vi.advanceTimersByTimeAsync(15000); await h.settle(); expect(h.container.textContent).toContain("Gespeichert"); h.remount(); await h.settle(); expect(h.container.textContent).toContain("Verbindung hergestellt"); h.setActive(); h.notify(); await h.settle(); expect(h.container.textContent).toContain("Gespeichert und aktiv"); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); });
  it("reconciles disconnected reloads and safely retries the exact saved target without resaving", async () => { const h=view("pending"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); h.host.connection.connected=false; h.notify(); h.reload.reject(new Error("disconnected")); await h.settle(); expect(h.container.textContent).toContain("Verbindung unterbrochen"); h.host.connection.connected=true; h.notify(); await h.settle(); await vi.advanceTimersByTimeAsync(70000); await h.settle(); expect(h.container.textContent).toContain("Aktivierung noch nicht bestätigt"); h.click("Gespeicherte Regeln anwenden"); await h.settle(); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(2); });
  it("offers safe overview navigation while pending and refuses to apply a changed target", async () => { const h=view("pending"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); h.reload.reject(new Error("drain timeout")); await h.settle(); await vi.advanceTimersByTimeAsync(70000); await h.settle(); h.click("Übersicht anzeigen"); expect(h.container.querySelector(".mg-steps [aria-current=step]")?.textContent).toBe("1  OneDrive"); expect([...h.container.querySelectorAll(".mg-steps button")].every(b=>(b as HTMLButtonElement).disabled)).toBe(true); h.changeSaved(); h.click("Gespeicherte Regeln anwenden"); await h.settle(); expect(h.container.textContent).toContain("anderswo geändert"); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(1); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); });
  it("does not dispatch reload after a delayed preflight when the page is disposed and administration revoked", async () => { const h=view("pending"); await h.settle(); await h.edit(); const preflight=deferred<unknown>(); h.setGet(()=>h.calls.filter(c=>c.method==="config.get").length>=4 ? preflight.promise : Promise.resolve(h.snapshot())); h.click("Speichern und anwenden"); await h.settle(); expect(h.calls.filter(c=>c.method==="config.get")).toHaveLength(4); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0); const fresh=h.snapshot(); h.dispose(); h.host.connection.canAdmin=false; preflight.resolve(fresh); await h.settle(); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); });
  it("discards a stale status response that began before saving", async () => { const h=view(); await h.settle(); await h.edit(); const old=h.snapshot(); const pending=deferred<unknown>(); h.setGet(()=>pending.promise); h.notify(); await h.settle(); h.click("Speichern und anwenden"); h.setGet(); pending.resolve(old); await h.settle(); await vi.advanceTimersByTimeAsync(5000); await h.settle(); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.container.textContent).toContain("Gespeichert und aktiv"); expect(h.container.textContent).not.toContain("anderswo geändert"); });
});
