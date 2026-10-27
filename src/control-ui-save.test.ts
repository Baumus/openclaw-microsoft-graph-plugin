// @vitest-environment happy-dom
import { createHash } from "node:crypto";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import plugin from "./control-ui.js";
import { planChangedToolExposure, planPolicyToolReconciliation, planToolExposure } from "./control-ui-tool-exposure.js";
import { canonicalPolicy } from "./config-policy-identity.js";
const initial = { version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: {} }, mail: { agents: { main: { operations: ["read"] } } }, todo: { agents: {} } } };
const clone = <T>(v: T): T => structuredClone(v);
const hash = (v: unknown) => createHash("sha256").update(canonicalPolicy(v)).digest("hex");
function deferred<T>() { let resolve!: (v: T) => void; let reject!: (v: unknown) => void; const promise = new Promise<T>((r,j) => { resolve=r; reject=j; }); return { promise, resolve, reject }; }
type Mode = "active" | "pending" | "persist-reject" | "reject" | "hang" | "persist-hang" | "unavailable" | "include-reject" | "policy-reject" | "tools-reject";
function harness(mode: Mode = "active", pluginsInclude = false, oneDriveGrant = false, healing = false) {
  let saved = clone(initial); let active = hash(saved); let revision = "r1"; let globalApplied = false;
  if (oneDriveGrant) saved.services.onedrive.allowed_roots = [{ label: "synthetic", path: "/Synthetic", drive_id: "synthetic-drive", item_id: "synthetic-item", include_descendants: true, permissions: { read: true, write: false, delete: false }, agents: { studio: { permissions: { read: true, write: false, delete: false } } } }] as never;
  active = hash(saved);
  const agentEntries: Record<string, { tools: { alsoAllow: string[] } }> = { main: { tools: { alsoAllow: ["outlook_mail_read", "microsoft_graph_capabilities"] } }, studio: { tools: { alsoAllow: [] } } };
  let page: { mount(container: HTMLElement, context: unknown): { dispose(): void } };
  let listener = () => {};
  const patch = deferred<unknown>(); const reload = deferred<unknown>();
  const calls: { method: string; params: Record<string, unknown> }[] = [];
  let intent: any = null; let loseClaimReply = false; let backendUnavailable = false;
  let getOverride: (() => Promise<unknown>) | undefined;
  const snapshot = () => ({ hash: revision, configRevisionHash: "unrelated-new", appliedConfigHash: globalApplied ? "unrelated-new" : "unrelated-old", parsed: pluginsInclude ? { plugins: { $include: "plugins.json" } } : undefined, config: { agents: { entries: clone(agentEntries) }, plugins: { entries: { "microsoft-graph": { enabled: true, config: { credentialVaultKey: "synthetic-secret-ref", policy: clone(saved) } } } } } });
  let host = { locale: "de", connection: { connected: true, canAdmin: true }, agents: { rows: [{ id: "main", name: "Main" }, { id: "studio", name: "Studio" }] },
    subscribe: (fn: () => void) => { listener = fn; return () => { listener = () => {}; }; },
    request: async (method: string, params: Record<string, unknown>) => {
      calls.push({ method, params });
      if (backendUnavailable && method.startsWith("microsoft-graph.configuration.")) throw Error("UNAVAILABLE");
      if (method === "microsoft-graph.configuration.activation") {
        if (params.action === "reconcile") {
          if (intent && intent.state !== "complete") throw Error("activation_pending");
          if (params.baseHash !== revision || active !== hash(saved)) throw Error("conflict");
          const plan = planPolicyToolReconciliation(saved, agentEntries);
          intent = { id: "synthetic-repair", targetHash: hash(saved), state: "issued", expected: plan.expected };
          if (loseClaimReply) throw Error("lost_repair_claim");
          return { record: clone(intent), ...plan, baseHash: revision };
        }
        if (params.action === "prepare") {
          const plan = planChangedToolExposure(params.baseline as any, params.target as any, agentEntries);
          const required = Object.fromEntries(Object.entries(plan.expected).map(([id, e]) => [id, e.value.filter(n => !agentEntries[id]!.tools.alsoAllow.includes(n))]));
          intent = { id: "synthetic-intent", targetHash: hash(params.target), required, state: "prepared" }; return clone(intent);
        }
        if (params.action === "status") return intent && intent.state !== "complete" && intent.targetHash === hash(saved) ? clone(intent) : null;
        if (params.action === "claim") {
          if (!intent || intent.state !== "prepared" || intent.id !== params.id || active !== intent.targetHash) throw Error("intent_conflict");
          const plan = planToolExposure(intent.required, agentEntries); intent.state="issued"; intent.expected=plan.expected;
          if (loseClaimReply) throw Error("lost_claim_reply");
          return { record: clone(intent), ...plan, baseHash: revision };
        }
        if (params.action === "cancel") { intent=null; return { cancelled:true }; }
        if (params.action === "complete") { intent.state="complete"; return { complete:true }; }
      }
      if (method === "config.get") return getOverride ? getOverride() : snapshot();
      if (method === "microsoft-graph.configuration.applicationStatus") return { activePolicyHash: active, activeVersion: "3.12.1", ...(healing ? { toolReconciliation: "control-ui-cas-v1" } : {}) };
      if (method === "microsoft-graph.configuration.validate") return { valid: true, requiredScopes: (params.policy as any)?.services?.contacts?.agents?.main ? ["Mail.Read", "Contacts.ReadWrite"] : ["Mail.Read"], policyHash: hash(params.policy) };
      if (method === "microsoft-graph.credentials.status") return { ok: true, value: { policyVersion: 2, credential: { result: "valid" } } };
      if (method === "microsoft-graph.configuration.calendars") return { defaultCalendarId: "me", calendars: [{ id: "me", name: "Primary" }] };
      if (method === "microsoft-graph.updateStatus") return { updateAvailable: false };
      if (method === "config.patch") {
        const raw = JSON.parse(params.raw as string); const changes = raw.plugins?.entries["microsoft-graph"].config.policy;
        if (mode === "tools-reject" && raw.agents && !changes) throw { code: "INVALID_REQUEST" };
        if (mode === "policy-reject" && changes) throw { code: "INVALID_REQUEST" };
        if (mode === "include-reject" || pluginsInclude && raw.agents && raw.plugins) throw { code: "UNAVAILABLE", message: "Config write would flatten $include-owned config: code=CONFIG_INCLUDE_OWNERSHIP" };
        if (mode !== "reject" && mode !== "hang" && mode !== "unavailable") { const merge = (before: any, patch: any): any => { if (!patch || typeof patch !== "object" || Array.isArray(patch)) return clone(patch); const out = { ...before }; for (const [key, value] of Object.entries(patch)) { if (value === null) delete out[key]; else out[key] = merge(before?.[key], value); } return out; };
          if (changes) saved = merge(saved, changes); for (const [id, entry] of Object.entries(raw.agents?.entries ?? {}) as [string, { tools: { alsoAllow: string[] } }][]) agentEntries[id]!.tools.alsoAllow = entry.tools.alsoAllow; revision = revision + "-next"; globalApplied=false; }
        if (mode === "active" || mode === "tools-reject") active = hash(saved);
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
  return { agentEntries, enableHealing: () => { healing=true; }, setBackendUnavailable: (v: boolean) => { backendUnavailable=v; }, loseClaim: () => { loseClaimReply=true; }, clearIntent: () => { intent=null; }, host, calls, container, settle, click, edit, patch, reload, snapshot, setInactive: () => { active=hash(initial); }, setMode: (next: Mode) => { mode=next; }, hardReload: () => { mounted.dispose(); controller.abort(); host={...host}; mount(); }, setGet: (fn?: () => Promise<unknown>) => { getOverride=fn; }, notify: () => listener(), setActive: () => { active=hash(saved); globalApplied=true; }, changeRevision: () => { revision="unrelated"; }, persistTarget: () => { saved.rules = { ...saved.rules, warningApprovalsByService: { onedrive: false } } as never; revision="late"; }, restoreBaseline: () => { saved = clone(initial); revision="rollback"; }, changeSaved: () => { saved.services.mail.agents.main.operations.push("draft"); revision="external"; }, remount: () => { mounted.dispose(); controller.abort(); mount(); }, dispose: () => { mounted.dispose(); controller.abort(); if (typeof stop === "function") stop(); container.remove(); } };
}
const cleanup: (() => void)[] = [];
function view(mode?: Mode, pluginsInclude = false, oneDriveGrant = false, healing = false) { const h=harness(mode, pluginsInclude, oneDriveGrant, healing); cleanup.push(h.dispose); return h; }
beforeEach(() => { localStorage.clear(); vi.useFakeTimers(); });
afterEach(() => { for (const fn of cleanup.splice(0)) fn(); vi.useRealTimers(); });
describe("policy save and application lifecycle", () => {
  it("does not repair unrelated existing OneDrive opt-ins during a rules-only save", async () => {
    const h=view("active", true, true); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle();
    const writes = h.calls.filter(c=>c.method==="config.patch"); expect(writes).toHaveLength(1);
    expect(JSON.parse(writes[0]!.params.raw as string).agents).toBeUndefined();
    expect(h.container.textContent).toContain("Microsoft-Regeln aktiv. Agent-Tools weichen ab.");
  });
  async function calendarDraft(h: ReturnType<typeof view>) {
    h.click("2  Dienste"); await h.settle();
    const calendar = [...h.container.querySelectorAll("fieldset")].find(x => x.querySelector("legend")?.textContent === "Kalender")!;
    const input = calendar.querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    input.checked=true; input.dispatchEvent(new Event("change", { bubbles: true }));
    h.click("4  Prüfen"); await h.settle();
    const confirm=h.container.querySelector<HTMLInputElement>(".mg-body input[type=checkbox]")!; confirm.checked=true; confirm.dispatchEvent(new Event("change", {bubbles:true}));
  }
  it("saves a real calendar grant across include ownership without touching unrelated OneDrive agents", async () => {
    const h=view("active", true, true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    await vi.advanceTimersByTimeAsync(1000); await h.settle();
    const writes=h.calls.filter(c=>c.method==="config.patch"); expect(writes).toHaveLength(2);
    expect(JSON.parse(writes[1]!.params.raw as string)).toEqual({ agents: { entries: { main: { tools: { alsoAllow: ["outlook_mail_read", "microsoft_graph_capabilities", "outlook_calendar_read", "outlook_calendar_write"] } } } } });
    const policy=JSON.parse(writes[0]!.params.raw as string); expect(policy.agents).toBeUndefined();
    expect(policy.plugins.entries["microsoft-graph"].config.policy.services.calendar.agents.main.operations).toEqual(["read", "create", "update", "respond", "attach", "delete"]);
    expect(writes[1]!.params.baseHash).toBe("r1-next");
    expect(h.container.textContent).toContain("Microsoft-Regeln aktiv. Agent-Tools weichen ab.");
  });
  it("reports wrapped include ownership refusal as failed, not saved or indefinitely unknown", async () => {
    const h=view("include-reject", true); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle();
    expect(h.container.textContent).not.toContain("Gespeichert und aktiv");
    expect(h.container.textContent).not.toContain("Speichern noch nicht bestätigt");
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect([...h.container.querySelectorAll(".mg-steps button")].some(b=>!(b as HTMLButtonElement).disabled)).toBe(true);
  });
  it("reconciles retained policy activation through remount before exposing tools once", async () => {
    const h=view("persist-hang", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    h.remount(); await h.settle(); await vi.advanceTimersByTimeAsync(3000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.container.textContent).not.toContain("Gespeichert und aktiv");
    h.setActive(); h.notify(); await vi.advanceTimersByTimeAsync(3000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(2);
    expect(h.container.textContent).not.toContain("Gespeichert und aktiv");
    h.patch.resolve({}); await h.settle();
    expect(h.container.textContent).toContain("Gespeichert und aktiv");
    expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0);
  });
  it("never sends the second stage if the policy changes during ambiguous tool preparation", async () => {
    const h=view("persist-hang", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    h.changeSaved(); await vi.advanceTimersByTimeAsync(3000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.container.textContent).toContain("anderswo geändert");
  });
  it("does not send the policy when administration is revoked during tool preparation", async () => {
    const h=view("persist-hang", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    h.host.connection.canAdmin=false; h.notify(); await vi.advanceTimersByTimeAsync(3000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
  });
  it("keeps an unpersisted tool write ambiguous without sending a policy or duplicate tool write", async () => {
    const h=view("hang", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    await vi.advanceTimersByTimeAsync(70000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.container.textContent).toContain("Speichern noch nicht bestätigt");
    expect(h.container.textContent).toContain("Speichern noch nicht bestätigt");
  });
  it("discloses persisted policy when the tool stage is rejected without claiming active access", async () => {
    const h=view("tools-reject", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle(); await vi.advanceTimersByTimeAsync(1000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(2);
    expect(h.container.textContent).toContain("Microsoft-Regeln sind gespeichert, aber die Agent-Tools konnten nicht freigeschaltet werden");
    expect(h.container.textContent).not.toContain("Gespeichert und aktiv");
    expect(h.snapshot().config.plugins.entries["microsoft-graph"].config.policy.services.calendar.agents).toHaveProperty("main");
  });
  it("automatically resumes only the original selected tools after a fresh host beyond the drain budget", async () => {
    const h=view("persist-hang", true, true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    await vi.advanceTimersByTimeAsync(75000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    h.hardReload(); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    h.setMode("active"); h.setActive();
    await vi.advanceTimersByTimeAsync(90000); await h.settle();
    const writes=h.calls.filter(c=>c.method==="config.patch"); expect(writes).toHaveLength(2);
    expect(Object.keys(JSON.parse(writes[1]!.params.raw as string).agents.entries)).toEqual(["main"]);
    expect(h.container.textContent).toContain("Microsoft-Regeln aktiv. Agent-Tools weichen ab.");
    expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0);
    h.hardReload(); await h.settle(); await vi.advanceTimersByTimeAsync(20000);
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(2);
  });
  it("rediscovers the durable intent when a fresh page opens while the backend is unavailable", async () => {
    const h=view("persist-hang", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    h.setBackendUnavailable(true); h.hardReload(); await h.settle();
    await vi.advanceTimersByTimeAsync(90000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    h.setBackendUnavailable(false); h.setActive(); h.setMode("active");
    await vi.advanceTimersByTimeAsync(30000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(2);
    expect(h.container.textContent).toContain("Gespeichert und aktiv");
    expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0);
  });
  it("keeps observing late activation without host notifications after the old polling budget", async () => {
    const h=view("persist-hang", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    await vi.advanceTimersByTimeAsync(90000); h.setActive(); h.setMode("active");
    await vi.advanceTimersByTimeAsync(16000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(2);
    expect(h.container.textContent).toContain("Gespeichert und aktiv");
  });
  it("does not replay a claimed dispatch after a lost reply and fresh host", async () => {
    const h=view("persist-hang", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    h.loseClaim(); h.setActive(); await vi.advanceTimersByTimeAsync(2000); await h.settle();
    h.hardReload(); await h.settle(); await vi.advanceTimersByTimeAsync(90000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.container.textContent).toContain("nicht blind wiederholt");
    expect(h.container.textContent).toContain("Aktuellen Stand laden");
    expect(h.container.textContent).not.toContain("Gespeichert und aktiv");
  });
  it("recovers missing selected calendar tools after hard refresh without rewriting policy or backfilling peers", async () => {
    const h=view("persist-hang", true, true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    h.clearIntent(); h.setActive(); h.hardReload(); await h.settle(); expect(h.container.textContent).not.toContain("Gespeichert und aktiv"); expect(h.container.textContent).toContain("Microsoft-Regeln aktiv. Agent-Tools weichen ab."); h.click("2  Dienste"); await h.settle();
    expect(h.container.textContent).toContain("Benötigte Agent-Tools fehlen noch");
    h.setMode("active"); h.click("Agent-Tools aktivieren"); await h.settle();
    const writes=h.calls.filter(c=>c.method==="config.patch"); expect(writes).toHaveLength(2);
    const repair=JSON.parse(writes[1]!.params.raw as string);
    expect(repair.plugins).toBeUndefined(); expect(Object.keys(repair.agents.entries)).toEqual(["main"]);
    expect(repair.agents.entries.main.tools.alsoAllow).toEqual(["outlook_mail_read", "microsoft_graph_capabilities", "outlook_calendar_read", "outlook_calendar_write"]);
    expect(h.container.textContent).toContain("Microsoft-Regeln aktiv. Agent-Tools weichen ab.");
  });
  it("recovers a known rejected tool stage through reload and explicit selected-service repair", async () => {
    const h=view("tools-reject", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle(); await vi.advanceTimersByTimeAsync(1000); await h.settle();
    expect(h.container.textContent).toContain("Zugang noch nicht vollständig aktiv");
    const previousConfirm=window.confirm; window.confirm=()=>true;
    h.click("Aktuellen Stand laden"); await h.settle(); window.confirm=previousConfirm; h.click("2  Dienste"); await h.settle();
    h.setMode("active"); h.click("Agent-Tools aktivieren"); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(3);
    expect(h.container.textContent).toContain("Gespeichert und aktiv");
  });
  it("does not repair selected tools until the exact saved policy is active", async () => {
    const h=view("persist-hang", true); await h.settle(); await calendarDraft(h); h.click("Speichern und anwenden"); await h.settle();
    h.clearIntent(); h.setActive(); h.hardReload(); await h.settle(); h.click("2  Dienste"); await h.settle();
    h.setInactive(); h.click("Agent-Tools aktivieren"); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.container.textContent).toContain("Agent-Tools konnten nicht aktiviert werden");
    expect(h.container.textContent).not.toContain("Gespeichert und aktiv");
  });
  it("permits a new scope grant to save before reauthorization while the old credential remains", async () => {
    const h=view(); await h.settle(); h.click("2  Dienste");
    const contacts = [...h.container.querySelectorAll("fieldset")].find(fieldset => fieldset.querySelector("legend")?.textContent === "Kontakte")!;
    const toggle = contacts.querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    toggle.checked = true; toggle.dispatchEvent(new Event("change", { bubbles: true }));
    h.click("4  Prüfen"); await h.settle();
    expect(h.container.textContent).toContain("benötigen nach dem Speichern eine Reautorisierung");
    const confirm = h.container.querySelector<HTMLInputElement>(".mg-body input[type=checkbox]")!;
    confirm.checked=true; confirm.dispatchEvent(new Event("change", { bubbles: true }));
    h.click("Speichern und anwenden"); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.calls.some(c=>c.method==="microsoft-graph.credentials.reauthorize-start")).toBe(false);
  });
  it("edits a policy under one top-level plugins include without changing the patch path", async () => { const h=view("active", true); await h.settle(); await h.edit(); expect(h.container.textContent).toContain("Konfigurationsquelle: plugins.json"); h.click("Speichern und anwenden"); await h.settle(); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(JSON.parse(h.calls.find(c=>c.method==="config.patch")!.params.raw as string).plugins.entries["microsoft-graph"].config.policy).toBeTruthy(); expect(h.container.textContent).toContain("Gespeichert und aktiv"); });
  it("proves Graph activation despite unrelated global revision mismatch and preserves connection", async () => { const h=view(); await h.settle(); expect(h.container.textContent).toContain("Verbindung hergestellt"); expect(h.container.textContent).not.toContain("Aktivierung noch nicht bestätigt"); expect(h.calls.filter(c=>c.method==="microsoft-graph.updateStatus")).toHaveLength(1); });
  it("locks editing immediately and issues one write for duplicate clicks, then confirms saved and active", async () => { const h=view(); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); h.click("Speichert …"); expect([...h.container.querySelectorAll(".mg-steps button")].every(b=>(b as HTMLButtonElement).disabled)).toBe(true); await h.settle(); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.container.textContent).toContain("Gespeichert und aktiv"); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false); });
  it("recognizes persistence after a rejected response without a duplicate automatic reload", async () => { const h=view("persist-reject"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); expect(h.container.textContent).toContain("Gespeichert"); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0); h.setActive(); h.notify(); await vi.advanceTimersByTimeAsync(250); await h.settle(); expect(h.container.textContent).toContain("Gespeichert und aktiv"); });
  it("confirms a persisted write before its delayed config.patch response", async () => { const h=view("persist-hang"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); expect(h.container.textContent).toContain("werden gespeichert"); await vi.advanceTimersByTimeAsync(1000); await h.settle(); expect(h.container.textContent).toContain("Gespeichert. Änderungen werden jetzt angewendet"); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0); });
  it("retains drafts after known precommit rejection and permits a corrected retry", async () => { const h=view("reject"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); expect(h.container.textContent).toContain("Nicht gespeichert. Dein Entwurf ist noch da."); expect(h.container.textContent).toContain("Ungespeicherte Änderungen"); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false); });
  it("does not claim saving succeeded after UNAVAILABLE with unchanged persistence", async () => { const h=view("unavailable"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); expect(h.container.textContent).toContain("Speichern noch nicht bestätigt"); expect(h.container.textContent).not.toContain("Deine Regeln sind gespeichert"); expect(h.container.textContent).not.toContain("Gespeichert."); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false); });
  it("bounds unknown late outcomes and never repeats a write", async () => { const h=view("hang"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); await vi.advanceTimersByTimeAsync(70000); await h.settle(); expect(h.container.textContent).toContain("Speichern noch nicht bestätigt"); const reads=h.calls.filter(c=>c.method==="config.get").length; await vi.advanceTimersByTimeAsync(120000); expect(h.calls.filter(c=>c.method==="config.get")).toHaveLength(reads); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false); });
  it("recognizes persisted timeout and remounts without losing the target or writing twice", async () => { const h=view("persist-hang"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); await vi.advanceTimersByTimeAsync(15000); await h.settle(); expect(h.container.textContent).toContain("Gespeichert"); h.remount(); await h.settle(); expect(h.container.textContent).toContain("Verbindung hergestellt"); h.setActive(); h.notify(); await vi.advanceTimersByTimeAsync(250); await h.settle(); expect(h.container.textContent).toContain("Gespeichert und aktiv"); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); });
  it("reconciles disconnection and manually retries the exact saved target without resaving", async () => { const h=view("pending"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); h.host.connection.connected=false; h.notify(); await h.settle(); expect(h.container.textContent).toContain("Verbindung unterbrochen"); h.host.connection.connected=true; h.notify(); await h.settle(); await vi.advanceTimersByTimeAsync(70000); await h.settle(); expect(h.container.textContent).toContain("Aktivierung noch nicht bestätigt"); h.click("Gespeicherte Regeln anwenden"); await h.settle(); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(1); expect(h.calls.filter(c=>c.method==="plugins.reload")[0]?.params).toEqual({ plugins: [{ pluginId: "microsoft-graph" }], waitForDrain: true }); });
  it("offers safe overview navigation while pending and refuses to apply a changed target", async () => { const h=view("pending"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); await vi.advanceTimersByTimeAsync(70000); await h.settle(); h.click("Übersicht anzeigen"); expect(h.container.querySelector(".mg-steps [aria-current=step]")?.textContent).toBe("1  OneDrive"); expect([...h.container.querySelectorAll(".mg-steps button")].every(b=>(b as HTMLButtonElement).disabled)).toBe(true); h.changeSaved(); h.click("Gespeicherte Regeln anwenden"); await h.settle(); expect(h.container.textContent).toContain("anderswo geändert"); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); });
  it("does not dispatch manual reload after a delayed preflight when the page is disposed and administration revoked", async () => { const h=view("pending"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); const preflight=deferred<unknown>(); const before=h.calls.filter(c=>c.method==="config.get").length; h.setGet(()=>preflight.promise); h.click("Gespeicherte Regeln anwenden"); await h.settle(); expect(h.calls.filter(c=>c.method==="config.get")).toHaveLength(before+1); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0); const fresh=h.snapshot(); h.dispose(); h.host.connection.canAdmin=false; preflight.resolve(fresh); await h.settle(); expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); });
  it("discards a stale status response that began before saving", async () => { const h=view(); await h.settle(); await h.edit(); const old=h.snapshot(); const pending=deferred<unknown>(); h.setGet(()=>pending.promise); h.notify(); await h.settle(); h.click("Speichern und anwenden"); h.setGet(); pending.resolve(old); await h.settle(); await vi.advanceTimersByTimeAsync(5000); await h.settle(); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.container.textContent).toContain("Gespeichert und aktiv"); expect(h.container.textContent).not.toContain("anderswo geändert"); });
  it("rebases an unchanged Graph draft onto the fresh revision without overwriting unrelated config", async () => {
    const h=view(); await h.settle(); await h.edit(); h.changeRevision(); h.click("Speichern und anwenden"); await h.settle();
    const writes=h.calls.filter(c=>c.method==="config.patch"); expect(writes).toHaveLength(1); expect(writes[0]!.params.baseHash).toBe("unrelated");
    expect(h.container.textContent).toContain("Gespeichert und aktiv");
  });
  it("continues reconciling after an ambiguous failure and accepts a late persisted target", async () => {
    const h=view("unavailable"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle();
    expect(h.container.textContent).toContain("Speichern noch nicht bestätigt"); expect(h.container.textContent).not.toContain("außerhalb dieser Seite geändert");
    h.persistTarget(); h.setActive(); await vi.advanceTimersByTimeAsync(5000); await h.settle();
    expect(h.container.textContent).toContain("Gespeichert und aktiv"); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
  });
  it("treats a rollback to baseline after proven persistence as a real conflict", async () => {
    const h=view("pending"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle();
    expect(h.container.textContent).toContain("Gespeichert");
    h.restoreBaseline(); h.notify(); await vi.advanceTimersByTimeAsync(250); await h.settle();
    expect(h.container.textContent).toContain("außerhalb dieser Seite geändert");
    expect(h.container.textContent).not.toContain("Speichern noch nicht bestätigt");
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.calls.filter(c=>c.method==="plugins.reload")).toHaveLength(0);
  });
  it("still refuses a real concurrent Graph policy change before writing", async () => {
    const h=view(); await h.settle(); await h.edit(); h.changeSaved(); h.click("Speichern und anwenden"); await h.settle();
    expect(h.container.textContent).toContain("anderswo geändert"); expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(0);
  });
  it("does not reread config or credential status for unrelated host updates once active", async () => {
    const h=view(); await h.settle(); const before=h.calls.length;
    for(let i=0;i<100;i++) h.notify(); await vi.advanceTimersByTimeAsync(5000); await h.settle();
    expect(h.calls).toHaveLength(before);
  });
  it("coalesces pending-status event bursts and avoids repeated credential reads once applied", async () => {
    const h=view("pending"); await h.settle(); await h.edit(); h.click("Speichern und anwenden"); await h.settle(); const before=h.calls.filter(c=>c.method==="config.get").length;
    h.setActive(); for(let i=0;i<100;i++) h.notify(); await vi.advanceTimersByTimeAsync(250); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.get")).toHaveLength(before+1); expect(h.container.textContent).toContain("Gespeichert und aktiv");
    const credentialReads=h.calls.filter(c=>c.method==="microsoft-graph.credentials.status").length;
    for(let i=0;i<100;i++) h.notify(); await vi.advanceTimersByTimeAsync(5000); await h.settle();
    expect(h.calls.filter(c=>c.method==="microsoft-graph.credentials.status")).toHaveLength(credentialReads);
  });

});

describe("automatic saved-policy tool healing", () => {
  it("repairs peers and stale exact tools on initial load using one root-only CAS even with includes", async () => {
    const h=view("active",true,true,true); await h.settle();
    const writes=h.calls.filter(c=>c.method==="config.patch"); expect(writes).toHaveLength(1);
    expect(writes[0]!.params.baseHash).toBe("r1");
    const raw=JSON.parse(writes[0]!.params.raw as string);
    expect(raw.plugins).toBeUndefined(); expect(raw.agents.entries.studio.tools.alsoAllow).toEqual(["onedrive_search","onedrive_list","onedrive_read","onedrive_download","microsoft_graph_capabilities"]);
    expect(h.container.textContent).toContain("Gespeichert und aktiv");
    h.hardReload(); await h.settle(); await vi.advanceTimersByTimeAsync(90000);
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false);
  });
  it("names the blocked agent and tools while repairing other agents without overriding denies", async () => {
    const h=view("active",true,true); await h.settle();
    h.agentEntries.main!.tools.alsoAllow=[]; (h.agentEntries.main!.tools as any).deny=["outlook_mail_*"];
    h.enableHealing(); h.hardReload(); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.agentEntries.main!.tools.alsoAllow).toEqual(["microsoft_graph_capabilities"]);
    expect(h.container.textContent).toContain("Main (main): durch Deny-Regeln gesperrt: outlook_mail_read");
    expect(h.container.textContent).toContain("Ausgewählter Agent: main");
    expect(h.container.textContent).toContain("Kein Speichervorgang läuft");
    expect(h.agentEntries.studio!.tools.alsoAllow).toContain("onedrive_read");
  });
  it("repairs fresh drift on status refresh and removes obsolete Graph tools without touching other tools", async () => {
    const h=view("active",false,true,true); await h.settle();
    h.agentEntries.main!.tools.alsoAllow=["read","outlook_mail_write"]; h.changeRevision();
    h.click("Status prüfen"); await h.settle();
    expect(h.agentEntries.main!.tools.alsoAllow).toEqual(["read","outlook_mail_read","microsoft_graph_capabilities"]);
    const writes=h.calls.filter(c=>c.method==="config.patch"); expect(writes).toHaveLength(2); expect(writes[1]!.params.baseHash).toBe("unrelated");
    expect(h.container.textContent).toContain("Gespeichert und aktiv");
  });
  it("observes a persisted lost response and does not resave or claim active tools prematurely", async () => {
    const h=view("persist-hang",true,true,true); await h.settle(); await vi.advanceTimersByTimeAsync(15000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.container.textContent).not.toContain("Gespeichert und aktiv");
    h.hardReload(); await h.settle(); h.setActive(); h.notify(); await vi.advanceTimersByTimeAsync(1000); await h.settle();
    expect(h.container.textContent).toContain("Gespeichert und aktiv");
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
  });
  it("never repairs during an unconfirmed or rejected save and does not loop on a CAS rejection", async () => {
    const h=view("reject",true,true,true); await h.settle(); await vi.advanceTimersByTimeAsync(90000); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(1);
    expect(h.container.textContent).not.toContain("Gespeichert und aktiv"); expect(h.calls.some(c=>c.method==="plugins.reload")).toBe(false);
  });
  it("does not dispatch after admin revocation or from an inactive policy generation", async () => {
    const h=view("active",true,true); await h.settle();
    h.enableHealing(); h.setInactive(); h.hardReload(); await h.settle();
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(0);
    h.host.connection.canAdmin=false; h.setActive(); h.notify(); await vi.advanceTimersByTimeAsync(90000);
    expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(0);
  });
});

it("does not label mismatched tools as active when generation proof is unavailable", async () => {
  const h=view("active",true,true); h.setBackendUnavailable(true); await h.settle();
  expect(h.container.textContent).toContain("Aktivierung nicht bestätigt");
  expect(h.container.textContent).not.toContain("Microsoft-Regeln aktiv. Agent-Tools weichen ab.");
  expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(0);
  expect(h.container.textContent).toContain("automatische Tool-Abgleich ist noch nicht verfügbar");
});

it("discovers a lost automatic repair claim without dispatching or retrying its write", async () => {
  const h=view("active",true,true); await h.settle(); h.loseClaim(); h.enableHealing(); h.hardReload(); await h.settle();
  await vi.advanceTimersByTimeAsync(90000); await h.settle();
  expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(0);
  expect(h.calls.filter(c=>c.method==="microsoft-graph.configuration.activation" && c.params.action==="reconcile")).toHaveLength(1);
  expect(h.container.textContent).toContain("nicht blind wiederholt");
});

it("explains why the last restrictive Graph allow entry cannot be automatically removed", async () => {
  const h=view("active",true); await h.settle();
  (h.agentEntries.studio as any).tools={allow:["outlook_calendar_read"]}; h.enableHealing(); h.hardReload(); await h.settle();
  expect(h.container.textContent).toContain("Studio (studio): Die letzte restriktive Tool-Freigabe bleibt erhalten.");
  expect(h.container.textContent).toContain("Eine leere Allow-Liste würde andere Tools freigeben.");
  expect(h.calls.filter(c=>c.method==="config.patch")).toHaveLength(0);
  expect((h.agentEntries.studio as any).tools.allow).toEqual(["outlook_calendar_read"]);
});
