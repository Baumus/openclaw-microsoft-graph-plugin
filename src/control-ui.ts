import { defineControlUiPlugin, type ControlUiHost } from "openclaw/plugin-sdk/control-ui";
import styles from "./control-ui.css" with { type: "text" };
import { localize, format, setLocale, isRtl } from "./control-ui-i18n.js";
import { isMicrosoftDeviceVerificationUri } from "./device-verification.js";
import { canonicalPolicy } from "./config-policy-identity.js";
import { isConnectionEstablished } from "./control-ui-status.js";
import { oneDriveToolExposureSaved, planOneDriveToolExposure } from "./control-ui-tool-exposure.js";
import { validClientId, validTenant } from "./app-identifiers.js";

type Grant = { operations: string[]; resources?: string[] };
type Root = { label: string; path: string; drive_id: string; item_id: string; include_descendants: true; agents_instructions?: "trusted"; permissions: Record<"read" | "write" | "delete", boolean>; agents: Record<string, { permissions: Partial<Record<"read" | "write" | "delete", boolean>> }> };
type Policy = { version: 2; rules: { default: "deny"; warningApprovalsByService?: Partial<Record<"onedrive" | "calendar" | "mail" | "contacts" | "todo", boolean>> }; services: { onedrive: { allowed_roots: Root[] }; calendar: { agents: Record<string, Grant> }; mail: { agents: Record<string, Grant> }; contacts?: { agents: Record<string, Grant> }; todo: { agents: Record<string, Grant> } } };
type CredentialStatus = { policyVersion: 2; credential: { result: "missing" | "valid" | "quarantined" | "unavailable"; missingScopes?: string[] } };
type DeviceStart = { sessionId: string; userCode: string; verificationUri: string; expiresAt: string; scopes: string[] };
type UpdateStatus = { updateAvailable: boolean; latestVersion?: string };
type CalendarCatalog = { defaultCalendarId: string; calendars: Array<{ id: string; name: string }> };
type DeviceStatus = { state: "pending" | "created" | "failed"; error?: string; scopes?: string[] };
type CredentialReply<T> = { ok: true; value: T } | { ok: false; error: string };
type ConfigSnapshot = { hash: string; config: { tools?: { deny?: string[] }; agents?: { entries?: Record<string, { tools?: { allow?: string[]; alsoAllow?: string[]; deny?: string[] } }> }; plugins?: { entries?: Record<string, { enabled?: boolean; config?: Record<string, unknown> }> } }; parsed?: { plugins?: { $include?: unknown; entries?: Record<string, { enabled?: boolean; config?: Record<string, unknown> }> } }; configRevisionHash?: string; appliedConfigHash?: string };
type SavePhase = "idle" | "validating" | "saving" | "applying" | "pending" | "unknown" | "failed" | "conflict" | "applied";
type SaveOperation = { target: Policy; baseline: Policy; toolExposure?: ReturnType<typeof planOneDriveToolExposure>["expected"]; targetHash?: string; phase: SavePhase; issued: boolean; settled: boolean; epoch: number; reloadAttempted?: boolean; reloadPending?: boolean; applicationFailed?: boolean; rejectedBeforeWrite?: boolean; persistenceConfirmed?: boolean };
// Private drafts live only in authenticated host memory, never browser storage.
const operationsByHost = new WeakMap<ControlUiHost, SaveOperation>();
const contextByHost = new WeakMap<ControlUiHost, { step: number; selectedAgent: string; credential?: CredentialStatus["credential"] }>();
// Core INVALID_REQUEST is a pre-commit validation/conflict rejection. A timeout,
// disconnect or UNAVAILABLE may follow a persisted write and is never retry proof.
function isPreCommitRejection(error: unknown): boolean {
  return !!error && typeof error === "object" && (error as { code?: unknown }).code === "INVALID_REQUEST";
}
const id = "microsoft-graph";
let configurationPageSequence = 0;
function authoredPolicySource(snapshot: ConfigSnapshot): string {
  const policy = snapshot.parsed?.plugins?.entries?.[id]?.config?.policy;
  if (policy && typeof policy === "object" && "$include" in policy) return `policy:${String((policy as { $include: unknown }).$include)}`;
  const plugins = snapshot.parsed?.plugins;
  if (plugins && "$include" in plugins) return `plugins:${String(plugins.$include)}`;
  return "";
}
const appIdsStorageKey = "microsoft-graph.app-ids.v1";
function savedAppIds(): { clientId: string; tenant: string } | undefined {
  try {
    const value = JSON.parse(localStorage.getItem(appIdsStorageKey) ?? "null") as unknown;
    if (value && typeof value === "object" && !Array.isArray(value)) {
      const ids = value as Record<string, unknown>;
      if (typeof ids.clientId === "string" && validClientId(ids.clientId) && typeof ids.tenant === "string" && validTenant(ids.tenant)) return { clientId: ids.clientId, tenant: ids.tenant };
    }
  } catch { /* Browser storage may be unavailable. The form remains usable. */ }
  return undefined;
}
function saveAppIds(clientId: string, tenant: string): void {
  try { localStorage.setItem(appIdsStorageKey, JSON.stringify({ clientId, tenant })); } catch { /* Non-secret convenience state is optional. */ }
}
const operations = { calendar: ["read", "create", "update", "respond", "attach", "delete"], mail: ["read", "draft", "update", "move", "mark", "send", "delete"], contacts: ["read", "create", "update", "delete"], todo: ["read", "create", "update", "delete"] } as const;
const serviceNames = { onedrive: "OneDrive", calendar: "Kalender", mail: "E-Mail", contacts: "Kontakte", todo: "To Do" } as const;
const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v)) as T;
const blankPolicy = (): Policy => ({ version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: {} }, mail: { agents: {} }, todo: { agents: {} } } });
function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = "", value?: string): HTMLElementTagNameMap[K] { const node = document.createElement(tag); if (className) node.className = className; if (value !== undefined) node.textContent = localize(value); return node; }
function append(parent: HTMLElement, ...children: HTMLElement[]) { parent.append(...children); }
function field(parent: HTMLElement, label: string, value: string, set: (value: string) => void, hint?: string) {
  const wrap = el("label", "mg-field"); const title = el("span", "mg-label", label); const input = el("input"); input.value = value; input.autocomplete = "off"; input.addEventListener("input", () => set(input.value)); input.addEventListener("change", () => set(input.value.trim())); append(wrap, title, input); if (hint) append(wrap, el("small", "mg-hint", hint)); append(parent, wrap); return input;
}
function checkbox(parent: HTMLElement, label: string, checked: boolean, set: (value: boolean) => void) {
  const wrap = el("label", "mg-check"); const input = el("input"); input.type = "checkbox"; input.checked = checked; input.addEventListener("change", () => set(input.checked)); append(wrap, input, el("span", "", label)); append(parent, wrap); return input;
}
function button(label: string, onClick: () => void, variant = "secondary") { const node = el("button", `mg-button ${variant}`, label); node.type = "button"; node.addEventListener("click", onClick); return node; }
function nestedArrayPaths(value: unknown, path: string, paths: string[]): void {
  if (Array.isArray(value)) { paths.push(path); return; }
  if (!value || typeof value !== "object") return;
  for (const [key, child] of Object.entries(value)) {
    if (key.includes(".")) throw new Error("Agent IDs containing dots cannot be saved with this editor");
    nestedArrayPaths(child, `${path}.${key}`, paths);
  }
}
function collectDiff(before: unknown, after: unknown, path = "", replacements: string[] = []): unknown {
  if (JSON.stringify(before) === JSON.stringify(after)) return undefined;
  if (Array.isArray(before) || Array.isArray(after)) { if (Array.isArray(before) && path) replacements.push(path); return after ?? null; }
  if (before && after && typeof before === "object" && typeof after === "object") {
    const out: Record<string, unknown> = {};
    for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
      if (key.includes(".")) throw new Error("Agent IDs containing dots cannot be saved with this editor");
      const next = collectDiff((before as Record<string, unknown>)[key], (after as Record<string, unknown>)[key], path ? `${path}.${key}` : key, replacements);
      if (next !== undefined) out[key] = next;
    }
    return Object.keys(out).length ? out : undefined;
  }
  if (after === undefined && before !== undefined) nestedArrayPaths(before, path, replacements);
  return after === undefined ? null : after;
}
function policySummary(policy: Policy): string[] {
  return [
    format("{count} OneDrive-Ordner", { count: policy.services.onedrive.allowed_roots.length }),
    ...(["calendar", "mail", "contacts", "todo"] as const).map(s => format("{count} Agenten für {service}", { count: Object.keys(policy.services[s]?.agents ?? {}).length, service: localize(serviceNames[s]) })),
  ];
}
class ConfigurationPage {
  private readonly helpIdPrefix = `mg-instructions-${++configurationPageSequence}`;
  private instructionsHelpRoot?: string;
  private readonly positionInstructionsHelp = () => {
    const panel = this.container.querySelector<HTMLElement>(".mg-instructions-help[data-open=true]");
    const trigger = panel && this.container.querySelector<HTMLElement>(`[aria-controls="${panel.id}"]`);
    if (!panel || !trigger) return;
    const anchor = trigger.getBoundingClientRect();
    const bounds = panel.getBoundingClientRect();
    const gap = 8, edge = 16;
    panel.style.left = `${Math.max(edge, Math.min(anchor.right - bounds.width, window.innerWidth - bounds.width - edge))}px`;
    const top = anchor.bottom + gap + bounds.height <= window.innerHeight - edge ? anchor.bottom + gap : anchor.top - bounds.height - gap;
    panel.style.top = `${Math.max(edge, Math.min(top, window.innerHeight - bounds.height - edge))}px`;
  };
  private snapshot?: ConfigSnapshot;
  private initial?: Policy;
  private policy = blankPolicy();
  private included = false;
  private includeName = "";
  private step = 0;
  private selectedAgent = "";
  private scopes: string[] = [];
  private initialScopes: string[] = [];
  private pluginEnabled = false;
  private newFolderPath = "";
  private editingAccess?: { rootLabel: string; agentId: string; permissions: Record<"read" | "write" | "delete", boolean> };
  private removedServiceGrants: Record<string, Grant> = {};
  private calendarCatalog?: CalendarCatalog;
  private calendarCatalogState: "idle" | "loading" | "ready" | "error" = "idle";
  private calendarLoadScheduled = false;
  private busy = false;
  private operation?: SaveOperation;
  private activePolicyHash?: string;
  private savedPolicyHash?: string;
  private activeProofAvailable = false;
  private statusInFlight = false;
  private statusEpoch = 0;
  private subscriptionStatusTimer?: ReturnType<typeof setTimeout>;
  private wasConnected = false;
  private validationEpoch = 0;
  private readonly guardEvent = (event: Event) => {
    if (!this.locked) return;
    const target = event.target as HTMLElement | null;
    if (target?.closest("[data-safe-status], summary")) return;
    if (target?.closest("button, input, select, textarea, a")) { event.preventDefault(); event.stopImmediatePropagation(); }
  };
  private readonly warnBeforeUnload = (event: BeforeUnloadEvent) => {
    if (this.operation && ["validating", "saving", "unknown"].includes(this.operation.phase)) { event.preventDefault(); event.returnValue = ""; }
  };
  private action(label: string, onClick: () => void, variant = "secondary") { return button(label, () => { if (!this.locked) onClick(); }, variant); }
  private field(parent: HTMLElement, label: string, value: string, set: (value: string) => void, hint?: string) { return field(parent, label, value, next => { if (!this.locked) set(next); }, hint); }
  private checkbox(parent: HTMLElement, label: string, checked: boolean, set: (value: boolean) => void) { return checkbox(parent, label, checked, next => { if (!this.locked) set(next); }); }
  private async read<T>(method: string, params: Record<string, unknown> = {}): Promise<T> {
    let timer: ReturnType<typeof setTimeout> | undefined;
    try { return await Promise.race([this.host.request<T>(method, params), new Promise<never>((_, reject) => { timer = setTimeout(() => reject(new Error("status_timeout")), 15000); })]); }
    finally { if (timer) clearTimeout(timer); }
  }
  private async loadCalendars() {
    if (!this.authorized || this.calendarCatalogState === "loading" || this.disposed) return;
    this.calendarCatalogState = "loading"; this.render();
    try {
      const result = await this.read<CalendarCatalog>("microsoft-graph.configuration.calendars");
      if (this.disposed || this.signal.aborted) return;
      if (!result || typeof result.defaultCalendarId !== "string" || !Array.isArray(result.calendars)) throw new Error("invalid_calendar_catalog");
      this.calendarCatalog = result; this.calendarCatalogState = "ready";
    } catch { if (!this.disposed) { this.calendarCatalog = undefined; this.calendarCatalogState = "error"; } }
    finally { if (!this.disposed) this.render(); }
  }
  private get locked() { return this.busy || !!this.operation && !["idle", "failed", "applied"].includes(this.operation.phase); }
  private error = "";
  private success = "";
  private statusError = "";
  private confirmedRemoval = false;
  private disposed = false;
  private statusTimer?: ReturnType<typeof setTimeout>;
  private statusChecksRemaining = 0;
  private authTimer?: ReturnType<typeof setTimeout>;
  private credential?: CredentialStatus["credential"];
  private updateStatus?: UpdateStatus;
  private updateCheckStarted = false;
  private authClientId = "";
  private authTenant = "";
  private editAppIds = false;
  private authFailureCode = "";
  private authFailureAt = "";
  private device?: DeviceStart;
  private authState: "idle" | "pending" | "created" | "failed" = "idle";
  private authError = "";
  private authBusy = false;
  private authEpoch = 0;
  private authRenewal = false;
  private readonly disclosureState = new Map<string, boolean>();
  private readonly unsubscribe: () => void;
  constructor(private container: HTMLElement, private host: ControlUiHost, private signal: AbortSignal) {
    this.operation = operationsByHost.get(host);
    const context = contextByHost.get(host);
    if (context) { this.step = context.step; this.selectedAgent = context.selectedAgent; this.credential = context.credential; }
    this.container.addEventListener("click", this.guardEvent, true);
    this.container.addEventListener("change", this.guardEvent, true);
    this.container.addEventListener("input", this.guardEvent, true);
    window.addEventListener("beforeunload", this.warnBeforeUnload);
    window.addEventListener("resize", this.positionInstructionsHelp);
    window.addEventListener("scroll", this.positionInstructionsHelp, true);
    const stored = savedAppIds();
    if (stored) { this.authClientId = stored.clientId; this.authTenant = stored.tenant; }
    this.unsubscribe = host.subscribe(() => {
      if (!this.authorized) {
        this.authEpoch++; this.authBusy = false;
        if (this.authTimer) clearTimeout(this.authTimer);
        this.authTimer = undefined;
      } else this.scheduleAuthCheck();
      setLocale(host.locale); if (this.authorized) void this.checkForUpdate(); if (!this.snapshot && !this.busy && this.authorized) void this.load();
      else {
        this.render();
        // The host publishes many unrelated session/agent updates. Coalesce
        // those notifications rather than reading the entire config per event.
        if (this.authorized && this.snapshot && !this.busy && !this.subscriptionStatusTimer
          && (!this.wasConnected || this.operation && !["applied", "failed", "conflict"].includes(this.operation.phase))) {
          this.subscriptionStatusTimer = setTimeout(() => { this.subscriptionStatusTimer = undefined; void this.checkApplication(); }, 250);
        }
      }
      this.wasConnected = this.authorized;
    });
    void this.load();
    void this.checkForUpdate();
    this.render();
  }
  dispose() { contextByHost.set(this.host, { step: this.step, selectedAgent: this.selectedAgent, credential: this.credential }); this.disposed = true; this.statusEpoch++; if (this.subscriptionStatusTimer) clearTimeout(this.subscriptionStatusTimer); this.container.removeEventListener("click", this.guardEvent, true); this.container.removeEventListener("change", this.guardEvent, true); this.container.removeEventListener("input", this.guardEvent, true); window.removeEventListener("beforeunload", this.warnBeforeUnload); window.removeEventListener("resize", this.positionInstructionsHelp); window.removeEventListener("scroll", this.positionInstructionsHelp, true); if (this.authTimer) clearTimeout(this.authTimer); this.stopStatusChecks(); this.unsubscribe(); this.container.replaceChildren(); }
  private async checkForUpdate() {
    if (!this.authorized || this.updateCheckStarted || this.disposed) return;
    this.updateCheckStarted = true;
    try {
      if (this.disposed || !this.authorized) return;
      const result = await this.host.request<UpdateStatus>("microsoft-graph.updateStatus", {});
      if (!this.disposed && !this.signal.aborted && result.updateAvailable === true && typeof result.latestVersion === "string") { this.updateStatus = result; this.render(); }
    } catch { /* Version information is optional; never block configuration. */ }
  }
  private get dirty() { return !!this.initial && canonicalPolicy(this.initial) !== canonicalPolicy(this.policy); }
  private get authorized() { return this.host.connection.connected && this.host.connection.canAdmin; }
  private get signInReady() {
    const entry = this.snapshot?.config?.plugins?.entries?.[id]?.config;
    const policy = this.initial;
    const grants = !!policy && (policy.services.onedrive.allowed_roots.some(root => Object.values(root.agents).some(agent => Object.values(agent.permissions).some(Boolean)))
      || (["calendar", "mail", "contacts", "todo"] as const).some(service => Object.keys(policy.services[service]?.agents ?? {}).length > 0));
    return !!entry?.credentialVaultKey && grants && !this.dirty && !this.statusError && this.applicationStatus === "applied";
  }
  private get applicationStatus(): "applied" | "pending" | "unknown" {
    if (this.savedPolicyHash && this.activePolicyHash) return this.savedPolicyHash === this.activePolicyHash ? "applied" : "pending";
    if (this.activeProofAvailable) return "unknown";
    // Older backends can only prove the whole saved revision is active.
    const { configRevisionHash, appliedConfigHash } = this.snapshot ?? {};
    return configRevisionHash && appliedConfigHash && configRevisionHash === appliedConfigHash ? "applied" : "unknown";
  }
  private stopStatusChecks() { if (this.statusTimer) clearTimeout(this.statusTimer); this.statusTimer = undefined; this.statusChecksRemaining = 0; }
  private scheduleStatusCheck() {
    if (this.disposed || this.signal.aborted || this.statusChecksRemaining <= 0 || this.statusTimer) return;
    // The first read can confirm a committed write while config.patch is still
    // waiting for runtime drain. Later reads stay paced to avoid RPC churn.
    const delay = this.statusChecksRemaining === 12 ? 1000 : 5000;
    this.statusTimer = setTimeout(() => { this.statusTimer = undefined; void this.checkApplication(); }, delay);
  }
  private async readActivePolicy() {
    try {
      const reply = await this.read<{ activePolicyHash: string | null }>("microsoft-graph.configuration.applicationStatus");
      return { hash: typeof reply.activePolicyHash === "string" && /^[a-f0-9]{64}$/.test(reply.activePolicyHash) ? reply.activePolicyHash : undefined, available: true };
    }
    catch { return { hash: undefined, available: false }; }
  }
  private async checkApplication() {
    if (!this.authorized || !this.snapshot || !this.initial || this.disposed || this.signal.aborted || this.statusInFlight) return;
    const epoch = this.statusEpoch;
    this.statusInFlight = true; this.render();
    const previouslyApplied = this.operation?.phase === "applied" || !this.operation && this.applicationStatus === "applied";
    try {
      const [snap, active] = await Promise.all([
        this.read<ConfigSnapshot>("config.get"),
        this.readActivePolicy(),
      ]);
      if (this.disposed || this.signal.aborted || epoch !== this.statusEpoch) return;
      const expected = this.operation?.target ?? this.initial;
      const saved = snap.config?.plugins?.entries?.[id]?.config?.policy;
      if (canonicalPolicy(saved) !== canonicalPolicy(expected) || !oneDriveToolExposureSaved(snap.config?.agents?.entries, this.operation?.toolExposure ?? {})) {
        if (this.operation?.issued && !this.operation.rejectedBeforeWrite && !this.operation.persistenceConfirmed && canonicalPolicy(saved) === canonicalPolicy(this.operation.baseline)) {
          // A transport failure is not an external edit or proof of rollback.
          // Keep observing the exact target for the bounded reconciliation window.
          this.statusError = "";
          this.operation.phase = this.operation.settled ? "unknown" : "saving";
        } else {
          this.stopStatusChecks();
          if (this.operation?.rejectedBeforeWrite && canonicalPolicy(saved) === canonicalPolicy(this.operation.baseline)) {
            this.snapshot = snap; this.operation.phase = "failed"; this.statusError = "";
          } else {
            this.statusError = "Die Regeln wurden außerhalb dieser Seite geändert. Der Entwurf bleibt erhalten.";
            if (this.operation) this.operation.phase = "conflict";
          }
        }
      } else {
        this.snapshot = snap; this.activePolicyHash = active.hash; this.activeProofAvailable = active.available; this.statusError = "";
        if (this.operation) { this.operation.persistenceConfirmed = true; this.initial = clone(expected); this.savedPolicyHash = this.operation.targetHash; this.operation.phase = this.applicationStatus === "applied" ? "applied" : this.operation.reloadPending || this.statusChecksRemaining > 0 ? "applying" : "pending"; }
        if (this.applicationStatus === "applied") { this.stopStatusChecks(); if (!previouslyApplied) await this.refreshCredential(); }
      }
    } catch {
      if (epoch === this.statusEpoch && !this.disposed) { this.statusError = "Verbindung unterbrochen. Der Status wird nach Wiederverbindung geprüft."; if (this.operation && !["applied", "applying", "pending"].includes(this.operation.phase)) this.operation.phase = "unknown"; }
    } finally {
      this.statusInFlight = false;
      if (epoch !== this.statusEpoch && !this.disposed && this.operation) this.scheduleStatusCheck();
      if (epoch === this.statusEpoch && !this.disposed) {
        if (this.statusChecksRemaining > 0) { this.statusChecksRemaining--; this.scheduleStatusCheck(); }
        if (this.operation && this.statusChecksRemaining === 0 && this.operation.phase === "applying") this.operation.phase = "pending";
        if (this.operation && this.statusChecksRemaining === 0 && this.operation.phase === "saving") this.operation.phase = "unknown";
        this.render();
      }
    }
  }
  private watchApplication() { this.stopStatusChecks(); if (this.applicationStatus !== "applied" || this.operation?.issued && !["applied", "failed", "conflict"].includes(this.operation.phase)) { this.statusChecksRemaining = 12; this.scheduleStatusCheck(); } }
  private async load() {
    if (!this.authorized || this.busy) return;
    this.busy = true; this.error = ""; this.render();
    try {
      const snap = await this.read<ConfigSnapshot>("config.get");
      if (this.disposed || this.signal.aborted) return;
      const pluginEntry = snap.config?.plugins?.entries?.[id];
      const entry = pluginEntry?.config ?? {};
      this.pluginEnabled = pluginEntry?.enabled === true && entry.enabled !== false;
      this.includeName = authoredPolicySource(snap);
      this.included = !!this.includeName;
      const policy = entry.policy as Policy | undefined;
      this.initial = policy ? clone(policy) : blankPolicy();
      this.policy = this.operation && this.operation.phase !== "applied" ? clone(this.operation.target) : clone(this.initial);
      this.snapshot = snap;
      this.watchApplication();
      this.selectedAgent ||= this.host.agents.rows[0]?.id ?? ""; this.removedServiceGrants = {};
      const [, baseline, active] = await Promise.all([
        this.refreshCredential(),
        this.read<{ requiredScopes: string[]; policyHash?: string }>("microsoft-graph.configuration.validate", { policy: this.initial }).catch(() => undefined),
        this.readActivePolicy(),
      ]);
      if (this.disposed || this.signal.aborted) return;
      this.initialScopes = baseline?.requiredScopes ?? []; this.scopes = this.initialScopes; this.savedPolicyHash = baseline?.policyHash;
      this.activePolicyHash = active.hash; this.activeProofAvailable = active.available;
      this.wasConnected = this.authorized;
      if (!this.operation && this.applicationStatus === "pending") {
        this.operation = { target: clone(this.initial), baseline: clone(this.initial), targetHash: this.savedPolicyHash, phase: "pending", issued: true, settled: true, epoch: this.statusEpoch };
        operationsByHost.set(this.host, this.operation);
      }
      if (this.operation) { this.statusChecksRemaining = 12; void this.checkApplication(); }
      else this.watchApplication();
    } catch { this.error = "Configuration could not be loaded. Check administrator access and Gateway connection."; }
    finally { this.busy = false; this.render(); }
  }
  private render() {
    if (this.disposed || this.signal.aborted) return;
    setLocale(this.host.locale);
    if (this.dirty && this.operation?.phase === "applied") { operationsByHost.delete(this.host); this.operation = undefined; }
    const main = el("main", "mg-ui"); main.dir = isRtl() ? "rtl" : "ltr"; const header = el("header", "mg-header");
    append(header, el("div", "mg-eyebrow", "Plugins / Connect Microsoft 365 to OpenClaw"), el("h1", "", "Connect Microsoft 365 to OpenClaw"), el("p", "mg-lead", "Lege fest, welcher Agent auf welche Microsoft-Daten zugreifen darf und wann eine Freigabe nötig ist."));
    append(main, header);
    if (!this.host.connection.connected) { append(main, el("p", "mg-message", this.operation ? "Verbindung unterbrochen. Der Status wird nach Wiederverbindung geprüft." : "Verbinde dich mit dem Gateway, um die Regeln zu bearbeiten.")); this.container.replaceChildren(main); return; }
    if (!this.host.connection.canAdmin) { append(main, el("p", "mg-message", "Zum Anzeigen und Ändern dieser Regeln brauchst du Administratorrechte.")); this.container.replaceChildren(main); return; }
    if (!this.snapshot) { append(main, el("p", "mg-message", this.error || "Regeln werden geladen…")); this.container.replaceChildren(main); return; }
    this.renderSetup(main);
    this.renderSignIn(main);
    const rail = el("nav", "mg-steps"); rail.setAttribute("aria-label", localize("Konfigurationsschritte"));
    ["OneDrive", "Dienste", "Freigaben", "Prüfen"].forEach((name, index) => { const tab = this.action(`${index + 1}  ${localize(name)}`, () => { this.step = index; this.render(); if (index === 3) void this.validate(); }, index === this.step ? "active" : "ghost"); tab.disabled = this.locked; tab.setAttribute("aria-current", index === this.step ? "step" : "false"); append(rail, tab); }); append(main, rail);
    if (this.included && (this.dirty || !!this.error)) append(main, el("p", "mg-banner", format("Konfigurationsquelle: {name}. Änderungen werden beim Speichern in diese Datei geschrieben.", { name: this.includeName.slice(this.includeName.indexOf(":") + 1) })));
    this.renderSaveStatus(main);
    const body = el("section", "mg-body"); if (this.step === 0) this.renderOneDrive(body); else if (this.step === 1) this.renderServices(body); else if (this.step === 2) this.renderApprovals(body); else this.renderReview(body); append(main, body);
    if (this.error) append(main, el("p", "mg-error", this.error)); if (this.success) append(main, el("p", "mg-success", this.success));
    const footer = el("footer", "mg-footer"); append(footer, el("span", "mg-dirty", this.dirty ? "Ungespeicherte Änderungen" : "Keine ungespeicherten Änderungen"));
    if (this.dirty) append(footer, this.action("Änderungen verwerfen", () => { if (window.confirm(localize("Alle ungespeicherten Änderungen verwerfen?"))) { this.policy = clone(this.initial!); this.error = ""; this.success = ""; this.render(); } }, "ghost"));
    if (this.step > 0) append(footer, this.action("Zurück", () => { this.step--; this.render(); }));
    if (this.step < 3) append(footer, this.action("Weiter", () => { this.step++; this.render(); if (this.step === 3) void this.validate(); }, "primary"));
    if (this.step === 3 && this.dirty) { const save = this.action(this.operation?.phase === "validating" || this.operation?.phase === "saving" ? "Speichert …" : "Speichern und anwenden", () => { void this.save(); }, "primary"); save.disabled = this.authBusy || this.authState === "pending"; append(footer, save); }
    if (this.step === 3 && this.dirty && !this.operation) append(footer, el("p", "mg-save-hint", "Wir speichern die Regeln und wenden sie automatisch an. Die Verbindung kann dabei kurz unterbrochen werden."));
    if (this.operation && !["failed", "applied"].includes(this.operation.phase)) {
      const saving = ["validating", "saving"].includes(this.operation.phase);
      append(footer, el("span", "mg-footer-status", saving ? "Speichert …" : this.operation.phase === "applying" ? "Wendet an …" : "Status noch nicht bestätigt"));
      if (["pending", "unknown", "conflict"].includes(this.operation.phase)) this.appendStatusButton(footer);
    }
    append(main, footer);
    main.setAttribute("aria-busy", String(this.busy || this.operation?.phase === "saving" || this.operation?.phase === "applying"));
    if (this.locked) for (const control of main.querySelectorAll<HTMLButtonElement | HTMLInputElement | HTMLSelectElement>("button:not([data-safe-status]), input, select, textarea")) control.disabled = true;
    const focused = this.container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : undefined;
    const controls = [...this.container.querySelectorAll<HTMLElement>("button, input, select, textarea, summary, [tabindex]")];
    const focusIndex = focused ? controls.indexOf(focused) : -1;
    const focusText = focused?.textContent;
    const focusLabel = focused?.closest("label")?.textContent;
    const selection = focused instanceof HTMLInputElement ? [focused.selectionStart, focused.selectionEnd] : undefined;
    // Background status/update renders must not collapse reading context.
    const disclosureKey = (node: HTMLDetailsElement) => node.querySelector("summary")?.textContent ?? "";
    for (const details of this.container.querySelectorAll("details")) this.disclosureState.set(disclosureKey(details), details.open);
    for (const details of main.querySelectorAll("details")) details.open = this.disclosureState.get(disclosureKey(details)) ?? details.open;
    const scrollTop = this.container.scrollTop;
    this.container.replaceChildren(main);
    this.container.scrollTop = scrollTop;
    if (!this.locked && this.instructionsHelpRoot) {
      const panel = [...main.querySelectorAll<HTMLElement>(".mg-instructions-help")].find(node => node.dataset.root === this.instructionsHelpRoot);
      if (panel) { panel.showPopover({ source: main.querySelector<HTMLElement>(`[aria-controls="${panel.id}"]`) ?? undefined }); this.positionInstructionsHelp(); }
      else this.instructionsHelpRoot = undefined;
    }
    if (this.step === 1 && this.calendarCatalogState === "idle" && !this.calendarLoadScheduled) {
      this.calendarLoadScheduled = true;
      queueMicrotask(() => { this.calendarLoadScheduled = false; if (this.step === 1 && this.calendarCatalogState === "idle" && !this.disposed) void this.loadCalendars(); });
    }
    if (focused) {
      const candidates = [...main.querySelectorAll<HTMLElement>("button, input, select, textarea, summary, [tabindex]")];
      const replacement = (focused.dataset.focusKey ? candidates.find(node => node.dataset.focusKey === focused.dataset.focusKey) : undefined) ?? candidates.find(node => node.tagName === focused.tagName && (focusLabel ? node.closest("label")?.textContent === focusLabel : node.textContent === focusText)) ?? candidates[focusIndex];
      if (replacement && !(replacement as HTMLButtonElement).disabled) { replacement.focus({ preventScroll: true }); if (selection && replacement instanceof HTMLInputElement && selection[0] !== null && selection[1] !== null) try { replacement.setSelectionRange(selection[0]!, selection[1]!); } catch { /* Checkbox inputs have no selection. */ } }
      else if (this.operation) main.querySelector<HTMLElement>(".mg-save-status")?.focus({ preventScroll: true });
    }

  }
  private async applySavedPolicy() {
    const operation = this.operation;
    if (!operation || !operation.issued || !operation.settled || operation.reloadPending || !["applying", "pending"].includes(operation.phase) || !this.authorized || this.disposed) return;
    operation.reloadPending = true; operation.reloadAttempted = true; operation.phase = "applying"; this.render();
    try {
      // Recheck the exact target. Do not rewrite config (including SecretRefs),
      // force a restart, or interrupt admitted work to make the badge green.
      const fresh = await this.read<ConfigSnapshot>("config.get");
      if (this.disposed || this.signal.aborted || !this.authorized || this.operation !== operation) return;
      if (canonicalPolicy(fresh.config?.plugins?.entries?.[id]?.config?.policy) !== canonicalPolicy(operation.target)) { operation.phase = "conflict"; return; }
      const result = await this.host.request<{ ok?: boolean; restartRequired?: boolean }>("plugins.reload", { plugins: [{ pluginId: id }], waitForDrain: true });
      if (!result.ok || result.restartRequired) { operation.applicationFailed = true; operation.phase = "pending"; }
    } catch { operation.applicationFailed = true; operation.phase = "pending"; }
    finally {
      operation.reloadPending = false;
      if (operation.phase === "applying") operation.phase = "pending";
      if (!this.disposed && this.operation === operation) { this.watchApplication(); await this.checkApplication(); this.render(); }
    }
  }
  private appendStatusButton(parent: HTMLElement) {
    const check = button(this.operation?.phase === "unknown" ? "Speicherstatus prüfen" : "Status prüfen", () => { void this.checkApplication(); });
    check.dataset.safeStatus = "true"; check.disabled = this.statusInFlight || !this.authorized; append(parent, check);
  }
  private renderSaveStatus(main: HTMLElement) {
    const phase = this.operation?.phase;
    if (!phase && this.applicationStatus === "applied" && !this.statusError) return;
    const status = el("section", phase === "applied" ? "mg-save-status mg-success" : phase === "failed" ? "mg-save-status mg-error" : "mg-save-status mg-warning");
    status.setAttribute("role", "status"); status.setAttribute("aria-live", "polite"); status.tabIndex = -1;
    let title = "Gespeichert. Aktivierung noch nicht bestätigt.";
    let detail = "Die bisherigen Regeln können noch gelten. Beim Wiederöffnen wird der Stand geprüft. Wenn die Verbindung während der Anwendung abbricht, kann ein neuer Anwendungsversuch nötig sein.";
    if (phase === "validating" || phase === "saving") { title = "Deine Änderungen werden gespeichert …"; detail = phase === "validating" ? "Zugriffe werden geprüft. Bitte warte; weitere Änderungen sind vorübergehend gesperrt." : "Bitte warte. Wir bestätigen zuerst, dass deine Änderungen gespeichert sind."; }
    else if (phase === "applying") { title = "Gespeichert. Änderungen werden jetzt angewendet …"; detail = "Die Aktivierung wird automatisch geprüft. Die bisherigen Regeln können noch gelten."; }
    else if (phase === "applied") { title = "Gespeichert und aktiv"; detail = "Die gespeicherten Microsoft-Regeln gelten jetzt. Du kannst wieder Änderungen bearbeiten."; }
    else if (phase === "unknown") { title = "Speichern noch nicht bestätigt"; detail = this.statusChecksRemaining > 0 ? "Bitte warte. Wir bestätigen zuerst, dass deine Änderungen gespeichert sind." : "Prüfung pausiert. Prüfe zuerst den Speicherstatus; wir speichern nicht erneut. Dein Entwurf bleibt in dieser Ansicht erhalten."; }
    else if (phase === "failed") { title = "Nicht gespeichert. Dein Entwurf ist noch da."; detail = "Korrigiere die Ursache und versuche es erneut."; }
    else if (phase === "conflict") { title = "Die Regeln wurden inzwischen anderswo geändert."; detail = "Dein Entwurf bleibt erhalten. Lade den aktuellen Stand neu, bevor du weitere Änderungen speicherst."; }
    append(status, el("strong", "", title), el("p", "mg-status-detail", detail));
    if (this.operation?.applicationFailed && ["pending", "applying"].includes(phase ?? "")) append(status, el("p", "mg-status-detail", "Die automatische Anwendung konnte nicht abgeschlossen werden. Deine Regeln sind gespeichert; laufende Arbeit wird nicht abgebrochen."));
    if (this.operation?.reloadPending) append(status, el("p", "mg-status-detail", "Anwendung wurde angefordert. Wir prüfen, wann die gespeicherten Regeln aktiv sind."));
    if (this.statusError) append(status, el("p", "mg-status-detail", this.statusError));
    if (!phase || ["pending", "unknown", "conflict"].includes(phase)) this.appendStatusButton(status);
    if (["pending", "applying"].includes(phase ?? "") && this.operation?.settled && !this.operation.reloadPending) {
      const apply = button("Gespeicherte Regeln anwenden", () => { void this.applySavedPolicy(); }); apply.dataset.safeStatus = "true"; apply.disabled = this.statusInFlight || this.busy; append(status, apply);
      append(status, el("p", "mg-hint", "Dabei werden nur die gespeicherten Regeln übernommen, nicht erneut gespeichert. Laufende Arbeit wird nicht abgebrochen."));
    }
    if (phase === "pending") {
      const overview = button("Übersicht anzeigen", () => { this.step = 0; this.render(); }); overview.dataset.safeStatus = "true"; append(status, overview);
    }
    if (phase === "conflict") { const reload = button("Aktuellen Stand laden", () => { if (window.confirm(localize("Alle ungespeicherten Änderungen verwerfen?"))) { operationsByHost.delete(this.host); this.operation = undefined; this.statusEpoch++; void this.load(); } }); reload.dataset.safeStatus = "true"; reload.disabled = this.busy || this.statusInFlight; append(status, reload); }
    append(main, status);
  }
  private renderSetup(main: HTMLElement) {
    const configured = this.snapshot?.config?.plugins?.entries?.[id]?.config?.credentialVaultKey !== undefined;
    const savedPolicy = this.initial ?? blankPolicy();
    const grants = savedPolicy.services.onedrive.allowed_roots.some(root => Object.values(root.agents).some(agent => Object.values(agent.permissions).some(Boolean)))
      || (["calendar", "mail", "contacts", "todo"] as const).some(service => Object.keys(savedPolicy.services[service]?.agents ?? {}).length > 0);
    const connected = this.credential?.result === "valid" && (this.credential.missingScopes?.length ?? 0) === 0;
    if (isConnectionEstablished({ secretRefConfigured: configured, savedGrantPresent: grants, applicationStatus: this.applicationStatus, credentialResult: this.credential?.result, missingScopes: this.credential?.missingScopes, statusError: !!this.statusError })) {
      const status = el("p", "mg-success mg-connected", "Verbindung hergestellt");
      status.setAttribute("role", "status");
      status.setAttribute("aria-label", localize("Verbindung hergestellt. Ein Lesezugriff durch einen berechtigten Agenten wurde nicht geprüft."));
      const row = el("div", "mg-top-status");
      append(row, status);
      if (this.updateStatus?.updateAvailable) {
        const badge = el("span", "mg-update-badge", "Update verfügbar");
        badge.setAttribute("role", "status");
        badge.setAttribute("aria-label", format("Update verfügbar: Version {version}", { version: this.updateStatus.latestVersion ?? "" }));
        append(row, badge);
      }
      append(main, row);
      return;
    }
    if (configured && grants && this.credential?.result === "missing") return;
    const card = el("section", "mg-section mg-setup");
    append(card, el("h2", "", "Einrichtung"));
    const steps = [
      ["Vault-SecretRef vorhanden", configured],
      ["Mindestens ein Agentenzugriff gespeichert", grants],
      ["Regeln im Gateway angewendet", grants && this.applicationStatus === "applied" && !this.statusError],
      ["Microsoft-Konto verbunden", connected],
    ] as const;
    const list = el("ol", "mg-setup-list");
    for (const [label, done] of steps) append(list, el("li", done ? "mg-setup-done" : "", `${done ? "✓" : "○"} ${localize(label)}`));
    append(card, list);
    let next: string;
    if (!configured) next = "Nächster Schritt: Vault-Schlüssel als SecretRef hinterlegen. Die Anleitung zeigt den Befehl.";
    else if (!grants && !this.dirty) next = "Nächster Schritt: Einen Agentenzugriff auswählen und die Regeln speichern.";
    else if (this.dirty) next = "Nächster Schritt: Änderungen unter Prüfen speichern.";
    else if (this.applicationStatus !== "applied") next = "Nächster Schritt: Warten, bis das Gateway die gespeicherten Regeln angewendet hat.";
    else if (this.credential?.result === "quarantined") next = "Der Zugang ist gesperrt. Stelle ihn über die Administrator-Wiederherstellung wieder her.";
    else if (this.credential?.result === "unavailable") next = "Der Zugangsstatus ist nicht verfügbar. Prüfe SecretRef und Gateway-Verbindung.";
    else if (this.credential?.result === "valid" && this.credential.missingScopes?.length) next = "Neue Microsoft-Berechtigungen fehlen. Folge unten der Reautorisierung; der bestehende Zugang bleibt bis zum erfolgreichen Abschluss erhalten.";
    else if (!connected) next = "Nächster Schritt: Unten mit Microsoft verbinden.";
    else next = "Konto verbunden. Prüfe mit einem berechtigten Agenten einen Lesezugriff; erst dann ist der Ablauf einsatzbereit.";
    append(card, el("p", connected ? "mg-status" : "mg-hint", next));
    if (!configured) { const link = el("a", "mg-button secondary", "Vault-Anleitung öffnen"); link.href = "https://clawhub.ai/packages/@baumus/openclaw-microsoft-graph"; link.target = "_blank"; link.rel = "noopener noreferrer"; append(card, link); }
    else if (!grants && !this.dirty) append(card, this.action("Zugriff festlegen", () => { this.step = 1; this.render(); }, "primary"));
    else if (this.dirty) append(card, this.action("Zum Prüfen", () => { this.step = 3; this.render(); void this.validate(); }, "primary"));
    else if (grants && this.applicationStatus === "applied" && !connected && this.credential?.result === "missing") append(card, this.action("Mit Microsoft verbinden", () => { this.container.querySelector("#microsoft-connect")?.scrollIntoView({ behavior: "smooth", block: "start" }); }, "primary"));
    else if (this.credential?.result === "unavailable") append(card, this.action("Berechtigungen erneut prüfen", () => { void this.refreshCredential().then(() => this.render()); }, "secondary"));
    append(main, card);
  }
  private async credentialCall<T>(method: string, params: Record<string, unknown>): Promise<T> {
    const reply = await this.host.request<CredentialReply<T>>(method, params);
    if (!reply || typeof reply !== "object" || typeof reply.ok !== "boolean") throw new Error("internal_error");
    if (!reply.ok) throw new Error(reply.error);
    return reply.value;
  }
  private async refreshCredential(isCurrent: () => boolean = () => true) {
    try {
      const status = await this.credentialCall<CredentialStatus>("microsoft-graph.credentials.status", {});
      if (!this.disposed && !this.signal.aborted && this.authorized && isCurrent()) this.credential = status.credential;
    } catch { if (!this.disposed && !this.signal.aborted && this.authorized && isCurrent()) this.credential = { result: "unavailable" }; }
  }
  private signInError(code: string): string {
    const messages: Record<string, string> = {
      credential_scope_missing: "Die Microsoft-Freigabe enthält nicht alle benötigten Berechtigungen. Prüfe die Einwilligung und starte erneut.",
      credential_vault_conflict: "Der Zugang hat sich geändert. Prüfe den Status und starte die Erneuerung erneut. Der Assistent überschreibt keinen ungeprüften Zugang.",
      credential_vault_conflict_exists: "Es gibt bereits einen Zugang. Prüfe den Status und nutze Microsoft-Berechtigungen erneuern statt einer neuen Verbindung.",
      credential_vault_conflict_identity: "App oder Mandant stimmen nicht mit dem bestehenden Zugang überein. Erneuere mit den vorhandenen Kennungen; lege keine neue App an.",
      credential_vault_conflict_changed: "Der Zugang wurde während der Anmeldung geändert, etwa durch eine Token-Erneuerung. Prüfe den Status und starte die Erneuerung erneut. Der aktuelle Zugang bleibt erhalten.",
      credential_vault_conflict_policy_changed: "Die benötigten Berechtigungen wurden während der Anmeldung geändert. Prüfe die angewendeten Regeln und starte die Erneuerung erneut.",
      credential_vault_conflict_quarantined: "Ein früherer Token-Austausch hat einen unklaren Ausgang. Der Zugang bleibt gesperrt; ein Admin muss die Wiederherstellung prüfen. Nicht löschen oder erneut verbinden.",
      credential_vault_locked: "Der Zugang wird gerade geändert. Bitte später erneut versuchen.",
      credential_vault_unavailable: "Der verschlüsselte Zugang ist nicht bereit. Prüfe den Vault-Schlüssel in der Plugin-Konfiguration.",
      device_authorization_declined: "Die Anmeldung wurde bei Microsoft abgelehnt. Du kannst erneut starten.",
      device_authorization_expired: "Der Anmeldecode ist abgelaufen. Starte die Anmeldung erneut.",
      device_authorization_cancelled: "Die Anmeldung wurde abgebrochen.",
      device_authorization_in_progress: "Eine Anmeldung läuft bereits. Warte auf ihren Abschluss oder brich sie ab.",
      invalid_policy: "Lege zuerst mindestens einen Agentenzugriff fest und speichere die Regeln.",
    };
    return messages[code] ?? "Die Anmeldung konnte nicht abgeschlossen werden. Prüfe App-ID, Tenant und Microsoft-Einwilligung; versuche es erneut.";
  }
  private authCurrent(epoch: number, sessionId?: string): boolean {
    return !this.disposed && !this.signal.aborted && this.authorized && epoch === this.authEpoch
      && (sessionId === undefined || this.device?.sessionId === sessionId);
  }
  private scheduleAuthCheck() {
    if (!this.authCurrent(this.authEpoch) || this.authBusy || this.authState !== "pending" || !this.device || this.authTimer) return;
    this.authTimer = setTimeout(() => { this.authTimer = undefined; void this.checkAuth(); }, 4000);
  }
  private async checkAuth() {
    if (!this.device || this.authState !== "pending" || this.authBusy || !this.authCurrent(this.authEpoch)) return;
    const epoch = this.authEpoch, device = this.device;
    const current = () => this.authCurrent(epoch, device.sessionId);
    try {
      const result = await this.credentialCall<DeviceStatus>("microsoft-graph.credentials.device-status", { sessionId: device.sessionId });
      if (!current() || this.authState !== "pending") return;
      this.authError = "";
      if (result.state === "created") { this.authState = "created"; this.authError = ""; await this.refreshCredential(current); }
      else if (result.state === "failed") { this.authState = "failed"; this.authFailureCode = result.error ?? ""; this.authFailureAt = new Date().toISOString(); this.authError = this.signInError(this.authFailureCode); }
      else if (Date.now() >= Date.parse(device.expiresAt)) { this.authState = "failed"; this.authFailureCode = "device_authorization_expired"; this.authFailureAt = new Date().toISOString(); this.authError = this.signInError(this.authFailureCode); }
      else this.scheduleAuthCheck();
    } catch {
      if (!current() || this.authState !== "pending") return;
      this.authError = "Statusprüfung unterbrochen. Verbindung prüfen; diese Seite versucht es erneut.";
      if (Date.now() >= Date.parse(device.expiresAt)) { this.authState = "failed"; this.authFailureCode = "device_authorization_expired"; this.authFailureAt = new Date().toISOString(); this.authError = this.signInError(this.authFailureCode); } else this.scheduleAuthCheck();
    }
    if (current()) this.render();
  }
  private async startAuth() {
    if (!this.authorized || this.authBusy || this.authState === "pending" || !["missing", "valid"].includes(this.credential?.result ?? "") || !this.signInReady) return;
    const renewal = this.credential?.result === "valid";
    if (!renewal && (!validClientId(this.authClientId.trim()) || !validTenant(this.authTenant.trim()))) { this.authError = "Bitte eine gültige Microsoft App-ID und Tenant-ID oder -Domain eingeben."; this.render(); return; }
    const epoch = ++this.authEpoch;
    const current = () => this.authCurrent(epoch);
    this.authRenewal = renewal;
    this.authBusy = true; this.authError = ""; this.render();
    try {
      const started = await this.credentialCall<DeviceStart>(renewal ? "microsoft-graph.credentials.reauthorize-start" : "microsoft-graph.credentials.device-start", renewal ? {} : { clientId: this.authClientId.trim(), tenant: this.authTenant.trim() });
      if (!current()) return;
      if (!isMicrosoftDeviceVerificationUri(started.verificationUri)) throw new Error("device_authorization_failed");
      if (!renewal) saveAppIds(this.authClientId.trim(), this.authTenant.trim());
      this.device = started; this.authState = "pending"; this.authFailureCode = ""; this.scheduleAuthCheck();
    } catch (error) {
      if (!current()) return;
      this.authState = "failed"; this.authFailureCode = error instanceof Error ? error.message : ""; this.authFailureAt = new Date().toISOString(); this.authError = this.signInError(this.authFailureCode);
    } finally { if (current()) { this.authBusy = false; this.scheduleAuthCheck(); this.render(); } }
  }
  private async cancelAuth() {
    if (!this.authorized || this.authBusy || !this.device || this.authState !== "pending") return;
    // Cancel supersedes any status request already in flight.
    const epoch = ++this.authEpoch, device = this.device;
    const current = () => this.authCurrent(epoch, device.sessionId);
    if (this.authTimer) clearTimeout(this.authTimer);
    this.authTimer = undefined;
    this.authBusy = true; this.render();
    try {
      const result = await this.credentialCall<DeviceStatus>("microsoft-graph.credentials.device-cancel", { sessionId: device.sessionId });
      if (!current()) return;
      if (result.state === "created") { this.authState = "created"; this.authError = ""; await this.refreshCredential(current); }
      else { this.authState = "failed"; this.authFailureCode = result.error ?? "device_authorization_cancelled"; this.authFailureAt = new Date().toISOString(); this.authError = this.signInError(this.authFailureCode); }
    } catch { if (current()) this.authError = this.signInError(""); }
    finally { if (current()) { this.authBusy = false; this.scheduleAuthCheck(); this.render(); } }
  }
  private renderSignIn(main: HTMLElement) {
    if (this.credential?.result === "valid" && this.authState !== "pending" && (this.authState !== "failed" || this.authFailureCode === "device_authorization_cancelled")) {
      const missing = this.credential.missingScopes ?? [];
      const card = missing.length ? el("section", "mg-section mg-auth") : el("details", "mg-admin-details mg-reauthorize-help");
      card.id = "microsoft-connect";
      append(card, missing.length ? el("h2", "", "Microsoft-Zugang erneut autorisieren") : el("summary", "", "Microsoft-Berechtigungen erneuern"));
      if (missing.length) append(card, el("p", "mg-warning", format("Diese Berechtigungen fehlen: {scopes}. Bis zur Erneuerung bleiben die betreffenden Aktionen gesperrt.", { scopes: missing.join(", ") })));
      append(card, el("p", "mg-hint", "Die vorhandene App und der vorhandene Mandant werden sicher wiederverwendet. Du musst keine Kennungen eingeben und keine neue App anlegen."));
      append(card, el("p", "mg-hint", "Melde dich auf der Microsoft-Seite mit demselben Konto an und prüfe die angeforderten Berechtigungen. Der bestehende Zugang bleibt bis zum erfolgreichen Abschluss erhalten. Teile Code und Zugangsdaten nicht im Chat."));
      append(card, el("p", "mg-hint", "Bei verweigertem Zugriff zuerst Regeln, Kontotyp und Einwilligung prüfen. Eine Anmeldung ersetzt keine erforderliche Administratorzustimmung."));
      const start = this.action("Microsoft-Berechtigungen erneuern", () => { void this.startAuth(); }, "primary");
      start.disabled = this.authBusy || !this.signInReady; append(card, start);
      append(card, this.action("Berechtigungen erneut prüfen", () => { void this.refreshCredential().then(() => this.render()); }, "secondary"));
      if (!this.signInReady) append(card, el("p", "mg-warning", "Schließe zuerst die Einrichtung oben ab: Vault-Schlüssel, Agentenzugriff und angewendete Regeln."));
      if (this.authError) append(card, el("p", "mg-error", this.authError));
      append(main, card); return;
    }
    const card = el("section", "mg-section mg-auth"); card.id = "microsoft-connect";
    const result = this.credential?.result;
    if (result === "quarantined" || result === "unavailable" || !result) return;
    append(card, el("p", "mg-kicker", "MICROSOFT-KONTO VERBINDEN"));
    if (this.authState === "pending" && this.device) {
      append(card, el("h2", "", "Microsoft-Anmeldung abschließen"), el("p", "mg-hint", "Öffne die Microsoft-Seite, gib den Code ein und kehre hierher zurück."));
      const link = el("a", "mg-button primary", "Microsoft-Anmeldeseite öffnen"); link.href = this.device.verificationUri; link.target = "_blank"; link.rel = "noopener noreferrer"; append(card, link);
      append(card, el("p", "mg-hint", "Öffnet einen neuen Tab. Diese Seite bleibt offen."));
      const code = el("p", "mg-auth-code", this.device.userCode); code.setAttribute("aria-label", localize("Microsoft-Anmeldecode"));
      append(card, el("h3", "", "Code auf der Microsoft-Seite eingeben"), code, el("p", "mg-hint", "Nur auf der Microsoft-Seite eingeben; nicht im Chat teilen."));
      append(card, this.action("Code kopieren", () => { if (!navigator.clipboard?.writeText) { this.authError = "Kopieren nicht möglich. Markiere den Code und gib ihn bei Microsoft ein."; this.render(); return; } void navigator.clipboard.writeText(this.device!.userCode).catch(() => { this.authError = "Kopieren nicht möglich. Markiere den Code und gib ihn bei Microsoft ein."; this.render(); }); }));
      append(card, el("h3", "", "Hier auf das Ergebnis warten"), el("p", "mg-hint", "Status: Anmeldung ausstehend. Diese Seite erkennt den Abschluss, falls Microsoft ihn liefert. Der Code läuft nach spätestens 15 Minuten ab."));
      const cancel = this.action("Anmeldung abbrechen", () => { void this.cancelAuth(); }, "ghost"); cancel.disabled = this.authBusy; append(card, cancel);
    } else if (this.authState === "failed" && this.authFailureCode !== "device_authorization_cancelled") {
      append(card, el("h2", "", "Anmeldung nicht abgeschlossen"), el("p", "mg-error", "Wir konnten die Anmeldung nicht als abgeschlossen erkennen."));
      append(card, el("p", "", this.authError || "Wir kennen die Ursache noch nicht. Bitte deinen Admin um Hilfe."));
      append(card, el("p", "mg-hint", "Ein erneuter Versuch ohne Änderung hilft möglicherweise nicht."));
      append(card, el("p", "mg-hint", "Teile Zeitpunkt und technischen Fehlercode mit. Falls Microsoft einen weiteren Code zeigt, gib ihn zusätzlich an."));
      const details = el("details", "mg-admin-details"); append(details, el("summary", "", "Für den Admin"));
      const logLink = el("a", "", "Anmeldeprotokolle öffnen"); logLink.href = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/SignInLogsBlade"; logLink.target = "_blank"; logLink.rel = "noopener noreferrer";
      append(details, el("p", "mg-hint", "Einwilligung, Kontotyp und Richtlinie prüfen."), logLink); append(card, details);
      const copy = this.action("Fehlerangaben für Admin kopieren", () => {
        const info = [`Microsoft-Anmeldung: nicht abgeschlossen`, `Zeitpunkt (UTC): ${this.authFailureAt || new Date().toISOString()}`, `Fehlercode: ${this.authFailureCode || "nicht verfügbar"}`, ...(this.authRenewal ? [] : [`Organisation: ${this.authTenant || "nicht angegeben"}`])].join("\n");
        if (!navigator.clipboard?.writeText) { this.authError = "Kopieren nicht möglich. Teile Zeitpunkt und Fehlercode manuell."; this.render(); return; }
        void navigator.clipboard.writeText(info).catch(() => { this.authError = "Kopieren nicht möglich. Teile Zeitpunkt und Fehlercode manuell."; this.render(); });
      }); append(card, copy);
      append(card, this.action("Erneut anmelden", () => { this.authState = "idle"; this.authError = ""; void this.refreshCredential().then(() => this.render()); }, "ghost"));
    } else {
      const stored = savedAppIds();
      if (stored && !this.editAppIds) {
        append(card, el("h2", "", "Konto verbinden"), el("p", "mg-lead", "App-Kennungen sind in diesem Browser gespeichert. Anmeldung und Freigabe prüfen wir beim Verbinden."));
        append(card, el("p", "mg-status", "Microsoft-Konto · Anmeldung noch nicht geprüft"));
        append(card, el("p", "mg-hint", "Du musst nichts kopieren. Danach öffnest du den Microsoft-Link und gibst dort den angezeigten Code ein."));
        const start = this.action("Mit Microsoft verbinden", () => { void this.startAuth(); }, "primary"); start.disabled = this.authBusy || !this.signInReady; append(card, start);
        append(card, this.action("Andere App verwenden oder Einrichtung ändern", () => { this.editAppIds = true; this.render(); }, "ghost"));
      } else {
        append(card, el("h2", "", "Konto verbinden"), el("p", "mg-lead", "Bitte deinen Admin um Anwendungs-ID und Verzeichnis-ID. Ein Client Secret, Kennwort oder Einmalcode gehört nicht hierher."));
        const client = this.field(card, "Anwendungs-ID (Client-ID)", this.authClientId, value => { this.authClientId = value; }); client.placeholder = "00000000-0000-0000-0000-000000000000";
        const tenant = this.field(card, "Verzeichnis-ID (Mandant) oder Domain", this.authTenant, value => { this.authTenant = value; }); tenant.placeholder = "example.onmicrosoft.com";
        const start = this.action("Anmeldung starten", () => { this.authClientId = client.value.trim(); this.authTenant = tenant.value.trim(); void this.startAuth(); }, "primary"); start.disabled = this.authBusy || !this.signInReady; append(card, start);
        const admin = el("details", "mg-admin-details");
        append(admin, el("summary", "", "Hinweis für Admins · App einrichten"));
        append(admin, el("p", "mg-hint", "Microsoft Entra Admin Center → Entra ID → App-Registrierungen → vorhandene App → Übersicht. „Anwendungs-ID (Client-ID)“ ist die App-ID; „Verzeichnis-ID (Mandant)“ ist die Tenant-ID. Nicht die Objekt-ID verwenden. Bei mehreren Verzeichnissen zuerst zum richtigen Mandanten wechseln."));
        append(admin, el("p", "mg-hint", "In der App öffentliche Clientanmeldung und benötigte delegierte Berechtigungen prüfen; falls nötig Admin-Einwilligung erteilen. Diese Schritte werden hier nicht automatisch geprüft."));
        const entra = el("a", "mg-button secondary", "Einrichtung für Admin öffnen"); entra.href = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade"; entra.target = "_blank"; entra.rel = "noopener noreferrer"; append(admin, entra); append(card, admin);
      }
      if (!this.signInReady) append(card, el("p", "mg-warning", "Schließe zuerst die Einrichtung oben ab: Vault-Schlüssel, Agentenzugriff und angewendete Regeln."));
    }
    if (this.authError && (this.authState !== "failed" || this.authFailureCode === "device_authorization_cancelled")) append(card, el("p", "mg-error", this.authError));
    append(main, card);
  }
  private renderAgentPicker(body: HTMLElement): string {
    const roster = this.host.agents.rows;
    const pick = el("label", "mg-field"); append(pick, el("span", "mg-label", "Agent auswählen"));
    const select = el("select");
    for (const agent of roster) { const option = el("option", "", agent.name ?? agent.id); option.value = agent.id; option.selected = agent.id === this.selectedAgent; append(select, option); }
    select.addEventListener("change", () => { this.selectedAgent = select.value; this.render(); }); append(pick, select); append(body, pick);
    return this.selectedAgent || roster[0]?.id || "";
  }
  private async addFolder() {
    if (this.locked) return;
    const path = this.newFolderPath.trim();
    if (!path.startsWith("/")) { this.error = "Bitte einen Ordnerpfad ab / eingeben."; this.render(); return; }
    this.busy = true; this.error = ""; this.render();
    try {
      const resolved = await this.host.request<{ path: string; drive_id: string; item_id: string }>("microsoft-graph.configuration.resolveFolder", { path });
      const roots = this.policy.services.onedrive.allowed_roots;
      if (roots.some(root => root.drive_id === resolved.drive_id && root.item_id === resolved.item_id)) { this.error = "Dieser Ordner ist bereits freigegeben. Wähle in der Ordnerkachel die Rechte für den Agenten."; return; }
      let label = "folder_" + (roots.length + 1); let suffix = roots.length + 1;
      while (roots.some(root => root.label === label)) label = "folder_" + ++suffix;
      roots.push({ label, path: resolved.path, drive_id: resolved.drive_id, item_id: resolved.item_id, include_descendants: true, permissions: { read: false, write: false, delete: false }, agents: {} });
      this.newFolderPath = ""; this.success = "Ordner geprüft. Wähle jetzt die Rechte für den Agenten.";
    } catch { this.error = "Ordner nicht gefunden oder nicht prüfbar. Prüfe den Pfad und die OneDrive-Verbindung."; }
    finally { this.busy = false; this.render(); }
  }
  private renderOneDrive(body: HTMLElement) {
    const roots = this.policy.services.onedrive.allowed_roots;
    const assignments = roots.reduce((total, root) => total + Object.values(root.agents).filter(grant => Object.values(grant.permissions).some(Boolean)).length, 0);
    const heading = el("div", "mg-onedrive-heading");
    const title = el("div");
    append(title, el("p", "mg-kicker", "OneDrive-Zugriff"), el("h2", "", "Wer darf auf welche OneDrive-Bereiche zugreifen?"), el("p", "mg-hint", "Klicke auf ein Recht, um es zu ändern. Ohne aktives Recht wird der Agent entfernt. Änderungen werden erst nach Prüfen und Speichern wirksam."));
    append(heading, title, el("span", "mg-count", format("{count} Zuweisungen", { count: assignments })));
    append(body, heading);
    const add = el("section", "mg-add-folder");
    append(add, el("h3", "", "Ordner hinzufügen"), el("p", "mg-hint", "Pfad in deinem OneDrive. Der Ordner wird vor dem Hinzufügen geprüft."));
    const input = this.field(add, "Ordnerpfad", this.newFolderPath, value => { this.newFolderPath = value; });
    input.placeholder = "/Projects/Clients";
    input.addEventListener("keydown", event => { if (event.key === "Enter" && !this.locked) { event.preventDefault(); this.newFolderPath = input.value.trim(); void this.addFolder(); } });
    const addButton = this.action("Ordner prüfen und hinzufügen", () => { this.newFolderPath = input.value.trim(); void this.addFolder(); }, "primary"); addButton.disabled = this.busy;
    append(add, addButton); append(body, add);
    if (!roots.length) { append(body, el("p", "mg-message", "Noch kein OneDrive-Ordner freigegeben.")); return; }
    const grid = el("div", "mg-root-grid");
    for (const [index, root] of roots.entries()) {
      const card = el("article", "mg-root-card");
      const header = el("div", "mg-root-header");
      const name = root.path.split("/").filter(Boolean).at(-1) || root.path;
      const identity = el("div", "mg-root-identity"); append(identity, el("h3", "", name), el("p", "mg-root-path", root.path));
      append(header, identity, el("span", "mg-descendants", "Inkl. Unterordner")); append(card, header);
      this.renderInstructionsControl(card, root, index);
      const agents = Object.entries(root.agents).filter(([, grant]) => Object.values(grant.permissions).some(Boolean));
      if (!agents.length) append(card, el("p", "mg-empty-grants", "Noch kein Agent berechtigt."));
      for (const [agentId, grant] of agents) {
        const row = el("div", "mg-agent-row");
        const agentName = this.host.agents.rows.find(agent => agent.id === agentId)?.name || agentId;
        const agentLabel = el("strong", "mg-agent-name", agentName);
        const rights = el("div", "mg-rights");
        for (const [op, label] of [["read", "Lesen"], ["write", "Schreiben"], ["delete", "Löschen"]] as const) {
          const enabled = grant.permissions[op] === true;
          const badge = this.action(`${enabled ? "✓" : "–"} ${localize(label)}`, () => {
            grant.permissions[op] = !enabled;
            if (!Object.values(grant.permissions).some(Boolean)) delete root.agents[agentId];
            for (const right of ["read", "write", "delete"] as const)
              root.permissions[right] = Object.values(root.agents).some(agent => agent.permissions[right] === true);
            this.render();
            const replacement = [...this.container.querySelectorAll<HTMLButtonElement>(".mg-right")]
              .find(control => control.dataset.root === root.label && control.dataset.agent === agentId && control.dataset.right === op);
            (replacement ?? [...this.container.querySelectorAll<HTMLButtonElement>(".mg-add-agent")]
              .find(control => control.dataset.root === root.label))?.focus();
          }, `mg-right ${enabled ? "is-allowed" : "is-denied"}`);
          badge.setAttribute("aria-label", format("{right} für {agent} auf {path}", { right: localize(label), agent: agentName, path: root.path }));
          badge.setAttribute("aria-pressed", String(enabled));
          badge.dataset.root = root.label; badge.dataset.agent = agentId; badge.dataset.right = op;
          append(rights, badge);
        }
        append(row, agentLabel, rights); append(card, row);
      }
      if (this.editingAccess?.rootLabel === root.label) this.renderAccessEditor(card, root);
      else {
        const available = this.host.agents.rows.filter(agent => !agents.some(([agentId]) => agentId === agent.id));
        if (available.length) {
          const addAgent = this.action("+ Agent hinzufügen", () => { this.editingAccess = { rootLabel: root.label, agentId: available[0]!.id, permissions: { read: false, write: false, delete: false } }; this.render(); }, "ghost mg-add-agent");
          addAgent.dataset.root = root.label; append(card, addAgent);
        }
        else if (!this.host.agents.rows.length) append(card, el("p", "mg-hint", "Keine Agenten gefunden."));
      }
      const actions = el("div", "mg-root-actions");
      append(actions, this.action("Ordner für alle Agenten entfernen", () => { if (window.confirm(format("Den Ordner {path} für alle Agenten entfernen?", { path: root.path }))) { roots.splice(index, 1); if (this.editingAccess?.rootLabel === root.label) this.editingAccess = undefined; this.render(); } }, "danger"));
      append(card, actions); append(grid, card);
    }
    append(body, grid);
  }
  private renderInstructionsControl(card: HTMLElement, root: Root, index: number) {
    const row = el("div", "mg-instructions-control");
    const input = this.checkbox(row, "AGENTS.md als Arbeitsanweisungen verwenden", root.agents_instructions === "trusted", checked => {
      if (checked) root.agents_instructions = "trusted";
      else delete root.agents_instructions;
      this.render();
    });
    input.dataset.focusKey = JSON.stringify(["instructions", "checkbox", root.label]);
    const panel = el("div", "mg-instructions-help");
    panel.id = `${this.helpIdPrefix}-${index}`;
    panel.dataset.root = root.label;
    panel.setAttribute("popover", "auto");
    panel.setAttribute("role", "note");
    panel.setAttribute("aria-label", localize("Informationen zu AGENTS.md-Arbeitsanweisungen"));
    append(panel,
      el("p", "", "Vor OneDrive-Zugriffen erhalten berechtigte Agenten die passenden AGENTS.md-Arbeitsanweisungen und müssen deren Empfang bestätigen. Aktiviere dies nur, wenn du den Bearbeitern dieser Dateien vertraust – auch zukünftige Änderungen können die Agenten anleiten."),
      el("p", "", "Zusätzliche Zugriffsrechte entstehen nicht; Sicherheitsregeln bleiben vorrangig. Die Bestätigung garantiert keine vollständige Befolgung."),
      el("p", "", "Voraussetzung: Im freigegebenen Hauptordner liegt eine AGENTS.md. Änderungen an Dateien können wegen des Caches bis zu einer Stunde verzögert erkannt werden."));
    // Declarative invocation lets the browser coordinate touch light-dismiss
    // with same-button toggling, Escape and focus return.
    const info = this.action("ⓘ", () => {}, "ghost mg-instructions-info");
    info.setAttribute("popovertarget", panel.id);
    info.dataset.focusKey = JSON.stringify(["instructions", "help", root.label]);
    info.setAttribute("aria-label", localize("Informationen zu AGENTS.md-Arbeitsanweisungen"));
    info.setAttribute("aria-controls", panel.id);
    info.setAttribute("aria-expanded", "false");
    panel.addEventListener("beforetoggle", event => {
      const open = (event as ToggleEvent).newState === "open";
      info.setAttribute("aria-expanded", String(open));
      panel.dataset.open = String(open);
      if (open) { this.instructionsHelpRoot = root.label; requestAnimationFrame(this.positionInstructionsHelp); }
      else if (panel.isConnected && this.instructionsHelpRoot === root.label) this.instructionsHelpRoot = undefined;
    });
    append(row, info, panel); append(card, row);
  }
  private renderAccessEditor(card: HTMLElement, root: Root) {
    const editing = this.editingAccess!;
    const panel = el("div", "mg-access-editor");
    append(panel, el("h4", "", "Agent hinzufügen"));
    const label = el("label", "mg-field"); append(label, el("span", "mg-label", "Agent"));
    const select = el("select");
    for (const agent of this.host.agents.rows.filter(agent => !Object.values(root.agents[agent.id]?.permissions ?? {}).some(Boolean))) {
      const option = el("option", "", agent.name || agent.id); option.value = agent.id; option.selected = agent.id === editing.agentId; append(select, option);
    }
    select.addEventListener("change", () => { editing.agentId = select.value; }); append(label, select); append(panel, label);
    const rights = el("div", "mg-editor-rights");
    for (const [op, label] of [["read", "Lesen"], ["write", "Schreiben"], ["delete", "Löschen"]] as const)
      this.checkbox(rights, label, editing.permissions[op], value => { editing.permissions[op] = value; });
    append(panel, rights, el("p", "mg-hint", "Ohne ausgewähltes Recht wird der Agent aus diesem Bereich entfernt."));
    const actions = el("div", "mg-editor-actions");
    append(actions, this.action("Abbrechen", () => { this.editingAccess = undefined; this.render(); }, "ghost"));
    append(actions, this.action("In Entwurf übernehmen", () => {
      if (Object.values(editing.permissions).some(Boolean)) root.agents[editing.agentId] = { permissions: { ...editing.permissions } };
      else delete root.agents[editing.agentId];
      for (const op of ["read", "write", "delete"] as const) root.permissions[op] = Object.values(root.agents).some(agent => agent.permissions[op] === true);
      this.editingAccess = undefined; this.render();
    }, "primary"));
    append(panel, actions); append(card, panel);
  }
  private renderServices(body: HTMLElement) {
    append(body, el("h2", "", "Microsoft-Dienste"), el("p", "", "Wähle, welche Dienste der Agent nutzen darf. Ein neuer Zugang umfasst alle Funktionen des Dienstes. Bestehende eingeschränkte Zugänge bleiben unverändert."));
    const agentId = this.renderAgentPicker(body);
    if (!agentId) { append(body, el("p", "mg-message", "Keine Agenten gefunden.")); return; }
    for (const service of ["calendar", "mail", "contacts", "todo"] as const) {
      const card = el("fieldset", "mg-section"); append(card, el("legend", "", serviceNames[service]));
      const grants = this.policy.services[service]?.agents ?? {}; const existing = grants[agentId];
      this.checkbox(card, format("{service} nutzen", { service: localize(serviceNames[service]) }), !!existing, checked => {
        const key = `${service}:${agentId}`;
        if (checked) {
          if (service === "contacts") this.policy.services.contacts ??= { agents: {} };
          this.policy.services[service]!.agents[agentId] = existing ?? this.removedServiceGrants[key] ?? { operations: [...operations[service]] };
        }
        else { if (existing) this.removedServiceGrants[key] = clone(existing); delete grants[agentId]; }
        this.render();
      });
      if (existing && existing.operations.length < operations[service].length) append(card, el("p", "mg-hint", "Bestehender Zugang ist auf einzelne Funktionen eingeschränkt. Diese Einschränkung bleibt erhalten."));
      if (service === "calendar") this.renderCalendarAccess(card, agentId, existing);
      append(body, card);
    }
  }
  private calendarChoices(): Array<{ id: string; name: string }> {
    return [{ id: "me", name: "Standardkalender" }, ...(this.calendarCatalog?.calendars ?? [])
      .filter(calendar => calendar.id !== this.calendarCatalog?.defaultCalendarId && calendar.id !== "me")];
  }
  private calendarName(id: string): string {
    return this.calendarChoices().find(calendar => calendar.id === id)?.name ?? (id === this.calendarCatalog?.defaultCalendarId ? localize("Standardkalender (direkte Kalender-ID)") : format("Kalender-ID: {id}", { id }));
  }
  private renderCalendarAccess(card: HTMLElement, agentId: string, grant?: Grant) {
    if (!grant) { append(card, el("p", "mg-hint", "Gesperrt · kein Kalender freigegeben.")); return; }
    const selected = grant.resources ?? ["me"];
    const choices = this.calendarChoices();
    const defaultId = this.calendarCatalog?.defaultCalendarId;
    const known = new Set([...choices.map(choice => choice.id), ...(defaultId ? [defaultId] : [])]);
    const unknown = selected.filter(resource => !known.has(resource));
    const summary = selected.length === 1 && selected[0] === "me" ? localize("Eingeschränkt · nur Standardkalender")
      : defaultId && selected.includes(defaultId) ? format("{count} Kalender-Ressourcen freigegeben (inkl. Standardkalender-ID)", { count: selected.length })
      : this.calendarCatalogState === "ready" && !unknown.length ? format("{selected} von {total} Kalendern freigegeben", { selected: selected.length, total: choices.length })
      : format("{count} Kalender freigegeben · Liste nicht vollständig prüfbar", { count: selected.length });
    append(card, el("p", "mg-calendar-summary", summary));
    if (this.calendarCatalogState === "loading") append(card, el("p", "mg-hint", "Verfügbare Kalender werden geladen …"));
    if (this.calendarCatalogState === "error") {
      append(card, el("p", "mg-warning", "Kalenderliste nicht ladbar. Bestehende Freigaben bleiben erhalten; weitere Kalender können erst nach erneutem Laden ausgewählt werden."));
      append(card, this.action("Kalenderliste erneut laden", () => { void this.loadCalendars(); }));
    }
    const list = el("div", "mg-calendar-list");
    const defaultIdChoice = defaultId && selected.includes(defaultId) ? [{ id: defaultId, name: localize("Standardkalender (direkte Kalender-ID)") }] : [];
    for (const choice of [...choices, ...defaultIdChoice, ...unknown.map(id => ({ id, name: this.calendarName(id) }))]) {
      const label = unknown.includes(choice.id) ? format("{name} · nicht verifizierbar", { name: choice.name }) : choice.name;
      this.checkbox(list, label, selected.includes(choice.id), checked => {
        const next = checked ? [...new Set([...selected, choice.id])] : selected.filter(id => id !== choice.id);
        if (next.length) grant.resources = next;
        else { delete this.policy.services.calendar.agents[agentId]; this.removedServiceGrants[`calendar:${agentId}`] = clone(grant); }
        this.confirmedRemoval = false; this.render();
      });
    }
    append(card, list, el("p", "mg-hint", "Die Auswahl gilt nur für diesen Agenten. Funktionsrechte bleiben unverändert; Änderungen werden erst nach Prüfen und Speichern aktiv."));
  }
  private renderApprovals(body: HTMLElement) {
    append(body, el("h2", "", "Wann ist eine Freigabe nötig?"), el("p", "", "Eine Freigabe ist eine Rückfrage vor der konkreten Aktion. Die Agentenrechte aus den ersten beiden Schritten gelten immer zusätzlich."));
    for (const service of ["onedrive", "calendar", "mail", "contacts", "todo"] as const) {
      const card = el("fieldset", "mg-section"); append(card, el("legend", "", serviceNames[service]));
      append(card, el("p", "", "Unkritisch · Lesen und Suchen: keine Rückfrage."));
      const configured = this.policy.rules.warningApprovalsByService ?? {};
      const baseline = this.snapshot?.config?.plugins?.entries?.[id]?.config?.warningApprovalsRequired !== false;
      this.checkbox(card, "Warn-Aktionen · Erstellen oder Ändern: vorher fragen", configured[service] ?? baseline, checked => {
        this.policy.rules.warningApprovalsByService = { ...configured, [service]: checked }; this.render();
      });
      const critical = service === "onedrive" ? "Löschen" : service === "calendar" ? "Löschen oder auf Termine antworten" : service === "mail" ? "Löschen oder E-Mail senden" : "Löschen";
      append(card, el("p", "", format("Kritisch · {action}: immer einzeln fragen.", { action: localize(critical) })));
      append(body, card);
    }
  }
  private renderReview(body: HTMLElement) {
    append(body, el("h2", "", "Änderungen prüfen"), el("p", "", "Prüfe die Zugriffe und Rückfragen. Benötigte OneDrive-Tools werden für berechtigte Agenten zusammen mit den Microsoft-Graph-Regeln freigeschaltet."));
    try {
      const exposure = planOneDriveToolExposure(this.policy.services.onedrive.allowed_roots, this.snapshot?.config?.agents?.entries, this.snapshot?.config?.tools?.deny);
      for (const [agentId, entry] of Object.entries(exposure.expected)) append(body, el("p", "mg-status", format("{agent}: zusätzliche OneDrive-Tools · {tools}", { agent: agentId, tools: entry.value.filter(name => !this.snapshot?.config?.agents?.entries?.[agentId]?.tools?.[entry.key]?.includes(name)).join(", ") })));
    } catch { append(body, el("p", "mg-warning", "OneDrive-Tool-Freigabe ist mit der aktuellen Agent-Policy nicht vereinbar. Speichern wird verhindert.")); }
    append(body, el("p", "mg-status", policySummary(this.policy).join(" · ")));
    for (const agentId of new Set([...Object.keys(this.initial?.services.calendar.agents ?? {}), ...Object.keys(this.policy.services.calendar.agents)])) {
      const before = this.initial?.services.calendar.agents[agentId]?.resources ?? (this.initial?.services.calendar.agents[agentId] ? ["me"] : []);
      const after = this.policy.services.calendar.agents[agentId]?.resources ?? (this.policy.services.calendar.agents[agentId] ? ["me"] : []);
      const added = after.filter(id => !before.includes(id)); const removed = before.filter(id => !after.includes(id));
      const name = this.host.agents.rows.find(agent => agent.id === agentId)?.name || agentId;
      if (added.length) append(body, el("p", "mg-calendar-review", format("{agent}: freigegeben · {calendars}", { agent: name, calendars: added.map(id => this.calendarName(id)).join(", ") })));
      if (removed.length) append(body, el("p", "mg-calendar-review", format("{agent}: entzogen · {calendars}", { agent: name, calendars: removed.map(id => this.calendarName(id)).join(", ") })));
    }

    const newScopes = this.scopes.filter(scope => !this.initialScopes.includes(scope));
    if (this.pluginEnabled && newScopes.length) append(body, el("p", "mg-warning", format("Neue Microsoft-Berechtigungen ({scopes}) benötigen nach dem Speichern eine Reautorisierung. Bis dahin bleiben diese Aktionen gesperrt.", { scopes: newScopes.join(", ") })));
    const approvals = this.policy.rules.warningApprovalsByService ?? {};
    for (const service of ["onedrive", "calendar", "mail", "contacts", "todo"] as const) if (approvals[service] === false) append(body, el("p", "mg-warning", format("{service}: Warn-Aktionen dürfen ohne Rückfrage ausgeführt werden. Kritische Aktionen benötigen weiterhin eine Freigabe.", { service: localize(serviceNames[service]) })));
    if (!this.dirty) { append(body, el("p", "", "Keine Änderungen zum Speichern.")); return; }
    if (JSON.stringify(this.initial) !== JSON.stringify(this.policy)) this.checkbox(body, "Ich habe die Zugriffsänderungen geprüft, auch entzogene Rechte", this.confirmedRemoval, v => { this.confirmedRemoval = v; });

  }
  private async validate(candidate = this.policy): Promise<boolean> {
    const epoch = ++this.validationEpoch;
    this.error = "";
    const roster = new Set(this.host.agents.rows.map(a => a.id));
    for (const service of ["calendar", "mail", "contacts", "todo"] as const) for (const agent of Object.keys(candidate.services[service]?.agents ?? {})) if (!roster.has(agent)) { this.error = format("Unknown configured agent: {agent}", { agent }); this.render(); return false; }
    for (const root of candidate.services.onedrive.allowed_roots) for (const agent of Object.keys(root.agents)) if (!roster.has(agent)) { this.error = format("Unknown configured agent: {agent}", { agent }); this.render(); return false; }
    try { const result = await this.read<{ valid: boolean; requiredScopes: string[]; policyHash?: string }>("microsoft-graph.configuration.validate", { policy: candidate }); if (epoch !== this.validationEpoch || this.disposed) return false; this.scopes = result.requiredScopes; if (this.operation && candidate === this.operation.target) this.operation.targetHash = result.policyHash; this.render(); return result.valid; }
    catch { if (epoch !== this.validationEpoch || this.disposed) return false; this.error = "Policy validation failed. Check root labels, paths, IDs, and grant resources."; this.render(); return false; }
  }
  private async save() {
    if (!this.dirty || this.locked || !this.snapshot || !this.initial || this.authBusy || this.authState === "pending") return;
    if (!this.confirmedRemoval) { this.error = "Bitte die Zugriffsänderungen vor dem Speichern bestätigen."; this.render(); return; }
    this.stopStatusChecks(); this.statusEpoch++; this.statusError = "";
    const operation: SaveOperation = { target: clone(this.policy), baseline: clone(this.initial), phase: "validating", issued: false, settled: false, epoch: this.statusEpoch };
    this.operation = operation; operationsByHost.set(this.host, operation);
    this.busy = true; this.error = ""; this.success = ""; this.render();
    try {
      if (!(await this.validate(operation.target))) { operation.phase = "failed"; return; }
      if (this.disposed || this.signal.aborted || !this.authorized) { operation.phase = "failed"; return; }
      const fresh = await this.read<ConfigSnapshot>("config.get");
      if (this.disposed || this.signal.aborted || !this.authorized) { operation.phase = "failed"; return; }
      if (canonicalPolicy(fresh.config?.plugins?.entries?.[id]?.config?.policy) !== canonicalPolicy(operation.baseline)) { this.error = "Configuration changed since this draft loaded. Reload and reapply your changes."; operation.phase = "conflict"; return; }
      if (authoredPolicySource(fresh) !== this.includeName) { this.error = "Die Konfigurationsquelle wurde geändert. Bitte neu laden."; operation.phase = "conflict"; return; }
      const exposure = planOneDriveToolExposure(operation.target.services.onedrive.allowed_roots, fresh.config?.agents?.entries, fresh.config?.tools?.deny);
      operation.toolExposure = exposure.expected;
      const replacements: string[] = [];
      const patchPolicy = collectDiff(operation.baseline, operation.target, "plugins.entries.microsoft-graph.config.policy", replacements);
      for (const [agentId, entry] of Object.entries(exposure.expected)) {
        const before = fresh.config?.agents?.entries?.[agentId]?.tools?.[entry.key];
        if (Array.isArray(before)) replacements.push(`agents.entries.${agentId}.tools.${entry.key}`);
      }
      const raw = JSON.stringify({ plugins: { entries: { [id]: { config: { policy: patchPolicy } } } }, ...(Object.keys(exposure.patch).length ? { agents: { entries: exposure.patch } } : {}) });
      operation.phase = "saving"; operation.issued = true; this.render();
      // Observe persistence independently: Core may drain admitted work before
      // replying to config.patch. Timeout is never evidence that no write occurred.
      this.statusChecksRemaining = 12; this.scheduleStatusCheck();
      const patch = this.host.request("config.patch", { raw, baseHash: fresh.hash, replacePaths: replacements, note: "Microsoft Graph configuration UI save" });
      // An RPC response is not the commit oracle. Release the UI immediately;
      // keep edits locked until config.get proves the target or a rejection is
      // reconciled. Never resend an ambiguous write.
      this.busy = false;
      void patch.then(() => { operation.settled = true; }, error => { operation.settled = true; operation.rejectedBeforeWrite = isPreCommitRejection(error); }).then(async () => {
        if (this.disposed || this.operation !== operation) return;
        await this.checkApplication();
      });
      // config.patch already requests the host's config activation. A second
      // automatic plugins.reload can contend with that drain and time out.
      // Leave an exact-target, manually initiated retry available if needed.
      this.render();
    } catch (error) {
      if (!operation.issued) { operation.phase = "failed"; this.error = error instanceof Error && /^(missing_agent_entry|unsupported_agent_id|conflicting_tool_policy|denied_tool):/.test(error.message) ? "OneDrive-Tool-Freigabe kann wegen der Agent-Tool-Policy nicht sicher gespeichert werden. Prüfe Agent-Eintrag und Deny-Regeln." : "Policy validation failed. Check root labels, paths, IDs, and grant resources."; }
      else { operation.phase = "unknown"; this.busy = false; await this.checkApplication(); }
    } finally {
      this.busy = false; this.render();
    }
  }

}

// Bundle the exact scoped stylesheet into the native entry: activation must not
// wait for a second plugin asset request. Share it across overlapping lifetimes.
let activeStyles: { node: HTMLStyleElement; users: number } | undefined;
function retainStyles(): () => void {
  if (!activeStyles) {
    const node = document.createElement("style");
    node.dataset.microsoftGraph = "styles";
    node.textContent = styles;
    document.head.append(node);
    activeStyles = { node, users: 0 };
  }
  const retained = activeStyles;
  retained.users++;
  return () => { if (--retained.users === 0) { retained.node.remove(); activeStyles = undefined; } };
}

export default defineControlUiPlugin({ id, activate(host) {
  setLocale(host.locale);
  const pageId = "configure";
  const disposers = [retainStyles()];
  const dispose = () => {
    let failure: unknown;
    for (const stop of disposers.splice(0).reverse()) { try { stop(); } catch (error) { failure ??= error; } }
    if (failure) throw failure;
  };
  try {
    disposers.push(host.ui.registerPage({ id: pageId, label: "Connect Microsoft 365 to OpenClaw", mount(container, context) { const view = new ConfigurationPage(container, context.host, context.signal); return { dispose: () => view.dispose() }; } }));
    disposers.push(host.ui.registerNavigation({ id: "configure", label: "Connect Microsoft 365 to OpenClaw", page: { id: pageId }, icon: "settings", order: 80 }));
    return dispose;
  } catch (error) { dispose(); throw error; }
} });
