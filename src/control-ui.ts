import { defineControlUiPlugin, type ControlUiHost } from "openclaw/plugin-sdk/control-ui";
import "./control-ui.css";
import { localize, format, setLocale, isRtl } from "./control-ui-i18n.js";
import { isMicrosoftDeviceVerificationUri } from "./device-verification.js";
import { isConnectionEstablished } from "./control-ui-status.js";

type Grant = { operations: string[]; resources?: string[] };
type Root = { label: string; path: string; drive_id: string; item_id: string; include_descendants: true; agents_instructions?: "trusted"; permissions: Record<"read" | "write" | "delete", boolean>; agents: Record<string, { permissions: Partial<Record<"read" | "write" | "delete", boolean>> }> };
type Policy = { version: 2; rules: { default: "deny"; warningApprovalsByService?: Partial<Record<"onedrive" | "calendar" | "mail" | "todo", boolean>> }; services: { onedrive: { allowed_roots: Root[] }; calendar: { agents: Record<string, Grant> }; mail: { agents: Record<string, Grant> }; todo: { agents: Record<string, Grant> } } };
type CredentialStatus = { policyVersion: 2; credential: { result: "missing" | "valid" | "quarantined" | "unavailable" } };
type DeviceStart = { sessionId: string; userCode: string; verificationUri: string; expiresAt: string; scopes: string[] };
type DeviceStatus = { state: "pending" | "created" | "failed"; error?: string; scopes?: string[] };
type CredentialReply<T> = { ok: true; value: T } | { ok: false; error: string };
type ConfigSnapshot = { hash: string; config: { plugins?: { entries?: Record<string, { enabled?: boolean; config?: Record<string, unknown> }> } }; parsed?: { plugins?: { entries?: Record<string, { enabled?: boolean; config?: Record<string, unknown> }> } }; configRevisionHash?: string; appliedConfigHash?: string };
const id = "microsoft-graph";
const operations = { calendar: ["read", "create", "update", "respond", "attach", "delete"], mail: ["read", "draft", "update", "move", "mark", "send", "delete"], todo: ["read", "create", "update", "delete"] } as const;
const serviceNames = { onedrive: "OneDrive", calendar: "Kalender", mail: "E-Mail", todo: "To Do" } as const;
const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v)) as T;
const blankPolicy = (): Policy => ({ version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: {} }, mail: { agents: {} }, todo: { agents: {} } } });
function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = "", value?: string): HTMLElementTagNameMap[K] { const node = document.createElement(tag); if (className) node.className = className; if (value !== undefined) node.textContent = localize(value); return node; }
function append(parent: HTMLElement, ...children: HTMLElement[]) { parent.append(...children); }
function field(parent: HTMLElement, label: string, value: string, set: (value: string) => void, hint?: string) {
  const wrap = el("label", "mg-field"); const title = el("span", "mg-label", label); const input = el("input"); input.value = value; input.autocomplete = "off"; input.addEventListener("change", () => set(input.value.trim())); append(wrap, title, input); if (hint) append(wrap, el("small", "mg-hint", hint)); append(parent, wrap); return input;
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
    ...(["calendar", "mail", "todo"] as const).map(s => format("{count} Agenten für {service}", { count: Object.keys(policy.services[s].agents).length, service: localize(serviceNames[s]) })),
  ];
}
class ConfigurationPage {
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
  private editingAccess?: { rootLabel: string; agentId: string; isNew: boolean; permissions: Record<"read" | "write" | "delete", boolean> };
  private removedServiceGrants: Record<string, Grant> = {};
  private busy = false;
  private error = "";
  private success = "";
  private statusError = "";
  private confirmedRemoval = false;
  private disposed = false;
  private statusTimer?: ReturnType<typeof setTimeout>;
  private statusChecksRemaining = 0;
  private authTimer?: ReturnType<typeof setTimeout>;
  private credential?: CredentialStatus["credential"];
  private authClientId = "";
  private authTenant = "";
  private device?: DeviceStart;
  private authState: "idle" | "pending" | "created" | "failed" = "idle";
  private authError = "";
  private authBusy = false;
  private readonly unsubscribe: () => void;
  constructor(private container: HTMLElement, private host: ControlUiHost, private signal: AbortSignal) {
    this.unsubscribe = host.subscribe(() => { setLocale(host.locale); if (!this.snapshot && !this.busy && this.authorized) void this.load(); else this.render(); });
    void this.load();
    this.render();
  }
  dispose() { this.disposed = true; if (this.authTimer) clearTimeout(this.authTimer); this.stopStatusChecks(); this.unsubscribe(); this.container.replaceChildren(); }
  private get dirty() { return !!this.initial && JSON.stringify(this.initial) !== JSON.stringify(this.policy); }
  private get authorized() { return this.host.connection.connected && this.host.connection.canAdmin; }
  private get applicationStatus(): "applied" | "pending" | "unknown" {
    const { configRevisionHash, appliedConfigHash } = this.snapshot ?? {};
    if (!configRevisionHash || !appliedConfigHash) return "unknown";
    return configRevisionHash === appliedConfigHash ? "applied" : "pending";
  }
  private stopStatusChecks() { if (this.statusTimer) clearTimeout(this.statusTimer); this.statusTimer = undefined; this.statusChecksRemaining = 0; }
  private scheduleStatusCheck() {
    if (this.disposed || this.signal.aborted || this.applicationStatus !== "pending" || this.statusChecksRemaining <= 0 || this.statusTimer) return;
    this.statusTimer = setTimeout(() => { this.statusTimer = undefined; void this.checkApplication(); }, 5000);
  }
  private async checkApplication() {
    if (!this.authorized || !this.snapshot || !this.initial || this.disposed || this.signal.aborted) return;
    if (this.busy) { this.scheduleStatusCheck(); return; }
    try {
      const snap = await this.host.request<ConfigSnapshot>("config.get", {});
      if (this.disposed || this.signal.aborted) return;
      if (JSON.stringify(snap.config?.plugins?.entries?.[id]?.config?.policy) !== JSON.stringify(this.initial)) {
        this.stopStatusChecks();
        this.statusError = "Die Regeln wurden außerhalb dieser Seite geändert. Bitte neu laden, um den aktuellen Stand zu sehen.";
        this.render();
        return;
      }
      this.snapshot = snap;
      this.statusError = "";
      if (this.applicationStatus !== "pending") this.stopStatusChecks();
      else { if (this.statusChecksRemaining === 0) this.statusChecksRemaining = 12; this.statusChecksRemaining--; this.scheduleStatusCheck(); }
    } catch { this.stopStatusChecks(); this.statusError = "Statusprüfung fehlgeschlagen. Verbindung prüfen und erneut versuchen."; }
    this.render();
  }
  private watchApplication() { this.stopStatusChecks(); if (this.applicationStatus === "pending") { this.statusChecksRemaining = 12; this.scheduleStatusCheck(); } }
  private async load() {
    if (!this.authorized || this.busy) return;
    this.busy = true; this.error = ""; this.render();
    try {
      const snap = await this.host.request<ConfigSnapshot>("config.get", {});
      if (this.disposed || this.signal.aborted) return;
      const pluginEntry = snap.config?.plugins?.entries?.[id];
      const entry = pluginEntry?.config ?? {};
      this.pluginEnabled = pluginEntry?.enabled === true && entry.enabled !== false;
      const authored = snap.parsed?.plugins?.entries?.[id]?.config?.policy;
      this.included = !!authored && typeof authored === "object" && "$include" in authored;
      this.includeName = this.included ? String((authored as { $include: unknown }).$include) : "";
      const policy = entry.policy as Policy | undefined;
      this.policy = policy ? clone(policy) : blankPolicy(); this.initial = clone(this.policy);
      this.snapshot = snap;
      this.watchApplication();
      await this.refreshCredential();
      this.selectedAgent = this.host.agents.rows[0]?.id ?? ""; this.removedServiceGrants = {};
      try { const baseline = await this.host.request<{ requiredScopes: string[] }>("microsoft-graph.configuration.validate", { policy: this.policy }); this.initialScopes = baseline.requiredScopes; this.scopes = baseline.requiredScopes; } catch { this.initialScopes = []; this.scopes = []; }
    } catch { this.error = "Configuration could not be loaded. Check administrator access and Gateway connection."; }
    finally { this.busy = false; this.render(); }
  }
  private render() {
    if (this.disposed || this.signal.aborted) return;
    setLocale(this.host.locale);
    const main = el("main", "mg-ui"); main.dir = isRtl() ? "rtl" : "ltr"; const header = el("header", "mg-header");
    append(header, el("div", "mg-eyebrow", "Plugins / Microsoft 365 for OpenClaw"), el("h1", "", "Microsoft 365 for OpenClaw"), el("p", "mg-lead", "Lege fest, welcher Agent auf welche Microsoft-Daten zugreifen darf und wann eine Freigabe nötig ist."));
    append(main, header);
    if (!this.host.connection.connected) { append(main, el("p", "mg-message", "Verbinde dich mit dem Gateway, um die Regeln zu bearbeiten.")); this.container.replaceChildren(main); return; }
    if (!this.host.connection.canAdmin) { append(main, el("p", "mg-message", "Zum Anzeigen und Ändern dieser Regeln brauchst du Administratorrechte.")); this.container.replaceChildren(main); return; }
    if (!this.snapshot) { append(main, el("p", "mg-message", this.error || "Regeln werden geladen…")); this.container.replaceChildren(main); return; }
    this.renderSetup(main);
    const rail = el("nav", "mg-steps"); rail.setAttribute("aria-label", localize("Konfigurationsschritte"));
    ["OneDrive", "Dienste", "Freigaben", "Prüfen"].forEach((name, index) => { const tab = button(`${index + 1}  ${localize(name)}`, () => { this.step = index; this.render(); if (index === 3) void this.validate(); }, index === this.step ? "active" : "ghost"); tab.disabled = this.busy; tab.setAttribute("aria-current", index === this.step ? "step" : "false"); append(rail, tab); }); append(main, rail);
    if (this.included && (this.dirty || !!this.error)) append(main, el("p", "mg-banner", format("Policy-Quelle: {name}. Änderungen werden beim Speichern in diese Datei geschrieben.", { name: this.includeName })));
    const application = this.statusError === "Die Regeln wurden außerhalb dieser Seite geändert. Bitte neu laden, um den aktuellen Stand zu sehen." ? "unknown" : this.applicationStatus;
    if (application !== "applied" || this.statusError) {
      const status = el("div", "mg-warning");
      status.setAttribute("role", "status");
      append(status, el("strong", "", application === "pending" ? "Regeln gespeichert – Anwendung noch ausstehend" : "Anwendung der Regeln nicht bestätigt"));
      append(status, el("p", "mg-status-detail", application === "pending" ? (this.statusChecksRemaining > 0 ? "Der Gateway hat die gespeicherte Version noch nicht übernommen. Diese Seite prüft den Status automatisch; bis dahin können die bisherigen Regeln gelten." : "Die Anwendung ist weiterhin nicht bestätigt. Die bisherigen Regeln können noch gelten; prüfe den Status erneut.") : "Der Gateway liefert derzeit keinen eindeutigen Anwendungsstatus. Die gespeicherten Regeln können bereits gelten, sind hier aber nicht bestätigt."));
      if (this.statusError) append(status, el("p", "mg-status-detail", this.statusError));
      const recheck = button("Anwendung erneut prüfen", () => { void this.checkApplication(); }, "secondary"); recheck.disabled = this.busy; append(status, recheck);
      append(main, status);
    }
    const body = el("section", "mg-body"); if (this.step === 0) this.renderOneDrive(body); else if (this.step === 1) this.renderServices(body); else if (this.step === 2) this.renderApprovals(body); else this.renderReview(body); append(main, body);
    if (this.error) append(main, el("p", "mg-error", this.error)); if (this.success) append(main, el("p", "mg-success", this.success));
    const footer = el("footer", "mg-footer"); append(footer, el("span", "mg-dirty", this.dirty ? "Ungespeicherte Änderungen" : "Keine ungespeicherten Änderungen"));
    if (this.dirty) append(footer, button("Änderungen verwerfen", () => { if (window.confirm(localize("Alle ungespeicherten Änderungen verwerfen?"))) { this.policy = clone(this.initial!); this.error = ""; this.success = ""; this.render(); } }, "ghost"));
    if (this.step > 0) append(footer, button("Zurück", () => { this.step--; this.render(); }));
    if (this.step < 3) append(footer, button("Weiter", () => { this.step++; this.render(); if (this.step === 3) void this.validate(); }, "primary"));
    if (this.step === 3 && this.dirty) append(footer, button("Änderungen speichern", () => { void this.save(); }, "primary"));
    append(main, footer); this.renderSignIn(main); this.container.replaceChildren(main);
  }
  private renderSetup(main: HTMLElement) {
    const configured = this.snapshot?.config?.plugins?.entries?.[id]?.config?.credentialVaultKey !== undefined;
    const savedPolicy = this.initial ?? blankPolicy();
    const grants = savedPolicy.services.onedrive.allowed_roots.some(root => Object.values(root.agents).some(agent => Object.values(agent.permissions).some(Boolean)))
      || (["calendar", "mail", "todo"] as const).some(service => Object.keys(savedPolicy.services[service].agents).length > 0);
    const connected = this.credential?.result === "valid";
    if (isConnectionEstablished({ secretRefConfigured: configured, savedGrantPresent: grants, applicationStatus: this.applicationStatus, credentialResult: this.credential?.result, statusError: !!this.statusError })) {
      const status = el("p", "mg-success mg-connected", "Verbindung hergestellt");
      status.setAttribute("role", "status");
      status.setAttribute("aria-label", localize("Verbindung hergestellt. Ein Lesezugriff durch einen berechtigten Agenten wurde nicht geprüft."));
      append(main, status);
      return;
    }
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
    else if (!connected) next = "Nächster Schritt: Unten mit Microsoft verbinden.";
    else next = "Konto verbunden. Prüfe mit einem berechtigten Agenten einen Lesezugriff; erst dann ist der Ablauf einsatzbereit.";
    append(card, el("p", connected ? "mg-status" : "mg-hint", next));
    if (!configured) { const link = el("a", "mg-button secondary", "Vault-Anleitung öffnen"); link.href = "https://clawhub.ai/packages/@baumus/openclaw-microsoft-graph"; link.target = "_blank"; link.rel = "noopener noreferrer"; append(card, link); }
    else if (!grants && !this.dirty) append(card, button("Zugriff festlegen", () => { this.step = 1; this.render(); }, "primary"));
    else if (this.dirty) append(card, button("Zum Prüfen", () => { this.step = 3; this.render(); void this.validate(); }, "primary"));
    else if (grants && this.applicationStatus === "applied" && !connected && this.credential?.result === "missing") append(card, button("Mit Microsoft verbinden", () => { this.container.querySelector("#microsoft-connect")?.scrollIntoView({ behavior: "smooth", block: "start" }); }, "primary"));
    append(main, card);
  }
  private async credentialCall<T>(method: string, params: Record<string, unknown>): Promise<T> {
    const reply = await this.host.request<CredentialReply<T>>(method, params);
    if (!reply || typeof reply !== "object" || typeof reply.ok !== "boolean") throw new Error("internal_error");
    if (!reply.ok) throw new Error(reply.error);
    return reply.value;
  }
  private async refreshCredential() {
    try {
      const status = await this.credentialCall<CredentialStatus>("microsoft-graph.credentials.status", {});
      if (!this.disposed && !this.signal.aborted) this.credential = status.credential;
    } catch { if (!this.disposed && !this.signal.aborted) this.credential = { result: "unavailable" }; }
  }
  private signInError(code: string): string {
    const messages: Record<string, string> = {
      credential_scope_missing: "Die Microsoft-Freigabe enthält nicht alle benötigten Berechtigungen. Prüfe die Einwilligung und starte erneut.",
      credential_vault_conflict: "Es gibt bereits einen Zugang. Dieser Assistent überschreibt ihn nicht.",
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
  private scheduleAuthCheck() {
    if (this.disposed || this.signal.aborted || this.authState !== "pending" || !this.device || this.authTimer) return;
    this.authTimer = setTimeout(() => { this.authTimer = undefined; void this.checkAuth(); }, 4000);
  }
  private async checkAuth() {
    if (!this.device || this.authState !== "pending" || this.disposed || this.signal.aborted) return;
    try {
      const result = await this.credentialCall<DeviceStatus>("microsoft-graph.credentials.device-status", { sessionId: this.device.sessionId });
      if (this.disposed || this.signal.aborted || this.authState !== "pending") return;
      this.authError = "";
      if (result.state === "created") { this.authState = "created"; this.authError = ""; await this.refreshCredential(); }
      else if (result.state === "failed") { this.authState = "failed"; this.authError = this.signInError(result.error ?? ""); }
      else if (Date.now() >= Date.parse(this.device.expiresAt)) { this.authState = "failed"; this.authError = this.signInError("device_authorization_expired"); }
      else this.scheduleAuthCheck();
    } catch { this.authError = "Statusprüfung unterbrochen. Verbindung prüfen; diese Seite versucht es erneut."; if (Date.now() >= Date.parse(this.device.expiresAt)) { this.authState = "failed"; this.authError = this.signInError("device_authorization_expired"); } else this.scheduleAuthCheck(); }
    this.render();
  }
  private async startAuth() {
    if (this.authBusy || this.authState === "pending" || this.credential?.result !== "missing" || this.dirty || this.applicationStatus !== "applied") return;
    if (!/^[0-9a-fA-F-]{36}$/.test(this.authClientId.trim()) || !/^[A-Za-z0-9.-]{1,253}$/.test(this.authTenant.trim())) { this.authError = "Bitte eine gültige Microsoft App-ID und Tenant-ID oder -Domain eingeben."; this.render(); return; }
    this.authBusy = true; this.authError = ""; this.render();
    try {
      const started = await this.credentialCall<DeviceStart>("microsoft-graph.credentials.device-start", { clientId: this.authClientId.trim(), tenant: this.authTenant.trim() });
      if (this.disposed || this.signal.aborted) return;
      if (!isMicrosoftDeviceVerificationUri(started.verificationUri)) throw new Error("device_authorization_failed");
      this.device = started; this.authState = "pending"; this.scheduleAuthCheck();
    } catch (error) { this.authState = "failed"; this.authError = this.signInError(error instanceof Error ? error.message : ""); }
    finally { this.authBusy = false; this.render(); }
  }
  private async cancelAuth() {
    if (!this.device || this.authState !== "pending") return;
    this.authBusy = true; this.render();
    try { const result = await this.credentialCall<DeviceStatus>("microsoft-graph.credentials.device-cancel", { sessionId: this.device.sessionId }); if (result.state === "created") { this.authState = "created"; this.authError = ""; await this.refreshCredential(); } else { this.authState = "failed"; this.authError = this.signInError(result.error ?? "device_authorization_cancelled"); } }
    catch { this.authError = this.signInError(""); }
    finally { this.authBusy = false; if (this.authTimer) clearTimeout(this.authTimer); this.authTimer = undefined; if (this.authState === "pending") this.scheduleAuthCheck(); this.render(); }
  }
  private renderSignIn(main: HTMLElement) {
    if (this.credential?.result === "valid") return;
    const card = el("section", "mg-section mg-auth"); card.id = "microsoft-connect";
    append(card, el("h2", "", "Mit Microsoft verbinden"));
    const result = this.credential?.result;
    if (result === "quarantined") { append(card, el("p", "mg-warning", "Der vorhandene Zugang ist gesperrt und muss separat wiederhergestellt werden. Dieser Assistent überschreibt ihn nicht.")); append(main, card); return; }
    if (result === "unavailable") { append(card, el("p", "mg-warning", "Zugangsstatus nicht verfügbar. Prüfe Vault-Schlüssel und Gateway-Verbindung.")); append(main, card); return; }
    if (!result) { append(card, el("p", "", "Zugangsstatus wird geladen…")); append(main, card); return; }
    append(card, el("p", "", "Melde dich in deinem Browser bei Microsoft an. Der Zugang wird danach direkt verschlüsselt gespeichert; du musst keinen Token kopieren."));
    if (this.authState === "pending" && this.device) {
      append(card, el("p", "mg-status", "1. Öffne Microsoft in einem neuen Tab. 2. Gib dort den Code ein. 3. Bestätige die angezeigten Berechtigungen. Diese Seite erkennt den Abschluss automatisch."));
      const link = el("a", "mg-button primary", "Microsoft-Anmeldung öffnen"); link.href = this.device.verificationUri; link.target = "_blank"; link.rel = "noopener noreferrer"; append(card, link);
      const code = el("p", "mg-auth-code", this.device.userCode); code.setAttribute("aria-label", localize("Microsoft-Anmeldecode")); append(card, el("p", "mg-hint", "Einmaliger Microsoft-Anmeldecode:"), code);
      append(card, button("Code kopieren", () => { if (!navigator.clipboard?.writeText) { this.authError = "Kopieren nicht möglich. Markiere den Code und gib ihn bei Microsoft ein."; this.render(); return; } void navigator.clipboard.writeText(this.device!.userCode).catch(() => { this.authError = "Kopieren nicht möglich. Markiere den Code und gib ihn bei Microsoft ein."; this.render(); }); }));
      append(card, el("p", "mg-hint", "Warte auf die Bestätigung hier. Der Code läuft nach spätestens 15 Minuten ab."));
      const cancel = button("Anmeldung abbrechen", () => { void this.cancelAuth(); }, "ghost"); cancel.disabled = this.authBusy; append(card, cancel);
    } else {
      append(card, el("p", "mg-hint", "Du benötigst die App-ID einer genehmigten öffentlichen Microsoft-Anwendung und deine Tenant-ID. Dein Administrator kann dir beide Werte geben. Ein Client Secret wird nicht benötigt."));
      const client = field(card, "Microsoft App-ID", this.authClientId, value => { this.authClientId = value; }); client.placeholder = "00000000-0000-0000-0000-000000000000";
      const tenant = field(card, "Tenant-ID oder Tenant-Domain", this.authTenant, value => { this.authTenant = value; }); tenant.placeholder = "example.onmicrosoft.com";
      const start = button("Anmeldung starten", () => { void this.startAuth(); }, "primary"); start.disabled = this.authBusy || this.dirty || this.applicationStatus !== "applied"; append(card, start);
      if (this.dirty || this.applicationStatus !== "applied") append(card, el("p", "mg-hint", "Speichere die Zugriffsregeln und warte, bis sie im Gateway angewendet sind."));
    }
    if (this.authError) append(card, el("p", "mg-error", this.authError));
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
    if (this.busy) return;
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
    append(title, el("p", "mg-kicker", "OneDrive-Zugriff"), el("h2", "", "Wer darf auf welche OneDrive-Bereiche zugreifen?"), el("p", "mg-hint", "Jeder Bereich zeigt die berechtigten Agenten und ihre Rechte. Änderungen werden erst nach Prüfen und Speichern wirksam."));
    append(heading, title, el("span", "mg-count", format("{count} Zuweisungen", { count: assignments })));
    append(body, heading);
    const add = el("section", "mg-add-folder");
    append(add, el("h3", "", "Ordner hinzufügen"), el("p", "mg-hint", "Pfad in deinem OneDrive. Der Ordner wird vor dem Hinzufügen geprüft."));
    const input = field(add, "Ordnerpfad", this.newFolderPath, value => { this.newFolderPath = value; });
    input.placeholder = "/Projects/Clients";
    input.addEventListener("keydown", event => { if (event.key === "Enter") { event.preventDefault(); this.newFolderPath = input.value.trim(); void this.addFolder(); } });
    const addButton = button("Ordner prüfen und hinzufügen", () => { this.newFolderPath = input.value.trim(); void this.addFolder(); }, "primary"); addButton.disabled = this.busy;
    append(add, addButton); append(body, add);
    if (!roots.length) { append(body, el("p", "mg-message", "Noch kein OneDrive-Ordner freigegeben.")); return; }
    const grid = el("div", "mg-root-grid");
    for (const [index, root] of roots.entries()) {
      const card = el("article", "mg-root-card");
      const header = el("div", "mg-root-header");
      const name = root.path.split("/").filter(Boolean).at(-1) || root.path;
      const identity = el("div", "mg-root-identity"); append(identity, el("h3", "", name), el("p", "mg-root-path", root.path));
      append(header, identity, el("span", "mg-descendants", "Inkl. Unterordner")); append(card, header);
      const agents = Object.entries(root.agents).filter(([, grant]) => Object.values(grant.permissions).some(Boolean));
      if (!agents.length) append(card, el("p", "mg-empty-grants", "Noch kein Agent berechtigt."));
      for (const [agentId, grant] of agents) {
        const row = el("div", "mg-agent-row");
        const agentName = this.host.agents.rows.find(agent => agent.id === agentId)?.name || agentId;
        const agentLabel = el("strong", "mg-agent-name", agentName);
        const rights = el("div", "mg-rights");
        for (const [op, label] of [["read", "Lesen"], ["write", "Schreiben"], ["delete", "Löschen"]] as const) {
          const enabled = grant.permissions[op] === true;
          const badge = el("span", `mg-right ${enabled ? "is-allowed" : "is-denied"}`, `${enabled ? "✓" : "–"} ${localize(label)}`);
          badge.setAttribute("aria-label", `${localize(label)}: ${localize(enabled ? "erlaubt" : "nicht erlaubt")}`);
          append(rights, badge);
        }
        const edit = button("Rechte ändern", () => { this.editingAccess = { rootLabel: root.label, agentId, isNew: false, permissions: { read: grant.permissions.read === true, write: grant.permissions.write === true, delete: grant.permissions.delete === true } }; this.render(); }, "ghost mg-edit-rights");
        edit.setAttribute("aria-label", format("Rechte für {agent} auf {path} ändern", { agent: agentName, path: root.path }));
        append(row, agentLabel, rights, edit); append(card, row);
      }
      if (this.editingAccess?.rootLabel === root.label) this.renderAccessEditor(card, root);
      else {
        const available = this.host.agents.rows.filter(agent => !agents.some(([agentId]) => agentId === agent.id));
        if (available.length) append(card, button("+ Agent hinzufügen", () => { this.editingAccess = { rootLabel: root.label, agentId: available[0]!.id, isNew: true, permissions: { read: false, write: false, delete: false } }; this.render(); }, "ghost mg-add-agent"));
        else if (!this.host.agents.rows.length) append(card, el("p", "mg-hint", "Keine Agenten gefunden."));
      }
      const menu = el("details", "mg-root-menu");
      append(menu, el("summary", "", "Weitere Aktionen"));
      append(menu, button("Ordner für alle Agenten entfernen", () => { if (window.confirm(format("Den Ordner {path} für alle Agenten entfernen?", { path: root.path }))) { roots.splice(index, 1); if (this.editingAccess?.rootLabel === root.label) this.editingAccess = undefined; this.render(); } }, "danger"));
      append(card, menu); append(grid, card);
    }
    append(body, grid);
  }
  private renderAccessEditor(card: HTMLElement, root: Root) {
    const editing = this.editingAccess!;
    const panel = el("div", "mg-access-editor");
    append(panel, el("h4", "", editing.isNew ? "Agent hinzufügen" : "Rechte konfigurieren"));
    if (editing.isNew) {
      const label = el("label", "mg-field"); append(label, el("span", "mg-label", "Agent"));
      const select = el("select");
      for (const agent of this.host.agents.rows.filter(agent => !Object.values(root.agents[agent.id]?.permissions ?? {}).some(Boolean))) {
        const option = el("option", "", agent.name || agent.id); option.value = agent.id; option.selected = agent.id === editing.agentId; append(select, option);
      }
      select.addEventListener("change", () => { editing.agentId = select.value; }); append(label, select); append(panel, label);
    } else append(panel, el("p", "mg-hint", this.host.agents.rows.find(agent => agent.id === editing.agentId)?.name || editing.agentId));
    const rights = el("div", "mg-editor-rights");
    for (const [op, label] of [["read", "Lesen"], ["write", "Schreiben"], ["delete", "Löschen"]] as const)
      checkbox(rights, label, editing.permissions[op], value => { editing.permissions[op] = value; });
    append(panel, rights, el("p", "mg-hint", "Ohne ausgewähltes Recht wird der Agent aus diesem Bereich entfernt."));
    const actions = el("div", "mg-editor-actions");
    append(actions, button("Abbrechen", () => { this.editingAccess = undefined; this.render(); }, "ghost"));
    append(actions, button("In Entwurf übernehmen", () => {
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
    for (const service of ["calendar", "mail", "todo"] as const) {
      const card = el("fieldset", "mg-section"); append(card, el("legend", "", serviceNames[service]));
      const grants = this.policy.services[service].agents; const existing = grants[agentId];
      checkbox(card, format("{service} nutzen", { service: localize(serviceNames[service]) }), !!existing, checked => {
        const key = `${service}:${agentId}`;
        if (checked) grants[agentId] = existing ?? this.removedServiceGrants[key] ?? { operations: [...operations[service]] };
        else { if (existing) this.removedServiceGrants[key] = clone(existing); delete grants[agentId]; }
        this.render();
      });
      if (existing && existing.operations.length < operations[service].length) append(card, el("p", "mg-hint", "Bestehender Zugang ist auf einzelne Funktionen eingeschränkt. Diese Einschränkung bleibt erhalten."));
      append(body, card);
    }
  }
  private renderApprovals(body: HTMLElement) {
    append(body, el("h2", "", "Wann ist eine Freigabe nötig?"), el("p", "", "Eine Freigabe ist eine Rückfrage vor der konkreten Aktion. Die Agentenrechte aus den ersten beiden Schritten gelten immer zusätzlich."));
    for (const service of ["onedrive", "calendar", "mail", "todo"] as const) {
      const card = el("fieldset", "mg-section"); append(card, el("legend", "", serviceNames[service]));
      append(card, el("p", "", "Unkritisch · Lesen und Suchen: keine Rückfrage."));
      const configured = this.policy.rules.warningApprovalsByService ?? {};
      const baseline = this.snapshot?.config?.plugins?.entries?.[id]?.config?.warningApprovalsRequired !== false;
      checkbox(card, "Warn-Aktionen · Erstellen oder Ändern: vorher fragen", configured[service] ?? baseline, checked => {
        this.policy.rules.warningApprovalsByService = { ...configured, [service]: checked }; this.render();
      });
      const critical = service === "onedrive" ? "Löschen" : service === "calendar" ? "Löschen oder auf Termine antworten" : service === "mail" ? "Löschen oder E-Mail senden" : "Löschen";
      append(card, el("p", "", format("Kritisch · {action}: immer einzeln fragen.", { action: localize(critical) })));
      append(body, card);
    }
  }
  private renderReview(body: HTMLElement) {
    append(body, el("h2", "", "Änderungen prüfen"), el("p", "", "Prüfe die Zugriffe und Rückfragen. Nur diese Microsoft-Graph-Regeln werden geändert."));
    append(body, el("p", "mg-status", policySummary(this.policy).join(" · ")));

    const newScopes = this.scopes.filter(scope => !this.initialScopes.includes(scope));
    if (this.pluginEnabled && newScopes.length) append(body, el("p", "mg-warning", format("Für neue Zugriffe kann eine Microsoft-Einwilligung nötig sein ({scopes}). Solange das Plugin aktiv ist, kann diese Seite erweiterte Rechte nicht speichern.", { scopes: newScopes.join(", ") })));
    const approvals = this.policy.rules.warningApprovalsByService ?? {};
    for (const service of ["onedrive", "calendar", "mail", "todo"] as const) if (approvals[service] === false) append(body, el("p", "mg-warning", format("{service}: Warn-Aktionen dürfen ohne Rückfrage ausgeführt werden. Kritische Aktionen benötigen weiterhin eine Freigabe.", { service: localize(serviceNames[service]) })));
    if (!this.dirty) { append(body, el("p", "", "Keine Änderungen zum Speichern.")); return; }
    if (JSON.stringify(this.initial) !== JSON.stringify(this.policy)) checkbox(body, "Ich habe die Zugriffsänderungen geprüft, auch entzogene Rechte", this.confirmedRemoval, v => { this.confirmedRemoval = v; });

  }
  private async validate(): Promise<boolean> {
    this.error = "";
    const roster = new Set(this.host.agents.rows.map(a => a.id));
    for (const service of ["calendar", "mail", "todo"] as const) for (const agent of Object.keys(this.policy.services[service].agents)) if (!roster.has(agent)) { this.error = format("Unknown configured agent: {agent}", { agent }); this.render(); return false; }
    for (const root of this.policy.services.onedrive.allowed_roots) for (const agent of Object.keys(root.agents)) if (!roster.has(agent)) { this.error = format("Unknown configured agent: {agent}", { agent }); this.render(); return false; }
    try { const result = await this.host.request<{ valid: boolean; requiredScopes: string[] }>("microsoft-graph.configuration.validate", { policy: this.policy }); this.scopes = result.requiredScopes; this.render(); return result.valid; }
    catch { this.error = "Policy validation failed. Check root labels, paths, IDs, and grant resources."; this.render(); return false; }
  }
  private async save() {
    if (!this.dirty || this.busy || !this.snapshot) return;
    if (JSON.stringify(this.initial) !== JSON.stringify(this.policy) && !this.confirmedRemoval) { this.error = "Bitte die Zugriffsänderungen vor dem Speichern bestätigen."; this.render(); return; }
    this.stopStatusChecks(); this.statusError = "";
    this.busy = true; this.error = ""; this.success = ""; this.render();
    try {
      if (!(await this.validate())) return;
      const newScopes = this.scopes.filter(scope => !this.initialScopes.includes(scope));
      if (this.pluginEnabled && newScopes.length) { this.error = format("Cannot save while the plugin is enabled: new delegated scopes require independent consent verification ({scopes}).", { scopes: newScopes.join(", ") }); return; }
      const fresh = await this.host.request<ConfigSnapshot>("config.get", {});
      if (fresh.hash !== this.snapshot.hash) { this.error = "Configuration changed since this draft loaded. Reload and reapply your changes."; return; }
      const freshAuthored = fresh.parsed?.plugins?.entries?.[id]?.config?.policy;
      const freshInclude = freshAuthored && typeof freshAuthored === "object" && "$include" in freshAuthored ? String((freshAuthored as { $include: unknown }).$include) : "";
      if (freshInclude !== this.includeName) { this.error = "Die Policy-Quelle wurde geändert. Bitte neu laden."; return; }
      const replacements: string[] = [];
      const patchPolicy = collectDiff(this.initial, this.policy, "plugins.entries.microsoft-graph.config.policy", replacements);
      const patch: Record<string, unknown> = {}; if (patchPolicy !== undefined) patch.policy = patchPolicy;
      const raw = JSON.stringify({ plugins: { entries: { [id]: { config: patch } } } });
      const result = await this.host.request<{ changedPaths?: string[] }>("config.patch", { raw, baseHash: fresh.hash, replacePaths: replacements, note: "Microsoft Graph configuration UI save" });
      const verify = await this.host.request<ConfigSnapshot>("config.get", {});
      const applied = verify.config?.plugins?.entries?.[id]?.config;
      if (JSON.stringify(applied?.policy) !== JSON.stringify(this.policy)) { this.error = "Configuration write returned, but the effective values could not be verified. Reload before retrying."; return; }
      this.snapshot = verify; this.initial = clone(this.policy);
      this.watchApplication();
      void result;
    } catch {
      try {
        const verify = await this.host.request<ConfigSnapshot>("config.get", {});
        if (JSON.stringify(verify.config?.plugins?.entries?.[id]?.config?.policy) === JSON.stringify(this.policy)) {
          this.snapshot = verify; this.initial = clone(this.policy);
          this.watchApplication();
        } else this.error = "Speichern nicht bestätigt. Der Entwurf bleibt erhalten; bitte vor einem erneuten Versuch neu laden.";
      } catch { this.error = "Speicherzustand unbekannt. Bitte Gateway-Status prüfen und die Seite neu laden."; }
    }
    finally { this.busy = false; if (this.applicationStatus === "pending" && !this.statusTimer && this.statusChecksRemaining === 0) this.watchApplication(); this.render(); }
  }
}

export default defineControlUiPlugin({ id, activate(host) {
  setLocale(host.locale);
  const pageId = "configure";
  const disposers = [
    host.ui.registerPage({ id: pageId, label: "Microsoft 365 for OpenClaw", mount(container, context) { const view = new ConfigurationPage(container, context.host, context.signal); return { dispose: () => view.dispose() }; } }),
    host.ui.registerNavigation({ id: "configure", label: "Microsoft 365 for OpenClaw", page: { id: pageId }, icon: "settings", order: 80 }),
  ];
  return () => { for (const dispose of disposers.reverse()) dispose(); };
} });
