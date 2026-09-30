import { defineControlUiPlugin, type ControlUiHost } from "openclaw/plugin-sdk/control-ui";
import "./control-ui.css";

type Grant = { operations: string[]; resources?: string[] };
type Root = { label: string; path: string; drive_id: string; item_id: string; include_descendants: true; agents_instructions?: "trusted"; permissions: Record<"read" | "write" | "delete", boolean>; agents: Record<string, { permissions: Partial<Record<"read" | "write" | "delete", boolean>> }> };
type Policy = { version: 2; rules: { default: "deny"; warningApprovalsByService?: Partial<Record<"onedrive" | "calendar" | "mail" | "todo", boolean>> }; services: { onedrive: { allowed_roots: Root[] }; calendar: { agents: Record<string, Grant> }; mail: { agents: Record<string, Grant> }; todo: { agents: Record<string, Grant> } } };
type ConfigSnapshot = { hash: string; config: { plugins?: { entries?: Record<string, { enabled?: boolean; config?: Record<string, unknown> }> } }; parsed?: { plugins?: { entries?: Record<string, { enabled?: boolean; config?: Record<string, unknown> }> } }; configRevisionHash?: string; appliedConfigHash?: string };
const id = "microsoft-graph";
const operations = { calendar: ["read", "create", "update", "respond", "attach", "delete"], mail: ["read", "draft", "update", "move", "mark", "send", "delete"], todo: ["read", "create", "update", "delete"] } as const;
const serviceNames = { onedrive: "OneDrive", calendar: "Kalender", mail: "E-Mail", todo: "To Do" } as const;
const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v)) as T;
const blankPolicy = (): Policy => ({ version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: {} }, mail: { agents: {} }, todo: { agents: {} } } });
function el<K extends keyof HTMLElementTagNameMap>(tag: K, className = "", value?: string): HTMLElementTagNameMap[K] { const node = document.createElement(tag); if (className) node.className = className; if (value !== undefined) node.textContent = value; return node; }
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
    `${policy.services.onedrive.allowed_roots.length} OneDrive-Ordner`,
    ...(["calendar", "mail", "todo"] as const).map(s => `${Object.keys(policy.services[s].agents).length} Agenten für ${serviceNames[s]}`),
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
  private removedServiceGrants: Record<string, Grant> = {};
  private busy = false;
  private error = "";
  private success = "";
  private statusError = "";
  private confirmedRemoval = false;
  private disposed = false;
  private statusTimer?: ReturnType<typeof setTimeout>;
  private statusChecksRemaining = 0;
  private readonly unsubscribe: () => void;
  constructor(private container: HTMLElement, private host: ControlUiHost, private signal: AbortSignal) {
    this.unsubscribe = host.subscribe(() => { if (!this.snapshot && !this.busy && this.authorized) void this.load(); else this.render(); });
    void this.load();
    this.render();
  }
  dispose() { this.disposed = true; this.stopStatusChecks(); this.unsubscribe(); this.container.replaceChildren(); }
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
      this.selectedAgent = this.host.agents.rows[0]?.id ?? ""; this.removedServiceGrants = {};
      try { const baseline = await this.host.request<{ requiredScopes: string[] }>("microsoft-graph.configuration.validate", { policy: this.policy }); this.initialScopes = baseline.requiredScopes; this.scopes = baseline.requiredScopes; } catch { this.initialScopes = []; this.scopes = []; }
    } catch { this.error = "Configuration could not be loaded. Check administrator access and Gateway connection."; }
    finally { this.busy = false; this.render(); }
  }
  private render() {
    if (this.disposed || this.signal.aborted) return;
    const main = el("main", "mg-ui"); const header = el("header", "mg-header");
    append(header, el("div", "mg-eyebrow", "Plugins / Microsoft Graph"), el("h1", "", "Microsoft Graph Zugriff"), el("p", "mg-lead", "Lege fest, welcher Agent auf welche Microsoft-Daten zugreifen darf und wann eine Freigabe nötig ist."));
    append(main, header);
    if (!this.host.connection.connected) { append(main, el("p", "mg-message", "Verbinde dich mit dem Gateway, um die Regeln zu bearbeiten.")); this.container.replaceChildren(main); return; }
    if (!this.host.connection.canAdmin) { append(main, el("p", "mg-message", "Zum Anzeigen und Ändern dieser Regeln brauchst du Administratorrechte.")); this.container.replaceChildren(main); return; }
    if (!this.snapshot) { append(main, el("p", "mg-message", this.error || "Regeln werden geladen…")); this.container.replaceChildren(main); return; }
    const rail = el("nav", "mg-steps"); rail.setAttribute("aria-label", "Konfigurationsschritte");
    ["OneDrive", "Dienste", "Freigaben", "Prüfen"].forEach((name, index) => { const tab = button(`${index + 1}  ${name}`, () => { this.step = index; this.render(); if (index === 3) void this.validate(); }, index === this.step ? "active" : "ghost"); tab.disabled = this.busy; tab.setAttribute("aria-current", index === this.step ? "step" : "false"); append(rail, tab); }); append(main, rail);
    if (this.included) append(main, el("p", "mg-banner", `Policy-Quelle: ${this.includeName}. Änderungen werden beim Speichern in diese Datei geschrieben.`));
    const application = this.statusError.startsWith("Die Regeln wurden außerhalb") ? "unknown" : this.applicationStatus;
    const status = el("div", application === "applied" ? "mg-success" : "mg-warning");
    status.setAttribute("role", "status");
    append(status, el("strong", "", application === "applied" ? "Gespeicherte Regeln im Gateway angewendet" : application === "pending" ? "Regeln gespeichert – Anwendung noch ausstehend" : "Anwendung der Regeln nicht bestätigt"));
    append(status, el("p", "mg-status-detail", application === "applied" ? "Gespeicherte und angewendete Konfigurationsversion stimmen überein." : application === "pending" ? (this.statusChecksRemaining > 0 ? "Der Gateway hat die gespeicherte Version noch nicht übernommen. Diese Seite prüft den Status automatisch; bis dahin können die bisherigen Regeln gelten." : "Die Anwendung ist weiterhin nicht bestätigt. Die bisherigen Regeln können noch gelten; prüfe den Status erneut.") : "Der Gateway liefert derzeit keinen eindeutigen Anwendungsstatus. Die gespeicherten Regeln können bereits gelten, sind hier aber nicht bestätigt."));
    if (this.statusError) append(status, el("p", "mg-status-detail", this.statusError));
    if (application !== "applied") { const recheck = button("Anwendung erneut prüfen", () => { void this.checkApplication(); }, "secondary"); recheck.disabled = this.busy; append(status, recheck); }
    append(main, status);
    const body = el("section", "mg-body"); if (this.step === 0) this.renderOneDrive(body); else if (this.step === 1) this.renderServices(body); else if (this.step === 2) this.renderApprovals(body); else this.renderReview(body); append(main, body);
    if (this.error) append(main, el("p", "mg-error", this.error)); if (this.success) append(main, el("p", "mg-success", this.success));
    const footer = el("footer", "mg-footer"); append(footer, el("span", "mg-dirty", this.dirty ? "Ungespeicherte Änderungen" : "Keine ungespeicherten Änderungen"));
    if (this.dirty) append(footer, button("Änderungen verwerfen", () => { if (window.confirm("Alle ungespeicherten Änderungen verwerfen?")) { this.policy = clone(this.initial!); this.error = ""; this.success = ""; this.render(); } }, "ghost"));
    if (this.step > 0) append(footer, button("Zurück", () => { this.step--; this.render(); }));
    if (this.step < 3) append(footer, button("Weiter", () => { this.step++; this.render(); if (this.step === 3) void this.validate(); }, "primary"));
    if (this.step === 3 && this.dirty) append(footer, button("Änderungen speichern", () => { void this.save(); }, "primary"));
    append(main, footer); this.container.replaceChildren(main);
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
      if (roots.some(root => root.drive_id === resolved.drive_id && root.item_id === resolved.item_id)) { this.error = "Dieser Ordner ist bereits freigegeben. Wähle unten die Rechte für den Agenten."; return; }
      let label = "folder_" + (roots.length + 1); let suffix = roots.length + 1;
      while (roots.some(root => root.label === label)) label = "folder_" + ++suffix;
      roots.push({ label, path: resolved.path, drive_id: resolved.drive_id, item_id: resolved.item_id, include_descendants: true, permissions: { read: false, write: false, delete: false }, agents: {} });
      this.newFolderPath = ""; this.success = "Ordner geprüft. Wähle jetzt die Rechte für den Agenten.";
    } catch { this.error = "Ordner nicht gefunden oder nicht prüfbar. Prüfe den Pfad und die OneDrive-Verbindung."; }
    finally { this.busy = false; this.render(); }
  }
  private renderOneDrive(body: HTMLElement) {
    append(body, el("h2", "", "OneDrive-Ordner"), el("p", "", "Wähle einen Agenten und seine Rechte je Ordner. Die Freigabe gilt immer auch für alle Unterordner. Ohne Häkchen hat der Agent keinen Zugriff."));
    const agentId = this.renderAgentPicker(body);
    if (!agentId) { append(body, el("p", "mg-message", "Keine Agenten gefunden.")); return; }
    const roots = this.policy.services.onedrive.allowed_roots;
    if (!roots.length) append(body, el("p", "mg-message", "Noch kein OneDrive-Ordner freigegeben."));
    for (const [index, root] of roots.entries()) {
      const card = el("fieldset", "mg-section"); append(card, el("legend", "", root.path));
      append(card, el("p", "mg-hint", "Gilt auch für alle Unterordner"));
      const row = el("div", "mg-check-grid");
      for (const [op, label] of [["read", "Lesen"], ["write", "Schreiben"], ["delete", "Löschen"]] as const) checkbox(row, label, root.agents[agentId]?.permissions[op] === true, checked => {
        const grant = root.agents[agentId] ?? { permissions: {} };
        if (checked) { grant.permissions[op] = true; root.agents[agentId] = grant; }
        else { delete grant.permissions[op]; if (Object.values(grant.permissions).some(Boolean)) root.agents[agentId] = grant; else delete root.agents[agentId]; }
        root.permissions[op] = Object.values(root.agents).some(a => a.permissions[op] === true);
        this.render();
      });
      append(card, row, button("Ordner für alle Agenten entfernen", () => { if (window.confirm(`Den Ordner ${root.path} für alle Agenten entfernen?`)) { roots.splice(index, 1); this.render(); } }, "danger")); append(body, card);
    }
    const add = el("section", "mg-section"); append(add, el("h3", "", "Ordner hinzufügen"));
    append(add, el("p", "mg-hint", "Pfad in deinem OneDrive, zum Beispiel /Projekte/Kunden. Der Ordner wird vor dem Hinzufügen geprüft; technische IDs werden automatisch ermittelt."));
    const input = field(add, "Ordnerpfad", this.newFolderPath, value => { this.newFolderPath = value; });
    input.placeholder = "/Projekte/Kunden";
    input.addEventListener("keydown", event => { if (event.key === "Enter") { event.preventDefault(); this.newFolderPath = input.value.trim(); void this.addFolder(); } });
    append(add, button("Ordner prüfen und hinzufügen", () => { void this.addFolder(); }, "primary")); append(body, add);
  }
  private renderServices(body: HTMLElement) {
    append(body, el("h2", "", "Microsoft-Dienste"), el("p", "", "Wähle, welche Dienste der Agent nutzen darf. Ein neuer Zugang umfasst alle Funktionen des Dienstes. Bestehende eingeschränkte Zugänge bleiben unverändert."));
    const agentId = this.renderAgentPicker(body);
    if (!agentId) { append(body, el("p", "mg-message", "Keine Agenten gefunden.")); return; }
    for (const service of ["calendar", "mail", "todo"] as const) {
      const card = el("fieldset", "mg-section"); append(card, el("legend", "", serviceNames[service]));
      const grants = this.policy.services[service].agents; const existing = grants[agentId];
      checkbox(card, `${serviceNames[service]} nutzen`, !!existing, checked => {
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
      append(card, el("p", "", `Kritisch · ${critical}: immer einzeln fragen.`));
      append(body, card);
    }
  }
  private renderReview(body: HTMLElement) {
    append(body, el("h2", "", "Änderungen prüfen"), el("p", "", "Prüfe die Zugriffe und Rückfragen. Nur diese Microsoft-Graph-Regeln werden geändert."));
    append(body, el("p", "mg-status", policySummary(this.policy).join(" · ")));

    const newScopes = this.scopes.filter(scope => !this.initialScopes.includes(scope));
    if (this.pluginEnabled && newScopes.length) append(body, el("p", "mg-warning", `Für neue Zugriffe kann eine Microsoft-Einwilligung nötig sein (${newScopes.join(", ")}). Solange das Plugin aktiv ist, kann diese Seite erweiterte Rechte nicht speichern.`));
    const approvals = this.policy.rules.warningApprovalsByService ?? {};
    for (const service of ["onedrive", "calendar", "mail", "todo"] as const) if (approvals[service] === false) append(body, el("p", "mg-warning", `${serviceNames[service]}: Warn-Aktionen dürfen ohne Rückfrage ausgeführt werden. Kritische Aktionen benötigen weiterhin eine Freigabe.`));
    if (!this.dirty) { append(body, el("p", "", "Keine Änderungen zum Speichern.")); return; }
    if (JSON.stringify(this.initial) !== JSON.stringify(this.policy)) checkbox(body, "Ich habe die Zugriffsänderungen geprüft, auch entzogene Rechte", this.confirmedRemoval, v => { this.confirmedRemoval = v; });

  }
  private async validate(): Promise<boolean> {
    this.error = "";
    const roster = new Set(this.host.agents.rows.map(a => a.id));
    for (const service of ["calendar", "mail", "todo"] as const) for (const agent of Object.keys(this.policy.services[service].agents)) if (!roster.has(agent)) { this.error = `Unknown configured agent: ${agent}`; this.render(); return false; }
    for (const root of this.policy.services.onedrive.allowed_roots) for (const agent of Object.keys(root.agents)) if (!roster.has(agent)) { this.error = `Unknown configured agent: ${agent}`; this.render(); return false; }
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
      if (this.pluginEnabled && newScopes.length) { this.error = `Cannot save while the plugin is enabled: new delegated scopes require independent consent verification (${newScopes.join(", ")}).`; return; }
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
  const pageId = "configure";
  const disposers = [
    host.ui.registerPage({ id: pageId, label: "Microsoft Graph Zugriff", mount(container, context) { const view = new ConfigurationPage(container, context.host, context.signal); return { dispose: () => view.dispose() }; } }),
    host.ui.registerNavigation({ id: "configure", label: "Microsoft Graph Zugriff", page: { id: pageId }, icon: "settings", order: 80 }),
  ];
  return () => { for (const dispose of disposers.reverse()) dispose(); };
} });
