import type { ControlUiHost } from "openclaw/plugin-sdk/control-ui";
import type { Account, Binding } from "./accounts.js";
import type { GraphPolicy } from "./policy.js";

type Status = { profileId: string; account: Account | null; ceiling: GraphPolicy; workers: string[]; warning: string; automatic?: { profileId: string; invalidated: boolean }; credential: { result: string } };
type SessionStatus = { binding: Binding | null; available: string[]; workerProfileId?: string };
const serviceLabels = { calendar: "Kalender", mail: "E-Mail", todo: "To Do" };
const actionLabels: Record<string, string> = { read: "Lesen", create: "Erstellen", update: "Bearbeiten", respond: "Antworten", attach: "Anhänge", delete: "Löschen", draft: "Entwürfe", move: "Verschieben", mark: "Markieren", send: "Senden", write: "Schreiben" };
const node = <K extends keyof HTMLElementTagNameMap>(tag: K, text?: string): HTMLElementTagNameMap[K] => { const el = document.createElement(tag); if (text) el.textContent = text; return el; };
const button = (text: string, action: () => void) => { const b = node("button", text); b.type = "button"; b.className = "mg-button secondary"; b.addEventListener("click", action); return b; };
const clone = <T>(x: T): T => structuredClone(x);
export class AccountPage {
  private status?: Status;
  private session?: SessionStatus;
  private policy?: GraphPolicy;
  private shared = false;
  private unattended = false;
  private busy = false;
  private error = "";
  private notice = "";
  private clientId = "";
  private tenant = "common";
  private device?: { sessionId: string; userCode: string; verificationUri: string; expiresAt: string };
  private poll?: ReturnType<typeof setTimeout>;
  private epoch = 0;
  private disposed = false;
  private sessionSelection = "";
  private readonly unsubscribe: () => void;
  constructor(private container: HTMLElement, private host: ControlUiHost, private signal: AbortSignal) {
    this.unsubscribe = host.subscribe(() => {
      if (!host.connection.connected) { this.status = undefined; this.session = undefined; this.policy = undefined; this.render(); }
      else if (this.sessionSelection !== host.sessions.selectedKey) { this.session = undefined; this.sessionSelection = host.sessions.selectedKey; void this.refreshSession(); }
    });
    this.sessionSelection = host.sessions.selectedKey;
    void this.load();
  }
  dispose() { this.disposed = true; this.epoch++; this.unsubscribe(); if (this.poll) clearTimeout(this.poll); this.container.replaceChildren(); }
  private active() { return !this.disposed && !this.signal.aborted; }
  private async request<T>(name: string, params: Record<string, unknown> = {}): Promise<T> {
    return this.host.request<T>(`microsoft-graph.accounts.${name}`, params);
  }
  private async task(action: () => Promise<void>) {
    if (this.busy) return;
    this.busy = true; this.error = ""; this.notice = ""; this.render();
    try { await action(); } catch (error) { this.error = error instanceof Error ? error.message : String((error as { message?: unknown })?.message ?? "Aktion nicht abgeschlossen"); }
    finally { this.busy = false; if (this.active()) this.render(); }
  }
  private async load() {
    const epoch = ++this.epoch;
    try {
      if (this.host.connection.canAdmin) {
        try { await this.request("automatic-migration"); }
        catch { this.notice = "Automatische Einbenutzer-Übernahme derzeit nicht verfügbar. Die ausdrückliche persönliche Einrichtung bleibt verfügbar; es gibt keinen globalen Rückfall."; }
      }
      const status = await this.request<Status>("status");
      if (!this.active() || epoch !== this.epoch) return;
      this.status = status; this.policy = status.account ? clone(status.account.policy) : undefined;
      this.shared = status.account?.shared ?? false; this.unattended = status.account?.unattended ?? false;
      await this.refreshSession();
    } catch (error) { if (this.active() && epoch === this.epoch) this.error = (error as { message?: string })?.message ?? "Persönliches Konto konnte nicht geladen werden"; }
    if (this.active()) this.render();
  }
  private target() {
    const row = this.host.sessions.rows.find(s => s.key === this.host.sessions.selectedKey);
    const agentId = row?.agentId ?? this.host.agents.selectedId;
    return row?.sessionId && agentId ? { agentId, sessionKey: row.key, sessionId: row.sessionId } : undefined;
  }
  private async refreshSession() {
    const target = this.target(); const selected = this.host.sessions.selectedKey;
    if (!target) { this.session = undefined; return; }
    try { const result = await this.request<SessionStatus>("session-status", target); if (this.active() && selected === this.host.sessions.selectedKey) this.session = result; }
    catch { if (this.active() && selected === this.host.sessions.selectedKey) this.session = undefined; }
    if (this.active()) this.render();
  }
  private check(parent: HTMLElement, label: string, value: boolean, set: (v: boolean) => void) {
    const wrapper = node("label"); wrapper.className = "mg-check"; const input = node("input"); input.type = "checkbox"; input.checked = value; input.disabled = this.busy; input.addEventListener("change", () => set(input.checked)); wrapper.append(input, node("span", label)); parent.append(wrapper);
  }
  private field(parent: HTMLElement, label: string, value: string, set: (v: string) => void) {
    const wrapper = node("label", label); wrapper.className = "mg-field"; const input = node("input"); input.value = value; input.autocomplete = "off"; input.addEventListener("input", () => set(input.value.trim())); wrapper.append(input); parent.append(wrapper);
  }
  private async save() {
    if (!this.status?.account || !this.policy) return;
    const expectedRevision = this.status.account.revision;
    await this.request("configure", { expectedRevision, policy: this.policy, shared: this.shared, unattended: this.unattended });
    this.notice = "Kontorechte gespeichert. Sie gelten unmittelbar für neue Aktionen.";
    await this.load();
  }
  private async signIn() {
    const started = await this.request<NonNullable<AccountPage["device"]>>("device-start", { clientId: this.clientId, tenant: this.tenant });
    this.device = started;
    const poll = async () => {
      if (!this.active() || !this.device) return;
      try {
        const result = await this.request<{ state: string; error?: string }>("device-status", { sessionId: started.sessionId });
        if (!this.active()) return;
        if (result.state === "created") { this.device = undefined; this.notice = "Persönliche Microsoft-Verbindung hergestellt."; await this.load(); return; }
        if (result.state === "failed") { this.device = undefined; this.error = result.error ?? "Anmeldung fehlgeschlagen"; this.render(); return; }
        if (Date.now() >= Date.parse(started.expiresAt)) { this.device = undefined; this.error = "Anmeldung abgelaufen"; this.render(); return; }
        this.poll = setTimeout(() => { void poll(); }, 3000);
      } catch (error) { this.error = (error as { message?: string })?.message ?? "Anmeldestatus unbekannt"; this.render(); }
    };
    this.poll = setTimeout(() => { void poll(); }, 3000);
  }
  render() {
    if (!this.active()) return;
    const main = node("main"); main.className = "mg-page"; main.append(node("h1", "Mein Microsoft-Konto"));
    main.append(node("p", "Deine persönliche Verbindung, ihre Agentrechte und das Konto der ausgewählten Session werden getrennt verwaltet. Bei einer nachgewiesenen Einbenutzer-Installation übernimmt der Assistent die bestehende Verbindung automatisch. Mehrbenutzer-Nutzung benötigt ausdrückliche Freigaben; private Ausgabegarantien für gemeinsame Sessions sind hier nicht nachgewiesen."));
    if (!this.host.connection.connected) { main.append(node("p", "Mit OpenClaw verbinden, um dein Konto zu verwalten.")); this.container.replaceChildren(main); return; }
    if (!this.status) { main.append(node("p", this.error || "Persönliches Konto wird geladen …")); this.container.replaceChildren(main); return; }
    const automatic = this.status.automatic?.profileId === this.status.profileId && !this.status.automatic.invalidated;
    main.append(node("p", automatic ? "Automatische Einbenutzer-Übernahme aktiv: Bestehende Agentrechte gelten für Sessions, Jobs und normale Delegation ohne zusätzliche Kontobindung. Das vollständige Benutzerverzeichnis wird vor neuen Vorgängen und an sensiblen asynchronen Grenzen erneut geprüft. Bereits versendete Microsoft-Anfragen können nicht zurückgerufen werden." : this.status.warning));
    if (this.status.automatic?.invalidated) main.append(node("p", "Automatischer Zugriff bleibt nach einer Benutzerverzeichnis- oder Kontorechteänderung ausgeschaltet. Nutze jetzt ausdrückliche Freigaben und feste Kontobindungen."));
    if (this.error) { const e = node("p", this.error); e.className = "mg-error"; main.append(e); }
    if (this.notice) { const e = node("p", this.notice); e.className = "mg-success"; main.append(e); }
    const account = this.status.account;
    if (!account) {
      main.append(button("Persönliches Konto einrichten", () => { void this.task(async () => { await this.request("provision"); await this.load(); }); }));
      if (this.host.connection.canAdmin) main.append(button("Bestehende Verbindung meinem Profil zuordnen", () => {
        if (window.confirm("Die bestehende Verbindung wird ausdrücklich deinem authentifizierten Profil zugeordnet. Gemeinsame und unbeaufsichtigte Nutzung bleiben ausgeschaltet. Fortfahren?")) void this.task(async () => { await this.request("migrate-legacy"); await this.load(); });
      }));
      this.container.replaceChildren(main); return;
    }
    main.append(node("p", `Verbindung: ${account.connected ? this.status.credential.result === "valid" ? "verbunden" : this.status.credential.result : "nicht verbunden"}`));
    if (!account.connected && !this.device) {
      this.field(main, "Microsoft-App-ID", this.clientId, v => this.clientId = v);
      this.field(main, "Mandant", this.tenant, v => this.tenant = v);
      main.append(button("Bei Microsoft anmelden", () => { void this.task(() => this.signIn()); }));
      main.append(node("small", "Zuerst Agentrechte auswählen und speichern; die Anmeldung fordert nur die benötigten Microsoft-Berechtigungen an."));
    }
    if (this.device) {
      main.append(node("p", `Microsoft-Anmeldecode: ${this.device.userCode}`));
      const link = node("a", "Microsoft-Anmeldung öffnen"); link.href = this.device.verificationUri; link.target = "_blank"; link.rel = "noopener noreferrer"; main.append(link);
      main.append(button("Anmeldung abbrechen", () => { void this.task(async () => { await this.request("device-cancel", { sessionId: this.device!.sessionId }); this.device = undefined; if (this.poll) clearTimeout(this.poll); }); }));
    }
    if (account.connected) main.append(button("Trennen / neu verbinden", () => {
      if (window.confirm("Deine Verbindung trennen? Alle laufenden Bindungen und Freigaben werden ungültig. Bereits ausgegebene Daten bleiben sichtbar.")) void this.task(async () => { await this.request("disconnect", { expectedRevision: account.revision }); this.device = undefined; await this.load(); });
    }));
    const permissions = node("section"); permissions.className = "mg-body"; permissions.append(node("h2", "Agentrechte für dieses Konto"));
    const policy = this.policy!;
    for (const service of ["calendar", "mail", "todo"] as const) {
      const group = node("fieldset"); group.append(node("legend", serviceLabels[service]));
      for (const [agent, ceiling] of Object.entries(this.status.ceiling.services[service].agents)) {
        const name = this.host.agents.rows.find(a => a.id === agent)?.name ?? agent;
        group.append(node("h3", name));
        for (const op of ceiling.operations) this.check(group, actionLabels[op] ?? op, policy.services[service].agents[agent]?.operations.includes(op) ?? false, enabled => {
          const previous = policy.services[service].agents[agent]?.operations ?? [];
          const ops = enabled ? [...new Set([...previous, op])] : previous.filter(x => x !== op);
          if (ops.length) policy.services[service].agents[agent] = { operations: ops, ...(ceiling.resources ? { resources: clone(ceiling.resources) } : {}) }; else delete policy.services[service].agents[agent];
        });
      }
      permissions.append(group);
    }
    const roots = node("fieldset"); roots.append(node("legend", "OneDrive: administrativ zugelassene Ordner"));
    for (const ceiling of this.status.ceiling.services.onedrive.allowed_roots) {
      roots.append(node("h3", ceiling.path));
      for (const [agent, grant] of Object.entries(ceiling.agents)) for (const op of ["read", "write", "delete"] as const) if (ceiling.permissions[op] && grant.permissions[op]) {
        this.check(roots, `${agent}: ${actionLabels[op]}`, policy.services.onedrive.allowed_roots.find(r => r.label === ceiling.label)?.agents[agent]?.permissions[op] ?? false, enabled => {
          let root = policy.services.onedrive.allowed_roots.find(r => r.label === ceiling.label);
          if (!root) { root = { ...clone(ceiling), agents: {} }; policy.services.onedrive.allowed_roots.push(root); }
          root.agents[agent] ??= { permissions: {} }; root.agents[agent].permissions[op] = enabled;
        });
      }
    }
    permissions.append(roots);
    if (automatic) permissions.append(node("p", "Das Speichern von Kontorechten beendet den automatischen Modus dauerhaft; anschließend gelten die ausdrücklichen Freigaben und Bindungen."));
    this.check(permissions, "Gemeinsame Nutzung dieses Kontos ausdrücklich freigeben", this.shared, v => this.shared = v);
    this.check(permissions, "Delegierte / unbeaufsichtigte Nutzung durch fest gebundene Worker freigeben", this.unattended, v => this.unattended = v);
    permissions.append(node("p", this.status.workers.length ? `Fest gebundene Worker: ${this.status.workers.join(", ")}` : "Keine fest gebundenen Worker konfiguriert."));
    permissions.append(button("Kontorechte speichern", () => { void this.task(() => this.save()); })); main.append(permissions);
    const context = node("section"); context.append(node("h2", "Ausführungskonto der ausgewählten Session"));
    if (this.session?.workerProfileId) context.append(node("p", `Festes Worker-Konto: ${this.session.workerProfileId}. Ein Kontowechsel ist nicht möglich.`));
    else if (this.session) {
      context.append(node("p", this.session.binding ? `Aktives Konto: ${this.session.binding.profileId}` : automatic ? "Dein automatisch übernommenes Konto wird verwendet; keine Session- oder Worker-Konfiguration nötig." : "Noch kein Konto gebunden. Es gibt keinen automatischen Rückfall auf ein anderes Konto."));
      const select = node("select"); select.setAttribute("aria-label", "Ausführungskonto auswählen");
      for (const id of this.session.available) { const option = node("option", id === this.status.profileId ? "Mein Konto" : id); option.value = id; select.append(option); }
      if (this.session.binding) select.value = this.session.binding.profileId;
      context.append(select, button("Dieses Konto für die Session verwenden", () => { const target = this.target(); if (!target) return; void this.task(async () => { await this.request("bind", { ...target, profileId: select.value, expectedRevision: this.session?.binding?.revision ?? null }); await this.refreshSession(); }); }));
      context.append(node("small", "Ein anderer Teilnehmer ändert das Konto nicht. Wechsel sind nur zwischen Aufträgen und ohne ausstehende Freigaben möglich."));
    } else context.append(node("p", "Eine bedienbare Session auswählen, um ihr ein Konto zuzuordnen."));
    main.append(context);
    if (this.busy) for (const c of main.querySelectorAll<HTMLButtonElement | HTMLInputElement | HTMLSelectElement>("button,input,select")) c.disabled = true;
    this.container.replaceChildren(main);
  }
}
