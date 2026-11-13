import { createHash } from "node:crypto";
import { readVaultCredential } from "./credential-vault.js";
import { tokenForAuthorizedOperation } from "./credential.js";
import type { GraphPolicy } from "./policy.js";

export type Service = "mail" | "calendar" | "contacts" | "todo" | "onedrive" | "unknown";
export type Health = { status: "unavailable" | "not_authorized" | "unsupported" | "checked"; reason?: string; evidence: "observed_request_failure" | "matching_microsoft_incident"; complete?: boolean; incidents?: Array<{ id: string; status: string; startDateTime: string; lastModifiedDateTime: string }> };
const cache = new Map<string, { until: number; value: Health }>();
const flights = new Map<string, Promise<Health>>();
const unavailable = (reason: string): Health => ({ status: "unavailable", reason, evidence: "observed_request_failure" });
const noGrant = (): Health => ({ status: "not_authorized", reason: "ServiceHealth.Read.All_admin_consent_required", evidence: "observed_request_failure" });
const serviceNames: Record<Service, string[]> = { mail: ["Exchange Online"], calendar: ["Exchange Online"], contacts: ["Exchange Online"], todo: ["Microsoft To-Do", "Microsoft To Do"], onedrive: ["OneDrive for Business"], unknown: [] };
export function matchingIncidents(value: unknown, service: Service, now = Date.now()): NonNullable<Health["incidents"]> {
  if (!Array.isArray(value)) return [];
  const active = new Set(["serviceDegradation", "serviceInterruption", "investigating", "restoringService", "extendedRecovery"]);
  return value.flatMap((item) => {
    if (!item || !serviceNames[service].includes(item.service) || !active.has(item.status) || item.isResolved !== false || item.origin !== "microsoft" || item.classification !== "incident" || item.endDateTime != null
      || typeof item.id !== "string" || !/^[A-Z]{2,8}[0-9]{1,20}$/.test(item.id)) return [];
    const start = Date.parse(item.startDateTime), updated = Date.parse(item.lastModifiedDateTime);
    if (!Number.isFinite(start) || !Number.isFinite(updated) || start > now || updated > now || updated < start || now - updated > 24 * 60 * 60_000) return [];
    // Exchange-wide incidents do not prove impact on a particular feature.
    const features: Record<Service, RegExp> = { mail: /^(mail|email|e-mail|outlook|exchange online)$/i, calendar: /^(calendar|calendaring)$/i, contacts: /^(contacts|people)$/i, todo: /^(tasks|to do|microsoft to do)$/i, onedrive: /^(onedrive|onedrive for business|files)$/i, unknown: /$a/ };
    if (typeof item.featureGroup !== "string" || !features[service].test(item.featureGroup)) return [];
    return [{ id: item.id, status: item.status, startDateTime: new Date(start).toISOString(), lastModifiedDateTime: new Date(updated).toISOString() }];
  }).slice(0, 10);
}
async function readStatus(token: string, service: Service, signal: AbortSignal, fetchFn: typeof fetch): Promise<Health> {
  let url = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues?$top=50";
  const incidents: NonNullable<Health["incidents"]> = [];
  for (let page = 0; page < 3; page++) {
    const response = await fetchFn(url, { headers: { authorization: "Bearer " + token, accept: "application/json" }, signal, redirect: "error" });
    if (response.status === 401 || response.status === 403) { await response.body?.cancel(); return { status: "not_authorized", reason: "status_endpoint_access_denied", evidence: "observed_request_failure" }; }
    if (!response.ok) { await response.body?.cancel(); return unavailable("status_endpoint_failed"); }
    if (!response.body) return unavailable("invalid_status_response");
    const reader = response.body.getReader(); let size = 0; const chunks: Uint8Array[] = [];
    try {
      while (true) { const part = await reader.read(); if (part.done) break; size += part.value.byteLength; if (size > 128 * 1024) { await reader.cancel(); return unavailable("status_response_too_large"); } chunks.push(part.value); }
    } finally { reader.releaseLock(); }
    const data = JSON.parse(Buffer.concat(chunks).toString("utf8"));
    if (!Array.isArray(data.value)) return unavailable("invalid_status_response");
    incidents.push(...matchingIncidents(data.value, service));
    const next = data["@odata.nextLink"];
    if (!next) return { status: "checked", evidence: incidents.length ? "matching_microsoft_incident" : "observed_request_failure", complete: true, incidents: incidents.slice(0, 10) };
    const parsed = new URL(next);
    if (parsed.origin !== "https://graph.microsoft.com" || parsed.pathname !== "/v1.0/admin/serviceAnnouncement/issues" || parsed.username || parsed.password || parsed.hash) return unavailable("invalid_status_continuation");
    url = parsed.href;
  }
  return { status: "checked", evidence: incidents.length ? "matching_microsoft_incident" : "observed_request_failure", complete: false, incidents: incidents.slice(0, 10) };
}
/** Additional scope is requested only when explicitly present in the operator-managed credential.
 * No interactive consent, scope widening, new tool, or automatic data-operation replay. */
export async function automaticServiceHealth(params: { config: { credentialVaultKey?: unknown }; policy: GraphPolicy; stateDir?: string; service: Service; signal?: AbortSignal; budgetMs?: number; fetchFn?: typeof fetch }): Promise<Health> {
  params.signal?.throwIfAborted();
  if (!params.stateDir) return unavailable("credential_unavailable");
  const stateDir = params.stateDir;
  if (params.service === "unknown") return { status: "unsupported", evidence: "observed_request_failure" };
  const budget = Math.max(1, Math.min(1500, params.budgetMs ?? 1500));
  const started = Date.now();
  let timer: ReturnType<typeof setTimeout> | undefined;
  let abort: (() => void) | undefined;
  const work = async (): Promise<Health> => {
    const record = await readVaultCredential(stateDir, params.config.credentialVaultKey);
    if (Date.now() - started >= budget) return unavailable("diagnostic_timeout");
    if (!record.credential.scopes.some(s => s.toLowerCase() === "servicehealth.read.all")) return noGrant();
    if (["consumers", "common"].includes(record.credential.tenant.toLowerCase())) return { status: "unsupported", reason: "organizational_tenant_required", evidence: "observed_request_failure" };
    // Vault binding isolates replacements/accounts without putting any token in a key.
    const key = createHash("sha256").update(JSON.stringify([params.stateDir, record.envelope.keyId, record.envelope.generation, record.digest, params.service])).digest("hex");
    const hit = cache.get(key); if (hit && hit.until > Date.now()) return hit.value;
    const current = flights.get(key); if (current) return current;
    const controller = new AbortController();
    if (flights.size >= 64) return unavailable("diagnostic_capacity");
    const task = (async () => {
      let expire: ReturnType<typeof setTimeout> | undefined;
      const timeout = new Promise<Health>(resolve => { expire = setTimeout(() => { controller.abort(); resolve(unavailable("diagnostic_timeout")); }, Math.max(1, budget - (Date.now() - started))); });
      try {
        // A rotating refresh must finish its durable vault transaction even after
        // the diagnostic waiter leaves. Use the normal 10s token deadline, never
        // the diagnostic/caller signal. Keep this flight/capacity slot until it
        // settles, and handle failures below without weakening vault quarantine.
        const token = await tokenForAuthorizedOperation({ config: params.config, policy: params.policy, allowedScopes: ["ServiceHealth.Read.All"], stateDir, fetchFn: params.fetchFn });
        if (controller.signal.aborted) return unavailable("diagnostic_timeout");
        return await Promise.race([timeout, readStatus(token, params.service, controller.signal, params.fetchFn ?? fetch)]);
      } catch { return unavailable("status_check_failed"); }
      finally { if (expire) clearTimeout(expire); }
    })();
    flights.set(key, task);
    void task.then(value => {
      if (cache.size >= 64) cache.delete(cache.keys().next().value!);
      cache.set(key, { until: Date.now() + 30_000, value }); flights.delete(key);
    });
    return task;
  };
  try {
    return await Promise.race([work().catch(() => unavailable("credential_unavailable")), new Promise<Health>((resolve, reject) => {
      timer = setTimeout(() => resolve(unavailable("diagnostic_timeout")), budget);
      abort = () => reject(params.signal!.reason);
      params.signal?.addEventListener("abort", abort, { once: true });
      if (params.signal?.aborted) abort();
    })]);
  } finally { if (timer) clearTimeout(timer); if (abort) params.signal?.removeEventListener("abort", abort); }
}
export function failureCategory(code: string, status?: number): string {
  if (status === 401 || code === "authentication_failed" || code.startsWith("credential_")) return "authentication";
  if (status === 403) return "authorization";
  if (code === "access_denied" || code.startsWith("approval_context_")) return "local_policy";
  if (status === 429 || code === "provider_throttled") return "throttled";
  if (status && status >= 500) return "transient_provider";
  if (code === "provider_unavailable") return "network";
  return "unknown";
}
