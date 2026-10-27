import { ActivationJournal } from "./config-activation.js";
import { createHash } from "node:crypto";
import { canonicalPolicy } from "./config-policy-identity.js";
import { requiredPolicyScopes } from "./credential-cli.js";
import { validatePolicy } from "./policy.js";
import { tokenForAuthorizedOperation } from "./credential.js";
import { canonicalGraphContinuation, graphRequest, safeId } from "./graph.js";

type ActivationClient = { authenticatedUserId?: string; connect?: { device?: { id?: string } } };
type HandlerContext = { client?: ActivationClient | null; context?: { configRevisionProjector?: { projectRawHash(rawHash: string): string } }; params: unknown; respond(ok: boolean, payload?: unknown, error?: unknown): void };
type GatewayApi = { version?: string; registerGatewayMethod(method: string, handler: (context: HandlerContext) => void | Promise<void>, options: { scope: "operator.admin" }): void };

/** Only host-authenticated handshake fields, never RPC parameters or a client class name.
 * Gateway verifies the device public-key digest and nonce-bound signature before dispatch.
 */
export function activationOwner(client?: ActivationClient | null): string {
  if (typeof client?.authenticatedUserId === "string" && client.authenticatedUserId.trim()) return `user:${client.authenticatedUserId}`;
  const deviceId = client?.connect?.device?.id;
  return typeof deviceId === "string" && /^[a-f0-9]{64}$/.test(deviceId) ? `device:${deviceId}` : "";
}

/** Admin-only validation and bounded activation intent. Never returns credential material. */
export function registerConfigurationUiMethods(api: GatewayApi, config?: { policy?: unknown; credentialVaultKey?: unknown; requestTimeoutMs?: number }, stateDir?: () => string): void {
  // Capture this generation, not a later mutated config object. No policy or
  // credential values leave this method, and no provider access is performed.
  const activeVersion = api.version ?? null;
  const digest = (value: unknown) => createHash("sha256").update(canonicalPolicy(validatePolicy(value))).digest("hex");
  let activePolicyHash: string | undefined;
  try { if (config?.policy) activePolicyHash = digest(config.policy); } catch { /* Invalid/missing active policy stays unconfirmed. */ }
  api.registerGatewayMethod("microsoft-graph.configuration.activation", async ({ params, respond, client, context }) => {
    try {
      if (!stateDir || !params || typeof params !== "object" || Array.isArray(params) || JSON.stringify(params).length > 512_000) throw Error("invalid_activation_parameters");
      const p = params as Record<string, unknown>;
      const owner = activationOwner(client);
      // 2026.10.8 projects public CAS revisions; older hosts use raw hashes.
      const projector = context?.configRevisionProjector;
      const journal = new ActivationJournal(stateDir(), undefined, owner, projector ? hash => projector.projectRawHash(hash) : undefined);
      let result: unknown;
      if (p.action === "status" && Object.keys(p).length === 1) result = await journal.status();
      else if (p.action === "prepare" && Object.keys(p).length === 4 && typeof p.baseHash === "string" && "baseline" in p && "target" in p) result = await journal.prepare(p.baseHash, p.baseline, p.target);
      else if (p.action === "claim" && Object.keys(p).length === 2 && typeof p.id === "string") result = await journal.claim(p.id, activePolicyHash);
      else if (p.action === "cancel" && Object.keys(p).length === 2 && typeof p.id === "string") result = await journal.cancel(p.id);
      else if (p.action === "complete" && Object.keys(p).length === 2 && typeof p.id === "string") result = await journal.complete(p.id, activePolicyHash);
      else throw Error("invalid_activation_parameters");
      respond(true, result);
    } catch { respond(false, undefined, { code: "INVALID_REQUEST", message: "Activation intent unavailable or changed; reload and review the saved access before continuing" }); }
  }, { scope: "operator.admin" });
  api.registerGatewayMethod("microsoft-graph.configuration.applicationStatus", ({ params, respond }) => {
    if (!params || typeof params !== "object" || Array.isArray(params) || Object.keys(params).length) {
      respond(false, undefined, { code: "INVALID_REQUEST", message: "Invalid application status parameters" }); return;
    }
    respond(true, { activePolicyHash: activePolicyHash ?? null, activeVersion });
  }, { scope: "operator.admin" });
  api.registerGatewayMethod("microsoft-graph.configuration.validate", ({ params, respond }) => {
    try {
      if (!params || typeof params !== "object" || Array.isArray(params)
        || Object.keys(params).length !== 1 || !("policy" in params)) throw new Error("invalid_rpc_parameters");
      const serialized = JSON.stringify((params as { policy: unknown }).policy);
      if (serialized.length > 256_000) throw new Error("invalid_rpc_parameters");
      const policy = validatePolicy((params as { policy: unknown }).policy);
      respond(true, { valid: true, requiredScopes: requiredPolicyScopes(policy), policyHash: digest(policy) });
    } catch {
      respond(false, undefined, { code: "INVALID_REQUEST", message: "Invalid Microsoft Graph policy" });
    }
  }, { scope: "operator.admin" });
  api.registerGatewayMethod("microsoft-graph.configuration.resolveFolder", async ({ params, respond }) => {
    try {
      if (!config?.policy || !stateDir || !params || typeof params !== "object" || Array.isArray(params) || Object.keys(params).length !== 1) throw new Error("invalid_rpc_parameters");
      const raw = (params as { path?: unknown }).path;
      if (typeof raw !== "string" || raw.length > 1024 || !raw.startsWith("/") || raw.includes("\\") || /[\u0000-\u001f\u007f]/.test(raw)) throw new Error("invalid_path");
      const segments = raw.split("/").filter(Boolean);
      if (segments.some(segment => segment === "." || segment === ".." || segment.length > 255)) throw new Error("invalid_path");
      const path = "/" + segments.join("/");
      const policy = validatePolicy(config.policy);
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 15000);
      try {
        const token = await tokenForAuthorizedOperation({ config, policy, allowedScopes: ["Files.Read"], stateDir: stateDir(), signal: controller.signal, requestTimeoutMs: config.requestTimeoutMs ?? 5000 });
        const selector = segments.length ? `/me/drive/root:/${segments.map(encodeURIComponent).join("/")}:` : "/me/drive/root";
        const item = await graphRequest(token, `${selector}?$select=id,name,folder,parentReference`, { signal: controller.signal, maxBytes: 16_384 });
        if (typeof item?.id !== "string" || !item.id || !item.folder || typeof item?.parentReference?.driveId !== "string" || !item.parentReference.driveId) throw new Error("not_folder");
        respond(true, { path, drive_id: item.parentReference.driveId, item_id: item.id });
      } finally { clearTimeout(timer); }
    } catch { respond(false, undefined, { code: "INVALID_REQUEST", message: "OneDrive folder could not be verified" }); }
  }, { scope: "operator.admin" });
  api.registerGatewayMethod("microsoft-graph.configuration.calendars", async ({ params, respond }) => {
    try {
      if (!config?.policy || !stateDir || !params || typeof params !== "object" || Array.isArray(params) || Object.keys(params).length) throw new Error("invalid_rpc_parameters");
      const policy = validatePolicy(config.policy);
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 15000);
      try {
        const token = await tokenForAuthorizedOperation({ config, policy, allowedScopes: ["Calendars.Read"], stateDir: stateDir(), signal: controller.signal, requestTimeoutMs: config.requestTimeoutMs ?? 5000 });
        const primary = await graphRequest(token, "/me/calendar?$select=id,name", { signal: controller.signal, maxBytes: 16_384 });
        if (typeof primary?.id !== "string" || !primary.id || typeof primary?.name !== "string") throw new Error("invalid_provider_response");
        const calendars: Array<{ id: string; name: string }> = [];
        const seen = new Set<string>();
        let path: string | undefined = "/me/calendars?$select=id,name&$top=100";
        for (let page = 0; path && page < 10; page++) {
          const data = await graphRequest(token, path, { signal: controller.signal, maxBytes: 128_000 });
          if (!Array.isArray(data?.value)) throw new Error("invalid_provider_response");
          for (const item of data.value) {
            if (typeof item?.id !== "string" || typeof item?.name !== "string" || item.name.length > 512) throw new Error("invalid_provider_response");
            try { safeId(item.id); } catch { throw new Error("invalid_provider_response"); }
            if (!seen.has(item.id)) { seen.add(item.id); calendars.push({ id: item.id, name: item.name }); }
            if (calendars.length > 500) throw new Error("calendar_list_too_large");
          }
          const next = data["@odata.nextLink"];
          path = next === undefined ? undefined : canonicalGraphContinuation(next, "/v1.0/me/calendars", "invalid_provider_response").slice("/v1.0".length);
          if (page === 9 && path) throw new Error("calendar_list_too_large");
        }
        respond(true, { defaultCalendarId: primary.id, calendars });
      } finally { clearTimeout(timer); }
    } catch { respond(false, undefined, { code: "INVALID_REQUEST", message: "Calendar list could not be loaded" }); }
  }, { scope: "operator.admin" });
}
