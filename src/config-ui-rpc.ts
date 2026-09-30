import { requiredPolicyScopes } from "./credential-cli.js";
import { validatePolicy } from "./policy.js";
import { tokenForAuthorizedOperation } from "./credential.js";
import { graphRequest } from "./graph.js";

type HandlerContext = { params: unknown; respond(ok: boolean, payload?: unknown, error?: unknown): void };
type GatewayApi = { registerGatewayMethod(method: string, handler: (context: HandlerContext) => void | Promise<void>, options: { scope: "operator.admin" }): void };

/** Admin-only, read-only validation. Never resolves or returns credential material. */
export function registerConfigurationUiMethods(api: GatewayApi, config?: { policy?: unknown; credentialVaultKey?: unknown; requestTimeoutMs?: number }, stateDir?: () => string): void {
  api.registerGatewayMethod("microsoft-graph.configuration.validate", ({ params, respond }) => {
    try {
      if (!params || typeof params !== "object" || Array.isArray(params)
        || Object.keys(params).length !== 1 || !("policy" in params)) throw new Error("invalid_rpc_parameters");
      const serialized = JSON.stringify((params as { policy: unknown }).policy);
      if (serialized.length > 256_000) throw new Error("invalid_rpc_parameters");
      const policy = validatePolicy((params as { policy: unknown }).policy);
      respond(true, { valid: true, requiredScopes: requiredPolicyScopes(policy) });
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
}
