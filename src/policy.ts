import { Type } from "typebox";
import { Compile } from "typebox/compile";

export type Service = "onedrive" | "calendar" | "mail" | "todo";
export type OneDriveOperation = "read" | "write" | "delete";
export type AllowedRoot = {
  label: string; path: string; drive_id: string; item_id: string; include_descendants: boolean;
  agents_instructions?: "trusted";
  permissions: Record<OneDriveOperation, boolean>;
  agents: Record<string, { permissions: Partial<Record<OneDriveOperation, boolean>> }>;
};
const LABEL = /^[a-z0-9][a-z0-9_-]{0,159}$/;
const OPERATIONS: Record<Service, Set<string>> = {
  onedrive: new Set(["read", "write", "delete"]),
  calendar: new Set(["read", "create", "update", "respond", "attach", "delete"]),
  mail: new Set(["read", "draft", "update", "move", "mark", "send", "delete"]),
  todo: new Set(["read", "create", "update", "delete"]),
};

type AgentGrant = { operations: string[]; resources?: string[] };
type GraphPolicyServices = {
  onedrive: { allowed_roots: AllowedRoot[] };
  calendar: { agents: Record<string, AgentGrant> };
  mail: { agents: Record<string, AgentGrant> };
  todo: { agents: Record<string, AgentGrant> };
};
type GraphPolicyBase = {
  rules: { default: "deny" };
  services: GraphPolicyServices;
};
export type GraphPolicy = GraphPolicyBase & { version: 2 };

const AgentId = Type.String({ minLength: 1 });
const Resource = Type.String({ minLength: 1 });
const agentGrant = (operations: readonly string[]) => Type.Object({
  operations: Type.Array(Type.String({ enum: [...operations] }), { minItems: 1 }),
  resources: Type.Optional(Type.Array(Resource, { minItems: 1 })),
}, { additionalProperties: false });
const agents = (operations: readonly string[]) => Type.Record(AgentId, agentGrant(operations));

const GraphPolicyServicesSchema = Type.Object({
  onedrive: Type.Object({
    allowed_roots: Type.Array(Type.Object({
      label: Type.String({ pattern: LABEL.source }),
      path: Type.String({ minLength: 1 }),
      drive_id: Type.String({ minLength: 1 }),
      item_id: Type.String({ minLength: 1 }),
      include_descendants: Type.Literal(true),
      agents_instructions: Type.Optional(Type.Literal("trusted")),
      permissions: Type.Object({ read: Type.Boolean(), write: Type.Boolean(), delete: Type.Boolean() }, { additionalProperties: false }),
      agents: Type.Record(AgentId, Type.Object({
        permissions: Type.Object({
          read: Type.Optional(Type.Boolean()),
          write: Type.Optional(Type.Boolean()),
          delete: Type.Optional(Type.Boolean()),
        }, { additionalProperties: false }),
      }, { additionalProperties: false })),
    }, { additionalProperties: false })),
  }, { additionalProperties: false }),
  calendar: Type.Object({ agents: agents(["read", "create", "update", "respond", "attach", "delete"]) }, { additionalProperties: false }),
  mail: Type.Object({ agents: agents(["read", "draft", "update", "move", "mark", "send", "delete"]) }, { additionalProperties: false }),
  todo: Type.Object({ agents: agents(["read", "create", "update", "delete"]) }, { additionalProperties: false }),
}, { additionalProperties: false });

export const GraphPolicySchema = Type.Object({
  version: Type.Literal(2),
  rules: Type.Object({ default: Type.Literal("deny") }, { additionalProperties: false }),
  services: GraphPolicyServicesSchema,
}, { additionalProperties: false });

const policyValidator = Compile(GraphPolicySchema);

function record(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === "object" && !Array.isArray(value);
}

export function validatePolicy(value: unknown): GraphPolicy {
  if (!policyValidator.Check(value)) throw new Error("invalid_policy");
  const services = value.services;
  const labels = new Set<string>();
  for (const root of services.onedrive.allowed_roots) {
    if (!record(root) || typeof root.label !== "string" || !LABEL.test(root.label) || labels.has(root.label) || typeof root.path !== "string" || !root.path.startsWith("/") || root.path.includes("\\") || root.path.split("/").includes("..") || typeof root.drive_id !== "string" || !root.drive_id || typeof root.item_id !== "string" || !root.item_id || root.include_descendants !== true || (root.agents_instructions !== undefined && root.agents_instructions !== "trusted") || !record(root.permissions) || !record(root.agents)) throw new Error("invalid_policy");
    labels.add(root.label);
    for (const op of ["read", "write", "delete"] as const) if (typeof root.permissions[op] !== "boolean") throw new Error("invalid_policy");
    for (const grant of Object.values(root.agents)) {
      if (!record(grant) || !record(grant.permissions)) throw new Error("invalid_policy");
      for (const [op, allowed] of Object.entries(grant.permissions)) if (!OPERATIONS.onedrive.has(op) || typeof allowed !== "boolean") throw new Error("invalid_policy");
    }
  }
  for (const service of ["calendar", "mail", "todo"] as const) {
    if (!record(services[service].agents)) throw new Error("invalid_policy");
    for (const [agentId, grant] of Object.entries(services[service].agents)) {
      if (!agentId || !record(grant) || !Array.isArray(grant.operations) || grant.operations.some((op) => typeof op !== "string" || !OPERATIONS[service].has(op))) throw new Error("invalid_policy");
      if (grant.resources !== undefined && (!Array.isArray(grant.resources) || grant.resources.some((resource) => typeof resource !== "string" || !resource))) throw new Error("invalid_policy");
    }
  }
  return value as GraphPolicy;
}

export function authorizeOperation(policy: GraphPolicy, agentId: string | undefined, service: Exclude<Service, "onedrive">, operation: string, resource = "me"): void {
  if (!agentId) throw new Error("trusted_agent_identity_required");
  const grant = policy.services[service].agents[agentId];
  if (!grant?.operations.includes(operation)) throw new Error("access_denied");
  const resources = grant.resources ?? ["me"];
  if (!resources.includes(resource)) throw new Error("access_denied");
}

export function authorizeRoot(policy: GraphPolicy, agentId: string | undefined, rootLabel: string, operation: OneDriveOperation): AllowedRoot {
  if (!agentId) throw new Error("trusted_agent_identity_required");
  const root = policy.services.onedrive.allowed_roots.find((candidate) => candidate.label === rootLabel);
  if (!root || root.permissions[operation] !== true || root.agents[agentId]?.permissions?.[operation] !== true) throw new Error("access_denied");
  return root;
}

export function normalizeRelativePath(value = ""): string {
  if (typeof value !== "string" || value.includes("\\") || value.startsWith("/") || /[\u0000-\u001f\u007f]/.test(value)) throw new Error("invalid_relative_path");
  const parts: string[] = [];
  for (const raw of value.split("/")) {
    const part = raw.trim();
    if (!part || part === ".") continue;
    if (part === "..") throw new Error("invalid_relative_path");
    parts.push(part);
  }
  return parts.join("/");
}

export function joinRootPath(rootPath: string, relativePath = ""): string {
  const root = `/${rootPath.split("/").filter(Boolean).join("/")}`;
  const relative = normalizeRelativePath(relativePath);
  return relative ? `${root}/${relative}` : root;
}
