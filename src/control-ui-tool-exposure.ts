type Permission = "read" | "write" | "delete";
type Root = {
  permissions: Partial<Record<Permission, boolean>>;
  agents: Record<string, { permissions: Partial<Record<Permission, boolean>> }>;
};
type AgentTools = { allow?: string[]; alsoAllow?: string[]; deny?: string[] };
type AgentEntry = { tools?: AgentTools };

const toolsByPermission: Record<Permission, readonly string[]> = {
  read: ["onedrive_search", "onedrive_list", "onedrive_read", "onedrive_download"],
  write: ["onedrive_upload", "onedrive_update", "onedrive_metadata_update", "onedrive_create_folder"],
  delete: ["onedrive_delete"],
};

function matches(pattern: string, tool: string): boolean {
  pattern = pattern.toLowerCase(); tool = tool.toLowerCase();
  if (pattern === tool || pattern === "*") return true;
  if (!pattern.includes("*")) return false;
  const escaped = pattern.replace(/[.+?^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*");
  return new RegExp(`^${escaped}$`).test(tool);
}

export function requiredOneDriveTools(roots: Root[]): Record<string, string[]> {
  const required = new Map<string, Set<string>>();
  for (const root of roots) for (const [agentId, grant] of Object.entries(root.agents)) {
    for (const permission of ["read", "write", "delete"] as const) {
      if (!root.permissions[permission] || !grant.permissions[permission]) continue;
      const tools = required.get(agentId) ?? new Set<string>();
      for (const name of toolsByPermission[permission]) tools.add(name);
      required.set(agentId, tools);
    }
  }
  return Object.fromEntries([...required].map(([id, names]) => [id, [...names]]));
}

export function planOneDriveToolExposure(roots: Root[], entries: Record<string, AgentEntry> | undefined, globalDeny: string[] = []) {
  return planToolExposure(requiredOneDriveTools(roots), entries, globalDeny);
}

export function planToolExposure(required: Record<string, string[]>, entries: Record<string, AgentEntry> | undefined, globalDeny: string[] = []) {
  const patch: Record<string, { tools: Partial<AgentTools> }> = {};
  const expected: Record<string, { key: "allow" | "alsoAllow"; value: string[] }> = {};
  for (const [agentId, names] of Object.entries(required)) {
    if (agentId.includes(".")) throw new Error(`unsupported_agent_id:${agentId}`);
    if (!entries || !Object.hasOwn(entries, agentId)) throw new Error(`missing_agent_entry:${agentId}`);
    const tools = entries[agentId]?.tools ?? {};
    if (tools.allow && tools.alsoAllow) throw new Error(`conflicting_tool_policy:${agentId}`);
    for (const name of names) if ([...globalDeny, ...(tools.deny ?? [])].some(pattern => ["group:plugins", "microsoft-graph"].includes(pattern.toLowerCase()) || matches(pattern, name))) throw new Error(`denied_tool:${agentId}:${name}`);
    const key = tools.allow ? "allow" : "alsoAllow";
    const before = tools[key] ?? [];
    const missing = names.filter(name => !before.some(pattern => matches(pattern, name)));
    if (!missing.length) continue;
    const value = [...before, ...missing];
    patch[agentId] = { tools: { [key]: value } };
    expected[agentId] = { key, value };
  }
  return { patch, expected };
}

export function oneDriveToolExposureSaved(entries: Record<string, AgentEntry> | undefined, expected: Record<string, { key: "allow" | "alsoAllow"; value: string[] }>): boolean {
  return Object.entries(expected).every(([id, { key, value }]) => JSON.stringify(entries?.[id]?.tools?.[key]) === JSON.stringify(value));
}


type AccessPolicy = { services: { onedrive: { allowed_roots: Root[] }; calendar: { agents: Record<string, { operations: string[]; resources?: string[] }> }; mail: { agents: Record<string, { operations: string[]; resources?: string[] }> }; contacts?: { agents: Record<string, { operations: string[]; resources?: string[] }> }; todo: { agents: Record<string, { operations: string[]; resources?: string[] }> } } };
/** Only the service/agent grants edited by this draft may opt tools in. */
export function planChangedToolExposure(before: AccessPolicy, after: AccessPolicy, entries: Record<string, AgentEntry> | undefined, globalDeny: string[] = []) {
  const required: Record<string, string[]> = {};
  for (const [agent, names] of Object.entries(requiredOneDriveTools(after.services.onedrive.allowed_roots))) {
    const footprint = (policy: AccessPolicy) => policy.services.onedrive.allowed_roots.filter(root => root.agents[agent]).map(({ agents, ...root }) => ({ ...root, grant: agents[agent] }));
    if (JSON.stringify(footprint(before)) !== JSON.stringify(footprint(after))) required[agent] = names;
  }
  const prefixes = { calendar: "outlook_calendar", mail: "outlook_mail", contacts: "outlook_contacts", todo: "microsoft_todo" };
  for (const service of ["calendar", "mail", "contacts", "todo"] as const) {
    for (const [agent, grant] of Object.entries(after.services[service]?.agents ?? {})) {
      if (JSON.stringify(grant) === JSON.stringify(before.services[service]?.agents[agent])) continue;
      const names = required[agent] ?? [];
      if (grant.operations.includes("read")) names.push(`${prefixes[service]}_read`);
      if (grant.operations.some(operation => operation !== "read")) names.push(`${prefixes[service]}_write`);
      if (names.length) required[agent] = [...new Set(names)];
    }
  }
  return planToolExposure(required, entries, globalDeny);
}

/** Explicit recovery only for the selected saved service/agent; never backfill peers. */
export function planSelectedServiceToolExposure(policy: AccessPolicy, service: "calendar" | "mail" | "contacts" | "todo", agent: string, entries: Record<string, AgentEntry> | undefined, globalDeny: string[] = []) {
  const before = structuredClone(policy);
  if (before.services[service]) delete before.services[service]!.agents[agent];
  return planChangedToolExposure(before, policy, entries, globalDeny);
}
