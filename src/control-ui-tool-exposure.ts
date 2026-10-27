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
  const patch: Record<string, { tools: Partial<AgentTools> }> = {};
  const expected: Record<string, { key: "allow" | "alsoAllow"; value: string[] }> = {};
  const required = requiredOneDriveTools(roots);
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
