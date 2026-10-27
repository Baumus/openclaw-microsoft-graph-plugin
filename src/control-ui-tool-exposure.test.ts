import { describe, expect, it } from "vitest";
import { planPolicyToolReconciliation, planSelectedServiceToolExposure, planChangedToolExposure, oneDriveToolExposureSaved, planOneDriveToolExposure, requiredOneDriveTools } from "./control-ui-tool-exposure.js";

const root = (permissions: { read: boolean; write: boolean; delete: boolean }, agentPermissions = permissions) => [{ permissions, agents: { studio: { permissions: agentPermissions } } }];

describe("OneDrive Control UI tool exposure", () => {
  it("derives only operations allowed by both the root and agent", () => {
    expect(requiredOneDriveTools(root({ read: true, write: false, delete: false }, { read: true, write: true, delete: true }))).toEqual({ studio: ["onedrive_search", "onedrive_list", "onedrive_read", "onedrive_download"] });
  });

  it("adds only missing read tools and preserves unrelated grants", () => {
    const entries = { studio: { tools: { alsoAllow: ["outlook_mail_read", "onedrive_list"] } } };
    const result = planOneDriveToolExposure(root({ read: true, write: false, delete: false }), entries);
    expect(result.patch.studio.tools.alsoAllow).toEqual(["outlook_mail_read", "onedrive_list", "onedrive_search", "onedrive_read", "onedrive_download"]);
    expect(oneDriveToolExposureSaved(entries, result.expected)).toBe(false);
    expect(oneDriveToolExposureSaved({ studio: { tools: { alsoAllow: result.patch.studio.tools.alsoAllow } } }, result.expected)).toBe(true);
  });

  it("uses an existing allow list rather than creating incompatible alsoAllow", () => {
    const result = planOneDriveToolExposure(root({ read: true, write: false, delete: false }), { studio: { tools: { allow: ["outlook_mail_read"] } } });
    expect(result.patch.studio.tools.allow).toContain("onedrive_list");
    expect(result.patch.studio.tools.alsoAllow).toBeUndefined();
  });

  it("fails closed on missing agents or explicit deny rules", () => {
    const roots = root({ read: true, write: false, delete: false });
    expect(() => planOneDriveToolExposure(roots, {})).toThrow("missing_agent_entry");
    expect(() => planOneDriveToolExposure(roots, { studio: { tools: { deny: ["onedrive_*"] } } })).toThrow("denied_tool");
    expect(() => planOneDriveToolExposure(roots, { studio: {} }, ["group:plugins"])).toThrow("denied_tool");
  });

  it("selects write and delete tools only for matching grants", () => {
    const result = planOneDriveToolExposure(root({ read: false, write: true, delete: true }), { studio: {} });
    expect(result.patch.studio.tools.alsoAllow).toEqual(["onedrive_upload", "onedrive_update", "onedrive_metadata_update", "onedrive_create_folder", "onedrive_delete"]);
  });
});


describe("changed service tool exposure", () => {
  const baseline = () => ({ services: { onedrive: { allowed_roots: root({read:true,write:true,delete:true}) }, calendar: { agents: {} }, mail: { agents: {} }, todo: { agents: {} } } });
  it("adds exactly read-only calendar exposure and leaves unrelated OneDrive agents unchanged", () => {
    const before=baseline(); const after={...before,services:{...before.services,calendar:{agents:{reader:{operations:["read"],resources:["me"]}}}}};
    expect(planChangedToolExposure(before,after,{reader:{},studio:{}}).patch).toEqual({reader:{tools:{alsoAllow:["outlook_calendar_read"]}}});
    expect(() => planChangedToolExposure(before,after,{reader:{tools:{deny:["outlook_calendar_*"]}},studio:{}})).toThrow("denied_tool");
  });
  it("does not add write exposure when narrowing an existing grant to read", () => {
    const before={...baseline(),services:{...baseline().services,calendar:{agents:{reader:{operations:["read","create"]}}}}};
    const after={...before,services:{...before.services,calendar:{agents:{reader:{operations:["read"]}}}}};
    expect(planChangedToolExposure(before,after,{reader:{}}).patch.reader.tools.alsoAllow).toEqual(["outlook_calendar_read"]);
  });
});

it("explicit recovery uses only the selected saved grant, preserving read-only and deny rules", () => {
 const policy={services:{onedrive:{allowed_roots:[]},calendar:{agents:{main:{operations:["read"]},studio:{operations:["read","create"]}}},mail:{agents:{}},todo:{agents:{}}}};
 const entries={main:{tools:{alsoAllow:["outlook_mail_read"]}},studio:{tools:{alsoAllow:[]}}};
 expect(planSelectedServiceToolExposure(policy,"calendar","main",entries).patch).toEqual({main:{tools:{alsoAllow:["outlook_mail_read","outlook_calendar_read"]}}});
 expect(()=>planSelectedServiceToolExposure(policy,"calendar","main",entries,["outlook_calendar_read"])).toThrow("denied_tool");
});

describe("saved policy reconciliation", () => {
  const policy = { services: { onedrive: { allowed_roots: root({ read: true, write: false, delete: false }) }, calendar: { agents: { main: { operations: ["read"] } } }, mail: { agents: { main: { operations: ["read", "draft"] } } }, contacts: { agents: { blocked: { operations: ["read"] } } }, todo: { agents: {} } } };
  it("synchronizes the policy union, removing only exact Graph names and preserving broad/operator grants", () => {
    const entries = { main: { tools: { allow: ["exec", "onedrive_delete", "outlook_mail_write", "microsoft_todo_*", "outlook_calendar_*"] } }, studio: {}, blocked: { tools: { deny: ["outlook_contacts_*"] } }, retired: { tools: { alsoAllow: ["outlook_mail_read", "read"] } } };
    const result = planPolicyToolReconciliation(policy, entries);
    expect(result.patch.main.tools.allow).toEqual(["exec", "outlook_mail_write", "microsoft_todo_*", "outlook_calendar_*", "outlook_mail_read"]);
    expect(result.patch.retired.tools.alsoAllow).toEqual(["read"]);
    expect(result.patch.studio.tools.alsoAllow).toEqual(["onedrive_search", "onedrive_list", "onedrive_read", "onedrive_download"]);
    expect(result.patch.blocked).toBeUndefined(); expect(result.drift.blocked.blocked).toEqual(["outlook_contacts_read"]);
    expect(entries.main.tools.allow).toContain("onedrive_delete");
    const repaired = structuredClone(entries) as any;
    for (const [id, patch] of Object.entries(result.patch)) repaired[id].tools = { ...repaired[id].tools, ...patch.tools };
    expect(planPolicyToolReconciliation(policy, repaired).patch).toEqual({});
  });
  it("respects global exact/wildcard/plugin-group denies without blocking revocation or peer diagnosis", () => {
    for (const deny of [["*"], ["group:plugins"], ["microsoft-graph"], ["outlook_*", "onedrive_*"]]) {
      const result = planPolicyToolReconciliation(policy, { main: { tools: { alsoAllow: ["onedrive_delete"] } }, studio: {}, blocked: {} }, deny);
      expect(result.patch.main.tools.alsoAllow).toEqual([]); expect(result.patch.studio).toBeUndefined(); expect(result.drift.main.blocked).toContain("outlook_calendar_read");
    }
  });
  it("reports missing/conflicting agents instead of inventing entries or policy precedence", () => {
    const result = planPolicyToolReconciliation(policy, { main: { tools: { allow: [], alsoAllow: [] } } });
    expect(result.drift.main.reason).toBe("conflicting_tool_policy"); expect(result.drift.studio.reason).toBe("missing_agent_entry"); expect(result.patch).toEqual({});
  });
});

it("never empties a restrictive allow list: real Core empty-allow semantics would broaden unrelated access", async () => {
  // Exercise the pinned Core's actual matcher, not a reimplementation. Internal
  // discovery is test-only: production uses only supported native config CAS.
  const { readdir } = await import("node:fs/promises");
  const { createRequire } = await import("node:module");
  const { dirname, resolve } = await import("node:path");
  const { pathToFileURL } = await import("node:url");
  const dist = resolve(dirname(createRequire(import.meta.url).resolve("openclaw/plugin-sdk/config-mutation")), "..");
  const filename = (await readdir(dist)).find(name => /^tool-policy-match-.*\.mjs$/.test(name));
  expect(filename).toBeTruthy();
  const core = await import(/* @vite-ignore */ pathToFileURL(resolve(dist, filename!)).href);
  const matcher = Object.values(core).find(value => typeof value === "function" && value.name === "createToolPolicyMatcher") as (policy: { allow: string[] }) => (name: string) => boolean;
  expect(matcher).toBeTypeOf("function");
  const policy = { services: { onedrive: { allowed_roots: [] }, calendar: { agents: {} }, mail: { agents: {} }, todo: { agents: {} } } };
  const entries = { restrictive: { tools: { allow: ["outlook_mail_read"] } }, additive: { tools: { alsoAllow: ["outlook_mail_read"] } }, restrictedPeer: { tools: { allow: ["read", "outlook_mail_read"] } } };
  expect(matcher({ allow: [] })("exec")).toBe(true);
  expect(matcher(entries.restrictive.tools)("exec")).toBe(false);
  const plan = planPolicyToolReconciliation(policy, entries);
  expect(plan.patch.restrictive).toBeUndefined();
  expect(plan.drift.restrictive.reason).toBe("empty_allow_would_broaden_access");
  expect(entries.restrictive.tools.allow).toEqual(["outlook_mail_read"]);
  expect(plan.patch.additive.tools.alsoAllow).toEqual([]);
  expect(plan.patch.restrictedPeer.tools.allow).toEqual(["read"]);
  expect(matcher({ allow: plan.patch.restrictedPeer.tools.allow! })("exec")).toBe(false);
});
