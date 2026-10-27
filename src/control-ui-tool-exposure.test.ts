import { describe, expect, it } from "vitest";
import { planChangedToolExposure, oneDriveToolExposureSaved, planOneDriveToolExposure, requiredOneDriveTools } from "./control-ui-tool-exposure.js";

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
