import type { GraphPolicy } from "../policy.js";

const fixture: GraphPolicy = {
  version: 2,
  rules: { default: "deny" },
  services: {
    onedrive: {
      allowed_roots: [{
        label: "synthetic_documents",
        path: "/Synthetic/Documents",
        drive_id: "synthetic-drive",
        item_id: "synthetic-root",
        include_descendants: true,
        agents_instructions: "trusted",
        permissions: { read: true, write: true, delete: false },
        agents: {
          "fixture-reader": { permissions: { read: true, write: false, delete: false } },
          main: { permissions: { read: false, write: true, delete: false } },
        },
      }],
    },
    calendar: {
      agents: {
        main: { operations: ["read", "create", "update", "respond", "attach", "delete"], resources: ["me", "synthetic-calendar"] },
        "secondary-agent": { operations: ["read", "create", "update", "respond", "attach", "delete"], resources: ["me", "synthetic-calendar"] },
      },
    },
    mail: { agents: { main: { operations: ["read", "draft", "update", "move", "mark", "send", "delete"], resources: ["me"] } } },
    todo: { agents: { main: { operations: ["read", "create", "update", "delete"], resources: ["me"] } } },
  },
};

export function graphPolicyFixture(): GraphPolicy {
  return structuredClone(fixture);
}
