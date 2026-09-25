import { describe, expect, it } from "vitest";
import { authorizeOperation, authorizeRoot, normalizeRelativePath, validatePolicy } from "./policy.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function fixture() {
  return validatePolicy(graphPolicyFixture());
}

describe("Microsoft Graph policy", () => {
  it("is default deny and preserves exact OneDrive grants", () => {
    const policy = fixture();
    expect(policy.rules.default).toBe("deny");
    expect(authorizeRoot(policy, "fixture-reader", "synthetic_documents", "read")).toMatchObject({ path: "/Synthetic/Documents", agents_instructions: "trusted" });
    expect(() => authorizeRoot(policy, "main", "synthetic_documents", "read")).toThrow("access_denied");
    expect(() => authorizeRoot(policy, "fixture-reader", "synthetic_documents", "write")).toThrow("access_denied");
  });

  it("requires trusted exact agent, operation, and me resource", () => {
    const policy = fixture();
    expect(() => authorizeOperation(policy, undefined, "calendar", "read")).toThrow("trusted_agent_identity_required");
    expect(() => authorizeOperation(policy, "fixture-reader", "calendar", "read")).toThrow("access_denied");
    expect(() => authorizeOperation(policy, "main", "todo", "delete", "other-user")).toThrow("access_denied");
    expect(() => authorizeOperation(policy, "main", "todo", "read", "me")).not.toThrow();
    const secondaryCalendar = policy.services.calendar.agents.main.resources?.find((resource) => resource !== "me");
    expect(secondaryCalendar).toBeTruthy();
    expect(() => authorizeOperation(policy, "main", "calendar", "create", secondaryCalendar)).not.toThrow();
    expect(() => authorizeOperation(policy, "secondary-agent", "calendar", "read", secondaryCalendar)).not.toThrow();
    expect(() => authorizeOperation(policy, "secondary-agent", "calendar", "create", secondaryCalendar)).not.toThrow();
    expect(() => authorizeOperation(policy, "main", "calendar", "create", "unknown-calendar")).toThrow("access_denied");
    expect(() => authorizeOperation(policy, "secondary-agent", "calendar", "create", "unknown-calendar")).toThrow("access_denied");
  });

  it("rejects path traversal and non-deny policy", () => {
    expect(() => normalizeRelativePath("../private")).toThrow("invalid_relative_path");
    expect(() => normalizeRelativePath("/absolute")).toThrow("invalid_relative_path");
    expect(() => validatePolicy({ ...graphPolicyFixture(), rules: { default: "allow" } })).toThrow("invalid_policy");
  });

  it("requires an immutable item-id pin for every OneDrive root", () => {
    const policy = fixture();
    const candidate = structuredClone(policy) as any;
    delete candidate.services.onedrive.allowed_roots[0].item_id;
    expect(() => validatePolicy(candidate)).toThrow("invalid_policy");
  });

  it("accepts only credential-free policy v2", () => {
    const current = fixture();
    const policy = validatePolicy({ version: 2, rules: current.rules, services: current.services });
    expect(policy.version).toBe(2);
    expect(policy).not.toHaveProperty("account");
    expect(() => validatePolicy({ ...policy, account: { credentials: {} } })).toThrow("invalid_policy");
  });

  it("accepts only an explicit trusted AGENTS.md opt-in", () => {
    const policy = fixture();
    const disabled = structuredClone(policy) as any;
    delete disabled.services.onedrive.allowed_roots[0].agents_instructions;
    expect(validatePolicy(disabled).services.onedrive.allowed_roots[0].agents_instructions).toBeUndefined();
    const invalid = structuredClone(policy) as any;
    invalid.services.onedrive.allowed_roots[0].agents_instructions = true;
    expect(() => validatePolicy(invalid)).toThrow("invalid_policy");
  });
});
