import { describe, expect, it } from "vitest";
import { isConnectionEstablished } from "./control-ui-status.js";

const complete = {
  secretRefConfigured: true,
  savedGrantPresent: true,
  applicationStatus: "applied" as const,
  credentialResult: "valid" as const,
  statusError: false,
};

describe("Control UI connection status", () => {
  it("keeps confirmed account setup through pending policy activation without claiming a read", () => {
    expect(isConnectionEstablished(complete)).toBe(true);
    expect(isConnectionEstablished({ ...complete, secretRefConfigured: false })).toBe(false);
    expect(isConnectionEstablished({ ...complete, savedGrantPresent: false })).toBe(false);
    expect(isConnectionEstablished({ ...complete, applicationStatus: "pending" })).toBe(true);
    expect(isConnectionEstablished({ ...complete, applicationStatus: "unknown" })).toBe(true);
    expect(isConnectionEstablished({ ...complete, statusError: true })).toBe(true);
    for (const credentialResult of ["missing", "quarantined", "unavailable"] as const) {
      expect(isConnectionEstablished({ ...complete, credentialResult })).toBe(false);
    }
  });
});
