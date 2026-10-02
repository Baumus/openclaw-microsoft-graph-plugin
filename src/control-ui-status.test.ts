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
  it("only collapses setup after every saved prerequisite is confirmed", () => {
    expect(isConnectionEstablished(complete)).toBe(true);
    expect(isConnectionEstablished({ ...complete, secretRefConfigured: false })).toBe(false);
    expect(isConnectionEstablished({ ...complete, savedGrantPresent: false })).toBe(false);
    expect(isConnectionEstablished({ ...complete, applicationStatus: "pending" })).toBe(false);
    expect(isConnectionEstablished({ ...complete, applicationStatus: "unknown" })).toBe(false);
    expect(isConnectionEstablished({ ...complete, statusError: true })).toBe(false);
    for (const credentialResult of ["missing", "quarantined", "unavailable"] as const) {
      expect(isConnectionEstablished({ ...complete, credentialResult })).toBe(false);
    }
  });
});
