import { describe, expect, it } from "vitest";
import { validClientId, validTenant } from "./app-identifiers.js";

describe("shared app identifier validation", () => {
  it("accepts GUIDs, tenant domains and documented Microsoft aliases", () => {
    const guid = "12345678-1234-1234-1234-1234567890ab";
    expect(validClientId(guid)).toBe(true);
    for (const tenant of [guid, "example.onmicrosoft.com", "example.com", "common", "consumers", "organizations"]) expect(validTenant(tenant), tenant).toBe(true);
  });
  it("rejects malformed IDs and non-domain or URL-like tenant input", () => {
    for (const value of ["-".repeat(36), "123456781234123412341234567890abcdef", "client-secret", ""]) expect(validClientId(value)).toBe(false);
    for (const value of ["", ".", "-", "example..com", "-example.com", "example-.com", "example.com.", "https://example.com", "example.com/path", "foo@bar.com", "a".repeat(64) + ".com"]) expect(validTenant(value), value).toBe(false);
  });
});
