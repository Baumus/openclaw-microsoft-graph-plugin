import { describe, expect, it } from "vitest";
import { boundedAttachmentSummary } from "./index.js";

/**
 * A successful attachment POST must not become an apparent failure because Graph's response
 * metadata uses a different size convention. A 201 receipt with a usable ID is authoritative;
 * the local content length remains the size exposed to callers.
 */
describe("bounded attachment summary size handling", () => {
  const plan = {
    mode: "direct" as const,
    name: "synthetic-brief.md",
    contentType: "text/markdown",
    content: Buffer.alloc(0),
    size: 17273,
  };

  it("accepts the larger size Graph actually reports for a real upload", () => {
    const summary = boundedAttachmentSummary(
      { id: "attachment-1", name: plan.name, contentType: plan.contentType, size: 17537 },
      plan,
    );
    expect(summary).toMatchObject({ id: "attachment-1", name: plan.name, contentType: plan.contentType });
    // The uploaded byte count is known locally, so it stays authoritative in the summary.
    expect(summary.size).toBe(plan.size);
  });

  it("accepts a response that omits size", () => {
    const summary = boundedAttachmentSummary({ id: "attachment-1", name: plan.name, contentType: plan.contentType }, plan);
    expect(summary.size).toBe(plan.size);
  });

  it("accepts an exact size, as the previous behaviour did", () => {
    const summary = boundedAttachmentSummary({ id: "attachment-1", name: plan.name, contentType: plan.contentType, size: plan.size }, plan);
    expect(summary.size).toBe(plan.size);
  });

  it.each([plan.size - 1, 17537.5, "17537", null])("ignores informational provider size %s", (reportedSize) => {
    const summary = boundedAttachmentSummary(
      { id: "attachment-1", name: plan.name, contentType: plan.contentType, size: reportedSize },
      plan,
    );
    expect(summary).toMatchObject({ id: "attachment-1", size: plan.size });
  });

  it("still rejects a malformed response envelope and bad field values", () => {
    expect(() => boundedAttachmentSummary(null, plan)).toThrow("invalid_provider_response");
    expect(() => boundedAttachmentSummary([], plan)).toThrow("invalid_provider_response");
    // Missing id has no local fallback, unlike name and contentType.
    expect(() => boundedAttachmentSummary({ name: plan.name, contentType: plan.contentType }, plan)).toThrow("invalid_provider_response");
    // Control characters are rejected.
    expect(() => boundedAttachmentSummary({ id: "bad\u0000id", name: plan.name, contentType: plan.contentType }, plan)).toThrow("invalid_provider_response");
  });
});
