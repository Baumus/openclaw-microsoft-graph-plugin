import { describe, expect, it } from "vitest";
import { boundedAttachmentSummary } from "./index.js";

/**
 * Regression: Microsoft Graph reports `fileAttachment.size` including attachment overhead, so it
 * is larger than the uploaded content. The response check required exact equality with the raw
 * buffer length, so every successful direct upload raised `invalid_provider_response` after the
 * attachment had already been created, reporting `mutationApplied: "unknown"` for a mutation that
 * had applied. Measured against a live tenant: 17273 bytes sent, 17537 reported.
 */
describe("bounded attachment summary size handling", () => {
  const plan = {
    mode: "direct" as const,
    name: "tamp-observer-ui-design-brief.md",
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

  it("still rejects a smaller size, which would indicate truncation", () => {
    expect(() => boundedAttachmentSummary(
      { id: "attachment-1", name: plan.name, contentType: plan.contentType, size: plan.size - 1 },
      plan,
    )).toThrow("invalid_provider_response");
  });

  it("still rejects a non integer size", () => {
    expect(() => boundedAttachmentSummary(
      { id: "attachment-1", name: plan.name, contentType: plan.contentType, size: 17537.5 },
      plan,
    )).toThrow("invalid_provider_response");
    expect(() => boundedAttachmentSummary(
      { id: "attachment-1", name: plan.name, contentType: plan.contentType, size: "17537" },
      plan,
    )).toThrow("invalid_provider_response");
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
