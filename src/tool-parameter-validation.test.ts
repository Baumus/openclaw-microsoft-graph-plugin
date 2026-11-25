import { describe, expect, it } from "vitest";
import { Type } from "typebox";
import { compileToolParameterCheck } from "./tool-parameter-validation.js";

describe("privacy-safe tool schema rejection", () => {
  const check = compileToolParameterCheck(Type.Object({
    action: Type.Literal("search_tasks"),
    listId: Type.String(),
    status: Type.Optional(Type.Union([Type.Literal("notStarted"), Type.Literal("completed")])),
    limit: Type.Optional(Type.Integer({ minimum: 1, maximum: 50 })),
    entries: Type.Optional(Type.Array(Type.Object({ title: Type.String() }, { additionalProperties: false }))),
  }, { additionalProperties: false }));
  const valid = { action: "search_tasks", listId: "private-list", limit: 50 };

  it("rejects unsupported arguments without echoing their names or values or changing the call", () => {
    const params = { ...valid, isCompleted: false, "private-token-name": "private-token-value" };
    const before = structuredClone(params);
    const reason = check(params)!;
    expect(reason).toContain("invalid_tool_parameters: Unsupported fields are not allowed.");
    expect(reason).toContain("tool_describe");
    expect(reason).toContain("Do not retry unchanged");
    expect(reason).not.toMatch(/private|isCompleted|false/);
    expect(params).toEqual(before);
    expect(check(valid)).toBeUndefined();
  });

  it("names schema-owned fields for enum, type, bound and nested failures without caller content", () => {
    const reason = check({ ...valid, status: "secret-status", limit: "secret-limit", entries: [{ title: 123, "secret-key": "secret-value" }] })!;
    expect(reason).toContain("status, limit, entries");
    expect(reason).not.toMatch(/secret|123|title|\/0/);
    for (const limit of [0, 51, 1.5]) expect(check({ ...valid, limit })).toContain("limit");
    expect(check({ ...valid, status: "completed" })).toBeUndefined();
    expect(check({ action: "search_tasks" })).toContain("tool_describe");
  });

  it("bounds diagnostics even with many failures and long unknown field names", () => {
    const schema = Type.Object(Object.fromEntries(Array.from({ length: 20 }, (_, i) => [`field${i}`, Type.String()])), { additionalProperties: false });
    const validate = compileToolParameterCheck(schema);
    const reason = validate(Object.fromEntries([
      ...Array.from({ length: 20 }, (_, i) => [`field${i}`, 0]),
      ["secret".repeat(1000), "private"],
    ]))!;
    expect(reason).toContain("field0, field1, field2, field3.");
    expect(reason).not.toMatch(/field4|secret|private/);
    expect(reason.length).toBeLessThan(512);
  });
});
