import { Compile } from "typebox/compile";
import type { TSchema } from "typebox";

export function compileToolParameterCheck(schema: TSchema & { properties?: Record<string, unknown> }): (value: unknown) => string | undefined {
  const checker = Compile(schema);
  const fields = Object.keys(schema.properties ?? {}).filter((field) => /^[A-Za-z][A-Za-z0-9_]{0,63}$/.test(field));
  const pointers = fields.map((field) => `#/properties/${field.replaceAll("~", "~0").replaceAll("/", "~1")}`);
  return (value) => {
    if (checker.Check(value)) return undefined;
    const invalid = new Set<string>();
    let unsupported = false;
    for (const error of checker.Errors(value)) {
      if (error.keyword === "additionalProperties") unsupported = true;
      const index = pointers.findIndex((pointer) => error.schemaPath === pointer || error.schemaPath.startsWith(`${pointer}/`));
      // Only schema-owned names may enter diagnostics, never caller paths, values or validator messages.
      if (index >= 0 && invalid.size < 4) invalid.add(fields[index]);
    }
    const details = [
      ...(unsupported ? ["Unsupported fields are not allowed."] : []),
      ...(invalid.size ? [`Check schema requirements for: ${[...invalid].join(", ")}.`] : []),
    ].join(" ");
    return `invalid_tool_parameters: ${details ? `${details} ` : ""}Call tool_describe for this tool's full schema; correct the arguments before retrying. Do not retry unchanged.`;
  };
}
