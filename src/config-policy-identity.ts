/** Stable identity; object insertion order is not part of a policy revision. */
export function canonicalPolicy(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonicalPolicy).join(",")}]`;
  if (value && typeof value === "object") return `{${Object.keys(value).sort().map(key => `${JSON.stringify(key)}:${canonicalPolicy((value as Record<string, unknown>)[key])}`).join(",")}}`;
  return JSON.stringify(value) ?? "null";
}
