// Only a bounded array of positive integer codes may cross the diagnostic boundary.
export function microsoftErrorCodes(value: unknown): number[] | undefined {
  if (!Array.isArray(value) || value.length === 0 || value.length > 8) return undefined;
  const codes: number[] = [];
  for (const code of value) {
    if (typeof code !== "number" || !Number.isInteger(code) || code < 1 || code > 2_147_483_647) return undefined;
    codes.push(code);
  }
  return codes;
}
