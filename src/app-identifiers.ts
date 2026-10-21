// Shared browser/Gateway validation. Identity aliases are never treated as
// interchangeable during existing-vault renewal: that comparison stays exact.
export function validClientId(value: string): boolean {
  return /^[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$/.test(value);
}
export function validTenant(value: string): boolean {
  return validClientId(value) || ["common", "organizations", "consumers"].includes(value)
    || (value.length <= 253 && value.includes(".") && value.split(".").every(label =>
      /^[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?$/.test(label)));
}
