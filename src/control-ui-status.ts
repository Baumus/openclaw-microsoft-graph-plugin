// Connection means setup prerequisites and the credential are confirmed; it
// does not establish that an agent can successfully read Microsoft data.
export function isConnectionEstablished(status: {
  secretRefConfigured: boolean;
  savedGrantPresent: boolean;
  applicationStatus: "applied" | "pending" | "unknown";
  credentialResult?: "missing" | "valid" | "quarantined" | "unavailable";
  statusError: boolean;
}): boolean {
  return status.secretRefConfigured
    && status.savedGrantPresent
    && status.credentialResult === "valid";
}
