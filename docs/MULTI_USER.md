# Personal Microsoft accounts and session-bound execution (4.0 candidate)

## Status and compatibility

Implemented in this isolated candidate against the **public OpenClaw 2026.10.1 SDK**. The build requires that host version for canonical authenticated profiles, exact session authority, and V2 `assertInvocationCurrent` tool factories. No private Core import or Core change is used.

Synthetic integration tests exercise the real registered plugin factories and handlers with two profile IDs and separate encrypted temporary vaults. They are not a production OAuth or live Control UI canary. No production credentials, configuration, install, reload, publication, or migration has been performed.

The accepted [design](MULTI_USER_CONCEPT.md) remains the semantic reference; this document states the implemented surface and conservative host fallbacks.

## Setup

1. An administrator configures `credentialVaultKey` as an existing structured SecretRef, enables the connector and gives named agents default-deny administrative policy ceilings. The existing `policy` is the fallback ceiling. Optional `accountPolicyCeilings` maps canonical profile IDs to different policies, allowing different agent operations and Microsoft resource IDs per account.
2. In the authenticated Control UI open **Mein Microsoft-Konto**. A non-admin with the host's operator read scope may manage only their own connection; the backend takes ownership from `client.authenticatedUserProfile.profileId`, never from form arguments.
3. Create your personal account record. Select account-specific agent rights within your administrative ceiling, and save. Team sharing and unattended-worker consent are separate, default-off choices. Sign in using your approved Microsoft public-client app and the device code shown for your own profile.
4. Select a controllable session, choose **Mein Konto** or an explicitly shared account and bind it. An unknown/unbound session, a new transcript generation, or a generic spawn/fork has no Graph account. It never falls back to the former global credential.

The existing `me` tools and their argument schemas are unchanged. Account management is not a model tool, and models cannot claim a profile or grant account rights.

### Administrative configuration

```json5
{
  enabled: true,
  credentialVaultKey: { source: "store", provider: "approved-provider", id: "graph-vault-key" },
  policy: { /* existing version-2 administrative ceiling */ },
  accountPolicyCeilings: {
    "canonical-profile-A": { /* version-2 ceiling for A */ },
    "canonical-profile-B": { /* version-2 ceiling for B */ }
  },
  accountWorkers: {
    "graph-a": "canonical-profile-A"
  }
}
```

A ceiling is not owner consent and does not connect an account. Administrative config cannot turn on sharing or unattended consent. Account policy cannot weaken a ceiling's warning-approval requirement or widen its agent/service/operation/resource/root permissions. The administrative editor keeps the existing global-ceiling UI; per-profile ceilings can be configured using supported host config editing. OneDrive folder discovery in that editor now authenticates the administrator's **own** profile and connection; no global credential is used.

## Account and session authority

Protected plugin state stores one personal account per canonical profile ID, a connection generation, policy revision and owner consents. Session bindings include exact host agent ID, session key, transcript ID, lifecycle revision and binding revision. Generic patchable `pluginExtensions` values are not consumed as authority.

Team sharing is explicit. Another host-authorized participant may steer the same shared session, but its account remains pinned. Changing accounts requires an exact authenticated UI operation. Active Graph work and pending or accepted native approvals prevent changes; no mutable global current-account variable is used.

Account-private use requires continuously proven host-private output/participation boundaries. The installed public runtime projection does not expose a complete current public-output boundary, so this candidate uses the accepted **share-or-deny fallback for every supported session and worker**, including draft sessions. A canonical owner or draft label alone does not authorize private execution. Personal ownership and connection management remain private/profile-scoped; Graph execution needs explicit team consent. No personal-output isolation is promised. Shared/read-only/suggest/public output and delegation can expose already returned data. Sharing a connection does not share OAuth tokens or connection-management rights.

### Fixed workers and scheduled work

`accountWorkers` binds a trusted host agent ID to one canonical profile. That binding is immutable from UI/model/session input, and takes precedence over session selection. A conflicting session binding is rejected with `graph_execution_account_conflict`. The account must consent to both shared execution and unattended use and retain the executing agent's current permissions. No current human creator is invented when the scheduler supplies none.

OpenClaw owns schedules and start/steering rights. An explicitly configured worker is the supported delegation/automation path. Generic spawn/fork does not implicitly transfer Graph authority. Interactive write approvals remain required where applicable; unattended consent cannot replace an approval.

## Revocation, secrets, approvals and transport

- Account-generation directories are derived from a hash of the exact canonical profile ID; OAuth secrets stay in separate encrypted existing vault implementations, not registry/config/browser state.
- Disconnect advances generation, removes session bindings, disables sharing/unattended consent and makes old vaults inaccessible through runtime resolution. Encrypted predecessors remain retained for recovery; they are not deleted and not automatically reused. Reconnect signs into a new generation.
- Every new effect verifies current account/policy/binding and host invocation authority. OAuth exchange, buffered/streaming Graph, large attachment uploads, direct file writes and resumable upload chunks all enforce the captured scope immediately before dispatch. Completed remote effects cannot be recalled by later revocation.
- Native approval snapshots, allow-always trust, continuation handles and OneDrive instructions/acknowledgements include account/generation/policy/binding identity. Callbacks capture the original scope even when the host resolves them outside the original async context. No approval, upload preflight or continuation is transferred to another account.
- Token refresh locks and cache identity include the scoped vault namespace. Vault refresh rotation retains existing fail-closed uncertain-outcome/quarantine behavior.
- Server handlers validate closed parameter shapes, exact session targets and current authenticated profile before publishing changes or returning personal status. No token export or foreign connection administration is registered.

## Explicit legacy migration and rollback

Version 4 intentionally breaks implicit-global account behavior. There is no first-visitor claim or transparent fallback.

An authenticated administrator explicitly selects **Bestehende Verbindung meinem Profil zuordnen**. The backend migrates only to that same canonical authenticated profile and copies the encrypted legacy credential into its new personal namespace. Migrated rights are the intersection of existing resource/agent grants and that profile's ceiling: no legacy-absent right is added, stricter personal limits apply, and mismatched OneDrive root identity or instruction settings omit that root. Warning approvals are disabled only where both policies permit it. Both new sharing consents remain disabled. The original encrypted legacy file stays unchanged; the protected registry records migration owner/source digest. A retry is idempotent for the same owner/generation; a second profile or replay after disconnect is rejected. New personal connection and migrated credential conflicts fail closed.

The original vault remains a backup, not runtime authority. Do **not** roll back to a plugin version that restores userless global access: leave Graph disabled until an explicitly reviewed rollback/migration is accepted. Restore only encrypted records plus their protected registry and separately protected key. Quarantined personal connections should reconnect; global token export/recovery RPC and legacy CLI commands are no longer runtime surfaces.

## Personal APIs (not model tools)

All `microsoft-graph.accounts.*` handlers use `profileAccess: required`:

| Handler | Purpose | Scope |
| --- | --- | --- |
| `status`, `provision` | Own sanitized connection/account state | operator.read |
| `configure` | Own revision-checked policy and consents within admin ceiling | operator.read |
| `device-start`, `device-status`, `device-cancel` | Own-profile device sign-in | operator.read |
| `disconnect` | Own generation revocation | operator.read |
| `session-status`, `bind` | Exact session account status/selection | operator.write plus host session write/tool policy |
| `migrate-legacy` | Explicit self migration only | operator.admin |
| `resolve-folder` | Own-profile folder discovery for administrator setup | operator.admin |

Source profile IDs, display labels, emails, request text and `for Robin` job prompts never substitute for authenticated authority. Raw provider error messages and credential material are not returned.

## Evidence and intentionally unsupported paths

Focused tests: `accounts.test.ts`, `account-runtime.integration.test.ts`, `control-ui-flow.test.ts`. Existing service, mutation, binary/streaming/upload, instruction, continuation, staging and receipt regressions isolate account admission with an explicit mock; the new unmocked integration suite verifies actual V2 registration, handler authority and account dispatch. Both lanes must pass.

Automatic personal DM/harness admission, arbitrary bind propagation across forks/spawns, private isolation against administrators/host code, and production/live provider behavior remain unverified or unsupported. Use authenticated UI binding, explicit shared consent, and approved fixed workers. No live activation is part of this implementation result.
