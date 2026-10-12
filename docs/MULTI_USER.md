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

## Automatic single-user assistant (snapshot semantics)

When an authenticated administrator loads **Mein Microsoft-Konto**, the page calls the admin-only `automatic-migration` handler automatically. No confirmation button, re-login for a valid vault, session binding, or worker configuration is required on the qualifying path. This is not a startup guess: before that authenticated entry the plugin cannot attribute a legacy connection to a person, and automatic execution is unavailable. A missing legacy connection remains normal personal setup; an existing explicit migration/personal account is not silently taken over.

Qualification reads the **complete actual host roster** through public `openclaw users list --json`, not the number of Graph accounts, observed participants, session metadata, or display names. Only the exact shared `gateway-owner` sentinel is excluded. Merged aliases must resolve consistently to existing canonical people. The sole canonical person must equal the RPC's currently authenticated profile; the host separately requires `operator.admin`. Unknown runtime/schema, malformed/incomplete/truncated output, foreign host, missing authority and unavailable reads fail closed.

The current adapter intentionally qualifies only the tested OpenClaw **2026.10.1** local loopback/token-auth shape with an explicit runtime port. It uses the exported `openclaw/cli-entry`, the public bounded argv runner, a sanitized child environment with an exact loopback URL and transient shared auth (never argv/log/model output), private disposable CLI scratch state instead of the real host store, joined child-process cleanup, and plugin-instance proof RPC readback before and after enumeration. Remote, TLS/custom-bind, anonymous/trusted-proxy/password, unsupported version and unresolved SecretRef configurations retain explicit personal migration. The roster call explicitly requests only `operator.read`; the separate proof RPC is read-only and requires `operator.read`, but the isolated public `gateway call` CLI uses its broader default connection scope set (`operator.admin`, `operator.read`, `operator.write`, `operator.approvals`, `operator.questions`, `operator.pairing`, `operator.talk.secrets`). No supported narrower proof transport is claimed. These fixed read requests use existing host-authorized shared auth, not a person impersonation or new model permission. No private dispatcher, private import, Core change or internal database read is used.

Migration writes one protected plugin-owned automatic receipt containing owner, account generation, roster fingerprint and a sticky invalidation flag. The common automatic/explicit writer holds both source and destination refresh-vault locks through registry publication, rejects any in-flight/quarantined or malformed marker and held/stale/compromised lock, and validates interrupted destination copies before granting authority. It never auto-recovers refresh state. The encrypted legacy vault and its existing markers stay byte-for-byte unchanged. Rights remain **legacy policy ∩ personal ceiling**; shared/unattended flags stay off. In absence of an explicit session binding or fixed worker, already authorized agents can use this exact personal account across continued/new sessions, scheduled jobs and ordinary delegation. Unapproved agents/resources/actions remain denied. Explicit bindings and fixed workers always take precedence and retain their consent requirements; this mode does not infer a human from cron or transfer account-management rights.

Every new automatic operation and the existing sensitive asynchronous checkpoints re-read the actual directory. A detected roster/identity change permanently invalidates the convenience receipt. During bootstrap the captured owner/generation/topology tombstone is serialized under the same registry lock as publication, independently of lost caller/profile authority; the revocation branch discards all positive account/grant changes. It stays disabled if the second user is later removed; no background retry/page remount resurrects it. Account-policy saves and disconnect also invalidate it. Explicit personal consent/session/worker selection remains available. Directory outages deny access without returning cached-singleton/global credentials.

**This is a snapshot-level candidate, not an atomic lifetime lease.** The public directory returns no catalog incarnation/revision capability. Addition/removal between snapshots (ABA), a mutation immediately after the last snapshot, and already issued provider requests cannot be continuously revoked or recalled by this adapter. Synchronous media commit guards retain account-generation/authority protection, but cannot independently enumerate a directory synchronously. The instance-proof sandwich likewise is not a single-connection atomic server/catalog lease. CLI authority admission can be slow: the automatic scope adds a finite **90-second local-directory allowance** after entering the tool action. Initial account/hook admission happens before that operation timer, so this is not a strict whole-wall-clock deadline and the configured total-latency budget is increased. Caller/native cancellation denies subsequent guarded effects, but the public directory runner does not receive that caller AbortSignal: cancellation does not immediately stop an already running directory read. Each child is separately bounded (10-second process / 5-second RPC) and requests joined process-tree cleanup. Per-provider request timeouts begin at actual network dispatch, never while waiting for local directory admission. This is not an unbounded timeout extension or a permission grant. No claim of immediate revocation, continuous private-output isolation, guaranteed zero-delay CLI startup, real-host transport correctness, or live OAuth migration is made. Public-CLI transport and a default registered tool with a 90s caller deadline are exercised only against an isolated synthetic Gateway and synthetic provider responses; this does not replace production validation. These precise limitations require delivery-owner/user review before activation; this candidate does not change the running installation.

## Explicit legacy migration and rollback

Version 4 intentionally breaks userless implicit-global account behavior. There is no first-visitor claim or arbitrary-account fallback. The bounded single-user assistant below can migrate an authenticated sole administrator automatically.

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
| `automatic-migration` | Automatic authenticated sole-user legacy assistant | operator.admin |
| `resolve-folder` | Own-profile folder discovery for administrator setup | operator.admin |

Source profile IDs, display labels, emails, request text and `for Robin` job prompts never substitute for authenticated authority. Raw provider error messages and credential material are not returned.

## Evidence and intentionally unsupported paths

Focused tests: `accounts.test.ts`, `account-runtime.integration.test.ts`, `control-ui-flow.test.ts`. Existing service, mutation, binary/streaming/upload, instruction, continuation, staging and receipt regressions isolate account admission with an explicit mock; the new unmocked integration suite verifies actual V2 registration, handler authority and account dispatch. Both lanes must pass.

Automatic multi-user personal DM/harness admission, arbitrary bind propagation across forks/spawns, private isolation against administrators/host code, and production/live provider behavior remain unverified or unsupported. The bounded singleton assistant is separate from those claims. Outside its qualifying path use authenticated UI binding, explicit shared consent, and approved fixed workers. No live activation is part of this implementation result.

The read-only `microsoft-graph.directory-host` (`operator.read`, empty parameters) exposes only an opaque process/config-generation instance ID for the public CLI exact-host proof; it does not expose profiles, credentials, or account administration.
