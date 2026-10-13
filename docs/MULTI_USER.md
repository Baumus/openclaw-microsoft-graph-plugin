# Personal Microsoft accounts and session-bound execution (4.0 candidate)

## Status and compatibility

Implemented in this isolated candidate against the **public OpenClaw SDK (minimum 2026.10.1; directory adapter verified on 2026.10.1 and 2026.10.2)**. The build requires that host version for canonical authenticated profiles, exact session authority, and V2 `assertInvocationCurrent` tool factories. No private Core import or Core change is used.

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

## Installation and upgrade completion

A qualifying single-user upgrade bootstraps automatically at Gateway startup without opening a browser, signing into a personal profile, acquiring an admin role or pressing a migration button. New connections, multi-user and unprovable directory states use the [explicit setup routine](../README.md#1-check-requirements-and-install). The personal UI retains genuine identity gates for account management; shared `gateway-owner` receives identity guidance, never another person's account forms.

The UI uses supported `registerAction` (header), `registerNavigation`, and `navigation.pageHref/openPage` contracts. The header and status reads do not migrate, provision, sign in or bind accounts. Entering the personal page only reads status and session state; it does not invoke automatic migration. The explicit migration button and the separately registered automatic-migration RPC retain the authenticated gates below. OpenClaw 2026.10.1's general feature-plugin install path supplies no authenticated personal post-install callback; model/provider setup descriptors are not a substitute. Consequently the installer cannot force-open an authenticated wizard or claim readiness. The packaged description, enabled-field hint, CLI guidance and README carry the required step even before native UI is available; UI contributions require **Settings → Labs → Custom plugin UI**.

## Automatic single-user assistant (snapshot semantics)

In full registration mode, an enabled plugin registers a lightweight version-2 service lifetime. The supported post-ready `gateway_start` hook schedules bounded migration only after the proof RPC is published. The host lifecycle authorizes preservation of the prior installation-owned connection, while a fresh exact-instance complete roster identifies its sole canonical personal owner. No account RPC, browser login, admin role or provider fetch is needed. Discovery/setup/CLI metadata modes never register this scheduler. Service stop, Gateway drain and same-config plugin replacement synchronously close admission and join scheduled work; a generation guard pins current plugin configuration, root and relevant host transport/auth configuration. After the first published host-ready hook, a retained service stop/start reacquires a fresh scheduler and re-proves the directory without replaying that hook, only while the original host drain signal is live. Initial service start never schedules before host publication; superseded registrations cannot reclaim ownership. Unrelated Gateway UI settings do not retire this authority. Directory readiness failures retry at most three times; conflicts and sticky invalidation do not retry or resurrect access. A missing legacy connection is a no-op with no account provisioning. Existing explicit accounts/migrations/session bindings are not silently promoted.

Qualification reads the **complete actual host roster** through the public `users.list` Gateway RPC, not the number of Graph accounts, observed participants, session metadata, or display names. Only the exact shared `gateway-owner` sentinel is excluded. Merged aliases must resolve consistently to existing canonical people. For the separately invoked authenticated automatic-migration RPC only, the sole canonical person must equal the RPC's currently authenticated profile and the host requires `operator.admin`. The trusted startup path does not require either personal login or that person's administrator role. Unknown runtime/schema, malformed/incomplete/truncated output, foreign host, missing authority and unavailable reads fail closed.

The portable adapter qualifies the verified public contracts in OpenClaw **2026.10.1 and 2026.10.2**. It has no personal ID, display-name, machine, home-directory, state-root or fixed-port dependency. It uses public `resolveGatewayPort` (configured/default/named-profile/environment ports) and `resolveGatewayAuth` (config-first credentials and explicit/inferred token/password mode). Local `loopback`, `lan` and `auto` listeners are contacted only through an exact loopback URL; instance proof must still pass. An environment `OPENCLAW_GATEWAY_URL` cannot redirect enumeration to another server. Transport auth, profile/port environment and host configuration are pinned across reads and startup authority. Each snapshot uses one fresh public SDK `GatewayClient` connection with ordinary `cli` identity/mode, operator role and only `operator.read`. Shared secrets stay in process memory, never argv, logs or model output. `deviceIdentity: null`, `sharedStateMode: "read-only"`, an empty transport environment and silent SDK log callbacks prevent device/token/config storage or inherited host paths. No child processes, scratch directories, reserved backend identity, persistent authenticated client or roster cache are used. The before/roster/after proof requests all use that same connection; any disconnect/reconnect retires the read. This also avoids hardcoded Unix/Windows home paths.

| Installation condition | Upgrade behavior |
| --- | --- |
| Verified .1/.2 local token/password listener, valid existing legacy vault, exactly one canonical personal user | Automatic at post-ready startup; no personal login/admin/click/reconnect required |
| Different canonical IDs, consistent merged aliases, default/custom/profile-selected ports or custom state root | Same automatic path; names and host paths do not decide identity |
| Fresh install, no existing valid legacy connection | No automatic Microsoft sign-in; explicit personal setup |
| Zero/multiple canonical users, unknown/incomplete/malformed directory, foreign instance | Never guess an owner; explicit authenticated setup |
| Remote, TLS/custom/tailnet-only listener, anonymous/trusted-proxy auth, unsupported release or unresolved active SecretRef | Conservative explicit setup; no unsafe auth/identity bypass |
| Existing explicit account/binding, interrupted conflicting migration, sticky revoked automatic receipt | No silent promotion or resurrection; explicit recovery/setup |

SecretRef values are not independently resolved or read from provider/files by this adapter. Resolved runtime configuration is used only when the public auth helper can safely supply its selected shared credential; an unresolved active reference cannot be replaced with an environment fallback. Future releases are not assumed compatible by a broad version comparison: adding one requires contract verification and the supported-host CI matrix.

The adapter uses the exported public `openclaw/plugin-sdk/gateway-runtime` client and joined connection cleanup, with plugin-instance proof RPC readback before and after enumeration on the same connection. Roster and proof request only `operator.read`. These fixed read requests use host-authorized shared auth, not person impersonation or a new model permission. No private dispatcher/import, Core change or internal database read is used.

Repeatable synthetic transport/startup verification is repository-owned: `npm run plugin:build && npm run migration:canary -- --auth password --owner synthetic-independent-person`. The harness derives the actual host version and public SDK from the selected entry's dependency tree (not a hardcoded machine path); `--entry /absolute/extracted/package/dist/index.js --receipt /temporary/receipt.json` tests a packed entry with its own supported OpenClaw dependencies. The CI matrix installs each exact .1/.2 release without updating the lockfile and runs the suite plus actual public SDK canary on both supported Node lines. Linux transport is exercised in CI; Windows-style state-path exclusion is unit-tested, not a native Windows runtime claim.

Migration writes one protected plugin-owned automatic receipt containing owner, account generation, roster fingerprint and a sticky invalidation flag. The common automatic/explicit writer holds both source and destination refresh-vault locks through registry publication, rejects any in-flight/quarantined or malformed marker and held/stale/compromised lock, and validates interrupted destination copies before granting authority. It never auto-recovers refresh state. The encrypted legacy vault and its existing markers stay byte-for-byte unchanged. Rights remain **legacy policy ∩ personal ceiling**; shared/unattended flags stay off. In absence of an explicit session binding or fixed worker, already authorized agents can use this exact personal account across continued/new sessions, scheduled jobs and ordinary delegation. Unapproved agents/resources/actions remain denied. Explicit bindings and fixed workers always take precedence and retain their consent requirements; this mode does not infer a human from cron or transfer account-management rights.

Every new automatic operation and the existing sensitive asynchronous checkpoints re-read the actual directory. A detected roster/identity change permanently invalidates the convenience receipt. During bootstrap the captured owner/generation/topology tombstone is serialized under the same registry lock as publication, independently of lost caller/profile authority; the revocation branch discards all positive account/grant changes. It stays disabled if the second user is later removed; no background retry/page remount resurrects it. Account-policy saves and disconnect also invalidate it. Explicit personal consent/session/worker selection remains available. Directory outages deny access without returning cached-singleton/global credentials.

**This is a snapshot-level candidate, not an atomic lifetime lease.** The public directory returns no catalog incarnation/revision capability. Addition/removal between snapshots (ABA), a mutation immediately after the last snapshot, and already issued provider requests cannot be continuously revoked or recalled by this adapter. Synchronous media commit guards retain account-generation/authority protection, but cannot independently enumerate a directory synchronously. The same-connection proof sandwich is not an atomic catalog lease. Fresh authorization snapshots are retained at admission and asynchronous sensitive boundaries; there is no TTL/cache window. Each read is bounded by a 5-second connection deadline, three 5-second RPC deadlines and a 1-second joined SDK close. Caller/native cancellation denies subsequent guarded effects; ordinary tool directory admission does not receive the caller AbortSignal. Startup admission and roster reads receive the owned lifecycle AbortSignal. After a complete roster arrives, the bounded final host proof finishes even during drain so observed topology mismatches can be recorded as revocation-only facts; live authority is still required for positive publication. The prior extra **90-second directory allowance is removed**: configured operation/provider request deadlines are unchanged. Initial hook/account admission still precedes the action timer, so it is not a strict whole-wall-clock deadline. The SDK enforces its own 25 MiB frame ceiling; the adapter rejects directory/proof JSON above 1 MiB before qualification. No claim of immediate revocation, continuously atomic isolation, guaranteed zero overhead, production transport correctness or live OAuth migration is made. Default registered tools and latency measurements use isolated synthetic Gateways and synthetic provider responses only; this does not replace production validation. Repeat comparable full-tool measurements with `npm run migration:canary -- --benchmark 3 --receipt /temporary/receipt.json` (optionally `--entry` for a packed baseline/candidate). Receipts include fresh roster/RPC/connection/process counts and individual tool durations, not a simulated SDK timer. These precise limitations require delivery-owner/user review before activation; this candidate does not change the running installation.

## Explicit legacy migration and rollback

Version 4 intentionally breaks userless implicit-global account behavior. There is no first-visitor claim or arbitrary-account fallback. The bounded single-user assistant below can migrate a proven sole personal user automatically at Gateway startup.

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

The read-only `microsoft-graph.directory-host` (`operator.read`, empty parameters) exposes only an opaque process/config-generation instance ID for the public SDK exact-host proof; it does not expose profiles, credentials, or account administration.
