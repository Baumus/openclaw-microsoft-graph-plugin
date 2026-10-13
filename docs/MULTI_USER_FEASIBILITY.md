# User-bound Microsoft OAuth: feasibility gate

> Scope update, 2026-10-12: [MULTI_USER_CONCEPT.md](MULTI_USER_CONCEPT.md)
> proposes explicit session-bound accounts and pre-authorized account workers.
> The gate below applies to the earlier always-current-human selection contract;
> it is not a blanket blocker for that separately defined shared-trust model.
> No multi-user implementation or OAuth migration has occurred.

**Status: implementation gate not satisfied; no multi-user capability shipped.**

Checked on 2026-10-12, Europe/Berlin, against OpenClaw **2026.10.1
(e178c39)** and plugin **3.12.1**, baseline
`ae8875b85c95e127d711859329cfcc35ec330e1f`. This document changes no runtime,
credential, permission, manifest, or package version.

## Required boundary

Exactly one Microsoft OAuth connection belongs to each verified OpenClaw user.
Existing tool names, schemas, and Graph `/me` semantics remain unchanged. Account
selection must use the host's current authenticated requester, never model input,
session ownership, names, or a shared default. Missing or ambiguous identity denies
access. Only documented public SDK interfaces are permitted; no Core changes.

## Findings

| Surface | Evidence and limitation |
| --- | --- |
| Personal Control UI RPCs | Public Gateway handlers receive an authenticated client profile. This is a viable foundation for own-account APIs, with live authorization checks; a finished isolated UI implementation has not been tested. |
| Tool factory and policy hook | Public tool context supplies trusted channel sender information, not a canonical user profile. A channel sender is not interchangeable with an OpenClaw profile. |
| Public request-scope accessor | `getPluginRuntimeGatewayRequestScope` is documented and exported by `openclaw/plugin-sdk/plugin-runtime`. Its public types expose the authenticated client profile and live caller authority. A direct, single-person WebChat run can retain this scope, but that does not prove the currently contributing human remains the original requester. |
| Cross-person steering | A second person can join an active run. The running attempt may retain the original caller scope/sender. Core's live personal-tool participant resolver handles this distinction, but no documented public resolver usable by this external plugin was established. The plugin must not import that private resolver. |
| Persisted participant history | `api.runtime.agent.session.getSessionEntry({ agentId, sessionKey, readConsistency: "latest" })` reads committed history. Participant recording is asynchronous and best-effort: accepted input does not await it, and recording failures do not abort the run. A one-person snapshot is therefore not proof that the active run has only one person. |
| Delegation and automation | Verified canonical requester propagation has not been established for these paths. They must not inherit a presumed person from session ownership, prompt text, or the last foreground caller. |

**Decisive gate:** neither the retained request scope nor persisted participants
prove the current requester's unambiguous, live authority after cross-person
steering. Combining them does not fix the asynchronous recording gap. Consequently
the complete account-selection/isolation goal cannot safely be implemented from
the interfaces verified here. This is a limitation of the proven contract, not a
claim that every future or undocumented host implementation is incapable.

The public message/prompt hooks reviewed provide sender IDs or model-context
snapshots; they do not supply a proven canonical, live participant capability for
every accepted active-run steer. Inferring identity from those snapshots is not
an approved replacement.

## Rejected alternatives

- Generic `api.runtime.gateway.request("users.…")`: the documented arbitrary
  in-process dispatcher rejects external plugins; it is not a requester resolver.
- `withUserProfileIdentity`: validates an already selected identity, not who is
  currently requesting the tool.
- Model-provided profile IDs, display names, session owner/creator, token-owner
  defaults, private Core imports, or automatic first-login ownership claims.
- Single-person persisted-history checks presented as a complete steering guard.
- Shipping a UI-only or single-operator subset as completion of the multi-user goal.

## Migration design to retain after the gate clears

The legacy vault/pass connection is globally unowned. Upgrade must not assign it
to the first visitor or silently fall back to it. Require an explicit,
authenticated administrative ownership attestation bound to the exact legacy
state and intended canonical profile before migration can publish that binding.

Migrate credential material and existing agent/resource permissions together,
using profile-bound storage and encryption, exclusive locking, durable atomic
publication, and an idempotent migration receipt. Preserve scopes, refresh-token
generation, rotation/quarantine state, and deny-by-default rules. Leave the legacy
state recoverable; incomplete, conflicting, or tampered migration state denies
access. No migration has run and no real credentials were read for this review.

Own-account status, sign-in, refresh, disconnect, and recovery APIs must derive
their owner server-side. Login sessions, cache keys, continuation handles,
approval state, and browser state must be isolated by user and connection
generation. Never disclose tokens, another person's account details, or global
credential status through personal UI APIs. Global administration is a separate
permission surface, not an implicit right to view personal OAuth details.

## Evidence required before implementation/release

1. A documented public SDK capability proves canonical requester identity and
   live ambiguity/revocation handling at tool execution, including active-run
   steering. Verify exact supported host versions without changing Core.
2. A real host-bound two-user synthetic test proves correct account selection or
   explicit denial for direct, steered, delegated, queued, and automated calls.
3. Isolation tests cover direct foreign-ID calls, concurrent logins, user switches,
   stale callbacks, caches, continuation/approval reuse, and missing identity.
4. Migration tests cover vault and pass backends, permission equivalence, owner
   attestation, retry/crash/conflict, rotation/quarantine, and rollback evidence.
5. Independent security review plus the repository's required package checks.

No version bump or full package test run is warranted for this documentation-only
candidate. Functional, UI, migration, and release acceptance remain **unverified**.

## Sources and local reproduction

Public contracts: [Tool policy hooks](https://docs.openclaw.ai/plugins/hooks/tool-policy),
[Tool plugins](https://docs.openclaw.ai/plugins/tool-plugins),
[Gateway/runtime helpers](https://docs.openclaw.ai/plugins/sdk-runtime/gateway-and-nodes),
[SDK compatibility policy](https://docs.openclaw.ai/plugins/sdk-migration/compatibility-policy),
[Gateway-auth route revalidation](https://docs.openclaw.ai/plugins/sdk-channel-plugins),
[Public session reader](https://docs.openclaw.ai/plugins/sdk-runtime/agent), and
[channel/profile identity links](https://docs.openclaw.ai/concepts/user-model#channel-identity-links).

Installed SDK/source inspection is diagnostic evidence only, not permission to
import private Core modules into the plugin. The separate evidence-directory
diagnostic extracts the installed participant-recording function and executes it
with synthetic worker/lifecycle fixtures: it demonstrates both a pending-record
gap and a failed-record gap. It makes no provider calls, reads no credentials or
live session data, and does not modify Core. It is not an end-to-end host test.
