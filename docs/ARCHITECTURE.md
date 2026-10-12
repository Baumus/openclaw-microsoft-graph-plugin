# Architecture

## Purpose and boundary

The plugin is one Gateway-owned boundary between optional OpenClaw tools and Microsoft Graph v1.0. It owns policy evaluation, credential retrieval, delegated token exchange, request construction, provider-response validation, bounded output, private-media transfer, and approval gates.

The plugin does not own user intent, business policy, data-retention decisions, Microsoft tenant administration, consent administration, or OpenClaw host security. It does not provide a generic Microsoft Graph proxy.

## Call flow

1. OpenClaw selects an optional plugin tool.
2. The `before_tool_call` hook classifies the action as read-only, warning-level, or critical.
3. The hook applies available fail-closed preflight checks. For OneDrive upload/update, it securely opens the protected artifact and verifies opened-file identity, exact lowercase SHA-256, and byte size before optional `AGENTS.md` discovery. Other OneDrive warning mutations retain exact root authorization and applicable discovery; discovery is a separate read and requires read authority.
4. Warning-level mutations require OpenClaw-native approval unless `warningApprovalsRequired` is explicitly `false`. Critical actions always require native approval. The hook returns the exact inspected parameters with an approval and binds host tool-call identity to an execution-time snapshot; composed parameter drift fails closed before the tool implementation proceeds.
5. The runtime validates action-specific parameters and trusted `toolContext` identity.
6. Execution validates the default-deny policy supplied in resolved plugin config and authorizes the exact agent, service, operation, and resource.
7. After authorization, the plugin validates the one resolved vault key, acquires the shared credential lock, and securely reads and decrypts the one vault record.
8. Under the credential lock, the plugin durably publishes and verifies an authenticated `in_flight` refresh marker before dispatching one operation-specific OAuth exchange. A provider-issued replacement refresh token is durably published and verified before the marker is cleared or the access token is cached or returned.
9. A fixed Graph v1.0 path is built from validated resource identifiers. Arbitrary Graph URLs are not accepted.
10. Responses are bounded, normalized, sanitized, and returned. Binary data is streamed to OpenClaw private media or reduced to a digest rather than emitted inline.

Every validation or authorization failure before step 7 avoids plugin-side key selection, vault access, OAuth, and provider calls.

## Trust boundaries

### OpenClaw host

The OpenClaw Gateway and plugin runtime are trusted to provide the real agent ID, session ID, authenticated requester context, private-media root, media-store API, approval mechanism, and plugin configuration. These identities are not accepted as model-controlled tool parameters. The plugin does not infer policy grants from local workspace content.

### Policy

The policy is an operator-controlled object at `plugins.entries["microsoft-graph"].config.policy`. OpenClaw may compose that host config with `$include`; the plugin receives only the resolved object, validates it for every tool call, and requires `rules.default: deny`. OneDrive roots bind a display label to immutable `drive_id` and `item_id` values; service grants bind agent IDs to closed operation sets and optional resource IDs.

### Credentials

Policy version 2 is credential-free. One static SecretRef supplies the AES-256-GCM key for one OAuth credential record below `<state>/plugin-data/microsoft-graph/credentials`. The vault is reached only after policy authorization. Access tokens stay process-local; refresh tokens are never returned or intentionally logged. There is no provider-side read/write isolation: the shared credential carries the union of consented scopes, while policy and approval gates govern ordinary plugin operations.

Local credential commands cross into the active Gateway through plugin-owned, operator-scoped RPC methods so only the Gateway's materialized config reaches vault operations. The read-scoped status method and admin-scoped device sign-in, migration, restore, and recovery methods accept closed parameter objects and return closed sanitized envelopes. Raw errors and secret-bearing values never cross this boundary. Gateway startup registers these methods even while `config.enabled` is `false`; that flag continues to block all Microsoft Graph tools.

### Microsoft Graph

Graph responses, continuation URLs, upload-session URLs, attachment metadata, and downloaded bytes are untrusted provider input. The plugin validates origins, paths, identifiers, lengths, MIME values, paging shape, and response structure before use.

### OneDrive `AGENTS.md`

This optional plugin feature applies OpenClaw's standard `AGENTS.md` instruction format to a policy-pinned OneDrive root. It is separate from, and never imports, the host agent's local workspace instructions. Remote content remains untrusted unless an operator explicitly sets `agents_instructions: trusted` after verifying who can write to that root, and it cannot override higher-priority OpenClaw instructions, host policy, or approval requirements.

Discovered instruction content has a one-hour absolute freshness limit, independent of the 12-hour idle cache limit. Repeated reads do not extend content freshness. Once it expires, approval validation makes no Graph request and requires rediscovery; execution rediscovery reloads the chain, invalidates the old receipt, and requires acknowledgement of the fresh content. The receipt remains bound to agent, session, and policy-pinned root.

## State and data handling

- OAuth access-token, continuation, warning approval trust, and OneDrive instruction caches are process-local and bounded. Warning `allow-always` trust is capped at 1,024 scopes, keyed by authenticated agent ID, exact tool name, and normalized action; it is never persisted and resets on plugin reload or process restart. The token cache is a 128-entry expiry-aware LRU keyed to durable vault identity, capped at one hour of residency with a 60-second expiry skew.
- The plugin persists encrypted OAuth credential envelopes and an authenticated secret-free write-ahead refresh marker (`in_flight`, then `quarantined` after uncertainty). It does not persist Graph content, access tokens, or continuation URLs.
- Private-media downloads use opaque `media://inbound/...` artifacts managed by OpenClaw.
- Uploads accept only canonical private-media URIs and revalidate containment, file identity, size, and links.
- Continuations are opaque random handles bound to caller, service, action, policy resource, and normalized criteria.

## Failure and retry behavior

- Reads may retry only narrowly classified throttling, selected 5xx, timeout, and transport failures within fixed budgets.
- OAuth and vault writes are not automatically retried. Marker publication and lock verification must complete before dispatch. Any uncertain outcome after dispatch retains authenticated blocking state and requires explicit operator recovery or reauthorization.
- Vault generation is authenticated but has no external monotonic anchor; replay of an older valid encrypted backup with its matching key remains an operator-managed residual risk.
- OneDrive creates use no-overwrite semantics; updates use ETag protection.
- Calendar multiwrite is capped at 100 operations, ordered, and non-atomic. It reports per-operation status and does not roll back or retry writes.
- A provider-confirmed write with a malformed verification response may be reported as applied but unverified; callers must not blindly repeat it.

## Build and packaging

`src/index.ts` is bundled to `dist/index.js` for Node 24 and newer supported host runtimes. Runtime dependencies remain external and are declared in `dependencies`, including exact-pinned `@openclaw/fs-safe` 0.21.1. `openclaw plugins build --check` verifies generated manifest/package metadata without rewriting it. The npm `files` allowlist ships only runtime code, static metadata, consumer documentation, the icon, synthetic JSON5 policy objects, and the bounded JSON helper.

## Version 4 bounded single-user migration candidate

`account-directory.ts` provides the whole-host topology oracle through the public exported CLI entry and bounded argv runner. Private disposable scratch state contains the public CLI’s startup bookkeeping, so the reader never selects the real host store for local CLI initialization. Exact local shared-token/instance proof and a pinned known JSON/runtime shape prevent foreign/default gateways or partial/unknown catalogs from qualifying. Neither private `user_profiles` storage nor restricted `api.runtime.gateway.request` is called. The latter restriction is not evaded through reserved device-less backend client identities.

An authenticated admin-only personal-page bootstrap reuses the same nonwidening legacy migration writer as the explicit multi-user operation. A protected owner/generation/topology receipt (not a fictional cron actor) then selects that account for legacy-authorized agents only where no explicit binding/fixed worker exists. Detected catalog or account-policy changes permanently revoke this convenience path without revoking independent explicit grants. The full roster itself, email identities and shared auth are never persisted or returned to models.

Fresh directory snapshots at admission and asynchronous sensitive boundaries are not a catalog lease. The public directory has no atomic incarnation/revision authority. Concurrent add/remove ABA, a change after the last snapshot and already-issued provider effects remain outside a strict immediate-revocation claim. Synchronous local media commit guards retain account-generation authority but cannot synchronously enumerate topology. Local CLI admission has a finite 90s whole-operation allowance; it does not start/extend the per-provider request timer or supersede native/caller aborts. Full behavior and activation limitations are documented in [MULTI_USER.md](MULTI_USER.md).

The common legacy writer enforces refresh-state admission under both source/destination vault locks through the registry commit. Any blocking or malformed marker, uncertain lock, or conflicting interrupted copy denies migration without resetting refresh safety. Bootstrap topology revocation is a captured owner/generation/digest-bound, safety-reducing registry commit under the same registry lock and independent of cancelled request grant authority; no partially mutated connection/grants are published by that branch.

The roster connection asks only for `operator.read`; the public proof CLI uses broader default connection scopes despite its read-only RPC. Directory subprocess cancellation is individually bounded, not caller-signal-immediate. The additional 90-second timer applies after tool entry, excluding initial hook/account admission; it is not an unchanged strict whole-wall-clock operation budget.
