# Architecture

## Purpose and boundary

The plugin is one Gateway-owned boundary between optional OpenClaw tools and Microsoft Graph v1.0. It owns policy evaluation, credential retrieval, delegated token exchange, request construction, provider-response validation, bounded output, private-media transfer, and approval gates.

The plugin does not own user intent, business policy, data-retention decisions, Microsoft tenant administration, consent administration, or OpenClaw host security. It does not provide a generic Microsoft Graph proxy.

## Call flow

1. OpenClaw selects an optional plugin tool.
2. The `before_tool_call` hook classifies the action as read-only, warning-level, or critical.
3. Warning-level actions require exact session-bound confirmation from the same originating OpenClaw session. Critical actions require OpenClaw's call-bound approval.
4. The runtime validates action-specific parameters and trusted `toolContext` identity.
5. The plugin validates the default-deny policy supplied in resolved plugin config and authorizes the exact agent, service, operation, and resource.
6. A OneDrive mutation is authorized for its exact requested root/action before optional `AGENTS.md` discovery. Discovery is a separate read and therefore also requires read authority.
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

Local credential commands cross into the active Gateway through plugin-owned, operator-scoped RPC methods so only the Gateway's materialized config reaches vault operations. The read-scoped status method and admin-scoped migration, restore, and recovery methods accept closed parameter objects and return closed sanitized envelopes. Raw errors and secret-bearing values never cross this boundary. Gateway startup registers these methods even while `config.enabled` is `false`; that flag continues to block all Microsoft Graph tools.

### Microsoft Graph

Graph responses, continuation URLs, upload-session URLs, attachment metadata, and downloaded bytes are untrusted provider input. The plugin validates origins, paths, identifiers, lengths, MIME values, paging shape, and response structure before use.

### OneDrive `AGENTS.md`

This optional plugin feature applies OpenClaw's standard `AGENTS.md` instruction format to a policy-pinned OneDrive root. It is separate from, and never imports, the host agent's local workspace instructions. Remote content remains untrusted unless an operator explicitly sets `agents_instructions: trusted` after verifying who can write to that root, and it cannot override higher-priority OpenClaw instructions, host policy, or approval requirements.

## State and data handling

- OAuth access-token, continuation, confirmation, and OneDrive instruction caches are process-local and bounded. The token cache is a 128-entry expiry-aware LRU keyed to durable vault identity, capped at one hour of residency with a 60-second expiry skew.
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

`src/index.ts` is bundled to `dist/index.js` for Node 24 and newer supported host runtimes. Runtime dependencies remain external and are declared in `dependencies`, including exact-pinned `@openclaw/fs-safe` 0.13.1. `openclaw plugins build --check` verifies generated manifest/package metadata without rewriting it. The npm `files` allowlist ships only runtime code, static metadata, consumer documentation, the icon, synthetic JSON5 policy objects, and the optional credential provisioning helper.
