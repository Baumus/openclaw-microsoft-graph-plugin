# Shared credential vault design

## Contract

Version 3 uses one Microsoft delegated OAuth credential for every operation, one encrypted vault record, one lock, one authenticated quarantine marker, and one SecretRef-only key at `credentialVaultKey`. Policy contains no credential locations and remains default deny.

Normal call order is strict:

1. Validate the tool action and trusted caller context, including available fail-closed preflight authorization and media checks.
2. Complete any required OpenClaw-native approval. Warning approval is configurable and defaults on; critical approval is mandatory.
3. Revalidate policy and authorize the exact agent, operation, and resource during execution.
4. Only then select the materialized key, acquire the lock, read/decrypt the vault, exchange OAuth, and call Graph.

Denied operations perform no plugin-side key selection, vault I/O, OAuth, or Graph request. OpenClaw may materialize declared SecretInputs while loading host configuration; this design does not claim to suppress host-level resolution.

## Storage

```text
<state>/plugin-data/microsoft-graph/credentials/
  .initialized
  credential.vault.json
  credential.vault.json.lock
  credential.quarantine.json   # authenticated write-ahead refresh state
```

The credential directory and files are private (`0700` and `0600` on POSIX). Secure reads are bounded and timed. Symlinks, hardlinks, unsafe ownership/filesystem conditions, malformed or duplicate JSON fields, non-canonical base64url, oversized input, and authentication failures fail closed.

The vault envelope is strict JSON containing format/version, generation, key ID, algorithm, nonce, ciphertext, and GCM tag. AES-256-GCM additional authenticated data binds the format, version, generation, algorithm, and key ID. The plaintext credential contains client ID, tenant, refresh token, and delegated scopes.

## Rotation, cache, and quarantine

All scope exchanges serialize under the same fail-closed multi-process lock. Before OAuth dispatch, the plugin durably creates and reads back an HMAC-authenticated v3 `in_flight` refresh-state marker while verifying that the lock remains held. The marker binds generation, key ID and exact vault digest, plus a reason and a `recoveryUsed` bit. Successful replacement first publishes a non-recoverable `publication` phase, durably replaces the credential and reads it back before clearing the marker and admitting the access token.

A definitive transient HTTP 429/500/502/503/504 token-endpoint rejection (or a bounded OAuth `server_error` / `temporarily_unavailable` rejection on HTTP 400/401) returns `credential_refresh_temporarily_unavailable` without quarantining an ordinary refresh. No internal retry or sleep is added. Definitive `invalid_grant`, `interaction_required`, `login_required`, `consent_required`, `invalid_client` or `unauthorized_client` returns `credential_reauthorization_required`; the verified OAuth-only marker is cleared so existing guided reauthorization remains available. Other HTTP 400/401 rejections remain sanitized authentication failures.

A transport loss or unusable successful OAuth response transitions only the exact unchanged vault to `quarantined` / `oauth_uncertain`. After a 60-second cooldown, a later already-authorized operation may claim exactly one recovery exchange under the same lock, first atomically publishing and verifying `in_flight` / `oauth_pending` with `recoveryUsed: 1`. It does not clear the marker before dispatch. A failed or interrupted claim, crash during recovery, second ambiguous outcome or transient recovery failure cannot reset that budget. A successful exchange follows the ordinary durable publication and cache-admission path; failure leaves exhausted or in-flight state for reviewed operator handling. Clock rollback delays recovery; no timer, daemon, immediate loop, credential-scope expansion or Graph data-operation replay exists. Cancellation is checked before locking and dispatch; cancellation after dispatch is treated as uncertain. A known cancellation before initial dispatch clears its unused marker, while an already-claimed recovery remains consumed. Each call has at most one exchange with the existing bounded request timeout (maximum 30 seconds), plus bounded secure-file I/O and lock acquisition; there is no nested recovery deadline or extra exchange.

Only authenticated v3 `quarantined` / `oauth_uncertain`, `recoveryUsed: 0` state matching all current binding fields is eligible. Legacy v1/v2 markers, unfinished `in_flight` markers, local write/readback/lock failures, publication phases, changed bindings and corrupt or wrong-key state never recover automatically. Legacy state cannot prove which side of publication failed. A package downgrade to a v1/v2-only parser rejects v3 markers fail-closed; never strip fields or delete state to bypass the durable retry budget.

Microsoft's [refresh-token lifecycle documentation](https://learn.microsoft.com/en-us/entra/identity-platform/refresh-tokens) says that issuing a new pair does not automatically revoke the previous refresh token. This provider-specific property supports the one bounded OAuth-only recovery, not fallback to a previous local secret or recovery from uncertain storage publication. Tests use synthetic credentials and responses only; no live account or mailbox is used.

OAuth has a bounded request timeout and a 64 KiB response parser. Successful responses require an RFC 6750 `b64token`/header-safe access token of 1 through 16 KiB and an `expires_in` JSON number that is a safe positive integer no greater than 86,400 seconds. Process-local cache residency is capped at one hour and still applies the 60-second pre-expiry skew. A replacement refresh token is written by durable atomic replacement and read back before its access token can enter that cache.

After OAuth dispatch, transport errors, malformed/empty/truncated/oversized HTTP-200 bodies, invalid token or expiry fields, and uncertain local publication retain or quarantine the authenticated marker bound to the observed vault generation and digest. Non-eligible later calls fail with `credential_reauthorization_required` before exchange; eligible state in cooldown returns `credential_refresh_retry_later`. Reviewed manual recovery still requires the current secret-free binding from `status`. Never invoke ordinary reauthorization against an unresolved quarantined vault.

## Sign-in and rollback

The local CLI preserves exact interactive confirmation for apply operations, then sends a closed parameter object to a plugin-owned Gateway RPC. The active Gateway executes against its materialized plugin config and returns only a validated sanitized envelope. Status is scoped to `operator.read`; device sign-in, cancellation, restore, and recovery (including dry-runs) require `operator.admin`. These methods remain registered while `config.enabled` is `false`, allowing credential setup before Microsoft Graph tools are enabled. Dry-runs perform no vault, lock, pass-destination, or quarantine mutation.

`sign-in --client-id <app-id> --tenant <tenant-id>` requests policy-derived delegated scopes plus `offline_access` using the Microsoft device-code flow. The Gateway owns polling and the refresh token; the Control UI receives only the one-time verification instruction and sanitized status. A missing or incomplete grant, failed consent, invalid vault, or existing vault fails closed. Sign-in is create-only.

The Control UI invokes the existing admin-scoped `microsoft-graph.credentials.reauthorize-start` RPC with an empty object to reuse app/tenant directly from the valid vault. Its older exact `{ clientId, tenant }` shape remains compatible for explicit-ID callers, but is not required by the UI. Status compares the stored Microsoft-returned scopes with the active policy and exposes only `missingScopes`. Codes, consent, cancellation, polling, and failures stay in the guided UI/Microsoft flow. Before OAuth dispatch, reauthorization verifies a valid, unquarantined vault and matching app/tenant, recording its sanitized binding. After Microsoft returns the full active-policy scope set, a vault lock and binding check guard atomic replacement. The old record remains until consent and validation succeed. An authenticated in-flight marker blocks concurrent token use during publication; on a verified pre-publication failure it is cleared, while an ambiguous publication remains fail-closed. A changed policy or vault binding aborts replacement. The old vault contains no independently verified account ID, so the operator must select the same account on Microsoft's page.

`restore-pass` requires `--destination <pass-ref>`. It writes and reads back the one current credential while vault operation remains unchanged. `unknown` means the local outcome could not be verified. No pass entry is deleted automatically.

Rollback cannot undo Microsoft-side refresh-token rotation or prove freshness of an offline encrypted backup. Vault generations have no external monotonic anchor, so replay of an older valid record with its matching key remains an operator-managed risk. Keep encrypted record/key backups separate, protect pass destinations, and prefer reauthorization after uncertainty.

## Security tradeoff

The shared credential reduces operational complexity. It does not provide provider-side read/write isolation: compromise of the credential, key, or a plugin bypass has the union of Microsoft scopes consented to the shared credential. Default-deny policy and write approvals remain the authority for ordinary plugin operations.
