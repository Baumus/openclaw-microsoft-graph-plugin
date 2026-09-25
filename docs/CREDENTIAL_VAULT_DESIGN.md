# Shared credential vault design

## Contract

Version 3 uses one Microsoft delegated OAuth credential for every operation, one encrypted vault record, one lock, one authenticated quarantine marker, and one SecretRef-only key at `credentialVaultKey`. Policy contains no credential locations and remains default deny.

Normal call order is strict:

1. Validate the tool action and trusted caller context.
2. Complete required chat confirmation or OpenClaw approval.
3. Validate policy and authorize the exact agent, operation, and resource.
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

All scope exchanges serialize under the same fail-closed lock. Before OAuth dispatch, the plugin durably creates and reads back an HMAC-authenticated `in_flight` refresh-state marker while verifying that the lock remains held. Existing `in_flight` or `quarantined` state blocks dispatch. In vault mode, only HTTP 400 and 401 token-endpoint responses are definitive authentication failures; every other non-2xx response after dispatch is uncertain. A definitive OAuth authentication failure or a fully verified durable success clears the marker; uncertain post-dispatch outcomes transition it atomically to `quarantined`. If transition publication or lock verification fails, the original write-ahead marker remains and later operations fail closed. Status and bound recovery treat either state as quarantined.

OAuth has a bounded request timeout and a 64 KiB response parser. Successful responses require an RFC 6750 `b64token`/header-safe access token of 1 through 16 KiB and an `expires_in` JSON number that is a safe positive integer no greater than 86,400 seconds. Process-local cache residency is capped at one hour and still applies the 60-second pre-expiry skew. A replacement refresh token is written by durable atomic replacement and read back before its access token can enter that cache.

After OAuth dispatch, transport errors, malformed/empty/truncated/oversized HTTP-200 bodies, invalid token or expiry fields, and uncertain local publication retain or quarantine the authenticated marker bound to the observed vault generation and digest. Later calls fail with `credential_reauthorization_required` before exchange. Recovery requires the current secret-free binding from `status`; reauthorization is safer whenever Microsoft may have rotated the remote token.

## Migration and rollback

The local CLI preserves exact interactive confirmation for apply operations, then sends a closed parameter object to a plugin-owned Gateway RPC. The active Gateway executes against its materialized plugin config and returns only a validated sanitized envelope. Status is scoped to `operator.read`; migration, restore, and recovery (including dry-runs) require `operator.admin`. These methods remain registered while `config.enabled` is `false`, allowing credential setup before Microsoft Graph tools are enabled. Dry-runs perform no vault, lock, pass-destination, or quarantine mutation.

`migrate-from-pass` requires `--source <pass-ref>`. It validates that this one selected credential covers every scope implied by current policy grants. It never merges credentials. Dry-run performs validation without creating the vault; apply uses create-only publication and emits a secret-free receipt.

`restore-pass` requires `--destination <pass-ref>`. It writes and reads back the one current credential while vault operation remains unchanged. `unknown` means the local outcome could not be verified. No pass entry is deleted automatically.

Rollback cannot undo Microsoft-side refresh-token rotation or prove freshness of an offline encrypted backup. Vault generations have no external monotonic anchor, so replay of an older valid record with its matching key remains an operator-managed risk. Keep encrypted record/key backups separate, protect pass destinations, and prefer reauthorization after uncertainty.

## Security tradeoff

The shared credential reduces operational and migration complexity. It does not provide provider-side read/write isolation: compromise of the credential, key, or a plugin bypass has the union of Microsoft scopes consented to the shared credential. Default-deny policy and write approvals remain the authority for ordinary plugin operations.
