# Microsoft Graph for OpenClaw

`@baumus/openclaw-microsoft-graph` is a community-maintained OpenClaw tool plugin for bounded Microsoft Graph v1.0 access to OneDrive, Outlook Calendar, Outlook Mail, and Microsoft To Do. It is not an official Microsoft or OpenClaw product and carries no enterprise, compliance, or security certification.

## Security model

- Every operation is checked against a credential-free, default-deny per-agent policy before plugin-side key selection, vault I/O, OAuth, or Graph access.
- Warning-level mutations require OpenClaw-native call approval by default. Operators may explicitly set `warningApprovalsRequired: false`; critical delete, send, and respond operations always retain native call-bound approval.
- One shared delegated OAuth credential serves every authorized read and write operation. The plugin requests only the operation-specific scope for each exchange.
- One AES-256-GCM encrypted vault record, one fail-closed lock, and one authenticated quarantine marker protect the rotating refresh-token lifecycle.
- Access-token caching is bounded and transactional: a token associated with a rotated refresh token is admitted only after durable vault replacement and verification.
- Every refresh durably publishes authenticated `in_flight` state before OAuth dispatch; uncertainty transitions it to `quarantined`, and either state blocks another exchange until verified completion, explicit recovery, or reauthorization.
- HTTP-200 OAuth responses require a 1–16 KiB RFC 6750 `b64token`/header-safe access token and a positive integer `expires_in` no greater than 86,400 seconds. Cache residency is capped at one hour with a 60-second skew.

This simpler shared-credential model has an explicit tradeoff: compromise of the credential, vault key, or a plugin bypass can expose the union of Microsoft delegated scopes granted to that credential. The policy and write-approval gates constrain normal plugin behavior; they are not provider-side credential isolation.

See [architecture](docs/ARCHITECTURE.md), [OAuth access matrix](docs/OAUTH_ACCESS_MATRIX.md), [vault design](docs/CREDENTIAL_VAULT_DESIGN.md), and [support](docs/SUPPORT.md).

## Setup guide

Version 3.4.0 uses one Microsoft delegated OAuth credential and one encrypted local vault. The policy contains authorization rules only; it never contains credential locations or credential material.

### 1. Check prerequisites

You need:

- Node.js `>=24.16.0 <25` or `>=26.1.0`.
- OpenClaw `>=2026.9.6` running as the OS account that will own the plugin state.
- A Microsoft Entra app registration and one delegated user grant for the account the plugin will use. Application permissions, client secrets, certificates, and daemon/service-principal flows are not supported.
- GNU `pass` and its GPG setup for the one-time migration source (and optional rollback destination). Version 3 does not read `pass` during normal tool calls.
- Permission to edit the OpenClaw configuration and to review the third-party plugin's declared capabilities.

Check the local versions before continuing:

```bash
node --version
openclaw --version
```

### 2. Prepare the Microsoft delegated credential

1. Register an application in Microsoft Entra and record its **Application (client) ID**. Select the tenant/account audience appropriate for your organization.
2. Configure a public-client redirect or device-code flow in accordance with your tenant policy. This plugin has no initial sign-in flow and does not need or accept a client secret.
3. Request delegated Microsoft Graph consent only for operations enabled by your policy. Also request `offline_access` so the authorization flow returns a refresh token. Use the [OAuth access matrix](docs/OAUTH_ACCESS_MATRIX.md) to calculate the set. For example, a policy that permits every workload's read and write operations needs the applicable `Files.ReadWrite`, `Calendars.ReadWrite`, `Mail.ReadWrite`, `Mail.Send`, and `Tasks.ReadWrite` grants; a read-only policy should use the corresponding read scopes instead.
4. Complete authorization with an operator-approved OAuth client and securely capture the resulting refresh token. Record the exact scopes represented by the grant. Tenant conditional-access and consent rules remain authoritative.

The migration source is one JSON document stored under a single `pass` reference:

```json
{
  "clientId": "replace-with-application-client-id",
  "tenant": "replace-with-tenant-id-or-common",
  "refreshToken": "replace-with-refresh-token",
  "scopes": ["Files.ReadWrite", "Calendars.ReadWrite", "offline_access"]
}
```

Create the entry interactively with `pass insert -m <pass-ref>`. Do not place this JSON in the OpenClaw config or policy, and never commit it or paste it into chat, issues, logs, screenshots, or test fixtures. Migration validates that this one credential covers every scope implied by the policy; it does not merge separate read and write credentials.

### 3. Install version 3.4.0

Version 3.4.0 is not yet available from npm. After publication, install the exact reviewed package and version:

```bash
openclaw plugins install npm:@baumus/openclaw-microsoft-graph@3.4.0 --pin
```

Review the package source, integrity, and declared capabilities before accepting the interactive consent prompt. Installation does not create credentials, consent Microsoft permissions, grant tool access, or make an incomplete configuration usable. OpenClaw may leave the plugin disabled until its required configuration is present.

### 4. Create the vault-key SecretRef

`credentialVaultKey` accepts a structured OpenClaw SecretRef only. A plaintext string in `openclaw.json` is rejected. The resolved value must be exactly 32 random bytes encoded as canonical, unpadded base64url (43 characters).

The built-in team secret store is the simplest supported provider. The following command generates the key without printing it and writes it as a secret-kind value:

```bash
node -e 'process.stdout.write(require("node:crypto").randomBytes(32).toString("base64url"))' \
  | openclaw secrets store set MICROSOFT_GRAPH_CREDENTIAL_VAULT_KEY --kind secret --value-file -
```

Reference it with the complete object shape:

```json5
credentialVaultKey: {
  source: "store",
  provider: "default",
  id: "MICROSOFT_GRAPH_CREDENTIAL_VAULT_KEY",
}
```

You may instead use an `env`, `file`, or `exec` SecretRef backed by a correctly configured OpenClaw secret provider. Do not replace the object with the resolved value. Back up the SecretRef provider independently from the encrypted vault; losing or changing this key makes the existing vault unreadable.

### 5. Create a version-2 default-deny policy

Copy [the shipped v2 example](examples/microsoft-graph-policy-v2.json5) to `microsoft-graph-policy.json5` beside `openclaw.json`, then replace every placeholder. `$include` paths are resolved relative to the file that contains them and normally must remain inside the top-level config directory.

A conservative OneDrive read-only policy looks like this:

```json5
{
  version: 2,
  rules: { default: "deny" },
  services: {
    onedrive: {
      allowed_roots: [{
        label: "documents",
        path: "/Documents",
        drive_id: "replace-with-immutable-drive-id",
        item_id: "replace-with-immutable-root-item-id",
        include_descendants: true,
        permissions: { read: true, write: false, delete: false },
        agents: {
          "replace-with-openclaw-agent-id": {
            permissions: { read: true, write: false, delete: false },
          },
        },
      }],
    },
    calendar: {
      agents: {
        // "replace-with-openclaw-agent-id": { operations: ["read"], resources: ["me"] },
      },
    },
    mail: { agents: {} },
    todo: { agents: {} },
  },
}
```

Use agent IDs from the OpenClaw host configuration; the plugin does not infer policy grants from local workspace content. For OneDrive, obtain the immutable drive and item IDs with Microsoft Graph Explorer or another tenant-approved administrative client; the display path alone is not an authorization boundary. Calendar grants may use `me` or exact calendar IDs. Mail and To Do use fixed signed-in-user `/me` routes. Add only the operations and resources you intend to permit, and add the corresponding delegated scopes to the one credential.

Keep `rules.default: "deny"`. `include_descendants` must be literal `true`; remove a root entirely if descendants must not be available. Leave `agents_instructions` absent unless trusted administrators control every writer to that root and you intentionally set `agents_instructions: "trusted"`. This optional plugin feature applies OpenClaw's standard `AGENTS.md` instruction format to policy-pinned OneDrive content; it does not read, copy, or modify the host agent's local workspace `AGENTS.md`.

### 6. Configure and enable the plugin

Add the following shape to `openclaw.json`. If you already use `tools.allow`, merge `microsoft-graph` into that existing allowlist instead of replacing unrelated entries.

```json5
{
  plugins: {
    entries: {
      "microsoft-graph": {
        enabled: true,
        config: {
          enabled: true,
          warningApprovalsRequired: true,
          credentialVaultKey: {
            source: "store",
            provider: "default",
            id: "MICROSOFT_GRAPH_CREDENTIAL_VAULT_KEY",
          },
          policy: { $include: "./microsoft-graph-policy.json5" },
        },
      },
    },
  },
  tools: { allow: ["microsoft-graph"] },
}
```

`$include` is OpenClaw host-config composition. The plugin receives the resolved policy object and does not read policy files itself. If the plugin remains disabled after saving valid configuration, enable it explicitly and accept capabilities only after review:

`warningApprovalsRequired` defaults to `true` when omitted. Set it to `false` only when policy-authorized warning-level mutations should run without an approval prompt. This setting never affects critical approvals and never bypasses policy authorization, parameter validation, managed-root instruction checks, protected-media checks, or write preconditions.

```bash
openclaw plugins enable microsoft-graph
```

### 7. Migrate the selected credential into the vault

First inspect status, then dry-run the one selected source. Substitute your own `pass` reference locally; it is intentionally not echoed in receipts.

```text
openclaw microsoft-graph credentials status
openclaw microsoft-graph credentials migrate-from-pass --source <pass-ref> --dry-run
openclaw microsoft-graph credentials migrate-from-pass --source <pass-ref> --apply
```

The dry run validates the policy, key, source credential, and required scope coverage without creating the vault. `--apply` is create-only, requires an interactive terminal, and asks you to type `MIGRATE MICROSOFT GRAPH CREDENTIAL` exactly. A successful receipt reports `result: "created"` plus sanitized generation, key ID, digest, binding, and timestamp fields. A second migration does not overwrite an existing vault.

Credential commands call plugin-owned RPC methods on the active Gateway so SecretRefs are resolved only in the Gateway's materialized configuration. The Gateway may perform migration while this plugin's `config.enabled` is `false`; ordinary Microsoft Graph tools remain disabled until it is set to `true`. Status requires `operator.read`; migration, recovery, and restore require `operator.admin`. RPC parameters and responses are closed, validated shapes and never include the vault key, pass contents, OAuth tokens, or raw operation errors.

Do not delete the source entry until you have completed validation and established your backup/recovery plan. Version 3 normal operation uses only the encrypted vault and no longer needs `pass`.

### 8. Validate the setup

Run the local checks in this order:

```bash
openclaw config validate
openclaw plugins doctor --json
openclaw plugins inspect microsoft-graph --runtime --json
openclaw secrets audit --check
openclaw microsoft-graph credentials status
```

Expected credential status is `result: "valid"` with secret-free metadata. Then, from an agent explicitly granted in policy, make one read-only request against an allowlisted resource and confirm that an ungranted agent or resource is denied. Do not start validation with a write or destructive action. `plugins list` or a cold manifest inspection alone does not prove that the running Gateway registered the plugin.

### 9. Understand approvals

| Class | Examples | Required user action |
| --- | --- | --- |
| none | Explicitly recognized read actions | No mutation approval; policy authorization still applies. |
| warning | Create/update/draft/move/mark operations | By default, use OpenClaw's native approval and choose `allow-once`, `allow-always`, or `deny`. With `warningApprovalsRequired: false`, no approval is requested. |
| critical | Delete, send, and calendar respond operations | Use OpenClaw's native call-bound approval and choose `allow-once` or `deny`. Warning configuration and legacy chat fields cannot downgrade or replace this approval. |

For warning requests, `allow-always` trusts only the authenticated agent ID, exact tool name, and normalized action (for example, `outlook_mail_write` + `mark_read`). It does not trust a resource, arbitrary future action, or another agent/tool/action. Trust is held only in plugin process memory and is revoked by plugin reload or process restart; it is not written to OpenClaw config or disk. Every trusted future call still runs authorization, validation, managed-root instruction discovery, protected-media validation, and execution preconditions. Approval-free trusted/configured warning calls also require the host's call identity so the plugin can bind and reverify their exact execution parameters. Critical calls never offer `allow-always`.

The deprecated `chatConfirmed` and `chatConfirmationToken` tool fields remain accepted for compatibility but are ignored and can never authorize execution. The plugin returns the exact parameters it inspected with each native approval so OpenClaw freezes that snapshot while approval is pending. It also binds the host tool-call identity to that snapshot and consumes it at execution; any later composed rewrite fails closed before plugin execution. If no approval route is available, or approval is denied, cancelled, malformed, or times out, the host blocks the call.

Native approval text identifies the mutation action, a minimized target, and the relevant risk without including message bodies, event bodies, subjects, recipient addresses, or To Do titles. OneDrive upload/update approvals include the allowlisted root, relative path, required SHA-256, and byte size; preflight and execution each securely open the protected artifact and fail before their downstream credential or Graph boundary if that exact content identity does not match. Calendar approvals include calendar/event identity, and multiwrite approvals include the operation count. Mail send approval identifies that recipients come from the stored draft and explicitly notes that the recipient count is unavailable from the send call itself.

OneDrive mutation authorization occurs before a warning approval request or approval-free warning continuation. For upload/update, the protected artifact's opened-file identity, exact lowercase SHA-256, and byte size are verified before any applicable managed-root instruction discovery. Discovery requires its own read authority before credential or Graph access. Execution independently reopens and revalidates the artifact plus policy authorization and all media/write preconditions before selecting the key, reading the vault, exchanging OAuth, or calling Graph.

### 10. Status, recovery, and rollback

`openclaw microsoft-graph credentials status` returns one of:

- `missing`: no vault exists; run the migration dry-run and apply steps.
- `valid`: the vault decrypts and has no authenticated refresh marker.
- `quarantined`: a dispatched refresh has an uncertain outcome. Normal exchanges remain blocked.
- `unavailable`: the key, record, permissions, ownership, filesystem, or authenticated envelope could not be validated. Stop and investigate; do not overwrite the vault.

For `quarantined`, reauthorization is the safest recovery whenever Microsoft may have rotated the refresh token. Only if an operator has independently established that the current vault credential remains usable should they copy the current **sanitized binding** from `status` and run:

```text
openclaw microsoft-graph credentials recover-refresh --expected-binding <binding> --dry-run
openclaw microsoft-graph credentials recover-refresh --expected-binding <binding> --apply
```

Apply requires the exact typed confirmation `RECOVER MICROSOFT GRAPH REFRESH`. The binding prevents recovery against a changed record. Do not blindly repeat a refresh or recovery after an uncertain result.

To create a verified emergency copy in a separate `pass` destination without changing vault operation:

```text
openclaw microsoft-graph credentials restore-pass --destination <pass-ref> --dry-run
openclaw microsoft-graph credentials restore-pass --destination <pass-ref> --apply
```

Apply requires `RESTORE MICROSOFT GRAPH CREDENTIAL`. A `complete` receipt means the destination was written and read back; `unknown` means the local result could not be verified and must be inspected before retrying. The plugin never deletes a `pass` entry automatically.

### Safety notes

- Keep the SecretRef provider, vault files, and any `pass` backup private and backed up separately. Never store the key beside an exported vault backup.
- Vault generations have no external monotonic anchor. Replaying an older valid encrypted record with its matching key cannot be detected cryptographically; compare sanitized bindings and prefer reauthorization after uncertainty.
- One shared credential has the union of its consented Microsoft scopes. Policy and approval gates constrain normal plugin use but do not provide provider-side read/write isolation after credential, key, host-account, or plugin compromise.
- OpenClaw may resolve declared SecretInputs while loading configuration. The plugin guarantees authorization before plugin-side key selection and vault/provider access, not suppression of host-level SecretRef materialization.
- Unknown actions, malformed resources, missing grants, missing credentials, unsafe files, and unsupported filesystem conditions fail closed.
- Process-local access-token, continuation, instruction, and warning allow-always trust is lost on restart. Retry from a fresh read/status check rather than assuming an interrupted mutation failed.

## Configuration UI (OpenClaw 2026.9.6+)

Administrators can enable **Settings → Labs → Custom plugin UI**, then open **Microsoft Graph** in the Control UI. Version 3.4.0 offers English, German, Spanish, and Arabic; other host locales fall back to English. The page edits OneDrive folder/agent rights, per-agent Calendar/Mail/To Do access, and warning-level approval choices for each service. New OneDrive paths are resolved to immutable drive/item IDs by an admin-only Gateway method. Critical delete, send, and respond actions always retain call-bound approval.

The page validates the policy and submits a revision-checked, policy-only `config.patch`. A supported single-file object-key `$include` is written through by OpenClaw; an unsupported include layout or concurrent change fails closed. The page re-reads the effective policy after saving and shows whether the saved configuration revision has been applied by the Gateway, is still pending, or cannot be confirmed. Pending application is checked automatically for up to one minute, with a manual recheck available. Do not treat “saved” as proof that the Gateway is using the new rules.

When new delegated scopes would be required, the page blocks saving while the plugin is enabled until Microsoft consent is verified through the separate operator workflow. Credential migration and recovery remain in the interactive CLI; no credential is entered in the browser. Generic plugin settings and CLI paths remain available when Custom plugin UI is disabled.

## Development

```bash
npm ci
npm test
npm run typecheck
npm run build
npm run plugin:build:check
npm run plugin:validate
npm run package:check
npx --no-install clawhub package validate .
```

Tests use synthetic fixtures and mocked provider boundaries. They must not use real credentials or call Microsoft Graph.

## License

Apache-2.0. See [LICENSE](LICENSE).
