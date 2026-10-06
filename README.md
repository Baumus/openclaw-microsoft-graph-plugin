# Microsoft 365 for OpenClaw

Your Outlook messages, calendar, OneDrive files, and Microsoft To Do tasks are part of your day. Bring the parts you choose into OpenClaw so your agents can help you find information and work with it—under access rules you control.

**Microsoft 365 for OpenClaw** is a community-maintained plugin for one signed-in Microsoft user. Grant named agents specific access to Outlook Mail, Outlook Calendar, OneDrive, and Microsoft To Do. Access is denied by default. Changes require OpenClaw approval by default; sending, deleting, and responding always require call-bound approval. An administrator can turn off warning-level approval, so review that setting before granting write access.

For example, an authorized agent can find a file in an approved OneDrive folder, check an upcoming event, find an Outlook message, or review permitted tasks. The plugin does not provide Teams, SharePoint, arbitrary Microsoft Graph access, or a service-principal connection.

**Start here:** [Check requirements and install from ClawHub](#1-check-requirements-and-install), then [choose access](#2-choose-access-and-connect-microsoft), then [verify a read-only request](#3-verify-your-first-request).

This is not an official Microsoft or OpenClaw product. It carries no enterprise, compliance, or security certification. [Review supported behavior and limits](docs/SUPPORT.md) before using it with important data.

## Get started in three steps

### 1. Check requirements and install

You need a supported OpenClaw host (`>=2026.9.6`), Node.js (`>=24.16.0 <25` or `>=26.1.0`), permission to edit its configuration, and an approved Microsoft Entra public-client app registration for the account you will connect. Check your local versions:

```bash
node --version
openclaw --version
```

Review the [ClawHub listing](https://clawhub.ai/packages/@baumus/openclaw-microsoft-graph), [source](https://github.com/Baumus/openclaw-microsoft-graph-plugin), and declared capabilities. Then install the published package:

```bash
openclaw plugins install clawhub:@baumus/openclaw-microsoft-graph
```

Installation alone does not sign you in, grant an agent access, or make an incomplete configuration usable. Existing installations should review changes before updating; this guide does not silently upgrade them.

### 2. Choose access and connect Microsoft

Create the vault-key SecretRef and a default-deny policy for named agents. Begin with one read-only grant. OneDrive access requires immutable drive and item IDs, not just a folder name. Configure and enable the plugin, then follow the setup checklist in **Plugins → Microsoft 365 for OpenClaw**. It shows the next actionable prerequisite before sign-in. Connect the approved Microsoft account only after the policy is saved and applied. Review the Microsoft permissions shown during sign-in.

The [detailed configuration guide](#configuration-reference) has the exact key command, policy example, host configuration, and sign-in steps. This plugin uses delegated access for one signed-in user; it does not support application permissions, client secrets, or unattended service-principal access.

### 3. Verify your first request

Wait until the Control UI confirms that the Gateway has applied your rules. Run the [local checks](#validate-the-setup), then ask an explicitly allowed agent for one read-only item within its grant. Confirm that an ungranted agent or resource is denied. A visible installed-plugin entry alone does not prove the running Gateway can serve the request.

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

## Configuration reference

Version 3.5.3 uses one Microsoft delegated OAuth credential and one encrypted local vault. The policy contains authorization rules only; it never contains credential locations or credential material.

### Detailed prerequisites

You need:

- Node.js `>=24.16.0 <25` or `>=26.1.0`.
- OpenClaw `>=2026.9.6` running as the OS account that will own the plugin state.
- A Microsoft Entra app registration and one delegated user grant for the account the plugin will use. Application permissions, client secrets, certificates, and daemon/service-principal flows are not supported.
- GNU `pass` and GPG only if you choose the optional emergency `restore-pass` backup. Do not invoke that export when credentials must remain exclusively in the local encrypted vault. Normal setup and tool calls do not use `pass`.
- Permission to edit the OpenClaw configuration and to review the third-party plugin's declared capabilities.

Check the local versions before continuing:

```bash
node --version
openclaw --version
```

### Prepare the Microsoft delegated credential

1. Register an application in Microsoft Entra and record its **Application (client) ID**. Select the tenant/account audience appropriate for your organization.
2. In **Authentication**, enable **Allow public client flows** for your own app registration. This device-code flow needs no redirect URI, HTTPS callback, public endpoint, or client secret. An administrator may need to grant delegated consent first; tenant conditional-access and consent rules remain authoritative.
3. Configure the policy and vault-key SecretRef below. Then open the Microsoft 365 for OpenClaw plugin page in the OpenClaw Control UI and follow **Connect to Microsoft** below. The Gateway derives requested delegated scopes from the configured policy and stores the refresh token directly in the encrypted vault. No token is copied into a file or command argument.

### Create the vault-key SecretRef

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

### Create a version-2 default-deny policy

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

### Configure and enable the plugin

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

### Sign in from the browser UI

Open **Plugins → Microsoft 365 for OpenClaw** in an administrator Control UI session. Save the access rules and wait until the page confirms that the Gateway has applied them. Under **Connect to Microsoft**, enter the approved public application's client ID and tenant ID/domain, then select **Start sign-in**. Open the Microsoft link, enter the one-time code shown on this page, and approve the displayed delegated permissions. The page detects completion automatically; you can cancel before authorization finishes.

The Gateway, not the browser, polls Microsoft's token endpoint, verifies the granted scopes against the applied policy, and writes the refresh token directly to an **empty encrypted local vault**. The device-code flow makes outbound requests to Microsoft; it does not operate or require an inbound OAuth callback service. The UI only receives a one-time code, scoped status, and sanitized errors. It never receives OAuth tokens or the vault key. The code expires within 15 minutes. A CLI alternative remains available from an interactive host terminal:

```text
openclaw microsoft-graph credentials status
openclaw microsoft-graph credentials sign-in --client-id <approved-app-id> --tenant <tenant-id>
```

Tenant consent and conditional-access rules remain authoritative. Sign-in will not overwrite an existing or quarantined vault; use a separately reviewed recovery or reauthorization procedure in that case. `status` requires `operator.read`; sign-in, cancellation, recovery, and optional `restore-pass` require `operator.admin`. Gateway RPC parameters and responses have closed validated shapes and do not include OAuth tokens, the vault key, `pass` contents, or raw errors.

### Validate the setup

Run the local checks in this order:

```bash
openclaw config validate
openclaw plugins doctor --json
openclaw plugins inspect microsoft-graph --runtime --json
openclaw secrets audit --check
openclaw microsoft-graph credentials status
```

Expected credential status is `result: "valid"` with secret-free metadata. Then, from an agent explicitly granted in policy, make one read-only request against an allowlisted resource and confirm that an ungranted agent or resource is denied. Do not start validation with a write or destructive action. `plugins list` or a cold manifest inspection alone does not prove that the running Gateway registered the plugin.

### Understand approvals

Agents can call `microsoft_graph_capabilities` first to see only their effective root labels, service actions/resources, and prerequisites. It never checks sign-in, accesses credentials, or calls Graph; connection and sign-in remain operator-owned. All concrete tools include action, exact-ID discovery, approval, result, and next-step guidance, so no bundled skill is required.

Mutation schemas accept optional `timeoutMs` as **OpenClaw outer transport metadata**, not a Graph parameter or approval. Set at least `180000` for calls that may wait through the native 120-second approval prompt, and include expected execution time up to the host's ordinary 600-second cap. The host default is 90 seconds; `timeoutSeconds` adds 30 seconds but is not part of this plugin's schema. The plugin strips `timeoutMs` before semantic validation, approval snapshot binding, and provider payload construction. Long OneDrive transfers configured for up to 24 hours cannot fit a single ordinary dynamic tool call; if the outer call ends, read back the exact target before retrying. A larger timeout never grants permission or extends an expired approval.

Results retain their previous fields and add `phase`, `code`, `nextAction`, `retrySafety`, and `mutationApplied` (`true`, `false`, or `"unknown"`). A send receipt means Graph accepted the request, **not** recipient delivery; inspect Sent Items before any uncertain retry. Paginated reads set `noResults` only when an empty page is complete; follow `continuation` or narrow a capped query when completeness is partial or unknown. Calendar multiwrite remains non-atomic: inspect each operation and retry only after readback.

| Class | Examples | Required user action |
| --- | --- | --- |
| none | Explicitly recognized read actions | No mutation approval; policy authorization still applies. |
| warning | Create/update/draft/move/mark operations | By default, use OpenClaw's native approval and choose `allow-once`, `allow-always`, or `deny`. With `warningApprovalsRequired: false`, no approval is requested. |
| critical | Delete, send, and calendar respond operations | Use OpenClaw's native call-bound approval and choose `allow-once` or `deny`. Warning configuration and legacy chat fields cannot downgrade or replace this approval. |

For warning requests, `allow-always` trusts only the authenticated agent ID, exact tool name, and normalized action (for example, `outlook_mail_write` + `mark_read`). It does not trust a resource, arbitrary future action, or another agent/tool/action. Trust is held only in plugin process memory and is revoked by plugin reload or process restart; it is not written to OpenClaw config or disk. Every trusted future call still runs authorization, validation, managed-root instruction discovery, protected-media validation, and execution preconditions. Approval-free trusted/configured warning calls also require the host's call identity so the plugin can bind and reverify their exact execution parameters. Critical calls never offer `allow-always`.

The deprecated `chatConfirmed` and `chatConfirmationToken` tool fields remain accepted for compatibility but are ignored and can never authorize execution. The plugin returns the exact parameters it inspected with each native approval so OpenClaw freezes that snapshot while approval is pending. It also binds the host tool-call identity to that snapshot and consumes it at execution; any later composed rewrite fails closed before plugin execution. If no approval route is available, or approval is denied, cancelled, malformed, or times out, the host blocks the call. For critical deletion, the host approval client must carry `operator.approvals`; a missing scope is an authorization failure, not Gateway unavailability. The plugin cannot grant this host scope or bypass a failed approval. After repairing the host approval route, request a fresh allow-once approval for the exact deletion.

Native approval text identifies the mutation action, a minimized target, and the relevant risk without including message bodies, event bodies, subjects, recipient addresses, or To Do titles. OneDrive upload/update approvals include the allowlisted root, relative path, plugin-computed SHA-256, and byte size. For an agent-created workspace file, agents make one `onedrive_upload` (or `onedrive_update`) call with `rootLabel`, destination `relativePath`, and `sourceWorkspacePath` relative to their own workspace, for example `reports/onepager.pdf`. The plugin securely stages, hashes, requests native approval, and transfers the file without a separate staging call. Workspace staging is limited to 64 MiB per file and 128 MiB and 64 files in aggregate across gateway processes sharing the same state directory (including concurrent reservations, pending approvals, and orphaned copies); larger files must already exist as private media. Plugin-created copies are removed after denial, cancellation, timeout, failed preflight, or execution. Files already returned as `media://inbound/...` use `sourceMediaUri` instead and are never deleted by this cleanup. Supply exactly one source; `sourceSha256` and `sourceByteSize` are optional paired assertions. Do not construct a `media://` URI from a workspace filename. Preflight and execution each securely open the protected artifact and fail before their downstream credential or Graph boundary if that exact content identity does not match. A failed preflight does not request approval. Calendar approvals include calendar/event identity, and multiwrite approvals include the operation count. Mail send approval identifies that recipients come from the stored draft and explicitly notes that the recipient count is unavailable from the send call itself.

Staged workspace copies live in a plugin-owned inbound subdirectory. On startup and hourly, the plugin reconciles marked runs from stopped processes, including interrupted uploads. Live processes retain their copies; when process identity cannot be checked (for example across hosts), only runs inactive for eight days are eligible. Shared staging-quota locks carry a process identity and are reclaimed when that owner is provably gone; an empty lock left before its owner marker was written is eligible only after eight days. Locks with unverifiable owners or foreign contents are left alone rather than stolen. Unmarked directories, foreign files, and caller-supplied inbound media are left alone. Cleanup failures are retried on later scans.

OneDrive mutation authorization occurs before a warning approval request or approval-free warning continuation. For upload/update, the protected artifact's opened-file identity, exact lowercase SHA-256, and byte size are verified before any applicable managed-root instruction discovery. Discovery requires its own read authority before credential or Graph access. Execution independently reopens and revalidates the artifact plus policy authorization and all media/write preconditions before selecting the key, reading the vault, exchanging OAuth, or calling Graph.

### 10. Status, recovery, and rollback

`openclaw microsoft-graph credentials status` returns one of:

- `missing`: no vault exists; use the browser sign-in above.
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

Administrators can enable **Settings → Labs → Custom plugin UI**, then open **Microsoft 365 for OpenClaw** in the Control UI. The page offers English, German, Spanish, and Arabic; other host locales fall back to English. When a newer security-clean stable release is published on ClawHub, an administrator-only, read-only Gateway check shows an update badge beside the established-connection status. It contacts ClawHub with no Microsoft credentials or policy data, and the badge never installs or activates an update. The page edits OneDrive folder/agent rights, per-agent Calendar/Mail/To Do access, and warning-level approval choices for each service. New OneDrive paths are resolved to immutable drive/item IDs by an admin-only Gateway method. Critical delete, send, and respond actions always retain call-bound approval.

The page validates the policy and submits a revision-checked, policy-only `config.patch`. A supported single-file object-key `$include` is written through by OpenClaw; an unsupported include layout or concurrent change fails closed. The page re-reads the effective policy after saving and shows whether the saved configuration revision has been applied by the Gateway, is still pending, or cannot be confirmed. Pending application is checked automatically for up to one minute, with a manual recheck available. Do not treat “saved” as proof that the Gateway is using the new rules.

When new delegated scopes would be required, the page blocks saving while the plugin is enabled until Microsoft consent is verified through the separate operator workflow. The browser sign-in is a guided Microsoft device-code flow; only the one-time code is displayed. Tokens and the vault key stay on the Gateway. Recovery remains in the interactive CLI. Generic plugin settings and CLI paths remain available when Custom plugin UI is disabled.

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
