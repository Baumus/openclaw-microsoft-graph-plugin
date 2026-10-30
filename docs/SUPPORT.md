# Support and limitations

## Support scope

This is a community-maintained plugin. Support is best-effort through the repository's GitHub issue tracker. Use [SECURITY.md](../SECURITY.md) for vulnerabilities and never attach credentials or private Microsoft data.

Version 3.5.0 declares OpenClaw host and plugin API compatibility from `2026.9.6` and is built/validated for that version. OpenClaw plugin APIs are experimental; future host versions may require changes.

## Intentional limitations

- Microsoft Graph v1.0 only; beta-only fields and endpoints are excluded.
- Delegated user permissions only; no application permissions or service-principal daemon flow.
- One signed-in account per policy/credential pair.
- Fixed `/me` workload routes; no shared-mailbox impersonation.
- Contacts tools cover only the signed-in user's default `/me/contacts` collection, not contact folders, shared contacts, or the organization directory. Client-side contact search scans at most 500 records and can report incomplete results.
- Existing credential-vault records are not automatically re-consented for new Contacts scopes. After saving and applying policy, use **Renew Microsoft permissions** in the Control UI. It reuses the existing app/tenant; no terminal commands or new app registration are required. Initial sign-in still refuses to overwrite a nonempty vault.
- No arbitrary Graph URL or raw payload tool.
- No sharing/permissions management, subscriptions, delta feeds, retention, legal hold, permanent deletion, OneDrive conversion, or unrelated Graph workloads.
- OneDrive content writes require OpenClaw private inbound media. Inline bytes and host-local paths are rejected.
- Calendar multiwrite is capped at 100 operations, non-atomic, and imposes no rollback.
- Search completeness can be limited by Microsoft Graph provider limits and bounded local scan budgets.
- Process-local continuation handles and warning allow-always trust are lost on process restart; warning trust also resets on plugin reload.
- The credential vault requires a local filesystem whose ownership, private permissions, regular-file identity, locking, rename, and durability behavior can be verified by `@openclaw/fs-safe`; unsupported or ambiguous filesystems fail closed.
- Vault encryption does not protect against an OS account or process that can both resolve the key and read or replace plugin state/code. Back up vault files and SecretRef keys independently.
- Vault generation is not externally anchored. Replaying an older valid encrypted backup with its matching key cannot be detected cryptographically; protect backup history and compare sanitized bindings before restore.
- A lost response or abort after OAuth refresh dispatch has an unknowable provider outcome. New authenticated OAuth-only uncertainty permits one recovery after a 60-second cooldown on a later authorized operation, with a durable attempt budget. Legacy markers, exhausted attempts, crashes and uncertain local publication still require reviewed operator recovery.
- Abrupt process or host loss leaves the durable in-flight marker fail-closed; no local mechanism can prove what the provider committed in that window. Recovery never replays Graph data operations.
- The optional OneDrive `AGENTS.md` extension is disabled unless a policy root explicitly opts into `agents_instructions: trusted`. It uses the OpenClaw instruction-file format for remote, policy-pinned content and does not read or expose the host agent's local workspace instructions.
- The package does not configure Microsoft Entra, consent permissions, GNU `pass`, SecretRef providers, OpenClaw tool policy, backups, or production monitoring.

## Size and timeout boundaries

- OneDrive individual-file ceiling: 250 GB, subject to Microsoft account and service limits.
- OneDrive simple upload: through 250 MB; larger writes use sequential upload sessions.
- Outlook mail/calendar file attachments: at most 150 MB.
- Microsoft To Do attachments: at most 25 MB.
- Model-facing text and To Do attachment output: 256 KiB by default, 1 MiB maximum.
- All provider requests and whole operations have configurable bounded timeouts.

Large transfers remain subject to host storage, network, Microsoft throttling, provider behavior, and OpenClaw private-media lifecycle.

## Not promised

The project does not claim official Microsoft or OpenClaw support, enterprise readiness, regulatory compliance, security certification, uninterrupted operation, or fitness for a particular tenant policy. Operators must perform their own review, testing, backup, incident-response, and access-governance work.


## Guided renewal troubleshooting

- In German Entra: **Entra ID → App-Registrierungen → existing app → Übersicht**. **Anwendungs-ID (Client-ID)** identifies the app; **Verzeichnis-ID (Mandant)** identifies the tenant. **Objekt-ID** is neither. Switch to the correct directory first. Renewal does not need these values to be copied.
- Initial setup needs an approved public client and delegated permissions. Required administrator consent and conditional-access rules are not bypassed by sign-in.
- `credential_vault_conflict_identity`: explicit app/tenant differs from the vault. Use the UI renewal action; do not register another app.
- `credential_vault_conflict_exists`: initial connection found a vault. Refresh status and choose renewal instead.
- `credential_vault_conflict_changed`: vault generation/binding changed while consent was pending (for example a normal refresh rotated a token). The current credential is retained. Refresh status, then start a fresh reviewed renewal. Do not automatically replay the old result.
- `credential_vault_conflict_policy_changed`: applied scopes changed during consent. Review the active rules before restarting.
- `credential_refresh_retry_later`: an eligible OAuth-only failure is cooling down; no Graph request was sent. A fresh authorized operation after at least 60 seconds may make its single recovery attempt. Do not automatically replay data writes.
- `credential_refresh_temporarily_unavailable`: Microsoft definitively rejected the token request with a transient error; no Graph request was sent. Wait before a fresh authorized operation.
- `credential_vault_conflict_quarantined`: authenticated in-flight/uncertain refresh state blocks renewal. An administrator must review recovery; this UI does not clear the marker or delete credentials.
- Generic `credential_vault_conflict` remains a safe fallback for an unexpected pre-publication race. No raw provider response, token, credential, or vault content is included in these errors.

Terminology/protocol checked against primary Microsoft documentation on 2026-10-21:
[Register an app (German)](https://learn.microsoft.com/de-de/entra/identity-platform/quickstart-register-app),
[Device authorization grant (German)](https://learn.microsoft.com/de-de/entra/identity-platform/v2-oauth2-device-code).
