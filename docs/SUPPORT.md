# Support and limitations

## Support scope

This is a community-maintained plugin. Support is best-effort through the repository's GitHub issue tracker. Use [SECURITY.md](../SECURITY.md) for vulnerabilities and never attach credentials or private Microsoft data.

Version 3.5.0 declares OpenClaw host and plugin API compatibility from `2026.9.6` and is built/validated for that version. OpenClaw plugin APIs are experimental; future host versions may require changes.

## Intentional limitations

- Microsoft Graph v1.0 only; beta-only fields and endpoints are excluded.
- Delegated user permissions only; no application permissions or service-principal daemon flow.
- One signed-in account per policy/credential pair.
- Fixed `/me` workload routes; no shared-mailbox impersonation.
- No arbitrary Graph URL or raw payload tool.
- No sharing/permissions management, subscriptions, delta feeds, retention, legal hold, permanent deletion, OneDrive conversion, or unrelated Graph workloads.
- OneDrive content writes require OpenClaw private inbound media. Inline bytes and host-local paths are rejected.
- Calendar multiwrite is capped at 100 operations, non-atomic, and imposes no rollback.
- Search completeness can be limited by Microsoft Graph provider limits and bounded local scan budgets.
- Process-local continuation handles and warning allow-always trust are lost on process restart; warning trust also resets on plugin reload.
- The credential vault requires a local filesystem whose ownership, private permissions, regular-file identity, locking, rename, and durability behavior can be verified by `@openclaw/fs-safe`; unsupported or ambiguous filesystems fail closed.
- Vault encryption does not protect against an OS account or process that can both resolve the key and read or replace plugin state/code. Back up vault files and SecretRef keys independently.
- Vault generation is not externally anchored. Replaying an older valid encrypted backup with its matching key cannot be detected cryptographically; protect backup history and compare sanitized bindings before restore.
- A lost response or abort after OAuth refresh dispatch has an unknowable provider outcome. The shared credential remains quarantined until explicit operator recovery or reauthorization.
- Abrupt process or host loss between OAuth dispatch and durable quarantine publication can leave no local outcome record; no local mechanism can prove what the provider committed in that window.
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
