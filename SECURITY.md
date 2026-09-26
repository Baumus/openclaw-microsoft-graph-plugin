# Security policy

## Reporting a vulnerability

Do not include credentials, refresh tokens, access tokens, private policy files, mailbox content, file content, or personal data in a public issue.

Report vulnerabilities privately through [GitHub Security Advisories](https://github.com/Baumus/openclaw-microsoft-graph/security/advisories/new). Do not open a public issue for a suspected vulnerability.

If GitHub does not show the private reporting form, contact the repository owner through GitHub first and ask for a private reporting route without disclosing vulnerability details or sensitive data publicly.

Include only the minimum reproducible information: affected version, impact, preconditions, sanitized steps, and a synthetic proof of concept. Remove host paths, tenant IDs, account IDs, drive IDs, message content, and tokens.

## Supported versions

Security fixes are expected on the latest released version only. Until a public release exists, the repository candidate is not a supported production release.

## Security boundaries

- One shared delegated OAuth credential is stored in one AES-256-GCM vault record under the resolved OpenClaw state directory. The key is one structured OpenClaw SecretRef. Credentials and keys are not tool parameters.
- OpenClaw may materialize the declared SecretInput while loading config; plugin-side key selection and use remain authorization-gated. Refresh-token rotation must be durably published before cache admission or use.
- An uncertain dispatched OAuth refresh, including an unusable HTTP-200 response, durably quarantines the credential until explicit bound recovery or reauthorization. Vault generation has no external monotonic anchor, so replay of an older valid encrypted backup with its matching key remains an operator-managed risk.
- There is no provider-side read/write credential isolation. Credential or plugin compromise has the union of granted Microsoft scopes; default-deny policy and write approvals constrain normal plugin behavior.
- The policy is default-deny and evaluated before credential or provider access.
- Credential migration, status, quarantine recovery, and restore are local operator CLI commands, not model-facing tools or chat commands. Migration requires one explicit all-scope source; restore targets one explicit destination and reports unverifiable outcomes as unknown.
- Warning-level mutations require OpenClaw-native call approval by default. An operator may explicitly disable warning approvals, while policy authorization and all preconditions remain mandatory. Process-local `allow-always` trust is scoped to the authenticated agent, exact tool, and normalized action. Approval and approval-free warning paths bind the inspected parameters to the host tool-call identity and fail closed if hook composition changes the parameters before execution.
- Destructive, send, and respond operations always require OpenClaw call-bound approval and offer only allow-once or deny.
- OneDrive content writes require an exact lowercase SHA-256 and byte-size claim. Approval preflight opens and verifies the protected artifact before managed-root instruction discovery or any credential, OAuth, or Graph boundary; execution independently reopens and verifies it before those provider boundaries.
- Provider continuation and upload URLs are origin/path constrained and are not returned to the model.
- Private media is represented by `media://inbound/...` references; host-local paths and inline file bytes are rejected.
- No security certification, compliance attestation, or enterprise-readiness claim is made.

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) and [docs/OAUTH_ACCESS_MATRIX.md](docs/OAUTH_ACCESS_MATRIX.md) for the detailed trust and permission model.
