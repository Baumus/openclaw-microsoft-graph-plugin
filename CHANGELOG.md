# Changelog

All notable changes to this project are documented here.

## Unreleased

## 3.1.0 - 2026-09-25

- Capped each calendar multiwrite at 100 operations in schema and runtime.
- Replaced bespoke chat-confirmation authority with configurable OpenClaw-native warning approval. Warning approval defaults on, supports process-lifetime agent/tool/action `allow-always` trust, and leaves critical mutations at `allow-once` or `deny`; legacy chat fields are inert compatibility inputs.

## 3.0.0 - 2026-09-25

- Initial public release.
- Added a credential-free, default-deny policy in resolved `plugins.entries["microsoft-graph"].config.policy`.
- Added one shared delegated OAuth credential, one encrypted vault record, one lock/quarantine lifecycle, and one SecretRef-only `credentialVaultKey`.
- Routed credential status, migration, restore, and recovery through strictly validated, operator-scoped Gateway RPC methods so local CLI loads never depend on unresolved SecretRef config.
- Added single-source migration and single-destination backup/restore receipts.
- Kept transactional bounded access-token caching, dedicated OAuth request timeouts, durable uncertainty quarantine, authorization/approval ordering, and fail-closed secure storage.
- Documented that vault generations have no external monotonic anchor and that replay of an older valid encrypted backup remains an operator-managed residual risk.
- Added exact-pinned `@openclaw/fs-safe` 0.13.1, a version-2 policy example, migration documentation, and deterministic vault/config/CLI/rotation tests.
- Added manifest and runtime schema validation for the complete default-deny policy while preserving authorization before GNU `pass` credential access.
- Added normal OpenClaw host-level `$include` composition guidance for splitting the policy into a JSON5 file.
- Added bounded OneDrive, Outlook Calendar, Outlook Mail, and Microsoft To Do tool coverage, private-media transfer, continuation handling, provider timeout/retry controls, and deterministic synthetic tests.
- Capped each confirmed calendar multiwrite at 100 operations in schema and runtime.

This changelog does not claim official Microsoft or OpenClaw support, security certification, or production suitability.
