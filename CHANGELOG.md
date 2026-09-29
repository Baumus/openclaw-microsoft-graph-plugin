# Changelog

All notable changes to this project are documented here.

## Unreleased

## 3.3.1 - 2026-09-29

- Refine the Microsoft Graph access UI with readable theme-aware controls, clearer hierarchy, responsive layout, and a non-overlapping footer.
- Rename the app navigation and page to “Microsoft Graph Zugriff”.

## 3.3.0 - 2026-09-29

- Narrow the native Microsoft access page to OneDrive folder/agent rights, per-agent Calendar/Mail/To Do use, and per-service warning approvals. Preserve critical call-bound approval and existing fine-grained grants.
- Resolve new OneDrive folder paths to immutable drive/item IDs through an admin-only, bounded, credential-backed Gateway query; no token or provider content enters the browser.
- Keep service warning approvals inside the policy so one Gateway config patch can write through a supported single-file `$include` boundary.


- Bound native approval copy, severity, decisions, and execution to one exact host-composed parameter snapshot. The plugin now makes its inspected parameters authoritative in the OpenClaw hook result and rejects any later execution-time rewrite, including warning-to-critical changes after process-local `allow-always` trust.

## 3.1.0 - 2026-09-25

- Capped each calendar multiwrite at 100 operations in schema and runtime.
- Replaced bespoke chat-confirmation authority with configurable OpenClaw-native warning approval. Warning approval defaults on, supports process-lifetime agent/tool/action `allow-always` trust, and leaves critical mutations at `allow-once` or `deny`; legacy chat fields are inert compatibility inputs.
- Added privacy-minimized native approval copy that names the action, target, and action-specific risk, including OneDrive root/path, calendar/event identity, multiwrite counts, and recipient-count context.
- Bound every OneDrive upload and replacement to required SHA-256 and byte-size claims, verified against an immutable opened artifact before credentials or Graph access.

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

## 3.2.0

- Add a native, configuration-only Microsoft Graph Control UI page with guided resources, per-agent permissions, review, admin-scoped validation, and safe include-file export.
- Save inline policy changes through revision-checked `config.patch`; keep credential migration and recovery in the existing CLI flows.
- Require OpenClaw 2026.9.6 for the tested native UI and browser asset contract.
