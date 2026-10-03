# Changelog

All notable changes to this project are documented here.

## Unreleased

## 3.7.0 - 2026-10-03

- Let OneDrive upload/update callers omit SHA-256 and byte size: the plugin derives both from protected inbound media before native approval and re-verifies the approved identity before transfer. Supplied values remain strict paired assertions.
- Explain the self-contained upload and critical-delete approval contracts in tool descriptions and documentation, including the host-owned `operator.approvals` requirement.

## 3.6.0 - 2026-10-02

- Redesign OneDrive access as responsive folder cards showing every authorized agent. Read/write/delete bubbles now toggle permissions directly; each card also exposes adding agents and a confirmed folder-removal action while preserving the review-and-save workflow.

## 3.5.4 - 2026-10-02

- Collapse the completed setup checklist to a compact connection status only when the vault SecretRef, saved agent grant, applied rules, and valid Microsoft credential are confirmed. This does not claim that an agent read has been tested.
- Keep actionable setup and sign-in states for missing or regressed prerequisites. Hide redundant connected sign-in and applied-version success cards, and show the policy source while editing or handling a save error.

## 3.5.3 - 2026-10-01

- Accept Microsoft's `https://login.microsoft.com/device` device-code verification page while keeping the browser destination on an exact allowlist.
- Accept the required `offline_access` scope in CLI sign-in start and completion results.
- Accept equivalent Microsoft Graph `mailFolders('id')` continuation paths only when they resolve to the exact expected mail collection; preserve origin and traversal checks.
- Add synthetic regression tests for sign-in validation and multi-page mail reads.

## 3.5.2 — Guided setup status

- Show a setup checklist and the next required step for the vault key, agent grant, applied policy, and Microsoft connection.
- Keep account connection distinct from verified agent access; the first read still requires an explicit authorized test.
- Present access rules before the sign-in card so the UI follows the actual setup dependency order.
- Clarify that a user-owned public-client Entra app with device-code sign-in needs no callback service or redirect URI; tokens remain in the local encrypted vault.

## 3.5.1 - 2026-10-01

- Refresh public positioning, installation guidance, and package metadata for Microsoft 365 for OpenClaw; no runtime permission or API changes.

## 3.5.0 - 2026-10-01

- Add a guided, four-language Control UI sign-in that opens Microsoft in the browser, displays the one-time code, tracks completion, and supports cancellation. Keep the terminal helper as an alternative.
- Remove the legacy `migrate-from-pass` command and Gateway method; preserve optional `restore-pass` emergency backup.

## 3.4.0 - 2026-09-30

- Localize the Microsoft Graph configuration UI for English, German, Spanish, and Arabic, with English fallback for other host locales.
- Resize the packaged plugin identity icon to OpenClaw’s recommended 512 × 512 pixels for reliable display.

## 3.3.2 - 2026-09-30

- Show a persistent saved-versus-applied Gateway status in Microsoft Graph Zugriff, recheck pending application automatically for one minute, and offer a manual status check without claiming success when revision evidence is missing.

## 3.3.1 - 2026-09-29

- Refine the Microsoft Graph access UI with readable theme-aware controls, clearer hierarchy, responsive layout, and a non-overlapping footer.
- Rename the app navigation and page to “Microsoft Graph Zugriff”.

## 3.3.0 - 2026-09-29

- Narrow the native Microsoft access page to OneDrive folder/agent rights, per-agent Calendar/Mail/To Do use, and per-service warning approvals. Preserve critical call-bound approval and existing fine-grained grants.
- Resolve new OneDrive folder paths to immutable drive/item IDs through an admin-only, bounded, credential-backed Gateway query; no token or provider content enters the browser.
- Keep service warning approvals inside the policy so one Gateway config patch can write through a supported single-file `$include` boundary.


## 3.2.0 - 2026-09-29

- Add a native, configuration-only Microsoft Graph Control UI page with guided resources, per-agent permissions, review, and admin-scoped validation.
- Save policy changes through revision-checked `config.patch` on a supported single-file include boundary; keep credential migration and recovery in the CLI.
- Require OpenClaw 2026.9.6 for the native UI and browser asset contract.

## 3.1.0 - 2026-09-25

- Bound native approval copy, severity, decisions, and execution to one exact host-composed parameter snapshot. The plugin now makes its inspected parameters authoritative in the OpenClaw hook result and rejects any later execution-time rewrite, including warning-to-critical changes after process-local `allow-always` trust.
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
