# Changelog

All notable changes to this project are documented here.

## 3.10.0 - 2026-10-06

- Add `gemacode-microsoft-graph-compact-operation/2` with a closed default-list
  To Do creation operation and optional due date/time zone.
- Require an explicitly configured non-interactive warning policy before that
  pre-model mutation lane can execute; retain the v1 compact-read bridge for
  compatibility.

## 3.9.0 - 2026-10-06

- Add the versioned, process-local `gemacode-microsoft-graph-compact-read/1`
  bridge for three exact read-only operations. The bridge reapplies Graph
  policy, credential, parameter and optional native-execution-permit controls;
  it is not exposed as a network service.
- Keep package, OpenClaw manifest and installation metadata on the same release
  version so a product compatibility gate can reject mixed installations.

## 3.8.9 - 2026-10-06

- Add compact, fail-closed OneDrive root-folder create and exact-delete tools for small local models, while preserving Graph policy, native approval, Native OS execution permits, instruction gating, and bounded exact-match scanning.
- Keep the compact folder schema to the exact name only so local models cannot invent instruction acknowledgement tokens; managed instruction roots still fail closed in preflight.

## 3.8.7 - 2026-10-06

- Add compact default-calendar event create/exact-delete and default-list To Do task create/exact-delete adapters for small local models. All four retain Microsoft Graph policy enforcement, OpenClaw-native approvals, and fail-closed exact target resolution.

## 3.8.6 - 2026-10-06

- Add a compact Microsoft To Do overview adapter that resolves authorized lists internally and returns pending task records rather than mistaking list containers for tasks.

## 3.8.5 - 2026-10-06

- Keep the compact OneDrive root-list result small for local models by returning only each item's name, folder flag, size, and MIME type while preserving bounded pagination status and QEL/Native DB verification.

## 3.8.4 - 2026-10-06

- Add a compact, read-only OneDrive root-list adapter for local models. It selects the root only when the caller has exactly one readable policy root and otherwise fails closed, avoiding guessed root labels without widening access.

## 3.8.3 - 2026-10-06

- Add a compact, read-only default-calendar day adapter for local models. It maps one validated `YYYY-MM-DD` date to the existing bounded calendar read path without widening permissions or bypassing policy, QEL, native execution, or receipt controls.

## 3.8.2 - 2026-10-06

- Add an optional, authenticated Unix-socket boundary for Native OS/Gemacode connected actions. It accepts only a closed high-level tool vocabulary, reuses existing policy and credential controls, and rejects stale, replayed, oversized, malformed, or incorrectly signed requests.
- Declare the Native boundary key as a separate SecretRef and document that this adapter does not replace LN, Rooted, QEL, NetKey, journal certainty, or Native Evidence.
- Normalize the closed `startDate`/`endDate` and `calendarId: default` aliases emitted by compact local models only for bounded calendar reads. Writes, permissions, and foreign-calendar targeting remain strict.

## 3.11.0 - 2026-10-07

- Guide Microsoft account connection from saved, non-secret app identifiers or two clearly labeled fields, with a separate Entra administrator path. Identifiers are stored only in the current browser after a successful device-code start; sign-in and consent remain unverified until the existing flow completes.
- Show explicit Microsoft device-code steps, pending state, and a user-facing failure handoff with copyable, bounded diagnostics. Keep existing vault, policy, and credential protections unchanged.
- Add responsive styling, reviewed translations, and DOM tests for the fast path, missing identifiers, and failure path.

## 3.10.1 - 2026-10-06

- Accept Microsoft's exact `https://www.microsoft.com/link` verification page for personal-account device-code sign-in, while continuing to reject lookalike and altered destinations. Thanks to [@DevilBehindTheSofa](https://github.com/DevilBehindTheSofa) for the clear report and verified workaround in [#28](https://github.com/Baumus/openclaw-microsoft-graph-plugin/issues/28).
- Replace the removed `infra-runtime` SDK import with the existing direct `@openclaw/fs-safe` dependency for protected local-file access.

## 3.10.0 - 2026-10-06

- Show a read-only update badge beside the established-connection status when ClawHub reports a newer, security-clean stable plugin release. The administrator-scoped Gateway check sends no credentials or policy data, is bounded and cached, and never installs an update.

## 3.9.0 - 2026-10-05

- Upload or replace a workspace-created file in one OneDrive call with `sourceWorkspacePath`. The plugin confines and streams the file into private inbound media, hashes it before native approval, verifies the staged bytes again before Graph access, and retains the existing `sourceMediaUri` path for already-staged files.
- Make source selection, path constraints, MIME inference, and recovery guidance explicit in the tool schemas and descriptions.
- Bind native mutation approval to the persisted session generation when the Host supplies only a session key; reject missing or reset sessions before approval or execution without changing OpenClaw Core.

## 3.8.1 - 2026-10-03

- Build and validate this release with OpenClaw 2026.9.8; keep the compatible host floor at 2026.9.6.
- Fix native approval snapshots across plugin module instances and resolve protected inbound media from the authoritative Gateway state root, preventing substituted or unreachable attachment inputs. Thanks to [@BrewingCoder](https://github.com/BrewingCoder) for the report and initial fix in [#22](https://github.com/Baumus/openclaw-microsoft-graph-plugin/pull/22).
- Accept a successful Microsoft Graph `201 Created` attachment response with an attachment ID regardless of provider-reported `size`. Report the locally known uploaded byte count without implying a content readback, and avoid misleading retry guidance after creation. Thanks to [@BrewingCoder](https://github.com/BrewingCoder) for the report and initial fix in [#24](https://github.com/Baumus/openclaw-microsoft-graph-plugin/pull/24).

## 3.8.0 - 2026-10-03

- Make every agent tool self-describing at runtime and add caller-scoped, read-only `microsoft_graph_capabilities` discovery without credentials or Graph calls.
- Add bounded `timeoutMs` transport metadata to mutation schemas, strip it before semantic approval and provider execution, and validate mutation syntax/policy before native approval. Critical managed-root deletion now requires cached instruction acknowledgement before prompting.
- Add additive lifecycle, retry-safety, mutation-certainty, send-acceptance, and pagination guidance to tool results. Legacy chat confirmation fields remain accepted but ignored.

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
