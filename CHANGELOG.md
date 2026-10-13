# Changelog

All notable changes to this project are documented here.

## Unreleased

## 4.0.0 - 2026-10-12 (candidate, not published)

- Guide explicit setup for new/multi-user/unprovable installations: a native Microsoft setup header action, incomplete/ready checklist and personal-page link, verified-profile display, shared-owner identity guidance, packaged Settings hints, and actionable CLI/README instructions. Qualifying single-user upgrades instead preserve the existing connection automatically at post-ready Gateway startup without any browser/account interaction.

- Breaking: remove implicit global Microsoft account authority and legacy credential RPC registration. Existing `me` tools and argument schemas are unchanged, but need an authenticated session binding, an explicitly consented fixed worker, or a qualified sole-user migration receipt.
- Add one encrypted personal connection per canonical OpenClaw profile, own-profile sign-in/status/disconnect/reconnect and explicit idempotent legacy migration. No first-visitor claim, team consent, worker consent, token export, or global fallback.
- Add protected persistent session-incarnation/account bindings, owner sharing and unattended consent, account-aware agent/resource policy restricted by administrative per-profile ceilings.
- Add personal native Control UI with account permissions and session selection; retain separate administrative ceiling editor.
- Bind native approvals, continuations, OneDrive instruction acknowledgements, vault/cache identities and every network dispatch to the account/generation/binding context. Recheck public V2 host invocation authority before effects.
- Require OpenClaw 2026.10.1 for public authenticated-profile/session authority and V2 tool factories. Automatic multi-user DM admission/fork propagation is not claimed; authenticated UI and fixed-worker paths are supported.
- Make automatic upgrades portable across verified OpenClaw 2026.10.1/2026.10.2 installations: public auth/port defaults, config-first token/password modes, loopback-capable LAN/auto binds, independent canonical IDs/aliases and state roots. Add a supported-host CI matrix and repeatable synthetic SDK/startup canary; unknown ownership or unresolved transport authority still fails closed.
- Add a bounded host-startup singleton legacy migration with a generation-owned v2 scheduler: full canonical host-directory snapshots through the public SDK, exact local instance/auth pinning, encrypted backup preservation, legacy∩ceiling policy, no manual singleton session/worker setup, and durable non-resurrecting topology/account-change invalidation.
- Replace CLI process starts with one fresh read-only, null-device public SDK connection per directory snapshot, requesting only `operator.read` and retaining same-connection before/roster/after instance proofs. Remove the extra 90-second directory allowance while preserving fresh snapshots and configured operation/provider deadlines; no roster cache or persistent client is introduced.
- Enforce existing refresh-marker admission under both source/destination vault locks in the common legacy migration writer; reject uncertain copies/locks without token laundering. Serialize detected bootstrap topology tombstones independently of request cancellation, publishing only revocation.
- Clearly separate fresh-operation/checkpoint snapshot checks from an atomic catalog lifetime lease; unsupported host/auth/runtime shapes retain explicit migration. No live activation or strict instant-revocation claim.


## 3.12.1 - 2026-10-10

- Distinguish saving from activation, lock duplicate writes immediately, and reconcile persisted policies even when the save response times out or disconnects. Keep drafts and operation state in authenticated host memory across page remounts.
- Confirm activation against the running Microsoft Graph policy hash instead of unrelated Gateway configuration changes. Provide a targeted, drain-aware application retry without rewriting saved rules or interrupting admitted work.
- Preserve confirmed Microsoft account connection while policy activation is pending; add bounded polling, explicit uncertainty, conflict recovery, and regression coverage.

## 3.12.0 - 2026-10-09

- Extend the existing Microsoft To Do write tool with explicit actual-default-list task creation. Safe bounded discovery requires independent read/create authorization, resolves the owned non-shared default before native approval, and binds execution to that exact list. Known-ID writes are unchanged; no new tool or post-write readback is introduced.
- Include the macOS workspace-staging repair from #38: canonical state-root identity checks, Darwin owner identity, targeted regressions, and actual macOS CI coverage.

## 3.11.4 - 2026-10-07

- Restore the feature-focused ClawHub subtitle: “Bring Outlook mail and calendar, OneDrive files, and Microsoft To Do into OpenClaw with per-agent access you control.” Keep the display name and README attribution unchanged.

## 3.11.3 - 2026-10-07

- Present the plugin as “Connect Microsoft 365 to OpenClaw” with the attribution “Independent community plugin by Baumus” across ClawHub metadata, README, and Control UI. This is a display-name and copy change only; the package name, plugin ID, permissions, tools, and configuration keys remain unchanged.

## 3.11.2 - 2026-10-07

- Restore the manifest tool contract required by OpenClaw 2026.9.8. The v3.11.1 package installed but left all Microsoft Graph agent tools unavailable on that host; v3.11.2 restores registration without changing Graph permissions, tool behavior, or policy.

## 3.11.1 - 2026-10-07

- Remove the unsupported generated tool contract from the plugin manifest while keeping runtime tool registration and supported tool metadata intact. The build now strips that field so future packages remain ClawHub-valid.

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
