# UI-guided reauthorization — producer evidence

## Candidate scope

Base: freshly fetched `origin/main` = `6d5b19bfb15f2319b657a294addd5441b871e789` (2026-10-21). Isolated branch: `fix/ui-guided-reauthorization`. Source-only repair candidate; version stays 3.16.0 because this is not a release. A release owner must select the version before publication. No compatibility floor or permission expansion.

## Findings and RCA boundaries

- Source-confirmed UI defect: each background render rebuilds the DOM with `replaceChildren`; disclosure expansion was not restored. Text fields persisted only on `change`, losing text typed before blur. Existing focus restoration alone did not preserve those drafts. New DOM regressions cover disclosure, input, focus, selection and background notifications/polling.
- Source-confirmed renewal usability defect: the UI offered a terminal command filled from browser convenience storage (or placeholders), although an administrator-scoped device renewal RPC already existed. Browser IDs were not authoritative vault IDs.
- Synthetic credential regressions reproduce distinct refusal causes: mismatched app/tenant before OAuth, existing vault in create mode, active/quarantined refresh state, changed vault binding after concurrent refresh rotation, and changed policy scopes during consent. Missing consent, denied/malformed responses and cancellation retain the prior credential; byte-preservation assertions cover failed renewal cases.
- The user's specific `credential_vault_conflict` cause is UNKNOWN: no real vault, credential, private logs, Microsoft account or runtime configuration was inspected. Multiple causes used the same generic code in the baseline. Do not claim a proven diagnosis of the user's incident.

## Implemented controls

The UI calls `microsoft-graph.credentials.reauthorize-start` with `{}`; the Gateway reuses app/tenant from the valid vault. Older exact `{clientId, tenant}` callers remain compatible. Strict closed parameters, `operator.admin` registration, UI authority checks, current policy scopes, same app/tenant comparison, snapshotted binding, quarantine rejection, lock and atomic publication remain intact. No automatic conflict rebase or marker clearing was added. Initial connection remains create-only. Shared GUID/domain/alias validation rejects malformed identifiers in browser and Gateway. Sanitized allowlisted conflict subcodes explain safe next steps without exposing raw errors or credentials.

Renewal is guided entirely through the UI: explicit start, trusted Microsoft link, displayed code, consent on Microsoft's page, polling, cancellation and completion/failure. No CLI entry or manual identifier entry is needed for renewal. German setup help distinguishes Anwendungs-ID (Client-ID), Verzeichnis-ID (Mandant), and Objekt-ID, and points to the existing app and correct directory.

## Initial candidate terminal receipts (2026-10-21)

- `openclaw --version`: OpenClaw 2026.10.5 (1b58cb9).
- `npm test`: **38 files / 575 tests passed**, final run 15:26:40 Europe/Berlin, 15.13 seconds, exit 0. Includes 10 happy-dom UI-flow tests, 9 device-code tests, 15 credential-CLI/RPC tests, 10 vault tests, and identifier/localization suites.
- `npm run build`: typecheck and esbuild passed. `npm run plugin:build` then reran typecheck/build and regenerated the immutable UI manifest; exit 0. A test-helper return-type mismatch found earlier was corrected before these terminal checks.
- `npm run plugin:build:check`: metadata up to date, exit 0.
- `npm run plugin:validate`: valid true, errors [], exit 0.
- `npm run package:check`: actual npm package matched the exact file allowlist, exit 0.
- `npx --no-install clawhub package validate .`: PASS, 0 breakages, 0 warnings, no findings, exit 0.
- `git diff --check`: exit 0.
- Security-governance `secret-scan-lite.sh` on the bounded tracked-source diff: no obvious secrets, exit 0. Newly added files contain only pure identifier validation/tests and this sanitized evidence.

Initial candidate Control UI revision: `f67ca16981d8043fc9913239f269e02a2c35ed46bc87ef86f11a36435295fc06` (superseded by the review correction below). Build outputs follow repository convention and are ignored; the updated manifest is tracked.

## Fresh primary documentation

Checked 2026-10-21: [Microsoft app registration, German](https://learn.microsoft.com/de-de/entra/identity-platform/quickstart-register-app), [Microsoft device authorization, German](https://learn.microsoft.com/de-de/entra/identity-platform/v2-oauth2-device-code). OpenClaw routes discovered through current `llms.txt`: [feature plugins](https://docs.openclaw.ai/plugins/feature-plugins), [building plugins](https://docs.openclaw.ai/plugins/building-plugins). Current docs describe authenticated `host.request` using the operator's actual scopes; installed SDK declarations and successful local builds confirm the reused Control UI / scoped Gateway registration boundaries (runtime-supported). No new host API was introduced; live dispatch/activation was not tested.

## Limitations, review and recovery

- Producer evidence is not independent security/usability approval. Next owner: requesting parent, to arrange that review on the frozen commit.
- UI evidence is synthetic happy-dom interaction coverage, not a screenshot, browser accessibility audit or connected Control UI activation test. No Microsoft sign-in or actual consent was attempted.
- Existing vault format has no independently verified account identity. The unchanged safeguard requires the operator to choose the same account on Microsoft's page. No new account-identity guarantee is claimed.
- Concurrent refresh rotation still intentionally aborts renewal on stale binding. The operator must refresh status and start a fresh renewal; atomicity is not weakened to make the conflict disappear.
- Quarantine recovery remains separately reviewed and blocked in ordinary renewal. The UI never deletes the vault or clears its marker.
- New guidance has German and English entries; other locales use the existing English fallback where translations are not reviewed.
- No real vault read/write, token access, live config or permission change, installation, reload, restart, publication, push, PR, or merge occurred. Rollback is to discard this candidate branch; no live state needs restoration.
- Next approval scope, only after independent review: separately authorize any version/release preparation and any deployment/activation; real account consent remains a distinct operator action. Neither is authorized by these producer checks.

## Independent-review correction: R-001 / R-002

Reviewed candidate: `8090a7709d41496d75c9a3f2fdf9cf40022ada77`. Both findings supplied by the parent's independent source reviewer are confirmed in-scope defects; producer fixes await targeted independent reverification, not self-approval.

- **R-001 (S2):** status/start/cancel requests now snapshot an auth operation epoch, and session-bound requests also snapshot the session ID. Cancellation immediately invalidates prior polling. Both success and rejection handlers, post-completion credential refresh, and final busy/render updates reject stale ownership. Authority loss invalidates outstanding requests and pauses timers; regained authority can resume the current session. Disposed/aborted views cannot accept a result. Polling does not run while start/cancel is busy.
- **R-002 (S3):** renewal records its vault-derived mode. Copied failure diagnostics omit Organisation entirely for that mode, rather than attributing the failure to browser-stored tenant IDs. No vault identifiers were exposed or added to responses.
- Changed files in this correction: `src/control-ui.ts`, `src/control-ui-flow.test.ts`, generated `openclaw.plugin.json`, and this evidence document. No backend, permission, vault format, binding, account or consent safeguard changed.
- Deterministic UI regressions use distinct A/B session IDs and codes. Deferred A status returning failed, rejecting, or returning created cannot replace B, change credential status, stop B polling or redirect B cancellation. Separate start/status/cancel cases cover success and rejection after authority loss/new session and after unmount. Renewal failure/copy with stale browser tenant A proves no tenant/app/code attribution is copied.
- Focused command `npm test -- src/control-ui-flow.test.ts src/control-ui-save.test.ts`: **33/33 passed**, exit 0 (20 flow tests, 13 save tests).
- Final `npm test`: **585/585 tests, 38 files passed**, exit 0; 15:42:05 Europe/Berlin on 2026-10-21, 21.52 seconds. Replaces the initial 575-test receipt for the corrected candidate.
- `npm run typecheck`: exit 0. `npm run plugin:build` reran typecheck/build, regenerated immutable UI metadata, exit 0. `plugin:build:check`: current, exit 0. `plugin:validate`: valid true/errors [], exit 0. ClawHub package validator: PASS, zero warnings/findings, exit 0.
- Package allowlist initially caught superseded ignored UI build outputs from the prior candidate. Only those two obsolete generated files were removed before the final package recheck; no source or live state was deleted. Final `npm run package:check`: exact allowlist passed, exit 0. Correction-diff secret scan: no obvious secrets, exit 0; `git diff --check`: exit 0.
- Corrected Control UI revision: `75cb753e0b7f52034a16c8e6c84415ec84bee82d2c0522a5bf4816cc28764edc`. Earlier UI revision/build receipt is superseded; unchanged backend/security evidence remains applicable.
- Residual limitation: authority loss/unmount does not cancel Microsoft consent server-side; the UI intentionally does not synthesize privileged cancellation. On return, an explicit start can recover the Gateway's existing pending flow. Real provider/browser activation remains untested.
- **READY FOR TARGETED REVERIFICATION:** parent to relay only R-001/R-002, the `8090a77..HEAD` fix diff and the above evidence to the same independent reviewer. Do not reopen unchanged implementation; independent reviewer owns the finding dispositions.
