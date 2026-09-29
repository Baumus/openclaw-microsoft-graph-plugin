# Microsoft Graph plugin configuration UI — concept

**Status:** implementation proposal, no UI shipped · **Baseline:** plugin 3.1.0 / OpenClaw 2026.9.6 · **Checked:** 2026-09-29 (Europe/Berlin)

## Scope and product decision

Build a **configuration-only** native page for the Microsoft Graph plugin inside the OpenClaw Control UI. It guides an operator through setup and edits the plugin's policy and safe settings. It is **not** an operations dashboard: no mail, calendar, file, or task browser; no activity feed, health monitor, approval inbox, or Graph actions. Operational status appears only where needed to validate a configuration step, e.g. “credential vault present” or “required scope missing”.

The primary user is the Gateway administrator. A non-admin sees a short “Administrator access required” state; the UI does not offer partial editing. Microsoft tenant consent and credential acquisition remain external prerequisites. The plugin must not promise that its UI can grant Microsoft permissions.

## Information architecture

One plugin-owned destination, **Microsoft Graph → Configure**. Use a single task-oriented page with a four-step progress rail; do not add four top-level routes or a second platform shell.

1. **Credential & prerequisites** — explain Entra delegated consent, the one shared credential, SecretRef for the vault key, and the existing safe migration path. Show only sanitized setup checks. Never show or accept a refresh token, vault key, client secret, or raw credential JSON in the browser.
2. **Services & resources** — select OneDrive, Calendar, Mail, and To Do access. OneDrive roots require label, display path, immutable drive ID and item ID; calendar resources are `me` or an exact calendar ID; Mail/To Do are fixed `/me`. The UI explains these boundaries at the field where they matter.
3. **Agent permissions** — choose existing agent IDs, then allowed operations by service/resource. Default deny. Group operations as Read, Change, and Critical with precise underlying operation labels. Show the delegated Microsoft scopes implied by the draft and distinguish those scopes from OpenClaw policy grants. Never silently widen permissions.
4. **Review & save** — human-readable before/after diff, affected agents/resources/scopes, warning-approval setting, validation errors, and one explicit **Save configuration** action. Saving is disabled until all required fields validate and the current config revision is fresh.

An **Advanced** disclosure holds bounded transport/time limits and `warningApprovalsRequired`; the default remains true. Explain that critical delete/send/respond actions always require native approval. Keep unrelated host settings (plugin install/enablement, `tools.allow`, secrets store, custom UI lab) on their owning OpenClaw Plugins/Settings pages, with direct links or precise instructions.

### Desktop sketch

```text
OpenClaw / Plugins / Microsoft Graph / Configure
[1 Credential] — [2 Services] — [3 Agent permissions] — [4 Review]

Agent permissions                                    Unsaved changes
Select agent: [main ▾]                               [Discard draft]

OneDrive · root “family-docs”
  [✓] Read   [ ] Write   [ ] Delete
  Bound to drive ID … · item ID …
Calendar · me
  [✓] Read   [✓] Create   [ ] Update   [ ] Respond   [ ] Delete

Required Microsoft consent for this draft
Files.Read · Calendars.ReadWrite · offline_access
These scopes do not grant agent access by themselves.

[Back]                                              [Review changes]
```

The example is illustrative, not the current installed policy. The review step shows exact effective changes and whether they are additions, removals, or replacements.

## Configuration journeys

- **Initial setup:** show prerequisites, link to OpenClaw Secrets for a secret-kind vault key, guide the existing `pass` migration with copyable CLI commands, then edit policy. The UI never receives a secret. Existing credential RPC `microsoft-graph.credentials.status` supplies only a sanitized readiness check; migration/recovery stay in their existing interactive CLI flows for this version.
- **Add an agent grant:** select an actual configured agent, choose the minimum operations/resource, see implied delegated scopes, review the diff, save. If Microsoft consent lacks a required scope, flag the mismatch before save and present the external consent step; do not present it as automatically resolved. Whether to block a policy-only save on that mismatch is an explicit product decision for implementation; default proposal: allow save only while the plugin remains disabled, otherwise block until coverage is verified.
- **Remove access:** show exactly which grants/roots disappear. Require deliberate confirmation of removals; never remove sibling agents or roots due to array-patch semantics.
- **Change approval setting:** `warningApprovalsRequired: false` gets a prominent consequence summary in Review. It changes only warning-level approval, not critical approval or default-deny policy.
- **Concurrent edit:** if config `baseHash` changed after the form loaded, stop saving, refresh the authored snapshot, show a field-level conflict, and let the operator reapply the draft. Do not force-write stale data.
- **Included policy file:** the current README recommends `$include` for `policy`. Do **not** assume a generic `config.patch` safely edits that included file or preserves its authorship. Detect this layout. Until an OpenClaw-supported include-aware write path is verified, offer a validated proposed policy export and an explicit handoff to the owning Config/file workflow, then re-read to verify; **do not display a misleading Save button**. A future include-aware editor may restore in-page save after its behavior is tested.

## Field ownership and save paths

| UI field/check | Read source | Save path / owner |
| --- | --- | --- |
| Vault readiness | Sanitized `microsoft-graph.credentials.status` RPC | No browser write; existing interactive credential CLI |
| Vault-key reference | Plugin config schema / SecretRef metadata, never resolved value | OpenClaw Secrets and plugin config Settings; no secret entry in this page |
| OneDrive roots, calendar resources, agent operations | Effective validated `policy` in `plugins.entries["microsoft-graph"].config` | Inline policy only: minimal `config.patch` with current `baseHash`; included policy: explicit file/config handoff |
| Warning approval setting | `warningApprovalsRequired` in the plugin config, default `true` | Same plugin-config-only patch after review |
| Time/size/concurrency limits | Plugin config schema and `config.schema.lookup` | Same plugin-config-only patch, Advanced section |
| Plugin install/enablement, `tools.allow`, lab setting, Microsoft consent | Their owning OpenClaw or Microsoft surfaces | No write from this page |

## UX specification

- Match the Control UI's typography, spacing, navigation, and form language. Use one clear page title, compact step rail, persistent unsaved-change indicator, and a sticky footer with Back/Review/Save. No generic card-in-card layout or marketing hero.
- Use a responsive two-column editor plus contextual help at desktop widths; on mobile stack fields and help, keep the primary action reachable, and turn permission tables into labeled rows. Long IDs wrap or use an accessible copy affordance. Never hide a required field behind horizontal overflow.
- Prefer small operation groups and explicit checkboxes over an enormous permission matrix. No “select all” that silently grants write/delete. New agents start with no permissions; new roots start with read/write/delete unchecked. Confirm before discarding a dirty draft or navigating away.
- Inline validation appears at the relevant field and is summarized at Review. Errors say what must change, not just “invalid config”. Preserve user input after validation failure; focus the first invalid field from the summary. Save success is shown only after the Gateway confirms application, not merely after request submission.
- Cover loading, disconnected, no admin scope, missing plugin config, invalid existing config, secret unavailable, credential missing/quarantined, insufficient Microsoft scopes, stale revision, included-file handoff, save in progress, saved-but-not-applied, and save failure. These are **configuration states**, not a monitoring dashboard. Do not poll Microsoft Graph.
- Keyboard-first behavior: semantic form groups/legends, labeled controls, visible focus, logical tab order, focus restoration after dialogs, non-color-only validation, accessible error summary and polite save announcement. Check reduced motion, high contrast, and desktop/mobile widths in a real browser.

## OpenClaw-compliant implementation boundary

**Documented and locally present:** `defineControlUiPlugin`, `host.ui.registerPage`/`registerNavigation`, `host.request`, and the browser source/build/manifest flow in [Feature plugins](https://docs.openclaw.ai/plugins/feature-plugins). `config.get`, `config.schema.lookup`, and `config.patch` with `baseHash`/`replacePaths` are documented in [Config RPC](https://docs.openclaw.ai/gateway/configuration/config-rpc). Plugin `configGroups`/`uiHints` are documented in [Manifest setup and auth](https://docs.openclaw.ai/plugins/manifest/setup-and-auth). These APIs are experimental; pin and test the host version.

- Add one browser entry (`package.json.openclaw.controlUi`) and generated manifest assets (`openclaw.plugin.json.controlUi`); one page/nav contribution, no replacement, session panel, widget, or composer action. Browser code owns its DOM, namespaced CSS, and lifecycle; do not import Control UI internals. Dispose subscriptions and check `context.signal` after asynchronous work.
- Use manifest `configGroups`/`uiHints` to improve the **existing generic Plugins Settings** fallback. A guided native form must consume the actual plugin config schema and validated policy model, not maintain a divergent list of fields/operations. `config.schema.lookup` can supply host field constraints/reload hints; plugin-owned policy validation remains authoritative.
- Backend validation must use existing `validatePolicy` and `requiredPolicyScopes`; expose a bounded, admin-scoped validation/query surface if the browser cannot safely obtain the needed derived data. Never create a second policy evaluator or return secret material. Validate once before Save, then let the Gateway validate again on write.
- For inline policy config, read the current authored config and opaque `hash` with `config.get`; construct a minimal patch confined to `plugins.entries["microsoft-graph"].config`. Use `config.patch`, never `config.apply` for this editor. Supply `baseHash`. Supply exact `replacePaths` for arrays whose removals/replacements are intentional; do not blanket-authorize array replacement. Re-read after the write and compare effective values and `configRevisionHash`/`appliedConfigHash` before saying “Applied”. A persisted-but-unapplied or deferred result needs a distinct message and recovery path.
- The native browser shares the signed-in operator's Gateway authority. Hide editing without admin capability **and enforce admin scope server-side** on validation/details/write paths; UI visibility is not authorization. `host.request` must not become a generic arbitrary-config editor inside this plugin. Scope new RPCs to this plugin's subtree and return bounded, sanitized data.
- The credential lifecycle is separate: the current sanitized status RPC is `operator.read`; migration, restore, and recovery are `operator.admin` with interactive typed confirmations in the CLI. Do not call the apply RPCs directly from v1 UI or replace their protections with `window.confirm`. A future browser credential flow needs a separate reviewed protocol for exact binding, unknown outcomes, and reauthorization.
- Native custom-plugin UI defaults off for separately installed plugins and requires HTTPS or trusted loopback for authenticated assets. The plugin still works without its browser UI. Do not make the UI the only possible way to configure it or add a separate cross-origin app.

## Security and data rules

- No OAuth tokens, vault key, secret-store value, raw Graph content, pass entry contents, raw config dump, or exception stack in the DOM, browser storage, telemetry, screenshots, or UI responses. Use only sanitized receipt/status fields needed for the current step.
- Treat OneDrive display paths/labels as descriptive; immutable drive/item IDs are the authorization boundary. `agents_instructions: trusted` is an advanced, explicit opt-in with a warning about who can write that root.
- The shared credential's consented scope union is not provider-side isolation. Agent access remains the policy's default-deny decision, and write approvals remain the execution gate.
- Only Microsoft Graph plugin config may be changed by this page. `tools.allow`, plugin enablement, secrets, Gateway lab settings, and Microsoft tenant consent are distinct effects controlled elsewhere.

## Implementation slices and acceptance

| Slice | Deliverable | Gate |
| --- | --- | --- |
| 0 | This concept and a field-to-source/save-path map, including `$include` behavior | No unsupported write assumption; exact admin/read scope decision |
| 1 | Schema-backed guided editor, validation, Review, inline-config save, safe included-file handoff | Config patch tests for concurrency/array removals/secret redaction; desktop/mobile/keyboard browser review |
| 2 | Setup guidance and generic Settings metadata polish | CLI/Settings fallback works with custom UI off; no credential material enters browser |
| Future, separate decision | Browser-managed OAuth/vault lifecycle or include-aware policy editing | Dedicated security design, authority/approval contract, migration/rollback, host-version tests |

Release acceptance: the UI only configures the plugin; every saved change is scoped to its config subtree and reviewable before commit; stale drafts and included policies cannot silently overwrite authored config; no secret or Microsoft content crosses into the browser; persisted vs applied state is accurate; the generic/CLI path remains available when the lab is off. When implemented, update package/manifest allowlists and compatibility/version metadata, run plugin build/check/validate plus focused backend and browser tests. **No runtime, config, permission, or plugin behavior changes are made by this concept.**

## Evidence and open decisions

- Canonical plugin: `Baumus/openclaw-microsoft-graph-plugin` main at `36cde1c`, package/manifest 3.1.0; no native UI entry yet. Existing policy, scope calculation, credential RPC, and setup instructions are the source baseline.
- Installed OpenClaw: `2026.9.6 (eb377ac)`; local SDK declarations and live official Feature plugins page checked 2026-09-29. **Runtime-supported:** page/nav, browser host request, asset build model. **Docs-current and to verify in implementation:** exact config patch behavior for this plugin's nested arrays and `$include` layout.
- **Implementation decisions still required:** whether a newly broadened policy can be saved while disabled without current Microsoft consent; exact admin-scoped validation response; include-aware editing support. These are stop gates for code, not reasons to imply the UI already exists.
