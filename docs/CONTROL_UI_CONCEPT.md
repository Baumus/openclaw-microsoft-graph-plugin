# Microsoft-Zugriff — focused Control UI

**Status:** Implemented in the 3.3.0 source worktree; runtime activation and visual browser check remain separate verification steps.

The native Control UI page has exactly three editable areas:

1. **OneDrive:** select a configured agent and grant Lesen, Schreiben, or Löschen for a displayed folder path. A grant always includes descendants. New paths are resolved by an admin-only Gateway method to immutable drive/item IDs before they enter policy; the browser never edits IDs. The add flow currently resolves paths in the signed-in account's own OneDrive (`/me/drive`). Removing a root affects all agents and requires an explicit confirmation.
2. **Dienste:** toggle Kalender, E-Mail, and To Do for each agent. Existing fine-grained operation/resource grants remain unchanged when a service stays on or is toggled off and back on in one draft. A genuinely new service grant includes all supported operations for that service and its default `me` resource.
3. **Freigaben:** for each of the four services, show that recognized reads need no mutation approval, choose whether warning-level changes ask first, and explain that critical delete/send/respond operations always require call-bound allow-once approval. These warning choices live in `policy.rules.warningApprovalsByService`; the older global `warningApprovalsRequired` remains a runtime fallback, not an editable field on this page.

There is one review/save step, but it introduces no fourth category of configuration. Credential handling, transport limits, plugin enablement, raw IDs, operation lists, and resource identifiers are not editable on this page. The generic OpenClaw plugin settings and CLI remain separate host surfaces.

## Source and save contract

The page reads `config.get`, validates the complete policy through the admin-scoped `microsoft-graph.configuration.validate` RPC, computes a minimal policy-only patch, and submits `config.patch` with the current `baseHash` and exact `replacePaths` for intentionally replaced arrays. It then re-reads effective values and distinguishes persisted from applied state. New scopes remain blocked while the plugin is enabled until Microsoft consent is verified through the existing operator workflow.

OpenClaw 2026.9.6 documents write-through for changes wholly owned by one single-file object-key `$include`. Keeping warning choices inside the policy makes this page's save one ownership boundary; the current `plugins.entries["microsoft-graph"].config.policy` include is a single-file object-key include. Unsupported include layouts and concurrent edits fail closed. No separate plugin RPC writes configuration files.

The folder-resolution RPC is `operator.admin`, accepts only one bounded path under `/me/drive`, uses the existing vault credential with `Files.Read`, and returns only normalized path and immutable drive/item IDs. It does not return tokens, file content, or other provider data. The UI never stores secrets.

## Remaining acceptance evidence

- Deterministic tests, typecheck, build, package check, and plugin validation must pass.
- A connected browser-capable Control UI client is needed for desktop/mobile/keyboard visual review.
- A non-disruptive activation window is needed to load the 3.3.0 backend. Source build success alone does not prove the installed 3.2.0 runtime has changed.
