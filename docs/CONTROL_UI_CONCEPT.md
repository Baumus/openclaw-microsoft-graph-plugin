# Microsoft Graph access — focused Control UI

**Status:** Implemented in version 3.5.0. This document describes the policy editor; guided Microsoft sign-in is a separate Control UI workflow.

The native Control UI page has exactly three editable areas:

1. **OneDrive:** select a configured agent and grant read, write, or delete access for a displayed folder path. A grant always includes descendants. New paths are resolved by an admin-only Gateway method to immutable drive/item IDs before they enter policy; the browser never edits IDs. The add flow currently resolves paths in the signed-in account's own OneDrive (`/me/drive`). Removing a root affects all agents and requires an explicit confirmation.
2. **Services:** toggle Calendar, Mail, and To Do for each agent. Existing fine-grained operation/resource grants remain unchanged when a service stays on or is toggled off and back on in one draft. A genuinely new service grant includes all supported operations for that service and its default `me` resource.
3. **Approvals:** for each of the four services, show that recognized reads need no mutation approval, choose whether warning-level changes ask first, and explain that critical delete/send/respond operations always require call-bound allow-once approval. These warning choices live in `policy.rules.warningApprovalsByService`; the older global `warningApprovalsRequired` remains a runtime fallback, not an editable field on this page.

There is one review/save step, but it introduces no fourth category of configuration. Credential handling, transport limits, plugin enablement, raw IDs, operation lists, and resource identifiers are not editable on this page. The generic OpenClaw plugin settings and CLI remain separate host surfaces.

## Source and save contract

The page reads `config.get`, validates the complete policy through the admin-scoped `microsoft-graph.configuration.validate` RPC, computes a minimal policy-only patch, and submits `config.patch` with the current `baseHash` and exact `replacePaths` for intentionally replaced arrays. It then re-reads effective values and distinguishes persisted from applied state. New scopes require the necessary Microsoft consent before the affected operation can succeed; the guided sign-in and operator CLI are separate credential workflows.

OpenClaw 2026.9.6 documents write-through for changes wholly owned by one single-file object-key `$include`. Keeping warning choices inside the policy makes this page's save one ownership boundary; the current `plugins.entries["microsoft-graph"].config.policy` include is a single-file object-key include. Unsupported include layouts and concurrent edits fail closed. No separate plugin RPC writes configuration files.

The folder-resolution RPC is `operator.admin`, accepts only one bounded path under `/me/drive`, uses the existing vault credential with `Files.Read`, and returns only normalized path and immutable drive/item IDs. It does not return tokens, file content, or other provider data. The UI never stores secrets.

## Release acceptance evidence

- Deterministic tests, typecheck, build, package check, and plugin validation must pass for the release commit.
- Browser visual verification and runtime activation require separate evidence from source validation.
