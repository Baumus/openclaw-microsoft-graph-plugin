# Local publication candidate — OneDrive native approval integrity

## Identity and ownership
- Origin: https://github.com/Baumus/openclaw-microsoft-graph-plugin.git
- Checkout: /tmp/onedrive-approval-candidate
- Branch: fix/onedrive-approval-integrity
- Base: 31a6d4285a1f01573db1b21504f2f04230240a42 (freshly fetched origin/main).
- Candidate commit: see git HEAD after local handoff commit.
- Main owns independent review, publication and delivery. BS owns actual Microsoft actions.
- No remote write, Microsoft canary, activation, install, policy/config/credential change performed.

## Proposed PR title
fix(onedrive): preserve original native approval input while privately binding staged content

## Proposed PR description
### Root cause
OneDrive preflight stages sourceWorkspacePath and returns augmented hook parameters.
The exact added fields for a minimal synthetic PDF upload are contentType,
sourceMediaUri, sourceSha256 and sourceByteSize. This is not a timeoutMs problem.
Core's Codex native relay correctly rejects a changed request immediately without deferred
approval and again after deferred approval resolution. The original app-server request
cannot safely execute the rewritten request. No Core change is needed.

### Change
Keep original OneDrive upload/update caller parameters intact across the hook.
Record original normalized caller criteria plus serialized, private artifact execution
parameters in the existing single-use approval snapshot. Hydrate those parameters
only after caller/tool/session/current-root and current-session checks succeed.
Approval text still derives from staged content, including computed fingerprint;
transfer still reopens/verifies protected media before credentials/network.
Version the shared snapshot protocol from v2 to v3 so mixed hot-loaded implementations
fail closed. Existing To Do handoff and Outlook mail staging behavior remain unchanged.

### Producer evidence
Before implementation, the two new minimal-input tests failed with exactly the four
added fields above (warning approval enabled and disabled). After implementation both
pass and reach the synthetic credential gate using unchanged original input.
Additional tests reject injection of staged URI, SHA-256, size and MIME type; existing
tests cover changed source selector, reset generations, denial cleanup, replay,
cross-module staged bytes, fingerprint assertions and protected-media mutation.
Earlier full suite: 55 files / 925 tests passed. Final suite after four injection tests and v2 isolation test update: 928 passed / 1 timed out (unchanged foreign-file quota-lock test, 5000 ms). Targeted unchanged staging-suite rerun: 31/31 PASS; formerly timed-out test passed in 2179 ms. Final full-suite timeout remains disclosed.
Build/typecheck: PASS. Generated plugin build: PASS, no tracked metadata change.
plugin:build:check: PASS; plugin:validate: valid=true; package:check: PASS.
check-include-save: PASS (expected rejected combined write; synthetic split writes preserve include).
ClawHub local package Inspector: PASS, zero breakages/warnings/findings.
git diff --check: PASS.

### Environment qualification / limitations
Node 26.13.1. Reused existing development node_modules via a local excluded symlink:
OpenClaw 2026.10.5 and Vitest 5.0.2 (not the package.json's current 2026.11.2/5.0.5).
A fresh npm ci from committed lockfile is still required in independent review/CI.
Initial host CLI checks refused newer live DB schema (22 versus supported 20);
rerun build/validation used only disposable OPENCLAW_STATE_DIR and OPENCLAW_CONFIG_PATH
under /tmp/onedrive-candidate-host-state. No database migration or live-state mutation.
Initial package check lacked generated UI; plugin:build generated it and recheck passed.
No native live end-to-end Codex invocation or real Microsoft operation was attempted.
Outlook attachment staging has a similar rewrite pattern but is deliberately outside this
OneDrive-only fix; do not claim that route is fixed.

## Proposed release notes (suggest patch 3.25.2; release owner chooses)
Fix OneDrive uploads and updates being blocked by native Codex approval parameter
integrity checks when private staging adds content metadata. Preserve original caller
input while binding the approved bytes privately; retain exact-call, session, root,
fingerprint and cleanup protections. No permission or configuration changes.

## Review focus
- Original caller binding versus serialized private artifact handoff; no caller injection.
- Single-use consumption, current-root/session validation and staging cleanup ordering.
- Cross-module v3 isolation and safe failure for mixed old/new loaded bundles.
- No loss of content hash verification or credential-before-validation protections.
Producer checks are not independent security/release approval.

## Release/distribution handoff
Repository CONTRIBUTING.md requires npm ci then npm test, npm run typecheck,
npm run build, npm run plugin:build:check, npm run plugin:validate,
npm run package:check, node scripts/check-include-save.mjs,
npx --no-install clawhub package validate ., git diff --check.
For changed runtime metadata use npm run plugin:build and inspect generated diff.
README.md declares clawhub:@baumus/openclaw-microsoft-graph and npm package
@baumus/openclaw-microsoft-graph; default installation is ClawHub.
Main must align package/lock/manifest/npmSpec version, changelog and reviewed source,
then complete authorized PR/merge/release/ClawHub publication and activation in its
owner-started runtime. This candidate intentionally leaves version unchanged.
No repository release automation or exact publish command was found; do not infer
publication success from local Inspector results.

## Recovery
Discard this isolated checkout before publication, or revert the candidate commit.
After release, Main can restore the prior package through its authorized install flow.
No data migration; pre-existing approval snapshots must not survive version transitions.
