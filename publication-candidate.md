# Local publication candidate — OneDrive native approval integrity (3.25.2)

## Identity and ownership
- Origin: https://github.com/Baumus/openclaw-microsoft-graph-plugin.git
- Checkout: /tmp/onedrive-approval-candidate
- Branch: fix/onedrive-approval-integrity
- Base: 31a6d4285a1f01573db1b21504f2f04230240a42
- Implementation commit: ad08a1f812273d15df9691f61323d2aa170c235f
- Final candidate: git HEAD containing this record (full hash returned in handoff).
- Main is sole next owner: obtain independent semantic review, then own authorized publication/delivery. BS owns Microsoft operations.
- No remote writes, Microsoft operations, live install/activation/config changes or integrity bypass performed.
- Read-only git ls-remote confirmed main remains at base and latest matching tag v3.25.1; v3.25.2 absent on 2026-11-16. Recheck before publication.

## Proposed PR title
fix(onedrive): preserve original native approval input while privately binding staged content

## Proposed PR description
### Root cause
OneDrive preflight stages sourceWorkspacePath and augments hook parameters with contentType,
sourceMediaUri, sourceSha256 and sourceByteSize for a minimal synthetic PDF upload.
Core correctly rejects changed native relay parameters, both immediately and after deferred
approval. This is not a timeoutMs problem. No Core change is needed.

### Change
Preserve original OneDrive upload/update caller input through the hook.
Store normalized original caller criteria plus serialized private artifact execution parameters
in the single-use approval snapshot. Hydrate only after caller/tool/session/current-root and
current-session checks succeed. Approval text retains staged content fingerprints; execution
reopens/verifies protected media before credentials/network. Advance shared protocol v2 to v3
to fail closed across mixed hot-loaded implementations. To Do handoff and Outlook mail
attachment staging remain unchanged. Align package, lock root, manifest, npmSpec and changelog
to patch 3.25.2; regenerate metadata using locked SDK (no other generated tracked changes).

### Regression proof
Red: two synthetic minimal-input tests failed with precisely the four added fields above.
Green: unchanged original input reaches the synthetic credential gate with warning approval
enabled/disabled. Added staged URI/hash/size/MIME injection rejection, cross-module staged-byte
execution and replay coverage. Existing tests cover source selector changes, reset generations,
denial cleanup, fingerprint assertions and protected media mutation.
Earlier reused-dependency run: 925/925 before extra cases, then 928/929 with an unchanged
foreign-file quota-lock timeout; isolated staging rerun 31/31. Superseded by fresh locked
dependency run below: 929/929 including the formerly flaky test (3442 ms, unchanged 5s gate).
No timeout relaxation, test exclusion or production security change.

## Final terminal evidence (2026-11-16)
Node 26.13.1, npm 11.20.0. Removed only excluded node_modules symlink with unlink, then npm ci
(exit 0) created an independent directory. Original linked dependencies were not modified.
Lock dependency resolution unchanged; only root package version advanced.
Actual dependencies: OpenClaw 2026.11.2, Vitest 5.0.5, TypeScript 7.0.2, esbuild 0.28.2,
ClawHub 0.25.1, happy-dom 21.0.4, typebox 1.4.6, fs-safe 0.28.0, json5 2.2.3.
- npm test -- --maxWorkers=2: exit 0, 55 files / 929 tests passed, 80.42s.
- npm run typecheck: exit 0.
- npm run plugin:build: exit 0; includes npm run build and its typecheck, esbuild runtime and UI/metadata generation.
- npm run plugin:build:check: exit 0, metadata up to date.
- npm run plugin:validate: exit 0, valid=true, no errors.
- npm run package:check: exit 0, exact expected 24 package files.
- node scripts/check-include-save.mjs: exit 0; expected combined cross-owner rejection, split writes PASS.
- npx --no-install clawhub package validate .: exit 0, PASS, 0 breakages/warnings/findings.
- git diff --check: exit 0.
- npm pack --ignore-scripts --json --pack-destination /tmp: exit 0 (already built artifact).
Host CLI build/validation used only OPENCLAW_STATE_DIR=/tmp/onedrive-fresh-host-state and
OPENCLAW_CONFIG_PATH=/tmp/onedrive-fresh-host-state/openclaw.json, not live state.
Terminal logs: /tmp/onedrive-fresh-evidence/{npm-ci,test,typecheck,build,plugin-build-check,plugin-validate,package-check,include-save,clawhub,pack}.log.
Inspector reports: reports/plugin-inspector-report.{json,md} and reports/plugin-inspector-issues.md (ignored).

## Exact artifacts (SHA-256)
- /tmp/baumus-openclaw-microsoft-graph-3.25.2.tgz (499793 bytes, 24 files):
  6a7f2798fe86a386e4784c412da6e4971a11e96986060817af59ca035da44336
- dist/index.js: 2a2f37e28fe36cb1b7eeefcf0b3a1acf163378eaecee36d6e507a5a0120bfbe4
- openclaw.plugin.json: 39691d536b435ec09f773820a9ff293d47e13bbb11d9e962811a9ac4278128af
- package.json: 96019d0a8a20ff004cd8271ae518994efa913591981f38b497d5667280010721
- package-lock.json: 0ed24add6d43e7f050596e7a9854ec323875a4a32c393026b4f86b8a20c6209b
- UI entry: dist/control-ui/d8d1f978ff4daf15d55ea71bb563b878139415dabd9414e9dd13de9f2056c521/index.js
Publication record is excluded by package files allowlist; updating it does not alter this artifact.

## Proposed release notes
Fix OneDrive uploads and updates blocked by native Codex approval parameter integrity checks.
Preserve original caller input while privately binding approved staged bytes. Retain exact-call,
session, root, fingerprint and cleanup protections. No permission or configuration changes.

## Changed-file map
- src/index.ts: private staged-parameter binding and protocol v3.
- src/{approval-snapshot-cross-realm.test.ts,content-identity-confirmation.integration.test.ts,one-call-staging.integration.test.ts,session-key-approval.integration.test.ts,workspace-context-cross-module.test.ts}: protocol isolation, caller injection, unchanged relay input, staged bytes/replay regression coverage.
- CHANGELOG.md, package.json, package-lock.json, openclaw.plugin.json: aligned release metadata.
- publication-candidate.md: review/publication evidence.

## Review focus, risks and limitations
Review caller binding versus private handoff, single-use consumption, root/session checks,
cleanup ordering, cross-module isolation, content revalidation before credentials.
Producer evidence is not independent approval. No live Codex/Graph end-to-end test performed.
Outlook attachment staging has a similar rewrite pattern and is not claimed fixed.
npm ci reported 28 dependency advisories (24 moderate, 4 high); no audit fix or dependency
mutation performed. npm also reported four install scripts not covered by allowScripts;
build and all checks nevertheless passed. Node 24/macOS/CI stable-host override matrix
not run locally. Configured builder route retained; no model escalation; exact resolved model
is not independently exposed in this execution evidence.

## Release/distribution and recovery
CONTRIBUTING.md gates are recorded above; generated metadata inspected.
Main must review exact final source, recheck version availability, then perform authorized
PR/merge/tag/release and ClawHub publication from its owner-started runtime.
Default distribution: clawhub:@baumus/openclaw-microsoft-graph; npm spec
@baumus/openclaw-microsoft-graph@3.25.2. No repository release automation/exact publish
command was found; local Inspector is not publication evidence.
Recovery: discard isolated checkout or revert both local commits. No data migration;
approval snapshots fail closed across protocol version transitions. Any post-release
rollback/install remains Main's authorized action.
