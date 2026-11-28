# Sandbox transfers (proposed; not release-ready)

## Required transfer parity

The sandbox path must cover every existing binary file transfer: OneDrive download, create and replace; Outlook mail and calendar attachment download; and sandbox file input for Outlook mail, Outlook calendar and Microsoft To Do attachment uploads. Existing operation-specific limits, configured limits, authorization and approval remain in force. The 1 MiB limit in the prototype is not a product requirement or an OpenClaw limit. Do not ship it or replace it with another sandbox-specific total file-size cap.

Keep memory and individual transport messages bounded by streaming or bounded chunks. The chunk size is not the file-size limit. Verify transfers above the former prototype ceiling, provider-limit enforcement, interruption cleanup and byte identity for each transfer path. Explicitly report unsupported sandbox capabilities or local resource failures instead of falling back to the host.

The ordinary Graph download remains unchanged. An optional sandbox copy may be created only from bytes produced by the same authorized download, never from an arbitrary private-media URI supplied by the model. An unavailable sandbox must not cause a host write. A failed optional copy returns the ordinary private download plus a bounded copy-error status.

## Runtime integration

The existing defineToolPlugin helper registers a v1 factory in both the pinned 2026.11.5 SDK and installed 2026.11.10 runtime. Introduce a typed contextVersion: 2 registration path for sandbox-aware tools, preserving manifest metadata and existing invocations on supported older hosts. Before wiring Graph, verify the actual host supplies sandbox.run to that factory and that invocation lifetime and session replacement are enforced. Do not infer this from workspaceDir or discover Docker containers by name.

Use a bounded fixed argv helper via sandbox.run. The SDK input/output is text, so transport binary internally with bounded encoding, never through tool schemas, transcripts or diagnostics. Select a dedicated private session-relative subtree. Reject traversal, links, non-regular files and over-limit inputs; open targets exclusively and check sizes and hashes. Do not expose raw host paths or direct access to the media store.

## Uploads

Use a sandbox-only staging action to snapshot one finished file into the existing managed artifact flow, or integrate staging into the same call only if the host provides the same invocation-bound sandbox capability before native approval. Do not register an arbitrary host path or a reusable sandbox capability. Bind staged artifacts to authenticated agent and session generation and enforce that binding at the consumption boundary. OneDrive create/replace and Outlook draft attachments must use the existing authorization, hash-bound native approval and ETag/no-overwrite rules; staging is not a Microsoft write. Revalidate before provider access. Preserve existing source selectors.

## Release proof

Test the unmodified default, byte-identical optional copy, missing/stale sandbox, sibling session refusal, traversal, links, limits, approval refusal and exact hash on mock provider writes. Exercise the real host runtime with synthetic XLSX data and no Microsoft mutation. Qualify the minimum supported OpenClaw release as well as the installed version; do not silently raise the package-wide compatibility floor. Then run the project pipeline, package the reviewed merged commit, install that exact archive locally independently of ClawHub, and verify public publication and local activation separately.
