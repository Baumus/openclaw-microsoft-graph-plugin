# Gemacode Native execution gate

`nativeExecutionRequired` is an optional production gate for Microsoft Graph
effects. It is disabled by default and is separate from the existing HMAC
Native provider boundary.

When enabled, each provider effect requires:

1. a preparation callback created by the QEL plugin for the same tool-call ID;
2. a valid Ed25519 permit signed by the pinned Native OS public key;
3. exact tool name, tool-call ID and final canonical parameter digest;
4. successful one-use consumption over the pinned Unix socket immediately
   before provider execution;
5. post-effect Native completion backed by a separate provider readback and
   Evidence.

The tool returns failure if preparation, signature verification, consumption,
observation or Evidence differs. Permits cannot be replayed. Their lifetime is
at most 30 seconds. The application never receives the Native private key.

Configuration fields:

- `nativeExecutionRequired`
- `nativeExecutionPublicKey` (protected configuration resolving to
  `b64u:<32 raw Ed25519 public-key bytes>`)
- `nativeExecutionSocketPath`
- `nativeExecutionSocketOwnerUid`

`microsoft_graph_capabilities` and `onedrive_agents_instructions` are local
control-plane reads, not external provider effects; they continue through QEL
and Native DB but do not consume an external-effect permit.

Keep this mode disabled on macOS until an Apple-specific Rooted attestation
profile is implemented and reviewed. The current production service depends on
Linux peer credentials and the Native OS measured-boot trust profile.
