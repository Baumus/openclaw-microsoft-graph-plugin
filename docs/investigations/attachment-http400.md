# Attachment list HTTP400 investigation

## Live evidence (2026-11-15)

The parent operator used the authorized connector diagnostic build against two distinct incoming messages. For both, baseline324, withoutTop, idWithTop, idOnly, name, contentType, size, isInline and lastModifiedDateTime returned HTTP200. Selecting id,contentId returned HTTP400 with allowlisted provider code BadRequest. No mailbox identifiers, provider prose or attachment contents are recorded here.

This isolates the derived-only contentId selection as the cause of the 3.25.0 query failure. It does not establish the cause of any historical 3.24.0 failure. PR90 workspace attachment upload and the separate OneDrive PDF issue are unrelated.

## Fix and regression

Remove contentId from the base attachment collection $select. Keep the existing bounded metadata output allowlist, including contentId when a provider returns it, and preserve top, continuation, policy and download boundaries. Remove the temporary attachmentDiagnostic tool field and matrix before stable publication. A regression simulates Graph rejecting contentId selection, verifies the exact base query and checks that contentBytes never reach public output.

Existing buffered transport still parses at most 16 KiB of error JSON and returns only allowlisted provider codes; the added HTTP400 codes never expose raw messages, URLs or bodies. Tests exercise transport and lifecycle sanitization.

## Security and recovery

User authorized diagnosis, PR, merge, release, ClawHub publication and local installation/activation. No scope, credential, consent, account or policy changes. Mail/provider data remains untrusted; no provider text is executed or published. Existing OAUTH_ACCESS_MATRIX.md governs unchanged access. Recovery is installation of the previous stable package; no data migrations or mailbox writes occur. Final regular connector list/download verification belongs to the parent operator, not the producer tests.

## References

https://learn.microsoft.com/en-us/graph/api/resources/attachment?view=graph-rest-1.0 lists base properties; contentId belongs to fileAttachment. Live comparisons, not documentation alone, establish this incident's cause.
