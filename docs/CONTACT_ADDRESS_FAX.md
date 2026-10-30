# Contact postal addresses and fax numbers

The contact tools remain limited to the signed-in user's default `/me/contacts` collection. No new permission, arbitrary Graph route, custom property identifier, contact folder, shared contact, or directory access is introduced.

## Tool fields

`outlook_contacts_read` returns `businessAddress`, `homeAddress`, and `otherAddress` with `street`, `city`, `state`, `postalCode`, and `countryOrRegion`. List, get, and bounded client-side search include these addresses and the normalized `businessFaxNumber`, `homeFaxNumber`, and `primaryFaxNumber` fields. Search also matches postal components and fax numbers; the existing 500-contact scan ceiling and partial-result handling remain in force.

`outlook_contacts_write` accepts the same named fields for create/update:

- Omit an entire address or fax field to leave it unchanged.
- An address is a **complete replacement**, not a component patch. Supply all five string components; use `""` for unused parts. Read the existing address first when changing one component. This requirement follows Microsoft's structured-property update semantics and prevents accidental partial-address loss.
- Pass `null` for an address to explicitly clear all five components.
- Pass `""` for a fax field to explicitly clear its stored number. Fax fields otherwise accept nonblank strings up to 100 characters; control characters and line breaks are rejected.
- Nested unknown fields and incorrect types are rejected before credentials, provider access, or approval. Address limits are 1,024 characters for street, 64 for postal code, and 255 for the other components.

For example, an address update can be requested as:

```json
{
  "action": "update",
  "contactId": "EXACT_CONTACT_ID_FROM_READ",
  "businessAddress": {
    "street": "1 Example Street",
    "city": "Berlin",
    "state": "BE",
    "postalCode": "10115",
    "countryOrRegion": "Germany"
  },
  "businessFaxNumber": "+49 30 12345"
}
```

## Native Outlook storage

Microsoft Graph v1.0 does **not** define a native `faxNumber` contact field. The plugin uses Graph's supported single-value legacy extended-property API with three fixed, Outlook-compatible MAPI properties:

| Tool field | Graph extended-property ID | Official MAPI property |
| --- | --- | --- |
| `businessFaxNumber` | `String 0x3A24` | `PidTagBusinessFaxNumber` |
| `homeFaxNumber` | `String 0x3A25` | `PidTagHomeFaxNumber` |
| `primaryFaxNumber` | `String 0x3A23` | `PidTagPrimaryFaxNumber` |

Fax values are not placed in voice-phone arrays or invented custom properties. Writes upsert only explicitly supplied IDs and do not send omitted fax IDs or unrelated extended properties. Reads expand the fixed allowlist and expose only normalized named fax fields, not raw extended-property collections. They use the existing delegated `Contacts.Read` / `Contacts.ReadWrite` permissions.

Graph omits new extended properties from mutation responses. After a fax write, the plugin therefore reads the exact saved contact and returns provider values (`verification: "read_back"`), rather than echoing input as evidence. A requested nonempty fax must be present and nonempty in the readback (provider formatting normalization is allowed); an explicitly cleared fax may be absent or empty. Missing requested values, duplicate IDs, malformed values, and a clear that remains nonempty are unverified. If readback fails after an accepted write, the result retains `mutationApplied: true` and reports `applied_with_warning`, `appliedButUnverified: true`, and `verification: "failed"`. Read the exact target before any retry; do not repeat the accepted write. A successful readback is evidence of the returned provider state, not a claim that every Outlook client displays the same fields. Real-account / Outlook-client roundtrip was not performed by the synthetic test suite.

## Authoritative references

Verified against Microsoft documentation on 2026-10-30:

- [Graph contact resource](https://learn.microsoft.com/en-us/graph/api/resources/contact?view=graph-rest-1.0)
- [Update contact, including full structured-property replacement](https://learn.microsoft.com/en-us/graph/api/contact-update?view=graph-rest-1.0)
- [Extended-property overview and proptag IDs](https://learn.microsoft.com/en-us/graph/api/resources/extended-properties-overview?view=graph-rest-1.0)
- [Create single-value extended properties in new/existing contacts](https://learn.microsoft.com/en-us/graph/api/singlevaluelegacyextendedproperty-post-singlevalueextendedproperties?view=graph-rest-1.0)
- [Read single-value extended properties](https://learn.microsoft.com/en-us/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0)
- [Business fax MAPI ID](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagbusinessfaxnumber-canonical-property)
- [Home fax MAPI ID](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtaghomefaxnumber-canonical-property)
- [Primary fax MAPI ID](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagprimaryfaxnumber-canonical-property)
