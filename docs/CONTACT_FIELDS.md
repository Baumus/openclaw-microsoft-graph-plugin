# Contact property coverage

Verified against the Microsoft Graph v1.0 [contact resource](https://learn.microsoft.com/en-us/graph/api/resources/contact?view=graph-rest-1.0) and [update contract](https://learn.microsoft.com/en-us/graph/api/contact-update?view=graph-rest-1.0), 2026-11-04.

All documented resource properties are returned by list/get/search when supplied by Graph. Search checks all exposed property values, including private landlines and notes, within the existing 500-record scan cap. Provider-only keys inside email/address objects are removed. No additional delegated scopes or policy grants are needed.

## Read and write

- Names: `assistantName`, `displayName`, `fileAs`, `generation`, `givenName`, `initials`, `middleName`, `nickName`, `surname`, `title`, `yomiCompanyName`, `yomiGivenName`, `yomiSurname`.
- Dates/family: `birthday`, `children`, `spouseName`.
- Work: `businessHomePage`, `companyName`, `department`, `jobTitle`, `manager`, `officeLocation`, `profession`.
- Phone and messaging: `businessPhones`, `homePhones`, `mobilePhone`, `imAddresses`.
- Email: `emailAddresses`, `primaryEmailAddress`, `secondaryEmailAddress`, `tertiaryEmailAddress`. Each slot is a complete `{address, name?}` object. Avoid conflicting slot/collection values in the same write; Graph owns their reconciliation.
- Postal: `businessAddress`, `homeAddress`, `otherAddress` (full replacements; see [address/fax contract](CONTACT_ADDRESS_FAX.md)).
- Organization: `categories`, `personalNotes`, `parentFolderId`. Changing `parentFolderId` changes folder membership; resolve the exact destination and retain the returned ID, which can change after moving. Contact discovery remains scoped to the default folder and its existing policy boundary.
- Outlook fax additions: `businessFaxNumber`, `homeFaxNumber`, `primaryFaxNumber` map to fixed MAPI properties as before.

Omit a field to preserve it. Arrays replace the entire collection; `[]` clears. Nullable scalar strings and email slots accept `null`; empty scalar strings also remain supported. `birthday` accepts a real `YYYY-MM-DD` calendar date (normalized to midnight UTC), or a valid timestamp with timezone, or `null` to clear. An address requires every component; `null` clears its components. Fax clears with `""`. Names are bounded to 255 characters, phone arrays to two entries, email collection to three, children/IM arrays to 100, notes to 32768 characters. These are bounded tool contracts, not claims about universal Outlook-client limits.

## Read only

`id`, `changeKey`, `createdDateTime`, `lastModifiedDateTime` are returned as provider metadata and rejected on mutation. `contactId` selects an existing contact; it does not overwrite its `id`.

## Relationships are not fields

`photo`, `extensions`, `singleValueExtendedProperties`, and `multiValueExtendedProperties` are Graph relationships with separate APIs, not arbitrary writable contact properties. They are not accepted or exposed wholesale. Only the three existing fixed fax MAPI properties are expanded and normalized. Photo and arbitrary custom-extension management are outside these contact tools.

Native approval, per-agent policy, credential separation, private-output boundaries, and create/update/delete classification remain unchanged. Tests use synthetic provider responses; no real-account or Outlook-client write was performed for this release.
