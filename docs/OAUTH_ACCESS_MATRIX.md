# OAuth delegated permissions and access matrix

The plugin uses delegated Microsoft Graph permissions only. It does not support application permissions, client secrets, certificates, arbitrary Graph scopes, or arbitrary Graph URLs.

`offline_access` is needed when initially obtaining a refresh token. For each authorized operation, the plugin exchanges the shared credential for one service-specific scope from the table below.

## Shared credential consent set

The one credential must contain every scope needed by enabled policy grants, plus `offline_access`. Write scopes may satisfy the corresponding read exchange where Microsoft defines that implication. Grant only workloads and operations actually enabled by policy.

This is not provider-side read/write isolation. A shared-credential compromise or plugin bypass has the union of consented scopes. Default-deny policy and write approvals remain the normal runtime authority.

## Tool-to-scope matrix

| Tool / actions | Delegated scope requested | Policy operation | Approval class |
| --- | --- | --- | --- |
| `microsoft_graph_capabilities` | none; local policy projection only | effective caller grants | none |
| `onedrive_search`, `onedrive_list`, `onedrive_read`, `onedrive_download`, `onedrive_agents_instructions` | `Files.Read` | OneDrive `read` | none |
| `onedrive_upload`, `onedrive_update`, `onedrive_metadata_update`, `onedrive_create_folder` | `Files.ReadWrite` | OneDrive `write` | warning |
| `onedrive_delete` | `Files.ReadWrite` | OneDrive `delete` | critical |
| `outlook_calendar_read`: all actions | `Calendars.Read` | Calendar `read` | none |
| `outlook_calendar_write`: `create` | `Calendars.ReadWrite` | Calendar `create` | warning |
| `outlook_calendar_write`: `update` | `Calendars.ReadWrite` | Calendar `update` | warning |
| `outlook_calendar_write`: `multiwrite` (maximum 100 operations) | `Calendars.ReadWrite` | Calendar `create` and/or `update` for every target calendar | warning |
| `outlook_calendar_write`: `attach` | `Calendars.ReadWrite` | Calendar `attach` | warning |
| `outlook_calendar_write`: `respond` | `Calendars.ReadWrite` | Calendar `respond` | critical |
| `outlook_calendar_write`: `delete` | `Calendars.ReadWrite` | Calendar `delete` | critical |
| `outlook_mail_read`: all actions | `Mail.Read` | Mail `read` | none |
| `outlook_mail_write`: `create_draft`, `update_draft`, `reply_draft`, `reply_all_draft`, `forward_draft` | `Mail.ReadWrite` | Mail `draft` | warning |
| `outlook_mail_write`: `update_properties`, `copy`, `add_attachment` | `Mail.ReadWrite` | Mail `update` | warning |
| `outlook_mail_write`: `add_message_attachment` (original MIME → existing draft) | `Mail.ReadWrite` | Mail `read` **and** `update` | warning |
| `outlook_mail_write`: `move` | `Mail.ReadWrite` | Mail `move` | warning |
| `outlook_mail_write`: `mark_read` | `Mail.ReadWrite` | Mail `mark` | warning |
| `outlook_mail_write`: `send_draft` | `Mail.Send` | Mail `send` | critical |
| `outlook_mail_write`: `delete` | `Mail.ReadWrite` | Mail `delete` | critical |
| `outlook_contacts_read`: list, search, get | `Contacts.Read` | Contacts `read` | none |
| `outlook_contacts_write`: `create`, `update` | `Contacts.ReadWrite` | Contacts `create`, `update` | warning |
| `outlook_contacts_write`: `delete` | `Contacts.ReadWrite` | Contacts `delete` | critical |
| `microsoft_todo_read`: all actions | `Tasks.Read` | To Do `read` | none |
| `microsoft_todo_write`: create/add actions | `Tasks.ReadWrite` | To Do `create` | warning |
| `microsoft_todo_write`: update actions | `Tasks.ReadWrite` | To Do `update` | warning |
| `microsoft_todo_write`: delete actions | `Tasks.ReadWrite` | To Do `delete` | critical |

Unknown or missing read actions fail closed before provider access. Unknown or missing mutation actions never receive a lower approval classification.

## Resource access

| Workload | Resource boundary |
| --- | --- |
| OneDrive | Exact policy-pinned `drive_id` + `item_id`. `include_descendants` must be literal `true`; omit the root rather than setting `false` when descendants must not be allowed. |
| Calendar | Default calendar resource `me` or an exact policy-allowlisted calendar ID |
| Mail | Signed-in user's `/me` mailbox only |
| Contacts | Signed-in user's default `/me/contacts` collection only; no folders, shared contacts, or organization directory |
| To Do | Signed-in user's `/me/todo` resources; writes reject shared/non-owned lists |

Microsoft tenant administrators remain responsible for consent policy, conditional access, token revocation, account lifecycle, and reviewing the effective delegated permissions.

Existing vault credentials do not gain newly configured Contacts scopes automatically. After the policy is saved and applied, credential status reports missing scopes and the Control UI offers **Renew Microsoft permissions**, using the existing vault app/tenant without manual identifiers or commands. The Gateway requests all scopes required by the active policy and atomically replaces a valid, binding-matched vault only after Microsoft returns the complete grant. A failed or denied reauthorization retains the prior credential; actions requiring missing scopes fail closed in the meantime. Initial sign-in remains create-only.

## Configuration activation recovery (3.18.4)

| RPC action | Authority | Data / effect |
| --- | --- | --- |
| `configuration.activation`: `prepare` | `operator.admin`, authenticated user or paired device | Validates the native saved baseline and fresh config CAS; writes one bounded intent, not config or provider data |
| `status` | Same admin identity and Gateway state directory | Returns only that identity's matching saved-policy intent, never a policy draft or credential |
| `claim` | Same admin; exact saved **and generation-active** policy, authored source, target-agent tool baseline and global deny | Atomically marks one tool dispatch issued; returns the exact reviewed tool additions and fresh native CAS |
| `complete` / `cancel` | Same admin identity and intent ID | Closes the journal; completion additionally verifies saved tool target and active policy. Cancellation changes no grants |

Method prefix: `microsoft-graph.`. Identity comes from the Gateway's authenticated user or signature-verified `client.connect.device.id`, never RPC parameters or the shared `pairedClientId` client class. Without a stable authenticated identity, durable preparation fails closed before policy is written. A different administrator cannot resume or replace an unfinished intent. Expiry is 24 hours; expired, cancelled or completed intents never dispatch tools. A fresh explicit review is required afterward.

The single intent lives in `<stateDir>/plugin-data/microsoft-graph/activation/pending.json` (directory `0700`, file `0600`), atomically published under a fail-closed filesystem lock. It contains hashes, opaque operation ID, timestamp and exact added tool names only: no policy/resource draft, credential, include path or user identifier. Existing grants, global and per-agent denies remain unchanged by journal operations. Config writes still use native `config.patch` and fresh CAS, with the included policy applied before root agent-tool exposure.

A fresh browser or plugin generation automatically resumes a prepared intent after exact policy activation. An issued but unconfirmed dispatch is **not** automatically repeated: the UI reads back persistence/activation, or offers cancellation and a new explicit review. This is intentional unknown-outcome handling, not a claim of automatic completion for an ambiguous dispatch. Closing the browser does not authorize a background daemon: continuation resumes when the same administrator reopens the page. No provider calls, tokens or live grant changes are needed to test recovery.

## Control UI tool self-healing boundary (3.19.0)

| Concern | Authority / permitted effect |
| --- | --- |
| Trigger | Authenticated `operator.admin` Control UI load/status check or completed save; no daemon/provider call |
| Desired permissions | Validated saved policy, matching the running plugin policy fingerprint |
| Repairable surface | Exact Microsoft Graph-owned names in existing agent `allow` or `alsoAllow`; union across every saved service/root |
| Preserved authority | Global/agent deny, wildcard/group grants, unrelated names, service/resource restrictions and native approval gates |
| Blocked state | Missing/conflicting agent configuration, final restrictive allow entry, denied tool, changed CAS revision/source, inactive policy or pending claim |
| Concurrency / recovery | Private locked administrator-bound journal claims exact targets before one native CAS; fresh pages observe issued claims without replay |
| Persistence / activation | Native config acknowledgement or applied config revision is required before tool activation is confirmed; no extra reload/restart |

Synthetic planner, journal, UI race and Chromium fixtures use no credentials or provider data. Journals contain only hashes, agent IDs and tool names, not policies/resources or credentials. Existing secure atomic-file permissions and lock checks remain unchanged. No backup/restore artifact is required: tests mutate only disposable fixture files, and the released feature uses the host's revision-checked configuration writer and existing current-state recovery UI. No new AI retrieval/output path, credential scope or Graph permission is introduced.


## Caller discovery contract (3.19.1)

`microsoft_graph_capabilities` is a local, credential-free policy projection. Every saved effective Microsoft grant includes its tool opt-in; existing installations receive the same deny-aware administrator Control UI reconciliation as operation tools. The activation journal admits the helper name. Removal prunes only exact owned helper names and never empties a restrictive allowlist. Explicit denies, wildcard/operator grants, and unchanged-peer boundaries remain authoritative.

OneDrive tools and schemas direct the caller to discover first and use the exact returned `rootLabel`, not an inferred agent ID, title, or path. The additive `rootPath` field is returned only after intersecting root and caller permissions; no foreign-agent roots/grants or drive/item IDs are projected. It identifies a configured folder, not a provider connectivity check. Local access denials keep a generic code and safe discovery guidance rather than disclosing whether a foreign root exists. Absent or ambiguous access requires operator review, not guessed-label retries or automatic consent repair.

`onedrive_agents_instructions` is included only for trusted roots with effective read permission. Trusted writes/deletes still require independently authorized instruction reads; capability discovery explains the unmet read prerequisite for write-only callers without granting it. No runtime agent can use discovery to change policy, impersonate another caller, or bypass native approvals.
