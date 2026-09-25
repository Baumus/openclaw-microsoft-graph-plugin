# OAuth delegated permissions and access matrix

The plugin uses delegated Microsoft Graph permissions only. It does not support application permissions, client secrets, certificates, arbitrary Graph scopes, or arbitrary Graph URLs.

`offline_access` is needed when initially obtaining a refresh token. For each authorized operation, the plugin exchanges the shared credential for one service-specific scope from the table below.

## Shared credential consent set

The one credential must contain every scope needed by enabled policy grants, plus `offline_access`. Write scopes may satisfy the corresponding read exchange where Microsoft defines that implication. Grant only workloads and operations actually enabled by policy.

This is not provider-side read/write isolation. A shared-credential compromise or plugin bypass has the union of consented scopes. Default-deny policy and write approvals remain the normal runtime authority.

## Tool-to-scope matrix

| Tool / actions | Delegated scope requested | Policy operation | Approval class |
| --- | --- | --- | --- |
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
| `outlook_mail_write`: `move` | `Mail.ReadWrite` | Mail `move` | warning |
| `outlook_mail_write`: `mark_read` | `Mail.ReadWrite` | Mail `mark` | warning |
| `outlook_mail_write`: `send_draft` | `Mail.Send` | Mail `send` | critical |
| `outlook_mail_write`: `delete` | `Mail.ReadWrite` | Mail `delete` | critical |
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
| To Do | Signed-in user's `/me/todo` resources; writes reject shared/non-owned lists |

Microsoft tenant administrators remain responsible for consent policy, conditional access, token revocation, account lifecycle, and reviewing the effective delegated permissions.
