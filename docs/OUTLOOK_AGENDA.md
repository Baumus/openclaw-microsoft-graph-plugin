# Get your Outlook agenda in OpenClaw with read-only access

Use one named agent to summarize tomorrow's events from one Outlook calendar. The agent does not create, edit, or respond to events.

## Set up the smallest useful grant

1. Follow the [plugin setup guide](../README.md#get-started-in-three-steps). In your default-deny policy, grant only calendar `read` to the named agent for the signed-in user's default calendar (`me`), or select one exact calendar ID. Do not grant calendar write operations or mail, contacts, OneDrive, or tasks for this example.
2. Save and apply the policy, then connect the Microsoft account. The plugin uses delegated `Calendars.Read` for calendar reads; see the [OAuth access matrix](OAUTH_ACCESS_MATRIX.md). The per-agent policy limits the plugin's access within that Microsoft scope.
3. Ask the allowed agent to run the prompt below. To verify the access boundary, ask an agent with no calendar grant to make the same read; access should be denied.

## Copy this prompt

> List tomorrow's appointments from my allowed Outlook calendar, in Europe/Berlin time. Use that calendar only. Show start and end times and titles, in chronological order. Don't change anything, respond to invitations, or send messages. If you can't read the calendar, say so rather than guessing.

Replace `Europe/Berlin` with your time zone. If you selected a specific calendar rather than the default, name it in the prompt.

Illustrative output (fictional events):

- 09:00–09:30 — Weekly planning
- 11:15–12:00 — Project review
- 14:00–14:30 — Vendor check-in

If the plugin is useful in your Microsoft 365 workflow, you can [sponsor its maintenance on GitHub](https://github.com/sponsors/Baumus).
