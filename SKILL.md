---
name: se-meeting-prep
description: |
  Weekly meeting prep for SE calendar. Reads Google Calendar for the current week,
  identifies customer meetings (prioritizing your VIP AE meetings), looks up
  attendee job titles from Snowflake SFDC, searches Google Drive for master_notes
  files, and writes a structured pre-call agenda in two places:
    1. An email to yourself (sent via SMTP) — easy to review before the meeting
    2. The calendar event description — original content preserved at top

  Triggers: meeting prep, prep my meetings, weekly prep, calendar prep,
  pre-call prep, meeting agenda, prep for this week, meeting notes,
  prep meetings, se prep, prepare for meetings, weekly calendar prep
created_date: 2026-04-20
---

# SE Meeting Prep Skill

When this skill is invoked, run the full automated meeting prep pipeline for the
current week — no stopping, no asking what to do next.

## Execution

Run the main script:

```bash
~/.snowflake/cortex/.mcp-servers/google-workspace/node \
  ~/.snowflake/cortex/skills/se-meeting-prep/prep_meetings.js
```

## What it does

1. Reads your Google Calendar (Mon–Sun) via macOS EventKit binaries
2. Filters to external/customer meetings + VIP AE meetings
3. Looks up attendee job titles from Snowflake SFDC via Snow CLI
4. Searches Google Drive for *master_notes* files, falls back to /memories/*.md
5. Sends an email to yourself with a formatted HTML prep agenda
6. Updates each calendar event description (original content preserved at top)

## Requirements

- macOS with Google Calendar synced (System Settings → Internet Accounts → Google)
- Terminal granted Full Calendar Access (System Settings → Privacy → Calendars)
- Cortex Code with Google Workspace MCP configured
- Snowflake CLI (`snow`) with a connection configured

See README.md and agents.md for full setup instructions.
