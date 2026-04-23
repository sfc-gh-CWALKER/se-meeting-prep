---
# SE Meeting Prep — Technical Notes & Troubleshooting
---

## Quick Reference

| What | Where |
|---|---|
| Run command | `~/.snowflake/cortex/.mcp-servers/google-workspace/node ~/.snowflake/cortex/skills/se-meeting-prep/prep_meetings.js` |
| Edit config | `prep_meetings.js` top section — ME, VIP_ATTENDEES, SNOW_CONNECTION, SNOW_WAREHOUSE |
| Skill dir | `~/.snowflake/cortex/skills/se-meeting-prep/` |
| Binaries | `read_cal_week`, `update_cal_event` (compile from `src/*.swift`) |

---

## Calendar Access: Swift EventKit Binaries

### Why not Google Calendar API?

The CoCo MCP uses a Snowflake-managed GCP project where the Google Calendar API is disabled. `serviceusage.services.enable` belongs to the MCP team, not individual users.

### Why not JXA / AppleScript?

macOS Calendar.app accumulates thousands of recurring event instances in the local store. JXA and AppleScript time out after 3+ minutes iterating `cal.events()`. The AppleScript `whose` clause returns error `-1728` on large collections.

### Why not CalDAV?

Snowflake Workspace admin blocks third-party calendar clients via CalDAV. The private ICS URL is also hidden by admin policy.

### The Working Approach: Swift EventKit

1. Add your Snowflake Google account to **System Settings → Internet Accounts → Google** with Calendars enabled
2. macOS syncs your Google Calendar to the local EventKit store
3. `read_cal_week` reads Mon–Sun events via `EKEventStore.events(matching:)` in ~0.1 seconds
4. `update_cal_event <eventId>` writes the event description back; macOS syncs to Google within ~30 sec

### One-time TCC permission

EventKit requires **Full Access** (not Write-Only) from Terminal:

> **System Settings → Privacy & Security → Calendars → Terminal → Full Access**

Status values returned by `EKEventStore.authorizationStatus(for: .event).rawValue`:
- `3` = fullAccess ← **required**
- `4` = writeOnly ← NOT enough (reads return 0 events)
- `2` = denied
- `0` = notDetermined

Verify your status:
```bash
swift - <<'EOF'
import EventKit; import Foundation
print("Auth status:", EKEventStore.authorizationStatus(for: .event).rawValue)
EOF
```

### Recompiling binaries

If binaries are missing (e.g. after cloning to a new machine):
```bash
swiftc src/read_cal_week.swift  -o read_cal_week
swiftc src/update_cal_event.swift -o update_cal_event
cp read_cal_week update_cal_event ~/.snowflake/cortex/skills/se-meeting-prep/
```

---

## Contact Title Lookup: Snowflake SFDC

External attendee job titles come from Snowflake:

```sql
SELECT LOWER(EMAIL) AS EMAIL, NAME, NULLIF(TITLE,'') AS TITLE
FROM FIVETRAN.SALESFORCE.CONTACT
WHERE LOWER(EMAIL) IN (...)
QUALIFY ROW_NUMBER() OVER (PARTITION BY LOWER(EMAIL) ORDER BY NULLIF(TITLE,'') DESC NULLS LAST) = 1
```

Called once per run via `snow sql --connection YOUR_CONNECTION --warehouse YOUR_WAREHOUSE --format json`. Falls back silently (titles blank) if unavailable.

Coverage: ~88% of customer contacts in SFDC have a resolvable title.

---

## Priority Sorting

| Type | Criteria | Prepared? |
|---|---|---|
| `vip-external` | VIP AE + external customer both attending | ✓ First |
| `vip-internal` | VIP AE, no external customer | ✓ Second |
| `external` | External customer, no VIP AE | ✓ Third |
| `internal` | Snowflake-only | ✗ Skipped |
| `skip` | Cancelled or declined | ✗ Skipped |

VIP AEs are configured in `VIP_ATTENDEES` in the CONFIG block.

---

## Notes Enrichment: Search Strategy & Fallback Levels

The script searches for account context in order and stops at the first hit. **The tool runs and produces useful prep at every level — notes are optional.**

### Search order

1. **VIP AE Drive folders** (`VIP_DRIVE_FOLDER_IDS`): looks for a subfolder matching the company name, then finds a `master_notes` file inside it. Requires `VIP_DRIVE_FOLDER_IDS` to be configured with folder IDs.

2. **Broad Drive search**: searches your entire Drive for files with `master_notes` + company name in the filename (e.g. `CW_Acme_Master_Notes`). Works with no extra config as long as files are named correctly.

3. **CoCo memory files**: checks `/memories/*.md` for any file containing the company name. Useful for accounts you've researched or captured notes for via CoCo chat.

4. **Title + attendees only**: if nothing is found, prep still runs — uses meeting title, attendee names, and SFDC job titles to build a relevant agenda.

### What you get at each level

| Context available | What the prep includes |
|---|---|
| Notes + SFDC titles | AI-synthesized recent SE notes, named use cases, open action items, attendee titles, audience-aware agenda with use case references |
| SFDC titles only | Full attendee table with names and job titles, executive vs. technical agenda path, discovery-oriented flow for new accounts |
| Title + attendees only | Attendee list from calendar, company name from title/domain, meeting-type agenda (QBR → business review; Demo → demo flow; Discovery → current state questions; etc.) |

### Naming your notes files

```
YourInitials_CompanyName_Master_Notes   ← Google Doc (exported as plain text)
YourInitials_CompanyName_Master_Notes   ← .docx, .txt also supported
```

Google Docs export cleanly as plain text. `.docx` files download via binary media. Spreadsheets export as CSV.

Falls back to `/memories/*.md` if no Drive file is found.

---

## Email Format

- **To / From:** Your own email (set `ME` in CONFIG)
- **Subject:** `Pre-call Prep - {Meeting Title} - {Day Mon D}` (ASCII only)
- **Body:** HTML with Snowflake blue (#29B5E8) styling, attendee table (Name | Title | Email | Company), account context, suggested agenda
- **Delivery:** Sent directly via Gmail SMTP XOAUTH2 (port 465) — arrives in your inbox, not Drafts

---

## Calendar Event Description

- **Method:** EventKit `store.save(ev, span: .thisEvent)` — only updates this occurrence of recurring events
- **Ordering:** Original content (Zoom link, etc.) stays at TOP; prep block appended below
- **Re-run safe:** Old prep block is stripped and replaced, original content untouched
- **Not organizer:** If you didn't create the meeting, `update_cal_event` exits non-zero → Gmail draft still created, calendar skipped

---

## Known Limitations

1. **Not the organizer**: You can't update descriptions for meetings you didn't create. Gmail draft is still generated.
2. **Company from domain**: Works best with company-specific email domains. Generic domains (gmail.com, hotmail.com) give poor company names.
3. **Drive search is name-based**: Only finds files with `master_notes` + company name in the filename.
4. **SFDC coverage**: ~88% of external contacts have titles in Snowhouse SFDC. Unknown contacts get titles inferred from their email prefix.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `read_cal_week` binary not found | `swiftc src/read_cal_week.swift -o read_cal_week` then copy to skill dir |
| 0 events returned | TCC: check auth status is `3` (fullAccess). Grant in System Settings → Privacy → Calendars |
| `Cannot find module 'keytar'` | Run with CoCo's Node: `~/.snowflake/cortex/.mcp-servers/google-workspace/node prep_meetings.js` |
| `No active warehouse` | Set SNOW_WAREHOUSE in CONFIG |
| `401 Unauthorized` on Gmail/Drive | Re-run Google Workspace MCP auth in CoCo |
| No Drive notes found | Name files like `YourName_CompanyName_Master_Notes` |
| `update_cal_event` exits 1 | You're not the event organizer — Gmail draft created, calendar skipped |
| No events showing up | Confirm Google account is in System Settings → Internet Accounts → Google with Calendars enabled |

---

## OAuth Scopes Required (Google Workspace MCP)

These scopes are already present if you've set up the Google Workspace MCP in CoCo:
- `https://mail.google.com/` — send email via SMTP XOAUTH2
- `https://www.googleapis.com/auth/drive` — search Drive for notes

Calendar is handled by EventKit, not OAuth — no calendar scope needed.

The script reads the OAuth token from macOS Keychain using `KEYCHAIN_SERVICE` and
`KEYCHAIN_ACCOUNT` constants defined at the top of `prep_meetings.js`. These are
Keychain lookup identifiers set by CoCo when you authorize the Google Workspace MCP.
To find your values: `security find-generic-password -a oauth_tokens 2>/dev/null | grep svce`

---

## File Structure

```
~/.snowflake/cortex/skills/se-meeting-prep/
├── SKILL.md              ← Cortex Code skill definition
├── agents.md             ← This file
├── prep_meetings.js      ← Main script
├── read_cal_week         ← Compiled Swift binary
└── update_cal_event      ← Compiled Swift binary
```
