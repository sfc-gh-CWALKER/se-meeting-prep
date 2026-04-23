# SE Meeting Prep — Cortex Code Skill

A **Cortex Code skill** for Snowflake Solutions Engineers that auto-generates structured pre-call prep for every external customer meeting in your calendar — once a week, one command, done.

---

## What it does

Runs every week (or on demand via Cortex Code) and for every external/customer meeting:

1. **Reads your Google Calendar** for the current Mon–Sun via macOS EventKit (no Google Calendar API needed)
2. **Looks up external attendee job titles** from Snowflake SFDC (`FIVETRAN.SALESFORCE.CONTACT`) via Snow CLI
3. **Searches Google Drive** for your `*master_notes*` files for the company, then falls back to `/memories/*.md`
4. **Writes the prep agenda in two places:**
   - An **email to yourself** (sent directly via SMTP) — formatted HTML with attendee table, account context, and suggested agenda
   - The **calendar event description** — plain text, written directly into the event so you see the prep when you open the meeting in any calendar app
5. **Tracks changes between runs** — if a meeting moves, an attendee is added, or the title changes, the next run detects it and re-preps automatically

Nothing is ever sent to customers or attendees.

---

## Sample Output

**Calendar event description** (after running):
```
[Your existing Zoom link and original description stays here]

=== PRE-CALL PREP (auto-generated) ===
Meeting: Acme Corp + Snowflake - Discovery Call
When: Mon, Apr 21, 10:30 a.m. EDT (30 min)

EXTERNAL ATTENDEES:
  Jane Smith | Head of Data Engineering | jane.smith@acme.com | acme
  Bob Lee | Senior Data Architect | bob.lee@acme.com | acme

SNOWFLAKE TEAM:
  Your AE | ae.name@snowflake.com
  You | your.name@snowflake.com

ACCOUNT CONTEXT:
  Use cases: Cortex Search, data sharing with partners...

SUGGESTED AGENDA:
  1. Intro / context recap (5 min)
  2. Review use cases and current state (15 min)
  3. Technical deep dive / demo (10 min)
  4. Next steps + action items (5 min)
=== END PREP ===
```

---

## Prerequisites

| Requirement | Notes |
|---|---|
| **macOS** | Uses Swift EventKit binaries to read/write calendar |
| **Cortex Code (CoCo)** | Needs CoCo installed with Google Workspace MCP configured |
| **Google Calendar synced to macOS** | System Settings → Internet Accounts → Google → Calendars ✓ |
| **Calendar Full Access granted to Terminal** | System Settings → Privacy & Security → Calendars → Terminal → Full Access |
| **Snowflake CLI (`snow`)** | `/Applications/SnowflakeCLI.app` — comes with CoCo |
| **Snowflake connection** | Any Snowflake connection with access to `FIVETRAN.SALESFORCE.CONTACT` (optional — falls back gracefully) |
| **Google Drive master_notes files** | Optional — files named like `CW_Acme_Master_Notes` |

---

## Setup

### 1. Clone this repo

```bash
git clone https://github.com/sfc-gh-CWALKER/se-meeting-prep.git
cd se-meeting-prep
```

### 2. Edit the CONFIG block in `prep_meetings.js`

Open `prep_meetings.js` and update the top section:

```javascript
const ME = 'your.name@snowflake.com';           // Your Snowflake email
const VIP_ATTENDEES = [                          // AEs you work with — their meetings get top priority
  'ae.one@snowflake.com',
  'ae.two@snowflake.com'
];
const VIP_DRIVE_FOLDER_IDS = [];                 // Optional: Google Drive folder IDs for AE account folders
const SNOW_CONNECTION = 'YOUR_CONNECTION_NAME';  // Your snow CLI connection name (snow connection list)
const SNOW_WAREHOUSE  = 'YOUR_WAREHOUSE';        // A warehouse you can use
```

Also update the keytar path if your CoCo installation is non-standard (see `agents.md`).

### 3. Compile the Swift EventKit binaries

```bash
swiftc src/read_cal_week.swift  -o read_cal_week
swiftc src/update_cal_event.swift -o update_cal_event
```

You only need to do this once. The compiled binaries are machine-specific and not included in the repo.

### 4. Grant Full Calendar Access to Terminal

> **Why?** macOS requires explicit permission for any app to read your calendar. `writeOnly` is not enough — you need `fullAccess`.

1. Open **System Settings → Privacy & Security → Calendars**
2. Find **Terminal** in the list
3. Set it to **Full Access** (not just "Write Only")

Verify it worked:
```bash
swift - <<'EOF'
import EventKit; import Foundation
print("Status:", EKEventStore.authorizationStatus(for: .event).rawValue)
// 3 = fullAccess ✓    4 = writeOnly ✗    2 = denied ✗
EOF
```

### 5. Install the skill into Cortex Code

```bash
cp -r . ~/.snowflake/cortex/skills/se-meeting-prep/
```

Or run setup.sh which handles this automatically:

```bash
chmod +x setup.sh
./setup.sh
```

### 6. Test it

```bash
~/.snowflake/cortex/.mcp-servers/google-workspace/node \
  ~/.snowflake/cortex/skills/se-meeting-prep/prep_meetings.js
```

Or trigger via Cortex Code: `prep my meetings`

---

## Running via Cortex Code

Once the skill is installed in `~/.snowflake/cortex/skills/se-meeting-prep/`, just type in Cortex Code:

```
prep my meetings
```

or

```
meeting prep for this week
```

---

## How it works under the hood

### Calendar access: Swift EventKit (not Google Calendar API)

The Google Calendar API is disabled in Snowflake's GCP OAuth project. CalDAV is blocked by the Workspace admin. JXA/AppleScript times out on calendars with 10,000+ recurring events.

**Solution**: Two compiled Swift binaries use macOS EventKit directly:
- `read_cal_week` — reads Mon–Sun events from the local EventKit store, outputs JSON (~0.1 sec)
- `update_cal_event <eventId>` — writes a new event description via stdin, syncs silently to Google Calendar within ~30 sec

### Contact title lookup: Snowflake SFDC

External attendee titles come from `FIVETRAN.SALESFORCE.CONTACT`, queried via the `snow sql` CLI. One batch query covers all external emails for the week. ~88% of customer contacts resolve to a name + title.

### Notes enrichment: Google Drive → memory → title-only fallback

**You don't need notes for the tool to be useful.** The script searches for context in order and stops at the first hit:

| Step | What it searches | Requires |
|---|---|---|
| 1 | Account subfolders inside `VIP_DRIVE_FOLDER_IDS` for a matching folder + `master_notes` file | `VIP_DRIVE_FOLDER_IDS` configured |
| 2 | Broad Drive search: files with `master_notes` + company name in the filename | Any Drive file named like `YourName_Acme_Master_Notes` |
| 3 | `/memories/*.md` files containing the company name | CoCo memory files from past sessions |
| 4 | **No notes found** — prep still runs | Nothing required |

**What you get at each level:**

- **With notes** (steps 1–3): Cortex AI synthesizes recent SE notes, active use cases, and open action items. The agenda references specific use cases by name and injects an action item review step when open items exist.

- **With SFDC titles only** (step 4, Snowflake connection configured): You still get a fully populated attendee table with names and job titles, an audience-aware agenda (executive vs. technical path), and a discovery-oriented flow for new accounts.

- **With nothing** (no notes, no SFDC, no Snowflake connection): You still get the attendee list from your calendar, company name inferred from the meeting title or email domain, and a relevant suggested agenda based on the meeting type (QBR, Demo, Intro, POC, etc.) detected from the title.

**Name your notes files like:** `YourInitials_CompanyName_Master_Notes` — Google Docs work best (exported as plain text automatically).

### Calendar description updates and change detection

Each meeting prep is written directly into the calendar event description so it's visible whenever you open the event — in Google Calendar, macOS Calendar, or your phone. The original invite text (Zoom link, organizer notes) is preserved above the prep block and never modified.

The script tracks a checksum of each meeting's key fields between runs:

| Change detected | What happens |
|---|---|
| New meeting (never prepped) | Full prep generated, calendar updated |
| Meeting time moved | Re-prep triggered automatically |
| Attendee added or removed | Re-prep triggered automatically |
| Title renamed | Re-prep triggered automatically |
| No changes | Skipped — calendar untouched, still included in summary email |

This means you can run `prep my meetings` multiple times during the week and only changed meetings get re-processed. The consolidated summary email always reflects the current state of all upcoming meetings.

---

## File structure

```
se-meeting-prep/
├── README.md                  ← This file
├── prep_meetings.js           ← Main script (edit CONFIG block before use)
├── src/
│   ├── read_cal_week.swift    ← Swift source: reads EventKit → JSON
│   └── update_cal_event.swift ← Swift source: writes event description
├── SKILL.md                   ← Cortex Code skill definition
├── agents.md                  ← Full technical documentation + troubleshooting
└── setup.sh                   ← One-command installer
```

After compiling, the binaries `read_cal_week` and `update_cal_event` live in the root of the skill directory alongside `prep_meetings.js`.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `read_cal_week` not found | Recompile: `swiftc src/read_cal_week.swift -o read_cal_week` |
| 0 events returned | Check TCC status — Terminal needs Full Access (value `3`), not Write Only (`4`) |
| `Cannot find module 'keytar'` | Run with CoCo's Node: `~/.snowflake/cortex/.mcp-servers/google-workspace/node prep_meetings.js` |
| `No active warehouse` | Add `--warehouse YOUR_WAREHOUSE` to the snow CLI call in CONFIG |
| `401 Unauthorized` on Gmail/Drive | Re-run Google Workspace MCP auth in CoCo |
| No notes found | Create Drive docs named like `YourName_CompanyName_Master_Notes` |
| `update_cal_event` exits 1 | You're not the event organizer — Gmail draft is still created, calendar skipped |

---

## Privacy & security

- All prep stays **private to you** — emails go to your own inbox only, calendar updates use `sendUpdates=none`
- No customer data leaves Snowflake internal systems
- OAuth token stored in macOS Keychain via CoCo's Google Workspace MCP
- Snowflake queries run under your own `snow` CLI credentials

---

## Author

Carson Walker — Sr. Solutions Engineer, Snowflake Canada
[GitHub: sfc-gh-CWALKER](https://github.com/sfc-gh-CWALKER)
