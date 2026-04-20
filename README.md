# SE Meeting Prep — Cortex Code Skill

A **Cortex Code skill** for Snowflake Solutions Engineers that auto-generates structured pre-call prep for every external customer meeting in your calendar — once a week, one command, done.

---

## What it does

Runs every week (or on demand via Cortex Code) and for every external/customer meeting:

1. **Reads your Google Calendar** for the current Mon–Sun via macOS EventKit (no Google Calendar API needed)
2. **Looks up external attendees** in Snowhouse SFDC (`FIVETRAN.SALESFORCE.CONTACT`) to get their job titles
3. **Searches Google Drive** for your `*master_notes*` files for the company, then falls back to `/memories/*.md`
4. **Writes the prep agenda in two places:**
   - A **Gmail draft to yourself** — formatted HTML with attendee table, account context, and suggested agenda
   - The **calendar event description** — plain text, prepended below the original Zoom/Meet link so nothing is lost

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
  Ali Roshanzamir | ali.roshanzamir@snowflake.com
  Carson Walker | carson.walker@snowflake.com

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
| **Snowhouse connection** | Any Snowflake connection with access to `FIVETRAN.SALESFORCE.CONTACT` |
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

### Contact title lookup: Snowhouse SFDC

External attendee titles come from `FIVETRAN.SALESFORCE.CONTACT` in Snowhouse, queried via the `snow sql` CLI. One batch query covers all external emails for the week. ~88% of customer contacts resolve to a name + title.

### Notes enrichment: Google Drive → memory fallback

For each customer company (derived from email domain), the script:
1. Searches Drive for `*master_notes*{company}*` files
2. If none found, checks `/memories/*.md` for a file containing the company name
3. Extracts use case context and key notes from the first match

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

- All prep stays **private to you** — drafts go to your own inbox, calendar updates use `sendUpdates=none`
- No customer data leaves Snowflake internal systems
- OAuth token stored in macOS Keychain via CoCo's Google Workspace MCP
- Snowhouse queries run under your own `snow` CLI credentials

---

## Author

Carson Walker — Solutions Engineer, Canada Growth  
[GitHub: sfc-gh-CWALKER](https://github.com/sfc-gh-CWALKER)
