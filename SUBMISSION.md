# Spreadsheet Submission — SE Meeting Prep Skill

## Field Values (copy-paste into the form)

---

### Developer
Carson Walker

---

### Use Case Solved
**Automated weekly pre-call prep for SE customer meetings**

Every Monday (or on demand), this Cortex Code skill reads your full Google Calendar for the week, finds every external/customer meeting, looks up attendee job titles from Snowhouse SFDC, pulls context from your Google Drive master_notes files, and writes a structured pre-call agenda in two places: (1) a Gmail draft to yourself for easy review, and (2) the calendar event description so you see the context right in the invite.

Zero manual effort — one trigger phrase in Cortex Code: "prep my meetings"

---

### Benefits

- **Saves 30–60 min/week** of manual prep work per SE
- **Shows customer job titles** for all external attendees automatically (pulled from Snowhouse SFDC, ~88% hit rate)
- **Surfaces account context** from your Google Drive master_notes files without switching tabs
- **Writes into both Gmail Drafts and calendar events** — prep is available wherever you look before a call
- **Preserves original event content** (Zoom links etc.) — never overwrites anything important
- **Prioritizes AE meetings** — Ali/Shaw + customer meetings always prepped first
- **100% private** — drafts to yourself only, no attendee notifications, no data leaves Snowflake-internal systems
- **Runs in ~5 seconds** — uses Swift EventKit for calendar access (10x faster than any other macOS calendar method)

---

### Reference Link (GitHub/Slack)
https://github.com/sfc-gh-CWALKER/se-meeting-prep

---

### CLI or Desktop?
**CLI (Cortex Code)**

Triggered by typing `prep my meetings` in Cortex Code. Also runnable directly from Terminal.

---

### Guidance / Notes

**Prerequisites (one-time setup ~10 min):**
1. Clone the repo: `git clone https://github.com/sfc-gh-CWALKER/se-meeting-prep.git`
2. Run `./setup.sh` — compiles Swift binaries and installs the skill
3. Edit 4 lines in `prep_meetings.js` CONFIG: your email, your AE(s), Snowhouse connection name, warehouse
4. Grant Terminal "Full Calendar Access" in System Settings → Privacy & Security → Calendars
5. Confirm Google Calendar is synced to macOS (System Settings → Internet Accounts → Google → Calendars ✓)

**macOS only** — uses Swift EventKit to read/write calendar events natively. The Google Calendar API and CalDAV are both blocked in Snowflake's Workspace environment, so EventKit is the only working approach.

**Snowhouse required** for job title enrichment (`FIVETRAN.SALESFORCE.CONTACT`). The skill still works without Snowhouse access — titles just show as blank.

**Notes files naming convention:** Name your account notes files in Google Drive like `YourInitials_CompanyName_Master_Notes` and the skill will find them automatically.

Full troubleshooting guide in `agents.md` in the repo.
