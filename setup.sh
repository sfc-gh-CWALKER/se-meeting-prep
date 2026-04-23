#!/bin/bash
set -e

SKILL_DIR="$HOME/.snowflake/cortex/skills/se-meeting-prep"
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "========================================"
echo "SE Meeting Prep — Setup"
echo "========================================"
echo ""

# ── 1. Check for CoCo Node ──────────────────────────────────────────────────
NODE="$HOME/.snowflake/cortex/.mcp-servers/google-workspace/node"
if [ ! -f "$NODE" ]; then
  echo "ERROR: Cortex Code Google Workspace MCP not found at:"
  echo "  $NODE"
  echo ""
  echo "Make sure you have:"
  echo "  1. Cortex Code (CoCo) installed"
  echo "  2. Google Workspace MCP configured in CoCo"
  exit 1
fi
echo "✓ CoCo Node found: $NODE"

# ── 2. Check for Snowflake CLI ───────────────────────────────────────────────
SNOW_CLI="/Applications/SnowflakeCLI.app/Contents/MacOS/snow"
if [ ! -f "$SNOW_CLI" ]; then
  echo "WARNING: Snowflake CLI not found at $SNOW_CLI"
  echo "  Contact title lookup from SFDC will be skipped."
  echo "  Install Snowflake CLI from: https://docs.snowflake.com/en/user-guide/snowflake-cli"
else
  echo "✓ Snowflake CLI found"
fi

# ── 3. Compile Swift binaries ────────────────────────────────────────────────
echo ""
echo "Compiling Swift EventKit binaries..."
swiftc "$REPO_DIR/src/read_cal_week.swift"  -o "$REPO_DIR/read_cal_week"
swiftc "$REPO_DIR/src/update_cal_event.swift" -o "$REPO_DIR/update_cal_event"
echo "✓ read_cal_week compiled"
echo "✓ update_cal_event compiled"

# ── 4. Create skill directory and copy files ─────────────────────────────────
echo ""
echo "Installing skill to $SKILL_DIR ..."
mkdir -p "$SKILL_DIR"
cp "$REPO_DIR/prep_meetings.js"    "$SKILL_DIR/"
cp "$REPO_DIR/read_cal_week"       "$SKILL_DIR/"
cp "$REPO_DIR/update_cal_event"    "$SKILL_DIR/"
cp "$REPO_DIR/SKILL.md"            "$SKILL_DIR/"
cp "$REPO_DIR/agents.md"           "$SKILL_DIR/"
echo "✓ Files installed"

# ── 5. Remind user about CONFIG ──────────────────────────────────────────────
echo ""
echo "========================================"
echo "NEXT STEPS:"
echo "========================================"
echo ""
echo "1. Edit CONFIG in $SKILL_DIR/prep_meetings.js:"
echo "     ME              = 'your.name@snowflake.com'"
echo "     VIP_ATTENDEES   = ['ae.name@snowflake.com']"
echo "     SNOW_CONNECTION = 'your-snowhouse-connection'"
echo "     SNOW_WAREHOUSE  = 'your-warehouse'"
echo ""
echo "2. Grant Full Calendar Access to Terminal:"
echo "     System Settings → Privacy & Security → Calendars → Terminal → Full Access"
echo "     (Must be 'Full Access', NOT 'Write Only')"
echo ""
echo "3. Ensure Google Calendar is synced to macOS:"
echo "     System Settings → Internet Accounts → Google → enable Calendars"
echo ""
echo "4. Test it:"
echo "     $NODE $SKILL_DIR/prep_meetings.js"
echo ""
echo "5. Or trigger via Cortex Code:  'prep my meetings'"
echo ""
echo "Done! See README.md and agents.md for full documentation."
