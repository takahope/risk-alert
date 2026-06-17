# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Google Apps Script (GAS)** web application that automates security vulnerability alert processing via Gmail. It scans incoming emails, matches against an asset list in Google Sheets, creates draft email responses, and serves an interactive web dashboard.

## Architecture

### Core Files
| File | Purpose |
|------|---------|
| `code.js` | All backend logic: email processing, asset matching, logging, `doGet()`, and all `google.script.run` API functions |
| `Dashboard.html` | Single-page dashboard UI (HTML/CSS/JS embedded together); communicates with backend via `google.script.run` |
| `env.js` | `CONFIG` object with environment-specific values (spreadsheet IDs, email addresses, webhook URLs). **Never commit real values.** |

### Data Flow
1. `processIncomingEmails()` (time-driven trigger) → scans Gmail for matching sender + subject keywords
2. Extracts Chinese-format warning info using regex (`警訊名稱[：:]`, `漏洞說明[：:]`)
3. Matches against Google Sheet asset list → creates Gmail drafts accordingly
4. Logs results to `SystemLogs` sheet with deduplication via Gmail Message ID

### Google Sheets Structure
- **Sheet1**: Asset list (columns A, B, C) used for vulnerability keyword matching
- **Settings**: Toggle switches at row 2; columns A–I and N–P for boolean flags ("是"/"否"), columns J–M for email lists (one email per row starting at row 2)
  - A2: scan read emails, B2: legacy auto-draft (deprecated, replaced by N/O/P), C2: chat notify, G2: notInUse email, H2: processed email, I2: asset hit notify
  - J: assetHitCc, K: notInUseCc, L: processedCc, M: assetHitRecipients
  - N2: assetHitAutoDraft, O2: notInUseAutoDraft, P2: processedAutoDraft (per-function auto-draft toggles)
- **SystemLogs**: Execution history; column G stores Gmail Message ID for deduplication; `sheetRow` values passed from frontend are actual 1-based spreadsheet row numbers (including header row)

## Key Patterns

### Frontend-Backend Communication
Dashboard calls backend exclusively via `google.script.run`. All backend functions called this way **must be top-level** (not inside objects or modules):
```javascript
google.script.run.withSuccessHandler(callback).functionName(args);
```

### Settings Storage
Boolean settings are stored as `"是"`/`"否"` strings. Email lists are stored one per cell in a column starting at row 2.

### `sheetRow` Convention
`updateUsageStatus(sheetRow, ...)` and `updateUsageStatuses(sheetRows, ...)` expect the actual spreadsheet row number (1-based, including header). `normalizeSheetRow()` validates and coerces this value.

## Development Workflow

There is no local build or test pipeline. All development is done manually via the Google Apps Script editor:

1. Copy updated file contents into the GAS editor (`code.js`, `Dashboard.html`, `env.js`)
2. Deploy as Web App: **Deploy → New deployment → Web app → Execute as: Me**
3. Set up a time-driven trigger for `processIncomingEmails()` (e.g., every 5 minutes)
4. Manually verify: trigger the function → check Google Sheets writes → reload dashboard

### Useful Search Commands
```bash
# Trace data flow between frontend and backend
rg "google.script.run|SystemLogs|fetchLogs" .

# Find all sheet column references
rg "getRange|setValue|column" code.js
```

## Coding Conventions
- 2-space indentation in both JS and HTML
- `camelCase` for functions/variables; `UPPER_SNAKE_CASE` for `CONFIG` keys
- Comment fixed column indexes when code depends on them (e.g., `// J 欄: 命中資產通知 CC`)
- When adding new sheet columns, update both `code.js` range reads and any Dashboard column references together

## Common Modifications
- **Add new asset columns**: Update `fetchComparisonData()` range (currently reads 3 columns from Sheet1)
- **Add new Settings toggle**: Add cell reference in `updateSystemSetting()` + matching key in `getSystemSettings()` + UI switch in Dashboard.html
- **Change email templates**: Modify `createDraftForPersonA()` or `createDraftReplyToSenderB()`
- **Add new SystemLogs columns**: Update the log write in `processIncomingEmails()`, adjust column offsets in Dashboard.html table rendering, and update `updateUsageStatus()` if the new column affects row reads
