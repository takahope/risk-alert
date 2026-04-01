# Repository Guidelines

## Project Structure & Module Organization
This repository is a small Google Apps Script web app.

- `code.js`: backend logic for Gmail processing, Google Sheets reads/writes, settings APIs, and `doGet()`.
- `Dashboard.html`: frontend dashboard UI, filtering, and `google.script.run` calls.
- `env.js`: local configuration constants such as spreadsheet IDs, email recipients, and webhook URLs. Treat as sensitive.
- `.github/copilot-instructions.md`: concise architecture notes and workflow hints.

There is no `src/` or `tests/` directory. Keep related frontend and backend changes aligned when adding fields, sheet columns, or new dashboard actions.

## Build, Test, and Development Commands
There is no local build pipeline in this repo. Development is mostly manual:

- `git status`: inspect current changes before editing.
- `git diff -- code.js Dashboard.html`: review backend/frontend changes together.
- `rg "functionName|SystemLogs|google.script.run" .`: trace data flow quickly.

Run and deploy through the Google Apps Script editor as a Web App. After changes, manually verify the dashboard, sheet writes, and Gmail-related flows in the Apps Script environment.

## Coding Style & Naming Conventions
Use 2-space indentation in JavaScript and HTML to match the existing files. Prefer:

- `camelCase` for functions and variables, for example `updateUsageStatus`.
- UPPER_SNAKE_CASE for config keys inside `CONFIG`.
- Descriptive sheet-column comments when code depends on fixed column indexes.

Do not wrap GAS-exposed functions in modules; functions called by `google.script.run` must remain top-level.

## Testing Guidelines
This project currently has no automated test suite. Validate changes manually:

1. Trigger the affected Apps Script function.
2. Confirm the expected write in Google Sheets, especially `Settings` and `SystemLogs`.
3. Reload the dashboard and verify the UI reflects persisted sheet data.

When fixing bugs, document the manual reproduction path in the PR or commit notes.

## Commit & Pull Request Guidelines
Recent history uses short, imperative English commit messages such as `Add keyword search functionality...` or `Update .gitignore...`. Follow that pattern:

- Start with a verb: `Add`, `Fix`, `Update`, `Refactor`.
- Keep the subject focused on one change.

PRs should include a brief summary, affected files, sheet/config impact, and screenshots for dashboard UI changes. Mention any required `env.js` or spreadsheet setup explicitly.

## Security & Configuration Tips
Never commit real secrets or production IDs casually. `env.js` is ignored; keep local values there and sanitize examples in shared documentation. Double-check any change that affects Gmail access, spreadsheet writes, or webhook destinations.
