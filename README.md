# DnD Scheduler GitHub Environment

This folder is a GitHub-ready `clasp` repository for the DnD Scheduler Google Apps Script project.

It is intentionally separate from the live working folder so you can version the script cleanly without pulling Excel files or other local artifacts into Git.

## Structure

```text
dnd-scheduler-github/
  .github/workflows/validate.yml
  .clasp.json.example
  .gitignore
  package.json
  README.md
  scripts/validate-repo.mjs
  src/
    Code.gs
    I18nData.gs
    Index.html
    appsscript.json
```

## What This Repo Assumes

- The scheduler is deployed as a standalone Google Apps Script web app.
- The spreadsheet is configured through the Apps Script script property `SPREADSHEET_ID`.
- The source files in `src/` are the files pushed to Apps Script.

## First-Time Setup

1. Create a new GitHub repository and copy this folder into it.
2. Copy `.clasp.json.example` to `.clasp.json`.
3. Replace `REPLACE_WITH_YOUR_APPS_SCRIPT_PROJECT_ID` with the real Apps Script project ID.
4. Review `src/appsscript.json` and confirm the timezone matches the production script.
5. Run `npm install`.
6. Run `npm run login`.
7. Run `npm run push`.

## Script Properties

After the first push, set this script property in Apps Script:

- `SPREADSHEET_ID`: the ID of the spreadsheet that stores Players, Campaigns, Slots, Availability, and related sheets.

If the script stays spreadsheet-bound instead of standalone, the code falls back to `SpreadsheetApp.getActive()`, but the standalone setup is the cleaner GitHub workflow.

## Recommended Workflow

- Make changes in `src/`.
- Run `npm run validate`.
- Commit and push to GitHub.
- Run `npm run push` to send the latest code to Apps Script.
- Create a new Apps Script version with `npm run version`.
- Deploy or update the web app from Apps Script.

## Notes

- `.clasp.json` is intentionally ignored because it contains the live project ID.
- The GitHub Actions workflow only performs lightweight repository validation. It does not deploy.
- The copied source files were taken from the current local DnD Scheduler project at the time this folder was created.
