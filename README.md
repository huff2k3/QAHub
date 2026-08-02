# QAHub

A lightweight QA bug tracker built entirely on Google Sheets + Google Apps Script. No external server, no
database — just a spreadsheet with a custom menu, automation, dropdowns, a dashboard, and one-click status
reports.

## Features

- **Add New Issue** — inserts a new row with an auto-generated Bug ID, creator, timestamps, and description/steps
  templates.
- **Dropdowns** for Status, Priority, Severity, and Reproducibility, backed by a `Lists` sheet.
- **Search** — jump to any issue by keyword.
- **Archive Closed** — moves closed issues off the active `Issues` sheet onto `Archive`.
- **Dashboard** — status/priority/severity breakdowns, weekly throughput, and a stale-issue list, each with a
  chart.
- **Generate Report** — a shareable Google Doc summarizing critical issues, throughput, and an overall quality
  status.

See [CLAUDE.md](./CLAUDE.md) for the full architecture writeup.

## Setup (from scratch)

This creates your **own** Google Sheet and Apps Script project — nothing here is tied to any specific Google
account. You'll need [Node.js](https://nodejs.org/) and a Google account.

1. **Clone the repo and install dependencies:**
   ```
   git clone https://github.com/huff2k3/QAHub.git
   cd QAHub
   npm install
   ```

2. **Enable the Apps Script API for your account** (a one-time setting, separate from OAuth login): go to
   [script.google.com/home/usersettings](https://script.google.com/home/usersettings) and turn **"Google Apps
   Script API"** to **On**. Skipping this causes `clasp create` in the next step to fail with `Insufficient
   Permission`.

3. **Log in and create your own project:**
   ```
   npm run login
   npm run create
   ```
   `npm run login` opens a browser for Google OAuth. `npm run create` provisions a brand-new Google Sheet with a
   bound Apps Script project and writes a local `.clasp.json` pointing at it (this file is gitignored — it's
   specific to your deployment, not shared).

4. **Push the code to it:**
   ```
   npm run push
   ```

5. **Open the new Sheet** (the URL was printed by `npm run create`, or run `npm run open` to jump to the script
   editor) and **reload the page** — the custom menu only appears after a fresh load.

6. Click **QA Tools → Initialize Workbook**. The first action you run will prompt an authorization screen
   ("This app isn't verified") — that's expected for a script you just created yourself; click **Advanced → Go to
   QAHub (unsafe) → Allow**. This builds the `Issues`/`Archive`/`Lists`/`Settings`/`Dashboard` sheets, headers,
   dropdowns, and filters.

7. You're set — try **QA Tools → Add New Issue**.

### If you'd rather connect to an *existing* QAHub deployment

Someone else's `scriptId` (skip `npm run create`, use this instead):
```
npm run login
npx clasp clone <scriptId>
npm run push
```

## Development

```
npm test          # run the full Jest suite (pure Node, no Google account needed)
npm run push       # deploy local changes to your connected Apps Script project
```

See [CLAUDE.md](./CLAUDE.md) for testing details, architecture notes, and known gotchas (e.g. Apps Script
authorization getting stuck, and how a stuck project was fixed by recreating it from scratch).
