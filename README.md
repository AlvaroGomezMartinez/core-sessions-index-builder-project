# Core Sessions Index Builder

Builds a hyperlinked index of every folder and file inside the **Core Sessions** Google Drive folder. The index is written to a Google Sheet so staff can quickly browse and jump to any resource.

## How It Works

The `listSubfoldersAndFiles` function runs on a nightly trigger and:

1. Opens the **Core Sessions** Drive folder.
2. Collects all top-level folders and files, then collects each folder's children.
3. Clears the existing index rows in the **Index** sheet.
4. Writes the full index in a single batch — column A for top-level items, column B for their children — using `HYPERLINK` formulas.

Drive shortcuts are automatically resolved to their target URLs.

## Project Links

| Resource | Link |
|---|---|
| Apps Script project | https://script.google.com/home/projects/1SXMhCbQ1Uun2f1FMRTkVmkhQ4uGCbfdX8nFFsU5egQ4FJad1xxbpo2nN/edit |
| Index spreadsheet | https://docs.google.com/spreadsheets/d/18jneeP5YgIKhfmiYj8vNHZYbhd_WTn8AN-hEXaO6K9k/edit |
| Core Sessions folder | https://drive.google.com/drive/folders/0B6Z4Z9quk8OJUmlwSXpMTER3d1k |

## Prerequisites

- A Google account with access to the Core Sessions folder and the Index spreadsheet.
- [Node.js](https://nodejs.org/) (v14 or later) installed locally.
- [clasp](https://github.com/google/clasp) — the Google Apps Script CLI.

## Local Setup

1. **Install clasp globally** (if you haven't already):

   ```bash
   npm install -g @google/clasp
   ```

2. **Log in to clasp:**

   ```bash
   clasp login
   ```

   This opens a browser window for Google OAuth. Use the account that has access to the project.

3. **Clone this repository** to your machine.

4. **Update `rootDir` in `.clasp.json`** to match your local path:

   ```json
   {
     "scriptId": "1SXMhCbQ1Uun2f1FMRTkVmkhQ4uGCbfdX8nFFsU5egQ4FJad1xxbpo2nN",
     "rootDir": "/your/local/path/to/core-sessions-index-builder-project"
   }
   ```

5. **Pull the latest code** from the Apps Script project (optional, to confirm everything is in sync):

   ```bash
   clasp pull
   ```

6. **Push local changes** to the Apps Script project:

   ```bash
   clasp push
   ```

## Setting Up the Trigger

The script is designed to run on a nightly time-driven trigger. To create one:

1. Open the [Apps Script project](https://script.google.com/home/projects/1SXMhCbQ1Uun2f1FMRTkVmkhQ4uGCbfdX8nFFsU5egQ4FJad1xxbpo2nN/edit).
2. Click the **Triggers** icon (clock) in the left sidebar.
3. Click **+ Add Trigger** and configure:
   - **Function:** `listSubfoldersAndFiles`
   - **Deployment:** Head
   - **Event source:** Time-driven
   - **Type:** Day timer
   - **Time of day:** Midnight to 1am
4. Click **Save**.

## Email Notifications

To receive an email when a run fails, set the `NOTIFY_EMAIL` constant in `Code.js` to your email address:

```javascript
const NOTIFY_EMAIL = "your.email@example.com";
```

Leave it as an empty string to disable notifications.

## Configuration

All configurable values are constants at the top of `Code.js`:

| Constant | Purpose |
|---|---|
| `CORE_SESSIONS_FOLDER_ID` | Google Drive ID of the Core Sessions folder |
| `INDEX_SPREADSHEET_ID` | Google Sheets ID of the index spreadsheet |
| `INDEX_SHEET_NAME` | Name of the sheet tab to write to |
| `START_ROW` | First row to write data (rows above are preserved as headers) |
| `NOTIFY_EMAIL` | Email address for failure alerts (empty = disabled) |
| `PROJECT_URL` | Link to the Apps Script project (included in alert emails) |

## Logging

The script logs to Google Cloud Logging (Stackdriver). To view logs:

1. Open the [Apps Script project](https://script.google.com/home/projects/1SXMhCbQ1Uun2f1FMRTkVmkhQ4uGCbfdX8nFFsU5egQ4FJad1xxbpo2nN/edit).
2. Click **Executions** in the left sidebar to see recent runs and their log output.

Log levels used:

- `console.info` — Milestones: run started, items collected, run complete.
- `console.warn` — Non-fatal issues like unresolvable shortcuts.
- `console.error` — Skipped items or full run failures.
