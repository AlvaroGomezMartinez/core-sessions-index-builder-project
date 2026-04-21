/*********************************************************
 *        CORE SESSIONS INDEX BUILDER PROJECT            *
 *                                                       *
 * Purpose:    Builds a hyperlinked index of all         *
 *             folders and files inside the "Core        *
 *             Sessions" Google Drive folder. Results    *
 *             are written to the "_Index" sheet         *
 *             (column A = top-level items,              *
 *             column B = their children).               *
 *                                                       *
 * Trigger:    Time-driven · Day timer · Midnight to 1am *
 *             Function: listSubfoldersAndFiles          *
 *             Deployment: Head                          *
 *                                                       *
 * Author:     Alvaro Gomez                              *
 *             +1-210-363-1577                           *
 ********************************************************/

// ── Configuration ────────────────────────────────────
const CORE_SESSIONS_FOLDER_ID = "0B6Z4Z9quk8OJUmlwSXpMTER3d1k";
const INDEX_SPREADSHEET_ID = "18jneeP5YgIKhfmiYj8vNHZYbhd_WTn8AN-hEXaO6K9k";
const INDEX_SHEET_NAME = "Index";
const START_ROW = 3;

// Set to an email address to receive failure alerts, or leave empty to disable.
const NOTIFY_EMAIL = "";
const PROJECT_URL =
  "https://script.google.com/home/projects/1SXMhCbQ1Uun2f1FMRTkVmkhQ4uGCbfdX8nFFsU5egQ4FJad1xxbpo2nN/edit";

// ── Helpers ──────────────────────────────────────────

/**
 * Builds a HYPERLINK formula string, escaping double quotes in the name.
 *
 * @param {string} url  - The URL to link to.
 * @param {string} name - The display text for the link.
 * @return {string} A Sheets HYPERLINK formula.
 */
function buildHyperlink(url, name) {
  const safeName = name.replace(/"/g, '""');
  return `=HYPERLINK("${url}", "${safeName}")`;
}

/**
 * Gets the actual URL of a file, resolving shortcuts when necessary.
 *
 * @param {GoogleAppsScript.Drive.File} file - The file object.
 * @return {string} The actual URL of the file or its shortcut target.
 */
function getActualUrl(file) {
  if (file.getMimeType() === MimeType.SHORTCUT) {
    try {
      const targetId = file.getTargetId();
      return DriveApp.getFileById(targetId).getUrl();
    } catch (e) {
      console.warn(`Could not resolve shortcut "${file.getName()}": ${e.message}`);
      return file.getUrl();
    }
  }
  return file.getUrl();
}

/**
 * Collects all immediate child folders and files of a given folder,
 * returned as a sorted array of { name, url } objects.
 * Individual items that fail to read are logged and skipped.
 *
 * @param {GoogleAppsScript.Drive.Folder} folder - The parent folder.
 * @return {Array<{name: string, url: string}>} Sorted child items.
 */
function getChildItems(folder) {
  const children = [];

  try {
    const folderIterator = folder.getFolders();
    while (folderIterator.hasNext()) {
      try {
        const f = folderIterator.next();
        children.push({ name: f.getName(), url: f.getUrl() });
      } catch (e) {
        console.error(`Skipping a subfolder in "${folder.getName()}": ${e.message}`);
      }
    }
  } catch (e) {
    console.error(`Failed to list subfolders of "${folder.getName()}": ${e.message}`);
  }

  try {
    const fileIterator = folder.getFiles();
    while (fileIterator.hasNext()) {
      try {
        const f = fileIterator.next();
        children.push({ name: f.getName(), url: getActualUrl(f) });
      } catch (e) {
        console.error(`Skipping a file in "${folder.getName()}": ${e.message}`);
      }
    }
  } catch (e) {
    console.error(`Failed to list files of "${folder.getName()}": ${e.message}`);
  }

  children.sort((a, b) =>
    a.name.toLowerCase().localeCompare(b.name.toLowerCase())
  );

  return children;
}

/**
 * Collects all root-level items (folders and files) inside the given folder.
 * Folders retain a reference so their children can be listed later.
 * Individual items that fail to read are logged and skipped.
 *
 * @param {GoogleAppsScript.Drive.Folder} folder - The root folder to scan.
 * @return {Array<{name: string, url: string, type: string, folder?: GoogleAppsScript.Drive.Folder}>}
 */
function getRootItems(folder) {
  const items = [];

  try {
    const subfolderIterator = folder.getFolders();
    while (subfolderIterator.hasNext()) {
      try {
        const subfolder = subfolderIterator.next();
        items.push({
          name: subfolder.getName(),
          url: subfolder.getUrl(),
          type: "folder",
          folder: subfolder,
        });
      } catch (e) {
        console.error(`Skipping a root subfolder: ${e.message}`);
      }
    }
  } catch (e) {
    console.error(`Failed to list root subfolders: ${e.message}`);
  }

  try {
    const fileIterator = folder.getFiles();
    while (fileIterator.hasNext()) {
      try {
        const file = fileIterator.next();
        items.push({
          name: file.getName(),
          url: getActualUrl(file),
          type: "file",
        });
      } catch (e) {
        console.error(`Skipping a root file: ${e.message}`);
      }
    }
  } catch (e) {
    console.error(`Failed to list root files: ${e.message}`);
  }

  items.sort((a, b) =>
    a.name.toLowerCase().localeCompare(b.name.toLowerCase())
  );

  return items;
}

/**
 * Sends a failure notification email if NOTIFY_EMAIL is configured.
 *
 * @param {Error} error - The error that caused the failure.
 */
function sendFailureNotification(error) {
  if (!NOTIFY_EMAIL) return;

  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: "Core Sessions Index Builder — Run Failed",
      body: [
        "The nightly index build failed.",
        "",
        `Error: ${error.message}`,
        "",
        `Stack trace:\n${error.stack}`,
        "",
        `Project: ${PROJECT_URL}`,
        "",
        `Timestamp: ${new Date().toISOString()}`,
      ].join("\n"),
    });
  } catch (mailError) {
    console.error(`Failed to send notification email: ${mailError.message}`);
  }
}

// ── Main ─────────────────────────────────────────────

/**
 * Lists all folders, subfolders, and files inside the "Core Sessions"
 * Google Drive folder and populates the _Index sheet with hyperlinks.
 */
function listSubfoldersAndFiles() {
  console.info("Index build started.");

  try {
    const folder = DriveApp.getFolderById(CORE_SESSIONS_FOLDER_ID);
    const spreadsheet = SpreadsheetApp.openById(INDEX_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(INDEX_SHEET_NAME);

    // Clear existing content from START_ROW down
    const lastRow = sheet.getLastRow();
    if (lastRow >= START_ROW) {
      sheet
        .getRange(START_ROW, 1, lastRow - START_ROW + 1, sheet.getMaxColumns())
        .clearContent();
    }

    // Collect root items
    const items = getRootItems(folder);
    console.info(`Collected ${items.length} root items.`);

    // Build a 2D data array for batch writing (columns A and B)
    const data = [];

    items.forEach((item) => {
      const linkFormula = buildHyperlink(item.url, item.name);

      if (item.type === "folder") {
        const children = getChildItems(item.folder);

        if (children.length === 0) {
          data.push([linkFormula, ""]);
        } else {
          data.push([linkFormula, buildHyperlink(children[0].url, children[0].name)]);

          for (let i = 1; i < children.length; i++) {
            data.push(["", buildHyperlink(children[i].url, children[i].name)]);
          }
        }

        data.push(["", ""]);
      } else {
        data.push([linkFormula, ""]);
        data.push(["", ""]);
      }
    });

    // Write everything to the sheet in a single batch call
    if (data.length > 0) {
      sheet.getRange(START_ROW, 1, data.length, 2).setValues(data);
    }

    console.info(`Index build complete. Wrote ${data.length} rows.`);
  } catch (error) {
    console.error(`Index build failed: ${error.message}\n${error.stack}`);
    sendFailureNotification(error);
    throw error; // Re-throw so Apps Script marks the execution as failed
  }
}
