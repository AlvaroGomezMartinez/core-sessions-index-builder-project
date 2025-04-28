/********************************************************
 *        CORE SESSIONS INDEX BUILDER PROJECT           *
 *                                                      *
 * Alvaro Gomez                                         *
 * Northside Independent School District                *
 * Academic Technology Coach                            *
 * alvaro.gomez@nisd.net                                *
 * +1-210-363-1577                                      *
 *                                                      *
 * Objective: This unbound Google Apps Script project   *
 *            creates an index of the folders           *
 *            and files located inside the root folder  *
 *            called, "Core Sessions". A trigger is     *
 *            set to run the main function every day    *
 *            at midnight.                              *
********************************************************/

/**
 * This is the main function that lists all folders,
 * subfolders, and files inside the "Core Sessions"
 * Google Drive folder, and populates the _Index sheet
 * with the hyperlinks to these items.
 */
function listSubfoldersAndFiles() {
  // Get the Google Folder named Core Sessions by its ID
  var folderId = "0B6Z4Z9quk8OJUmlwSXpMTER3d1k";
  var folder = DriveApp.getFolderById(folderId);

  // Get the Google Spreadsheet and the specific sheet named Index by its ID
  var spreadsheetId = "18jneeP5YgIKhfmiYj8vNHZYbhd_WTn8AN-hEXaO6K9k";
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName("Index");

  // Clear the Index sheet from row 3 down
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(3, 1, lastRow - 2, sheet.getMaxColumns()).clearContent();
  }

  var row = 3; // Start at row 3

  // Get a list of all subfolders and files inside Core Sessions
  var items = [];
  var subfolderIterator = folder.getFolders();

  while (subfolderIterator.hasNext()) {
    var subfolder = subfolderIterator.next();
    items.push({
      name: subfolder.getName(),
      url: subfolder.getUrl(), // No need to resolve shortcuts for folders
      type: "folder",
      folder: subfolder,
    });
  }

  // Get files that are directly inside Core Sessions
  var fileIterator = folder.getFiles();

  while (fileIterator.hasNext()) {
    var file = fileIterator.next();
    items.push({
      name: file.getName(),
      url: getActualUrl(file), // Use helper to resolve shortcuts
      type: "file",
    });
  }

  // Sort the combined list of subfolders and files alphabetically by name
  items.sort(function (a, b) {
    return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
  });

  // List items (both subfolders and files) in the Index sheet
  items.forEach(function (item) {
    sheet
      .getRange(row, 1)
      .setValue('=HYPERLINK("' + item.url + '", "' + item.name + '")');

    if (item.type === "folder") {
      // If the item is a subfolder, list its contents in column B
      var fileRow = row; // Start at the same row for files and subfolders

      // Get the sub-subfolders in the subfolder
      var subSubfolders = [];
      var subSubfolderIterator = item.folder.getFolders();

      while (subSubfolderIterator.hasNext()) {
        var subSubfolder = subSubfolderIterator.next();
        subSubfolders.push({
          name: subSubfolder.getName(),
          url: subSubfolder.getUrl(), // No need to resolve shortcuts for folders
        });
      }

      // Sort the sub-subfolders alphabetically by name
      subSubfolders.sort(function (a, b) {
        return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
      });

      // Loop through each sorted sub-subfolder
      subSubfolders.forEach(function (subSubfolderInfo) {
        sheet
          .getRange(fileRow, 2)
          .setValue(
            '=HYPERLINK("' +
              subSubfolderInfo.url +
              '", "' +
              subSubfolderInfo.name +
              '")'
          );
        fileRow++; // Move to the next row for the next sub-subfolder
      });

      // Get the files in the subfolder
      var files = item.folder.getFiles();

      // Loop through each file in the subfolder
      while (files.hasNext()) {
        var file = files.next();
        sheet
          .getRange(fileRow, 2)
          .setValue(
            '=HYPERLINK("' + getActualUrl(file) + '", "' + file.getName() + '")'
          );
        fileRow++; // Move to the next row for the next file
      }

      row = fileRow + 1; // Move to the next row after the last file for the next subfolder
    } else {
      row++; // Move to the next row for the next item
      sheet.insertRowAfter(row - 1); // Insert an empty row after each root file
      row++; // Skip the empty row
    }
  });
}

/**
 * Gets the actual URL of a file, handling shortcuts if necessary.
 *
 * @param {GoogleAppsScript.Drive.File} file - The file object from which to get the URL.
 * @return {string} The actual URL of the file. If the file is a shortcut, returns the target URL.
 */
function getActualUrl(file) {
  if (file.getMimeType() === MimeType.SHORTCUT) {
    try {
      var targetId = file.getTargetId();
      return DriveApp.getFileById(targetId).getUrl();
    } catch (e) {
      // If there's an error (e.g., invalid ID or no permission), return the shortcut URL itself
      return file.getUrl();
    }
  } else {
    return file.getUrl();
  }
}
