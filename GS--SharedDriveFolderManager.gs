/**
 * This Google Apps Script allows users to import and manage the folder structure of a shared drive.
 * 
 * Requirements to run this script:
 * 1. Google Drive API must be enabled. 
 *    - Go to Extensions -> Apps Script.
 *    - In the script editor, go to Resources -> Advanced Google Services.
 *    - Turn on the Drive API.
 *    - Click on the Google Cloud Platform API Console link at the bottom.
 *    - In the API Console, enable the Google Drive API.
 * 2. The Google Sheets document should have the sheet name as the shared drive ID.
 * 
 * What this script does:
 * 1. Adds a custom menu with two options: "Import Folder Structure" and "Push Changes".
 * 2. "Import Folder Structure" reads the folder structure from a shared drive and lists it in the active sheet.
 *    - The top-level folder is marked as "Drive" in column A.
 *    - Folder names are listed starting from column A.
 *    - Column A will always be in italic style, and column B will always be in bold style.
 * 3. "Push Changes" reads the folder structure from the sheet and creates any missing folders in the shared drive.
 *    - It only creates folders; it does not rename,delete or move existing folders.
 * 
 * Written by Jonas Lund // 2024
 */

// Function to add a custom menu to the Google Sheets interface
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Folder Manager Menu')
      .addItem('Import Folder Structure', 'importSharedDriveStructure')
      .addItem('Push Folder Changes', 'pushChanges')
      .addToUi();
}

// Function to read the shared drive ID from the current active sheet name and export the folder structure to Google Sheets
function importSharedDriveStructure() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var sheetName = sheet.getName(); // Get the current active sheet name

  // Assume the sheet name is the shared drive ID
  var driveId = sheetName;

  try {
    var drive = DriveApp.getFolderById(driveId);
  } catch (e) {
    Browser.msgBox("Failed to access the shared drive. Please ensure you have the necessary permissions and the sheet name is a valid shared drive ID.");
    return;
  }

  // Clear previous data if any
  sheet.clear();

  // Write "Drive" as the top level in column A
  sheet.getRange(1, 1).setValue("Drive").setFontStyle("italic");

  // Get the folder structure with the shared drive name as the root
  var folderTree = getFolderTree(drive);

  // Write the folder structure to the Google Sheets document
  writeFolderTreeToSheet(folderTree, sheet, 2, 1, []); // Start from row 2, column A (index 1)

  // Apply bold style to column B
  sheet.getRange("B:B").setFontWeight('bold');
}

// Function to recursively get the folder structure
function getFolderTree(folder) {
  var folderTree = {
    name: folder.getName(),
    subfolders: []
  };

  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    folderTree.subfolders.push(getFolderTree(subfolder));
  }

  return folderTree;
}

// Function to write the folder structure to the Google Sheets document with hierarchical columns
function writeFolderTreeToSheet(folderTree, sheet, row, column, path) {
  path.push(folderTree.name);
  var range = sheet.getRange(row, column, 1, path.length).setValues([path]);

  // Apply italic style to column A
  sheet.getRange(row, 1).setFontStyle("italic");

  var nextRow = row + 1;
  for (var i = 0; i < folderTree.subfolders.length; i++) {
    nextRow = writeFolderTreeToSheet(folderTree.subfolders[i], sheet, nextRow, column, path.slice());
  }

  return nextRow;
}

// Function to push changes from the sheet to the shared drive
function pushChanges() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var sheetName = sheet.getName(); // Get the current active sheet name

  // Assume the sheet name is the shared drive ID
  var driveId = sheetName;

  try {
    var drive = DriveApp.getFolderById(driveId);
  } catch (e) {
    Browser.msgBox("Failed to access the shared drive. Please ensure you have the necessary permissions and the sheet name is a valid shared drive ID.");
    return;
  }

  // Read the folder structure from the sheet, starting from column B
  var sheetData = readFolderTreeFromSheet(sheet);

  // Push the folder structure to the shared drive
  createMissingFolders(drive, sheetData);
}

// Function to read the folder structure from the sheet, starting from column B
function readFolderTreeFromSheet(sheet) {
  var data = sheet.getDataRange().getValues();
  var folderTree = {
    name: 'root',
    subfolders: []
  };

  var pathMap = {};
  pathMap['root'] = folderTree;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var parentNode = folderTree;

    for (var j = 1; j < row.length; j++) { // Start from column B (index 1)
      var cell = row[j];
      if (cell !== '') {
        var path = row.slice(1, j + 1).join('/'); // Skip column A
        if (!pathMap[path]) {
          var newNode = {
            name: cell,
            subfolders: []
          };
          pathMap[path] = newNode;
          parentNode.subfolders.push(newNode);
        }
        parentNode = pathMap[path];
      }
    }
  }

  return folderTree;
}

// Function to create missing folders in the shared drive based on the folder structure from the sheet
function createMissingFolders(parentFolder, folderTree) {
  if (folderTree.name !== 'root') { // Skip the 'root' node
    var folderName = folderTree.name;

    var folders = parentFolder.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = parentFolder.createFolder(folderName);
    }

    for (var i = 0; i < folderTree.subfolders.length; i++) {
      createMissingFolders(folder, folderTree.subfolders[i]);
    }
  } else {
    for (var i = 0; i < folderTree.subfolders.length; i++) {
      createMissingFolders(parentFolder, folderTree.subfolders[i]);
    }
  }
}
