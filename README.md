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
