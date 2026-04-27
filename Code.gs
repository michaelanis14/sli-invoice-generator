/*================================================================================================================*
Invoice Generator
================================================================================================================
Version:      1.2.0
Project Page: https://github.com/Sheetgo/invoice-generator
Copyright:    (c) 2018 by Sheetgo

License:      GNU General Public License, version 3 (GPL-3.0)
http://www.opensource.org/licenses/gpl-3.0.html
----------------------------------------------------------------------------------------------------------------
Changelog:

1.0.0  Initial release
1.1.0  Auto configuration
1.2.0  - Replaced Drive.Files.remove() with DriveApp.setTrashed(true) to fix
         "API call to drive files delete failed with error: Empty response"
       - Added retryOnDriveError() wrapper for transient "Service error: Drive"
       - Hardened cleanup in createInvoices() catch block (no more ReferenceError
         when invoiceId never got assigned)
       - convertPDF() rewritten to avoid the deprecated v2 Drive.Files.update
         signature; now uses DriveApp directly with retries
       - Fixed createSystem() typos: SETTINGS.col.systemCreated -> SystemCreated,
         SETTINGS.col.count -> Count, replaced non-existent templateId/folderId
         keys with Original_ID / Original_Folder_ID
       - sendInvoice() now logs failures per row but continues processing
         remaining rows instead of swallowing errors silently
       - Minor: tightened scoping, removed dead code
*================================================================================================================*/

/**
 * Project Settings
 * @type {Object}
 */
SETTINGS = {

  // Spreadsheet name
  sheetName: "Data",

  // Document Url
  documentUrl: null,

  // Template Url
  templateUrl: '14oTfL_zUbBdRD4VXY8U0NAJjQ4cKNxHGBax-bfH5NDs',

  // Set name spreadsheet
  spreadsheetName: 'Invoice data',

  // Set name document
  documentName: 'Invoice Template',

  // Sheet Settings
  sheetSettings: "Settings",

  // Authorised editors for protected rows
  editors: ["anis@sli-eg.com", "george@sli-eg.com"],

  // Retry config for transient Drive errors
  retry: {
    maxAttempts: 3,
    baseDelayMs: 2000
  },

  // Column Settings (cell references in the Settings sheet)
  col: {
    Count: "B1",
    Original_ID: "B2",
    Original_Folder_ID: "B3",
    Draft_ID: "B4",
    Draft_Folder_ID: "B5",
    Copy_ID: "B6",
    Copy_Folder_ID: "B7",
    Deleted_ID: "B8",
    Deleted_Folder_ID: "B9",
    Puplic_ID: "B10",
    Puplic_Folder_ID: "B11",
    SystemCreated: "B12"
  }
};

/* =================================================================================
 *  UI / Menu
 * ================================================================================= */

function testUI() {
  SpreadsheetApp.getUi().alert("Hello! The UI is connected.");
}

/**
 * Runs when the spreadsheet is opened. Creates the custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Invoice Generator')
    .addItem('Generate Drafts', 'createDraft')
    .addItem('Generate Invoices', 'sendInvoice')
    .addToUi();
}

/* =================================================================================
 *  Retry helper for transient Drive failures
 * ================================================================================= */

/**
 * Re-runs `fn` up to maxAttempts times with linear backoff when it throws.
 * Targets transient Drive failures like "Service error: Drive" or "Empty response".
 *
 * @param {Function} fn      Function to invoke (no args, returns a value)
 * @param {string}   label   Short label for logs
 * @returns {*} Whatever fn returns on a successful attempt
 */
function retryOnDriveError(fn, label) {
  var maxAttempts = SETTINGS.retry.maxAttempts;
  var baseDelayMs = SETTINGS.retry.baseDelayMs;
  var lastError;

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return fn();
    } catch (e) {
      lastError = e;
      Logger.log('retryOnDriveError [' + label + '] attempt ' + attempt +
                 '/' + maxAttempts + ' failed: ' + e.message);
      if (attempt < maxAttempts) {
        Utilities.sleep(baseDelayMs * attempt);
      }
    }
  }
  throw lastError;
}

/* =================================================================================
 *  System setup
 * ================================================================================= */

/**
 * Initial one-time setup: creates the Drive folder structure and copies the template.
 */
function createSystem() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSettings = ss.getSheetByName(SETTINGS.sheetSettings);

    // Has the system already been created?
    var systemCreated = sheetSettings.getRange(SETTINGS.col.SystemCreated);
    if (!systemCreated.getValue()) {
      systemCreated.setValue('True');
    } else {
      showUiDialog('Warning', 'Solution has already been created!');
      return;
    }

    // Initialise Count cell if empty
    var count = sheetSettings.getRange(SETTINGS.col.Count);
    if (!count.getValue()) {
      count.setValue(0);
    }

    // Create the solution folder structure on the user's Drive
    var invoiceFolder = DriveApp.createFolder('Invoice Folder');
    var folder = invoiceFolder.createFolder('Invoices');

    // Surface the folder URL on the Instructions tab
    ss.getSheetByName('Instructions').getRange('C15').setValue(invoiceFolder.getUrl());

    // Move the active spreadsheet into the new folder
    var file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
    file.setName(SETTINGS.spreadsheetName);
    moveFile(file, invoiceFolder);

    // Copy the template doc into the new folder
    var doc = DriveApp.getFileById(SETTINGS.templateUrl);
    var docCopy = doc.makeCopy(SETTINGS.documentName);
    sheetSettings.getRange(SETTINGS.col.Original_ID).setValue(docCopy.getId());
    moveFile(docCopy, invoiceFolder);

    // Record the invoices folder ID
    sheetSettings.getRange(SETTINGS.col.Original_Folder_ID).setValue(folder.getId());

    showUiDialog('Success', 'Your solution is ready');
    return true;
  } catch (e) {
    showUiDialog('Something went wrong', e.message);
  }
}

/**
 * Creates the per-year Original / Copy subfolders and writes their IDs back to Settings.
 */
function init() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS.sheetSettings);

  if (!settingsSheet) {
    throw new Error('Settings sheet not found. Please ensure the sheet is named correctly.');
  }

  const currentYear = new Date().getFullYear();

  // Locate the parent of the existing original-document folder (two levels up)
  const settingsSheetValues = settingsSheet.getDataRange().getValues();
  const existingOriginalDocumentId = settingsSheetValues[2][1];

  const originalDocument = DriveApp.getFileById(existingOriginalDocumentId);
  const originalDocumentParent = originalDocument.getParents().next();
  const pdfFolder = originalDocumentParent.getParents().next();

  // Find or create the year folder
  let newYearFolder;
  const yearFolders = pdfFolder.getFoldersByName(currentYear.toString());
  if (yearFolders.hasNext()) {
    newYearFolder = yearFolders.next();
  } else {
    newYearFolder = pdfFolder.createFolder(currentYear.toString());
  }

  const originalFolderId = newYearFolder.createFolder("Original").getId();
  const copyFolderId = newYearFolder.createFolder("Copy").getId();

  settingsSheet.getRange(SETTINGS.col.Original_Folder_ID).setValue(originalFolderId);
  settingsSheet.getRange(SETTINGS.col.Copy_Folder_ID).setValue(copyFolderId);

  SpreadsheetApp.getUi().alert('Folders created and settings updated successfully!');
}

/* =================================================================================
 *  Main entry: generate invoices for every unprocessed row
 * ================================================================================= */

function sendInvoice() {
  Logger.log('Script Started');

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName(SETTINGS.sheetName);
    var sheetSettings = ss.getSheetByName(SETTINGS.sheetSettings);

    var sheetValues = dataSheet.getDataRange().getValues();

    var pdfIndex      = sheetValues[0].indexOf("Original_Url");
    var pdf_ID_Index  = sheetValues[0].indexOf("orginal_ID");
    var copyUrlIdx    = sheetValues[0].indexOf("Copy_Url");
    var copyIdIdx     = sheetValues[0].indexOf("copy_ID");
    var serialIndex   = sheetValues[0].indexOf("Serial_Number");
    var draftIndex    = sheetValues[0].indexOf("draft");
    var deleteIndex   = sheetValues[0].indexOf("delete");

    var settingsSheetValues = sheetSettings.getDataRange().getValues();

    // Has setup been completed?
    var control = settingsSheetValues[11][1];
    if (!control) {
      showUiDialog('Warning', 'Run "Install Solution" in tab Instructions');
      return;
    }

    var counter           = settingsSheetValues[0][1];
    var originalsDocument = settingsSheetValues[1][1] + "";
    var originalsFolder   = settingsSheetValues[2][1] + "";
    var copyDocument      = settingsSheetValues[5][1] + "";
    var copyFolder        = settingsSheetValues[6][1] + "";

    Logger.log('Loaded all data');

    var invoiceNumCount;
    var failures = [];

    for (var i = 1; i < sheetValues.length; i++) {
      Logger.log('Loop row ' + i + ' of ' + (sheetValues.length - 1));

      // Skip rows that already have a PDF, are flagged as drafts, or are deleted
      if (sheetValues[i][pdfIndex] || sheetValues[i][draftIndex] || sheetValues[i][deleteIndex]) {
        if (sheetValues[i][draftIndex]) {
          Logger.log('Hit draft row ' + i + ', stopping');
          break;
        }
        continue;
      }

      Logger.log('Creating invoice for row ' + i);

      // Increment counter only when there is no existing serial number
      var existingSerial = (sheetValues[i][serialIndex] != null)
          ? sheetValues[i][serialIndex].toString().trim()
          : '';

      if (existingSerial.length === 0) {
        invoiceNumCount = counter + 1;
        sheetSettings.getRange(SETTINGS.col.Count).setValue(invoiceNumCount);
        counter = invoiceNumCount;
      } else {
        invoiceNumCount = sheetValues[i][serialIndex];
      }

      try {
        Logger.log('Create Original for row ' + i);
        createInvoices(dataSheet, sheetValues, i, originalsDocument,
                       invoiceNumCount, originalsFolder, "Original",
                       pdfIndex, pdf_ID_Index);

        Logger.log('Create Copy for row ' + i);
        createInvoices(dataSheet, sheetValues, i, copyDocument,
                       invoiceNumCount, copyFolder, "Copy",
                       copyUrlIdx, copyIdIdx);

        Logger.log('Set serial for row ' + i);
        dataSheet.getRange(i + 1, serialIndex + 1).setValue(invoiceNumCount);
        dataSheet.getRange(i + 1, draftIndex + 1).setValue("FALSE");
        dataSheet.getRange(i + 1, deleteIndex + 1).setValue("FALSE");

        Logger.log('Protect row ' + i);
        protectRow(ss, i);
      } catch (rowError) {
        Logger.log('Row ' + i + ' failed: ' + rowError.message);
        failures.push('Row ' + (i + 1) + ': ' + rowError.message);
        // Continue with the next row instead of aborting the whole batch
      }
    }

    if (failures.length > 0) {
      showUiDialog('Finished with errors',
                   'Some rows failed:\n\n' + failures.join('\n'));
    } else {
      Logger.log('All rows processed successfully');
    }
  } catch (e) {
    showUiDialog('Finished Invoice Generation', e.message);
  }
}

/* =================================================================================
 *  Per-row invoice creation
 * ================================================================================= */

/**
 * Copies the template, fills it with row data, converts to PDF, then trashes the temp doc.
 */
function createInvoices(dataSheet, sheetValues, rowIndex, docId, invoiceNumCount,
                        folderId, linkText, clmnLinkIndex, clmnIdIndex) {
  var invoiceId = null;

  try {
    Logger.log('createInvoices [' + linkText + '] - copy template');
    invoiceId = retryOnDriveError(function () {
      return DriveApp.getFileById(docId).makeCopy(linkText + "_Template").getId();
    }, 'makeCopy ' + linkText);

    Logger.log('createInvoices [' + linkText + '] - fill document');
    var newFileTitle = createDocument(sheetValues, rowIndex, invoiceId,
                                      invoiceNumCount, linkText);

    Logger.log('createInvoices [' + linkText + '] - read existing PDF id');
    var existingFileId = dataSheet.getRange(rowIndex + 1, clmnIdIndex + 1).getValue();

    Logger.log('createInvoices [' + linkText + '] - convert to PDF');
    var pdfInvoice = retryOnDriveError(function () {
      return convertPDF(invoiceId, folderId, newFileTitle, existingFileId);
    }, 'convertPDF ' + linkText);

    Logger.log('createInvoices [' + linkText + '] - write back URL & id');
    dataSheet.getRange(rowIndex + 1, clmnLinkIndex + 1)
             .setValue(createHyperlinkString(pdfInvoice[0], linkText));
    dataSheet.getRange(rowIndex + 1, clmnIdIndex + 1).setValue(pdfInvoice[1]);

    Logger.log('createInvoices [' + linkText + '] - trash template copy');
    safeTrash(invoiceId);
  } catch (error) {
    Logger.log('createInvoices [' + linkText + '] error: ' + error.message);
    // Best-effort cleanup of the temp template copy
    safeTrash(invoiceId);
    // Rethrow so sendInvoice can record this row as a failure
    throw error;
  }
}

/**
 * Trashes a Drive file by ID, swallowing any errors. Safe to call with null/undefined.
 */
function safeTrash(fileId) {
  if (!fileId) return;
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (cleanupErr) {
    Logger.log('safeTrash failed for ' + fileId + ': ' + cleanupErr.message);
  }
}

/**
 * Replaces template placeholders with row data and saves the document.
 * Returns the title that was set on the document.
 */
function createDocument(sheetValues, rowIndex, invoiceId, invoiceNumCount, linkText) {
  var key, values, invoiceNumber, invoiceDate;

  var doc = DocumentApp.openById(invoiceId);
  var docBody = doc.getBody();

  // Columns whose values must be inserted as-is (no numeric formatting)
  var rawKeys = {
    'mawb_no': 1, 'hawb_no': 1, 'no_pieces': 1, 'discharge_txt': 1,
    'natural': 1, 'decimal': 1, 'cash': 1, 'rout': 1,
    'weight': 1, 'gross_weight': 1
  };

  for (var j = 0; j < sheetValues[rowIndex].length; j++) {
    key    = sheetValues[0][j];
    values = sheetValues[rowIndex][j];

    if (key === "date") {
      var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      invoiceDate = Utilities.formatDate(values, timezone, "d/M/yyyy");
      replace('%date%', invoiceDate, docBody);
    } else if (values) {
      if (rawKeys[key]) {
        replace('%' + key + '%', values, docBody);
      } else if (!isNaN(parseFloat(values)) && isFinite(values)) {
        replace('%' + key + '%', financial(values), docBody);
      } else {
        replace('%' + key + '%', values, docBody);
      }
    } else {
      replace('%' + key + '%', '', docBody);
    }
  }

  // Format the invoice number and rename the file
  invoiceNumber = invoiceNumCount.padLeft(7, '0') + "/" + invoiceDate.split("/")[2];
  replace('%number%', invoiceNumber, docBody);

  var newFileTitle = invoiceNumber + " - " + linkText;
  doc.setName(newFileTitle).saveAndClose();

  return newFileTitle;
}

/* =================================================================================
 *  Row protection
 * ================================================================================= */

function protectRow(ss, rowIndex) {
  Logger.log('protectRow ' + rowIndex);
  var rangeString = "Data!" + (rowIndex + 1) + ":" + (rowIndex + 1);
  var range = ss.getRange(rangeString);
  var protectionDescription = 'INVOICE CREATED ROW ' + (rowIndex + 1);

  removeProtections(ss, protectionDescription);
  var protection = range.protect().setDescription(protectionDescription);

  protection.removeEditors(protection.getEditors());
  for (var k = 0; k < SETTINGS.editors.length; k++) {
    protection.addEditor(SETTINGS.editors[k]);
  }
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

function removeProtections(ss, protectionDescription) {
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() == protectionDescription) {
      protections[i].remove();
    }
  }
}

function removeSpecificProtections() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var protections = sheets[i].getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var j = 0; j < protections.length; j++) {
      var description = protections[j].getDescription();
      if (description && description.match(/^INVOICE CREATED ROW \d+$/)) {
        protections[j].remove();
      }
    }
  }
}

/* =================================================================================
 *  Drive helpers
 * ================================================================================= */

/**
 * Move a file from one folder into another.
 */
function moveFile(file, dest_folder, isFolder) {
  if (isFolder === true) {
    dest_folder.addFolder(file);
  } else {
    dest_folder.addFile(file);
  }
  var parents = file.getParents();
  while (parents.hasNext()) {
    var folder = parents.next();
    if (folder.getId() != dest_folder.getId()) {
      if (isFolder === true) {
        folder.removeFolder(file);
      } else {
        folder.removeFile(file);
      }
    }
  }
}

/**
 * Convert a Google Doc into a PDF file. If existingFileID is provided, the
 * existing PDF is replaced (old one trashed, new one created in invFolder);
 * otherwise a new PDF is created in invFolder.
 *
 * @param {string} templateGdocId  ID of the source Google Doc
 * @param {string} invFolder       ID of the destination folder
 * @param {string} fileName        PDF filename
 * @param {string} existingFileID  Optional ID of an existing PDF to replace
 * @returns {[string, string]}     [url, id]
 */
function convertPDF(templateGdocId, invFolder, fileName, existingFileID) {
  var docBlob = DocumentApp.openById(templateGdocId).getAs('application/pdf');
  docBlob.setName(fileName);

  // If an existing PDF is referenced, validate and trash it.
  // Avoids the deprecated v2 Drive.Files.update signature that triggers
  // "Empty response" errors on some accounts.
  if (existingFileID) {
    try {
      DriveApp.getFileById(existingFileID).setTrashed(true);
    } catch (e) {
      Logger.log('convertPDF: existingFileID ' + existingFileID +
                 ' could not be trashed (' + e.message + '), continuing.');
    }
  }

  var newFile = DriveApp.getFolderById(invFolder).createFile(docBlob);
  return [newFile.getUrl(), newFile.getId()];
}

/* =================================================================================
 *  Misc helpers
 * ================================================================================= */

/**
 * Replace a placeholder in the document body with the supplied text.
 */
function replace(key, text, body) {
  return body.editAsText().replaceText(key, text);
}

/**
 * Right-align by left-padding to a fixed length.
 */
Number.prototype.padLeft = function (n, str) {
  return Array(n - String(this).length + 1).join(str || '0') + this;
};

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('iframe.html')
    .setWidth(200)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Creating Solution..');
}

function createHyperlinkString(link, text) {
  return '=HYPERLINK("' + link + '", "' + text + '")';
}

function showUiDialog(title, message) {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (e) {
    // No UI context available (e.g., triggered run) — log instead.
    Logger.log('showUiDialog: ' + title + ' - ' + message);
  }
}

function financial(x) {
  return Number.parseFloat(x).toFixed(2);
}
