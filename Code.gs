/*================================================================================================================*
Invoice Generator
================================================================================================================
Version:      1.0.0
Project Page: https://github.com/Sheetgo/invoice-generator
Copyright:    (c) 2018 by Sheetgo

License:      GNU General Public License, version 3 (GPL-3.0)
http://www.opensource.org/licenses/gpl-3.0.html
----------------------------------------------------------------------------------------------------------------
Changelog:

1.0.0  Initial release
1.1.0  Auto configuration
*================================================================================================================*/

/**
* Project Settings
* @type {JSON}
*/
const SETTINGS = {

  // Spreadsheet name
  sheetName: "Data",

  // Document Url
  documentUrl: null,

  // Template Url
  templateUrl: '14oTfL_zUbBdRD4VXY8U0NAJjQ4cKNxHGBax-bfH5NDs',

  // Set name spreadsheet
  spreadsheetName: 'Invoice data',

  //Set name document
  documentName: 'Invoice Template',

  // Sheet Settings
    sheetSettings: "Settings",


  // Column Settings
  col: {
    Count: "B1",
    Original_ID: "B2",
    Original_Folder_ID:"B3",
    Draft_ID:"B4",
    Draft_Folder_ID:"B5",
    Copy_ID:"B6",
    Copy_Folder_ID:"B7",
    Deleted_ID:"B8",
    Deleted_Folder_ID:"B9",
    Public_ID:"B10",
    Public_Folder_ID:"B11",
    SystemCreated: "B12"
  }
};

/**
* This funcion will run when you open the spreadsheet. It creates a Spreadsheet menu option to run the spript
*/
function onOpen() {

  // Adds a custom menu to the spreadsheet.
  SpreadsheetApp.getUi()
  .createMenu('Invoice Generator')
  .addItem('Generate Invoices', 'sendInvoice')
  .addToUi();
}

/**
* This function Create system
*/
function createSystem() {

  try {

    var ss = SpreadsheetApp.getActiveSpreadsheet();


    // Get name tab
    var sheetSettings = ss.getSheetByName(SETTINGS.sheetSettings);

    // Checks function createSystem is run
    var systemCreated = sheetSettings.getRange(SETTINGS.col.SystemCreated);
    if (!systemCreated.getValue()){
      systemCreated.setValue('True');
    } else {
      showUiDialog('Warnning','Solution has already been created!');
      return;
    }

    // Checks if cell Count exists
    var count = sheetSettings.getRange(SETTINGS.col.Count);
    if(!count.getValue()){
      count.setValue(0);
    }

    // Create the Solution folder on users Drive
    var invoiceFolder = DriveApp.createFolder('Invoice Folder');
    var folder = invoiceFolder.createFolder('Invoices');

    // Set URL Invoice Folder in tab Instructions
    ss.getSheetByName('Instructions').getRange('C15').setValue(invoiceFolder.getUrl());

    // Move the current Dashboard spreadsheet into the Solution folder
    var file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
    file.setName(SETTINGS.spreadsheetName);

    // Move the sheet for invoice folder
    moveFile(file, invoiceFolder);

    // Move the current Dashboard template into the Solution folder
    var doc = DriveApp.getFileById(SETTINGS.templateUrl);
    var docCopy = doc.makeCopy(SETTINGS.documentName);

    // Set tab settings document ID (Original template — Copy template is set up separately via init())
    sheetSettings.getRange(SETTINGS.col.Original_ID).setValue(docCopy.getId());

    // Move an copy for invoice folder
    moveFile(docCopy, invoiceFolder);

    // Set folder ID
    sheetSettings.getRange(SETTINGS.col.Original_Folder_ID).setValue(folder.getId());


    // End process
    showUiDialog('Success', 'Your solution is ready');

    return true;
  } catch (e) {

    // Show the error
    showUiDialog('Something went wrong', e.message);

  }
}


/**
* Reads the spreadsheet data and creates the PDF invoice
*/
function sendInvoice() {
    Logger.log('Script Started');

//var response = SpreadsheetApp.getUi().alert('Do you want to generate document?', SpreadsheetApp.getUi().ButtonSet.YES_NO);    // Opens the spreadsheet and access the tab containing the data
  //if (response == SpreadsheetApp.getUi().Button.YES) {
  try {

    Logger.log('YES to start script');


    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName(SETTINGS.sheetName);
    var sheetSettings = ss.getSheetByName(SETTINGS.sheetSettings);

   /*
    // Checks if cell Count exists
    var count = sheetSettings.getRange(SETTINGS.col.Count).getValue();
    if(!count){
      sheetSettings.getRange(SETTINGS.col.Count).setValue(0);
    }
    */


    // Gets all values from the instanciated tab
    var sheetValues = dataSheet.getDataRange().getValues();

    var pdfIndex = sheetValues[0].indexOf("Original_Url");
    var pdf_ID_Index = sheetValues[0].indexOf("orginal_ID");
    var copyUrlIdx = sheetValues[0].indexOf("Copy_Url");
    var copyIdIdx = sheetValues[0].indexOf("copy_ID");


    var serialIndex = sheetValues[0].indexOf("Serial_Number");
    var draftIndex = sheetValues[0].indexOf("draft");
    var deleteIndex = sheetValues[0].indexOf("delete");

    // Gets the user's name (will be used as the PDF file name)
    // var clientNameIndex = sheetValues[0].indexOf("client_name");

    var counter, invoiceNumCount;

    var settingsSheetValues = sheetSettings.getDataRange().getValues();

    // Checks function createSystem is run
    var control = settingsSheetValues[11][1];
    if (!control){
      showUiDialog('Warnning','Run "Install Solution" in tab Instructions');
      return;
    }

    // Duplicate teh template on Google Drive to manipulate the data
     counter = settingsSheetValues[0][1];


    //var originalsDocument = sheetSettings.getRange(SETTINGS.col.Original_ID).getValue();
    var originalsDocument =  settingsSheetValues[1][1]+"";
  //SpreadsheetApp.getUi().alert('indexOf method on a string in google app originalsDocument '+originalsDocument)
    //var originalsFolder = sheetSettings.getRange(SETTINGS.col.Original_Folder_ID).getValue();
    var originalsFolder =  settingsSheetValues[2][1]+"";

    //var copyDocument = sheetSettings.getRange(SETTINGS.col.Copy_ID).getValue();
    var copyDocument =  settingsSheetValues[5][1]+"";

   // var copyFolder = sheetSettings.getRange(SETTINGS.col.Copy_Folder_ID).getValue();
    var copyFolder =  settingsSheetValues[6][1]+"";

    Logger.log('Loaded all data');

    for (var i = 1; i < sheetValues.length; i++) {


    Logger.log('Loop for sheetValues '+ i +' of '+ sheetValues.length);

      // Creates the Invoice
      if (!sheetValues[i][pdfIndex] && !sheetValues[i][draftIndex] && !sheetValues[i][deleteIndex]) {
        Logger.log('Creating Invoice '+ i );

    //    var response = SpreadsheetApp.getUi().alert('Generate Invoice '+(counter + 1) + ' ?', SpreadsheetApp.getUi().ButtonSet.YES_NO);    // Opens the spreadsheet and access the tab containing the data
    //     if (response == SpreadsheetApp.getUi().Button.NO) continue;

       // Get last invoice count from the tab 'Count'
       // counter = sheetSettings.getRange(SETTINGS.col.Count);
      // if(!sheetValues[i+1][serialIndex+1])
       // SpreadsheetApp.getUi().alert(sheetValues[i][serialIndex] +' & '+dataSheet.getRange(i + 1, serialIndex + 1).getValue());

      //  SpreadsheetApp.getUi().alert((sheetValues[i+1][serialIndex]?.trim()?.length || 0) === 0);


       if((sheetValues[i][serialIndex]?.toString().trim()?.length || 0) === 0){ // avoid incremental the counter
        Logger.log('NO Serial '+ i );

            invoiceNumCount = counter + 1;
            sheetSettings.getRange(SETTINGS.col.Count).setValue(invoiceNumCount);
            counter = invoiceNumCount;

        } else invoiceNumCount = sheetValues[i][serialIndex];

        Logger.log('CreateOriginal '+ i );
        //OriginalInvoice
        var originalOk = createInvoices(dataSheet,sheetValues,i,originalsDocument,invoiceNumCount,originalsFolder,"Original",pdfIndex,pdf_ID_Index);
        if (!originalOk) {
          Logger.log('Original failed for row '+i+'; skipping Copy and row completion');
          continue;
        }

        Logger.log('CreateCopy '+ i );
        //CopyInvoice
        var copyOk = createInvoices(dataSheet,sheetValues,i,copyDocument,invoiceNumCount,copyFolder,"Copy",copyUrlIdx,copyIdIdx);
        if (!copyOk) {
          Logger.log('Copy failed for row '+i+'; row left unprotected so it can be retried after clearing the Original_Url cell');
          continue;
        }

        Logger.log('SetSerial '+ i );
        //set the serial number in the sheet
        dataSheet.getRange(i + 1, serialIndex + 1).setValue(invoiceNumCount);
        dataSheet.getRange(i + 1, draftIndex + 1).setValue("FALSE");
        dataSheet.getRange(i + 1, deleteIndex + 1).setValue("FALSE");


        Logger.log('protectRow '+ i );
        //protect row from futher editing
        protectRow(ss,i);


      }
      else if(sheetValues[i][draftIndex]){
        Logger.log('Break END '+ i );
        break;
      }
    }
  } catch (e) {

    // Show the error
    showUiDialog('Finished Invoice Generation', e.message);

  }

//}
}

function createInvoices(dataSheet,sheetValues,rowIndex,docId,invoiceNumCount,folderId,linkText,clmnLinkIndex,clmnIdIndex){
 var invoiceId;
 try{
  Logger.log('createInvoices - getFile and make copy');
  invoiceId = withDriveRetry(function() {
    return DriveApp.getFileById(docId).makeCopy(linkText+"_Template").getId();
  }, 'makeCopy');
  Logger.log('createInvoices - create document');
  var newFileTitle = createDocument(sheetValues,rowIndex,invoiceId,invoiceNumCount,linkText);
  Logger.log('createInvoices - read existing id');
  var existingFileId = dataSheet.getRange(rowIndex + 1, clmnIdIndex + 1).getValue();
  // Convert the Invoice Document into a PDF file
  Logger.log('createInvoices - convert to pdf');
  var pdfInvoice = convertPDF(invoiceId,folderId,newFileTitle,existingFileId);
   // Set the PDF url into the spreadsheet
  Logger.log('createInvoices - add urls');
  dataSheet.getRange(rowIndex + 1, clmnLinkIndex + 1).setValue(createHyperlinkString(pdfInvoice[0],linkText));
  Logger.log('createInvoices - create pdf invoice');
  dataSheet.getRange(rowIndex + 1, clmnIdIndex + 1).setValue(pdfInvoice[1]);
  Logger.log('createInvoices - delete template');
    // Delete the original document (will leave only the PDF)
  safeDeleteFile(invoiceId);
  invoiceId = null;
  return true;
 }catch (error) {
    Logger.log('createInvoices ['+linkText+'] failed at row '+(rowIndex+1)+': '+error.message);
    showUiDialog('Invoice Generation Error: '+linkText, error.message);
    if (invoiceId) {
      safeDeleteFile(invoiceId);
    }
    return false;
  }

}

/**
* Run a Drive operation with retry-and-backoff for transient failures
* like "Service error: Drive" and "Empty response". Re-throws after
* the final attempt so callers still see the error.
* @param {function():*} fn  The Drive call to invoke.
* @param {string} label     Short name for log lines.
* @returns {*} Whatever fn returns.
*/
function withDriveRetry(fn, label) {
  var maxAttempts = 4;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return fn();
    } catch (e) {
      var msg = (e && e.message) || String(e);
      var transient = /Service error|Empty response|Internal error|rate limit|backendError|timed? ?out/i.test(msg);
      Logger.log('withDriveRetry[' + label + '] attempt ' + attempt + ' failed: ' + msg);
      if (attempt === maxAttempts || !transient) {
        throw e;
      }
      Utilities.sleep(1000 * Math.pow(2, attempt - 1)); // 1s, 2s, 4s
    }
  }
}

/**
* Safely trash a Drive file. Avoids the "Empty response" error from
* Drive.Files.remove() by using DriveApp + setTrashed, with retry on
* transient failures. Failure here is non-fatal — the template copy
* will end up in Trash even if the call partially fails.
* @param {string} fileId
*/
function safeDeleteFile(fileId) {
  if (!fileId) return;
  var maxAttempts = 3;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      var file = DriveApp.getFileById(fileId);
      if (file.isTrashed()) return;
      file.setTrashed(true);
      return;
    } catch (e) {
      Logger.log('safeDeleteFile attempt ' + attempt + ' failed for ' + fileId + ': ' + e.message);
      if (attempt === maxAttempts) {
        Logger.log('safeDeleteFile giving up on ' + fileId);
        return;
      }
      Utilities.sleep(500 * attempt);
    }
  }
}

function createDocument(sheetValues,rowIndex,invoiceId,invoiceNumCount,linkText){

        var key, values , invoiceNumber, invoiceDate;

        // Instantiate the document
        var docBody = DocumentApp.openById(invoiceId).getBody();

        // Iterates over the spreadsheet columns to get the values used to write the document
        for (var j = 0; j < sheetValues[rowIndex].length; j++) {

          // Key and Values to be replaced
          key = sheetValues[0][j];
          values = sheetValues[rowIndex][j];

        if (key === "date") {
          // Use Utilities.formatDate with the spreadsheet's timezone to avoid date shifting.
          // Guard against empty cells or string-formatted dates so a single bad row
          // doesn't abort the whole batch.
          if (values instanceof Date) {
            var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
            invoiceDate = Utilities.formatDate(values, timezone, "d/M/yyyy");
          } else {
            invoiceDate = values ? String(values) : '';
          }
          replace('%date%', invoiceDate, docBody);
        }else if (values) {

            // Everything else appart from date values
             if (key === "mawb_no" ||
             key === "hawb_no" ||
             key === "no_pieces" ||
             key === "discharge_txt" ||
             key ==="natural"||
             key ==="decimal"||
              key ==="cash"||
              key ==="rout" ||
              key ==="weight"||
              key ==="gross_weight"
             ) {
                  replace('%' + key + '%', values, docBody);
            } else {
              if (!isNaN(parseFloat(values)) && isFinite(values)) {
                  replace('%' + key + '%',  financial(values), docBody); // Replace values
              } else{
              replace('%' + key + '%', values, docBody);
            }
            }
          }

           else {
            replace('%' + key + '%', '', docBody); // Replace empty string
          }
        }

         var newFileTitle;
        // Format invoice name pdf
        invoiceNumber = String(invoiceNumCount).padStart(7, '0') + "/" + invoiceDate.split("/")[2];
        replace('%number%', invoiceNumber, docBody);
        newFileTitle = invoiceNumber+" - "+linkText;
        // Rename the invoice document
        DocumentApp.openById(invoiceId).setName(newFileTitle).saveAndClose();

return newFileTitle;

}

function protectRow(ss,rowIndex){
       Logger.log('protectRow - start');
       var rangeString = "Data!"+(rowIndex+1)+":"+(rowIndex+1);
        // Protect range....
        var range = ss.getRange(rangeString);
        var protectionDescription = 'INVOICE CREATED ROW '+(rowIndex+1);
        Logger.log('protectRow - removeProtections');
        removeProtections(ss,protectionDescription);
        Logger.log('protectRow - range.protect');
        var protection = range.protect().setDescription(protectionDescription);

        // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
        // permission comes from a group, the script throws an exception upon removing the group.
        Logger.log('protectRow - removeEditors');
        protection.removeEditors(protection.getEditors());
        protection.addEditor("anis@sli-eg.com");
        protection.addEditor("george@sli-eg.com");
        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }

}

function removeProtections(ss,protectionDescription)
{
 Logger.log('removeProtections - get protections');

  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
   Logger.log('removeProtections - Loop');

for (var i = 0; i < protections.length; i++) {
  if (protections[i].getDescription() == protectionDescription) {
     protections[i].remove();
    }
  }
}
/**
* Move a file from one folder into another
* @param {Object} file A file object in Google Drive
* @param {Object} dest_folder A folder object in Google Drive
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
* Convert a Google Docs into a PDF file
* @param {string} id - File Id
* @returns {*[]}
*/
function convertPDF(templateGdocId,invFolder,fileName,existingFileID) {
  var docBlob = DocumentApp.openById(templateGdocId).getAs('application/pdf');
  docBlob.setName(fileName); // Add the PDF extension
 // for(var i = 0; i < file.getEditors().length;i++){
 //     file.revokePermissions(file.getEditors()[i]);
 // }
  //file.addViewer("nevineezzat@sli-eg.com")



 var currentFile, newFile, url,id;

 if(existingFileID){
  currentFile = DriveApp.getFileById(existingFileID);
  }
  if (currentFile) {//If there is a truthy value for the current file
    withDriveRetry(function() {
      return Drive.Files.update({
        title: fileName, mimeType: currentFile.getMimeType()
      }, currentFile.getId(), docBlob);
    }, 'Drive.Files.update');
    // Drive.Files.update() returns a Drive API resource (field .id, no .getId() method),
    // so we read id/url from currentFile to stay consistent with the create branch.
    id = currentFile.getId();
    url = currentFile.getUrl();
  }else{
    newFile = withDriveRetry(function() {
      return DriveApp.getFolderById(invFolder).createFile(docBlob);
    }, 'createFile');
    id = newFile.getId();
    url = newFile.getUrl();
  }

  return [url, id];
}

/**
* Replace the document key/value
* @param {String} key - The document key to be replaced
* @param {String} text - The document text to be inserted
* @param {Body} body - the active document's Body.
* @returns {Element}
*/
function replace(key, text, body) {
  return body.editAsText().replaceText(key, text);
}


/**
* Loads the showDialog
*/
function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('iframe.html')
  .setWidth(200)
  .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Creating Solution..');
}


function createHyperlinkString(link,text){
  return `=HYPERLINK("${link}", "${text}")`;
}
/**
* Show an UI dialog
* @param {string} title - Dialog title
* @param {string} message - Dialog message
*/
function showUiDialog(title, message) {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (e) {
    // pass
  }
}

function financial(x) {
  return Number.parseFloat(x).toFixed(2);
}

function removeSpecificProtections() {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get all sheets in the spreadsheet
  var sheets = ss.getSheets();

  // Loop through each sheet
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];

    // Get all protections on the current sheet
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

    // Loop through each protection
    for (var j = 0; j < protections.length; j++) {
      var protection = protections[j];

      // Check if the protection's description matches the pattern
      var description = protection.getDescription();
      if (description && description.match(/^INVOICE CREATED ROW \d+$/)) {
        protection.remove();
      }
    }
  }
}
function init() {
  // Get the active spreadsheet and its sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS.sheetSettings);

  // Check if the settings sheet exists
  if (!settingsSheet) {
    throw new Error('Settings sheet not found. Please ensure the sheet is named correctly.');
  }

  // Get the current year
  const currentYear = new Date().getFullYear();

  // Get existing folder IDs from the settings sheet
  const settingsSheetValues = settingsSheet.getDataRange().getValues();
  // SETTINGS.col.Original_ID is B2 → 0-indexed [1][1]
  const existingOriginalDocumentId = settingsSheetValues[1][1];

  // Get the parent folder of the existing original document
  const originalDocument = DriveApp.getFileById(existingOriginalDocumentId);
  const originalDocumentParent = originalDocument.getParents().next(); // Get the immediate parent folder

  // Get the parent folder of the original document's parent (two levels up)
  const pdfFolder = originalDocumentParent.getParents().next();

  // Check if the folder for the current year exists
  let newYearFolder;
  const yearFolders = pdfFolder.getFoldersByName(currentYear.toString());
  if (yearFolders.hasNext()) {
    newYearFolder = yearFolders.next();
  } else {
    newYearFolder = pdfFolder.createFolder(currentYear.toString());
  }

  // Create "Original" and "Copy" subfolders within the year folder
  const originalFolderId = newYearFolder.createFolder("Original").getId();
  const copyFolderId = newYearFolder.createFolder("Copy").getId();

  // Update the settings sheet with the new folder IDs
  settingsSheet.getRange(SETTINGS.col.Original_Folder_ID).setValue(originalFolderId);
  settingsSheet.getRange(SETTINGS.col.Copy_Folder_ID).setValue(copyFolderId);

  // Show a confirmation dialog
  SpreadsheetApp.getUi().alert('Folders created and settings updated successfully!');
}