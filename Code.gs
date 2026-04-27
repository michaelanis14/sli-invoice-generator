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
function testUI() {
  SpreadsheetApp.getUi().alert("Hello! The UI is connected.");
}

SETTINGS = {

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
    Puplic_ID:"B10",
    Puplic_Folder_ID:"B11",
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
  .addItem('Generate Drafts', 'createDraft')
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
    var systemCreated = sheetSettings.getRange(SETTINGS.col.systemCreated);
    if (!systemCreated.getValue()){
      systemCreated.setValue('True');
    } else {
      showUiDialog('Warnning','Solution has already been created!');
      return;
    }

    // Checks if cell Count exists
    var count = sheetSettings.getRange(SETTINGS.col.count);
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

    // Set tab settings document ID
    sheetSettings.getRange(SETTINGS.col.templateId).setValue(docCopy.getId());

    // Move an copy for invoice folder
    moveFile(docCopy, invoiceFolder);

    // Set folder ID
    sheetSettings.getRange(SETTINGS.col.folderId).setValue(folder.getId());


    // End process
    showUiDialog('Success', 'Your solution is ready');

    return true;
  } catch (e) {

    // Show the error
    showUiDialog('Something went wrong', e.message)

  }
}


/**
* Reads the spreadsheet data and creates the PDF invoice
*/
function sendInvoice() {
    Logger.log('144: Script Started');

  const EDITOR = "anis@sli-eg.com";
//var response = SpreadsheetApp.getUi().alert('Do you want to generate document?', SpreadsheetApp.getUi().ButtonSet.YES_NO);    // Opens the spreadsheet and access the tab containing the data
  //if (response == SpreadsheetApp.getUi().Button.YES) {
  try {

    Logger.log('151: YES to start script');


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
    var control = settingsSheetValues[11][1]
    if (!control){
      showUiDialog('Warnning','Run "Install Solution" in tab Instructions');
      return;
    }

    // Duplicate teh template on Google Drive to manipulate the data
     counter = settingsSheetValues[0][1]


    //var originalsDocument = sheetSettings.getRange(SETTINGS.col.Original_ID).getValue();
    var originalsDocument =  settingsSheetValues[1][1]+""
  //SpreadsheetApp.getUi().alert('indexOf method on a string in google app originalsDocument '+originalsDocument)
    //var originalsFolder = sheetSettings.getRange(SETTINGS.col.Original_Folder_ID).getValue();
    var originalsFolder =  settingsSheetValues[2][1]+""

    //var copyDocument = sheetSettings.getRange(SETTINGS.col.Copy_ID).getValue();
    var copyDocument =  settingsSheetValues[5][1]+""

   // var copyFolder = sheetSettings.getRange(SETTINGS.col.Copy_Folder_ID).getValue();
    var copyFolder =  settingsSheetValues[6][1]+""

    Logger.log('210: Loaded all data');

    for (var i = 1; i < sheetValues.length; i++) {


    Logger.log('215: Loop for sheetValues'+ i +'of '+ sheetValues.length);

      // Creates the Invoice
      if (!sheetValues[i][pdfIndex] && !sheetValues[i][draftIndex] && !sheetValues[i][deleteIndex]) {
        Logger.log('220: Creating Invoice'+ i );

    //    var response = SpreadsheetApp.getUi().alert('Generate Invoice '+(counter + 1) + ' ?', SpreadsheetApp.getUi().ButtonSet.YES_NO);    // Opens the spreadsheet and access the tab containing the data
    //     if (response == SpreadsheetApp.getUi().Button.NO) continue;

       // Get last invoice count from the tab 'Count'
       // counter = sheetSettings.getRange(SETTINGS.col.Count);
      // if(!sheetValues[i+1][serialIndex+1])
       // SpreadsheetApp.getUi().alert(sheetValues[i][serialIndex] +' & '+dataSheet.getRange(i + 1, serialIndex + 1).getValue());

      //  SpreadsheetApp.getUi().alert((sheetValues[i+1][serialIndex]?.trim()?.length || 0) === 0);


       if((sheetValues[i][serialIndex]?.toString().trim()?.length || 0) === 0){ // avoid incremental the counter
        Logger.log('233: NO Serial'+ i );

            invoiceNumCount = counter + 1;
            sheetSettings.getRange(SETTINGS.col.Count).setValue(invoiceNumCount);
            counter = invoiceNumCount

        } else invoiceNumCount = sheetValues[i][serialIndex];

        Logger.log('242: CreateOriginal'+ i );
        //OriginalInvoice
        createInvoices(dataSheet,sheetValues,i,originalsDocument,invoiceNumCount,originalsFolder,"Original",pdfIndex,pdf_ID_Index)

        Logger.log('246: CreateCopy'+ i );
        //CopyInvoice
        createInvoices(dataSheet,sheetValues,i,copyDocument,invoiceNumCount,copyFolder,"Copy",copyUrlIdx,copyIdIdx)

        Logger.log('250: SetSerial'+ i );
        //set the serial number in the sheet
        dataSheet.getRange(i + 1, serialIndex + 1).setValue(invoiceNumCount);
        dataSheet.getRange(i + 1, draftIndex + 1).setValue("FALSE");
        dataSheet.getRange(i + 1, deleteIndex + 1).setValue("FALSE");





        Logger.log('260: protectRow'+ i );
        //protect row from futher editing
        protectRow(ss,i);


      }
      else if(sheetValues[i][draftIndex]){
        Logger.log('267: Break END'+ i );
        break;
      }
    }
  } catch (e) {

    // Show the error
    showUiDialog('Finished Invoice Generation', e.message)

  }

//}
}

function createInvoices(dataSheet,sheetValues,rowIndex,docId,invoiceNumCount,folderId,linkText,clmnLinkIndex,clmnIdIndex){
 var invoiceId;
 try{
  Logger.log('282: createInvoices - getFile and make copy' );
  invoiceId = DriveApp.getFileById(docId).makeCopy(linkText+"_Template").getId();
  Logger.log('284: createInvoices - create document' );
  var newFileTitle = createDocument(sheetValues,rowIndex,invoiceId,invoiceNumCount,linkText);
  Logger.log('286: read existing id' );
  var existingFileId = dataSheet.getRange(rowIndex + 1, clmnIdIndex + 1).getValue();
  // Convert the Invoice Document into a PDF file
  Logger.log('289: createInvoices - convert to pdf' );
  var pdfInvoice = convertPDF(invoiceId,folderId,newFileTitle,existingFileId);
   // Set the PDF url into the spreadsheet
   Logger.log('292: createInvoices - add urls' );
  dataSheet.getRange(rowIndex + 1, clmnLinkIndex + 1).setValue(createHyperlinkString(pdfInvoice[0],linkText));
  Logger.log('294: createInvoices - create pdf invoice' );
  dataSheet.getRange(rowIndex + 1, clmnIdIndex + 1).setValue(pdfInvoice[1]);
  Logger.log('296: createInvoices - delete templete' );
    // Delete the original document (will leave only the PDF)
  safeDeleteFile(invoiceId);
  invoiceId = null;
 }catch (error) {
    showUiDialog('Finished Invoice Generation CI',error.message);
    if (invoiceId) {
      safeDeleteFile(invoiceId);
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
          // Use Utilities.formatDate with the spreadsheet's timezone to avoid date shifting
          var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
          invoiceDate = Utilities.formatDate(values, timezone, "d/M/yyyy");
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
                  replace('%' + key + '%', values, docBody)
            } else {
              if (!isNaN(parseFloat(values)) && isFinite(values)) {
                  replace('%' + key + '%',  financial(values), docBody); // Replace values
              } else{
              replace('%' + key + '%', values, docBody)
            }
            }
          }

           else {
            replace('%' + key + '%', '', docBody) // Replace empty string
          }
        }

         var newFileTitle;
        // Format invoice name pdf
        invoiceNumber = invoiceNumCount.padLeft(7, '0') + "/" + invoiceDate.split("/")[2];
        replace('%number%', invoiceNumber, docBody);
        newFileTitle = invoiceNumber+" - "+linkText;
        // Rename the invoice document
        DocumentApp.openById(invoiceId).setName(newFileTitle).saveAndClose();

return newFileTitle;

}

function protectRow(ss,rowIndex){
       Logger.log('372: protectRow - start' );
       var rangeString = "Data!"+(rowIndex+1)+":"+(rowIndex+1);
        // Protect range....
        var range = ss.getRange(rangeString);
        var protectionDescription = 'INVOICE CREATED ROW '+(rowIndex+1);
        Logger.log('377: removeProtection - removeProtections' );
        removeProtections(ss,protectionDescription);
        Logger.log('377: protect - range.protect' );
        var protection = range.protect().setDescription(protectionDescription);

        // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
        // permission comes from a group, the script throws an exception upon removing the group.
        Logger.log('384: protection - removeEditors' );
        protection.removeEditors(protection.getEditors());
        protection.addEditor("anis@sli-eg.com");
        protection.addEditor("george@sli-eg.com");
        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }

}

function removeProtections(ss,protectionDescription)
{
 Logger.log('396: removeProtections - removeProtections get protections' );

  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
   Logger.log('399: removeProtections - Loop' );

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
    dest_folder.addFolder(file)
  } else {
    dest_folder.addFile(file);
  }
  var parents = file.getParents();
  while (parents.hasNext()) {
    var folder = parents.next();
    if (folder.getId() != dest_folder.getId()) {
      if (isFolder === true) {
        folder.removeFolder(file)
      } else {
        folder.removeFile(file)
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
    newFile = Drive.Files.update({
      title: fileName, mimeType: currentFile.getMimeType()
    }, currentFile.getId(), docBlob);
    url = DriveApp.getFileById(existingFileID).getUrl();

  }else{
    newFile = DriveApp.getFolderById(invFolder).createFile(docBlob);
    url = newFile.getUrl();

  }

     id = newFile.getId();

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
* Returns a new string that right-aligns the characters in this instance by padding them with any string on the left,
* for a specified total length.
* @param {Number} n - Number of characters to pad
* @param {String} str - The string to be padded
* @returns {string}
*/
Number.prototype.padLeft = function (n, str) {
  return Array(n - String(this).length + 1).join(str || '0') + this;
};

/**
* Loads the showDialog
*/
function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('iframe.html')
  .setWidth(200)
  .setHeight(150)
  SpreadsheetApp.getUi().showModalDialog(html, 'Creating Solution..')
}


function createHyperlinkString(link,text){
  return value = `=HYPERLINK("${link}", "${text}")`;
}
/**
* Show an UI dialog
* @param {string} title - Dialog title
* @param {string} message - Dialog message
*/
function showUiDialog(title, message) {
  try {
    var ui = SpreadsheetApp.getUi()
    ui.alert(title, message, ui.ButtonSet.OK)
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
  const existingOriginalDocumentId = settingsSheetValues[2][1]; // Assuming Original_ID is in row 2, column 2

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