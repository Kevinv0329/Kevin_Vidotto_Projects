//Purpose: Add correct ACL access for domain and all staff group on Kinder version
//
//Prerequisites: Need Drive API V2 Installed on AppScript (found in Services)
//
//Process: Redacted
//

function PDFShareLinkandEditor() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WSLP-PDF Spring');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var docId = data[i][138];
    var editorEmail = data[i][137];
    var processed = data[i][136]; 
    
    if (docId !== "" && editorEmail !== "" && processed === "") {
      try {
        // Set view access to everyone in your Google Workspace domain with the link
        var file = DriveApp.getFileById(docId);
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);

        // Add an editor without sending notification emails
        Drive.Permissions.insert(
          {
            'role': 'writer',
            'type': 'user',
            'value': editorEmail
          },
          docId,
          {
            'sendNotificationEmails': false
          }
        );

        sheet.getRange(i + 1, 137).setValue("Y"); // i + 1 because the loop is 0-based, but sheets are 1-based

        Logger.log('PDF ID ' + docId + ' shared within domain and editor ' + editorEmail + ' added without sending notification.');

      } catch (e) {
        Logger.log('Error processing PDF ID ' + docId + ': ' + e.message);
      }
    } 
  }
}


function GDocShareLinkandEditor() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WSLP-GDoc Spring');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var docId = data[i][138];
    var editorEmail = data[i][137];
    var processed = data[i][136]; 
    
    if (docId !== "" && editorEmail !== "" && processed === "") {
      try {
        var file = DriveApp.getFileById(docId);
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);

        // Add an editor without sending notification emails
        Drive.Permissions.insert(
          {
            'role': 'writer',
            'type': 'user',
            'value': editorEmail
          },
          docId,
          {
            'sendNotificationEmails': false
          }
        );

        sheet.getRange(i + 1, 137).setValue("Y"); // i + 1 because the loop is 0-based, but sheets are 1-based

        Logger.log('GDoc ID ' + docId + ' shared within domain and editor ' + editorEmail + ' added without sending notification.');

      } catch (e) {
        Logger.log('Error processing GDoc ID ' + docId + ': ' + e.message);
      }
    }
  }
}


function GetIndices() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  for (var i = 0; i < headerRow.length; i++) {
    var columnName = headerRow[i];
    var columnIndex = i; 
    Logger.log("Column Name: " + columnName + ", Index: " + columnIndex);
  }
}