/*
 * script to export data in all sheets in the current spreadsheet as individual csv files
 * files will be named according to the name of the sheet
 * Author: Tuur Dutoit
 * Original author: Michael Derazon
*/

function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem("Export whole spreadsheet as CSV", "openSidebar");
  menu.addToUi();
}

function openSidebar() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("download"));
}

function saveAsCSV(details) {
  Logger.log("Saving");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var csvBlobs = [];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var fileName = sheet.getName() + ".csv";
    var csv = convertRangeToCsv(sheet, details.separator, details.includeSepLine);
    var csvBlob = Utilities.newBlob(csv, "text/csv", fileName);
    csvBlobs.push(csvBlob);
  }
  
  var zip = Utilities.zip(csvBlobs, details.fileName);
  var zipFile = DriveApp.createFile(zip);
  var downloadUrl = zipFile.getDownloadUrl().slice(0, -8); // remove gd=true that somehow makes download fail
  details.downloadUrl = downloadUrl;
  details.zipId = zipFile.getId();
  
  return details;
}

function removeZip(zipId) {
  var zipFile = DriveApp.getFileById(zipId);
  DriveApp.removeFile(zipFile);
}

function convertRangeToCsv(sheet, separator, includeSepLine) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  var data = activeRange.getValues();
  var csv = includeSepLine ? "sep=" + separator + "\r\n" : "";
  
  // loop through the data in the range and build a string with the csv data
  for (var row = 0; row < data.length; row++) {
    for (var col = 0; col < data[row].length; col++) {
      var val = data[row][col].toString();
      
      // Put values containing the separator in double quotes
      // and escape double quotes
      if (val.indexOf(separator) != -1) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }
      
      data[row][col] = val;
    }
    
    // join each row's columns
    // add a carriage return to end of each row, except for the last one
    if (row < data.length - 1) {
      csv += data[row].join(separator) + "\r\n";
    }
    else {
      csv += data[row];
    }
  }
  
  return csv;
}
