// in html code
google.script.run.withSuccessHandler(displayData).getDataFromSheet();

// Html file with sheet
function doGet(){
  return HtmlService.createHtmlOutputFromFile('index').setSandboxMode(HtmlService.SandboxMode.IFRAME)
}

// To get all the data from the sheet
function doGet(){
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  let velues = sheet1.getDataRange().getValues();
  return ContentService.createTextOutput(JSON.stringify(velues)).setMimeType(ContentService.MimeType.JSON)
}

// To get only one value from the sheet
function doGet(){
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  let velues = sheet1.getDataRange().getValues();
  let oneValue = velues[1][2];
  return ContentService.createTextOutput(JSON.stringify(oneValue)).setMimeType(ContentService.MimeType.JSON)
}

// To get only one value from the sheet by row and column
function doGet(){
  let rowNumber =1;
  let columnNumber =1;
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  let title = sheet1.getRange(rowNumber, columnNumber,1).getValue();
  return ContentService.createTextOutput(JSON.stringify(title)).setMimeType(ContentService.MimeType.JSON);
}

// To get limited values from the sheet
function doGet(){
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  let range = sheet1.getRange("B2:B111");
  let values = range.getValues();
  return ContentService.createTextOutput(JSON.stringify(values)).setMimeType(ContentService.MimeType.JSON)
}

// To get only one row
function doGet(){
  let rowNumber =1;
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  let values = sheet1.getRange(rowNumber, 1, 1, sheet1.getLastColumn()).getValues()[0]; 
  return ContentService.createTextOutput(JSON.stringify(values)).setMimeType(ContentService.MimeType.JSON);
}

// To get one row , skip one and take one
function doGet(){
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  let numRows = sheet1.getLastRow();
  let values = [];

  for (let rowNumber = 1; rowNumber <= numRows; rowNumber += 2) {
    let rowValues = sheet1.getRange(rowNumber, 1, 1, sheet1.getLastColumn()).getValues()[0];
    values.push(rowValues);
  }

  return ContentService.createTextOutput(JSON.stringify(values)).setMimeType(ContentService.MimeType.JSON);
}


