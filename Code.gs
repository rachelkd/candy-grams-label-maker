/*
  Create Candy Gram Labels using this script in a Google Doc
  1. Edit URL_SS to the Google Sheet containing all data
  2. Edit the total rows (including header row) for each school. Make sure your first data point is in the 2nd row
  3. Edit sheet names within spreadsheet for each school. Sometimes this is 'Sheet1,' 'Sheet2,' etc.
  
  IMPORTANT! Google Sheet Requirements:
  First column: Sender name,
  Second column: Recipient name,
  Third: Recipient grade
  Fourth: Message
*/
APP_TITLE = "Candy Grams Label Maker";
URL_SS = "https://docs.google.com/spreadsheets/d/1bDnvK_JaRu8guYf2VnaHFJK6fV9xfP8fcuZwNk4p__A/edit#gid=1153281473";
SHEET_CHS = "Crofton";
SHEET_SGS = "Saints";
SHEET_YHS = "York";
SHEET_VC = "VC"
ENTRIES_CHS = 212;
ENTRIES_SGS = 0;
ENTRIES_YHS = 0;
ENTRIES_VC = 0;
START_ENTRY_CHS = 192;
START_ENTRY_SGS = 63;
START_ENTRY_YHS = 2;
START_ENTRY_VC = 24;

function run() {
  var body = DocumentApp.getActiveDocument().getBody();
  
  // VC
  tagMaker(getData(SHEET_VC, ENTRIES_VC,START_ENTRY_VC));
  // York
  tagMaker(getData(SHEET_YHS, ENTRIES_YHS,START_ENTRY_YHS));
  // Saints
  tagMaker(getData(SHEET_SGS, ENTRIES_SGS,START_ENTRY_SGS));
  // CHS
  tagMaker(getData(SHEET_CHS, ENTRIES_CHS,START_ENTRY_CHS));
}

function tagMaker(cells) {
  
  var body = DocumentApp.getActiveDocument().getBody();

  // Styles
  var style = {};
  style[DocumentApp.Attribute.FONT_SIZE] = '11';
  style[DocumentApp.Attribute.FONT_FAMILY] = 'IM Fell DW Pica';

  // Build a table from the cells.
  var table1 = body.insertTable(0, cells);
  
  // Set Background Colour of every cell to be orange
  for(var r = 0; r < cells.length; r++) {
    for(var c = 0; c < 2; c++) {
      table1.getCell(r,c).setBackgroundColor('#fcd292');
    }
  }
  // Set styles
  table1.setAttributes(style);
}

function getData(sheetName, sheetRows, startEntry) {
  console.log(`Application setup for: ${APP_TITLE}`)
  // Open Spreadsheet
  var ss = SpreadsheetApp.openByUrl(URL_SS);
  // Set sheet
  var sheet = ss.getSheetByName(sheetName);
  console.log("Sheet ID: " + sheet.getSheetId());
  // Declare cell variables for Doc table
  var cellsRow = [];
  var cells = [];
  
  // Create content for each cell
  for(var r = startEntry; r <= sheetRows; r++) {
    // Table Content
    var tagContent = "Halloween Candy Grams!\n\nFrom: ";
    // Append all pieces of info for tag together
    for(var c = 1; c <= 4; c++) {
      var range = sheet.getRange(r,c); 
      var data = range.getValue();
      if(c == 2) {
        tagContent += "To: ";
      } else if (c == 3) {
        tagContent += "Grade: ";
      }
      // If name empty, set name to be anonymous
      if(c == 1 && data == "") {
        tagContent += "Anonymous\n";
      } else {
        tagContent += data + "\n";
      }

    }
      // Push tagContent to cellsRow
      // if cellsRow has < 2 elements
      if(cellsRow.length < 2) {
        cellsRow.push(tagContent);
      } else {
        cells.push(cellsRow);
        cellsRow = [];
        cellsRow.push(tagContent);
      }
      // Check to push last iteration
      if(r == sheetRows) {
        while(cellsRow.length < 2) {
          cellsRow.push("");
        }
        cells.push(cellsRow);
    }

    // Testing parameters
    // console.log(tagContent);
    // console.log(r + " " + cellsRow);
  }

  // console.log(cells);
  return cells;
}
