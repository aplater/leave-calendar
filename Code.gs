var PROCESSED = "Processed";
var ENTER = "Enter a new leave record into the Calendar";
var DELETE = "Delete a previously entered Leave from the Calendar"

// modify SPREADSHEET_ID to refer to new spreadsheet
var SPREADSHEET_ID = "1JiyR3mgaekPijSsW8UGk-HJayYLBPkesOqKVYv3QD2g"
var BASE_SPREADSHEET = SpreadsheetApp.openById(SPREADSHEET_ID);
var LEAVE_SHEET = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("records");
var EVENT_SHEET = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("events");

// modify CALENDAR_ID to refer to new calendar
var CALENDAR_ID = "ocvqlkmnnadl5mfpauuk2v34mg@group.calendar.google.com"
var cal = CalendarApp.getCalendarById(CALENDAR_ID);

//This is the row where the data starts (2 since there is a header row)
var START_ROW = 2; 


function onSubmit() {
  var dataRange = LEAVE_SHEET.getRange(2, 1, LEAVE_SHEET.getMaxRows() + 1, 11);
 
  // Create one JavaScript object per row of data.
  objects = getRowsData(LEAVE_SHEET, dataRange, 1);
  
  // For every leave record, check if it needs to go the calendar
  for (var i=0; i<objects.length; i++) {
     var row = objects[i];
   
    //  Browser.msgBox("value of row.added = " + row.added );
    if (row.processed != PROCESSED) {
      if(row.youWantTo == ENTER){
        enterCalendar(row);
        LEAVE_SHEET.getRange(START_ROW + i, 9).setValue(PROCESSED);
        }
     
      if(row.youWantTo == DELETE){
        deleteCalendar(row);
        LEAVE_SHEET.getRange(START_ROW + i, 9).setValue(PROCESSED);
       }
      }
    }
  }


// enterCalendar reads the data for a given row object,
// creates an All Day Event given the start date and end date specified in the sheet.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - row: row data object
function enterCalendar(row) {
  var start_date = new Date(row.startDateOfYourNewLeaveRecord);
  var no_of_days = calculateDateDiff(row.startDateOfYourNewLeaveRecord,row.endDateOfYourNewLeaveRecord) + 1;
  
  for (var i=0; i<no_of_days; i++){
    var event = cal.createAllDayEvent(row.whoAreYou + " - " + row.typeOfLeave,start_date, {description:row.typeOfLeave});
    var eventid = event.getId();
    var lastRow = EVENT_SHEET.getLastRow();
    var values = [[ row.timestamp, row.whoAreYou,start_date,eventid]];

    EVENT_SHEET.getRange(lastRow + 1, 1, 1, 4).setValues(values); 
    start_date.setDate(start_date.getDate() + 1);
   }
}


// deleteCalendar reads the data for a given row object,
// iterates and deletes the Event Series using the Event ID
// by referencing from the given the start date and end date specified in the sheet.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - row: row data object
function deleteCalendar(row) {
 // delete from calendar
 var end_date = new Date(row.endDateOfLeaveYouWishToDelete);  
 var no_of_days = calculateDateDiff(row.startDateOfLeaveYouWishToDelete,row.endDateOfLeaveYouWishToDelete) + 1;
 var dataRange_events = EVENT_SHEET.getRange(2, 1, EVENT_SHEET.getMaxRows() + 1, 4);

 // create one JavaScript object per row of data.
 objects_events = getRowsData(EVENT_SHEET, dataRange_events, 1);
 
 for(var k=0; k<no_of_days; k++) 
  { 
    for (var j=0; j<objects_events.length; j++) { 
      var row_events = objects_events[j];
      if(row.whoAreYou == row_events.whoAreYou && end_date.getTime() == row_events.startDateOfYourNewLeaveRecord.getTime()) {
        var event = cal.getEventSeriesById(row_events.eventId);
        event.deleteEventSeries();

        var row_no = getRowNo(row_events.eventId);
        EVENT_SHEET.deleteRow(row_no);
        } 
      }
      end_date.setDate(end_date.getDate() - 1);
     }
}


// calculateDateDiff gets the difference in number of days given 2 dates
// Arguments:
//   - date1/date2: Date object
function calculateDateDiff(date1,date2) {
  var start_date = new Date(date1);
  var end_date = new Date(date2);
  var timeDiff = Math.abs(end_date.getTime() - start_date.getTime());
  var no_of_days = Math.ceil(timeDiff / (1000 * 3600 * 24));

  return no_of_days;
}


// getRowNo returns the row_no given eventID
// Arguments:
//   - eventID: string
function getRowNo(eventId)
{
var range = EVENT_SHEET.getRange("D1:D"); 
var values = range.getValues();

// examine the values in the array
var i = []; 
for (var y=0; y<values.length; y++) {
   if(values[y] == eventId){
      i.push(y);
      }
  }
var row_no = Number(i)+Number(range.getRow());

return row_no;
}


// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];

  return getObjects(range.getValues(), normalizeHeaders(headers));
}


// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i=0; i<data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j=0; j<data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}


// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}


// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}


// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}


// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}


// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}
