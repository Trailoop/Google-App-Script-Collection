/*************************************************************************************************************************************
                                                           Support Form Email
                                                         Google App Script (GAS)
                                                        Created by Scott Lawrence
                                                 Solution Engineer - Business Intelligence
                                                   Contact: scott.lawrence@ahold.com
 
                                                     Scheduling GAS - on form submit  
                                                     Function - sendAcknowledgement
                                  
 ======================================================================================================================================*/



//Global Scope Varibles

/* Identifies the first none form filled column 
This column is used to control the flow of triggering 
the email to the support group by sendAcknowledgement
*/
var ACKNOWLEDGRMRNTSENT = 12;  //Column that identifies the first non-form filled column



/*
!----------------Triggered by a form submission-------------------!
Each time a new form response is recorded to the spreadsheet this function is executed to email
support group with the details from the submission
*/
function sendAcknowledgement() {
  //Get DataSource from responses
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Form Responses 1");
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, dataSheet.getMaxColumns());

  
  var templateSheet = ss.getSheetByName("email"); //sheet that stores templated data
  var emailSubject = templateSheet.getRange("B2").getValue();  // email subject
  var emailTemplate = templateSheet.getRange("B3").getValue(); //body of email fill form
  var emailAddress = "ausa.microstrategysupport.group@ahold.com"; //to field for email
 
  // Create one JavaScript object per row of data. -each form response
  //properties name equal row 1 values
  //i.e objects[0].emailaddress
  var objects = getRowsData(dataSheet, dataRange);
  
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];
    rowData.rowNumber = i + 2;
    
    //flow control ensures only new form submissions are emailed
    if ( !rowData.acknowledgementSent ) {
    
   //Logger.log(rowData);      
      // Generate a personalized email.
        // Given a template string, replace markers (for instance ${"First Name"}) with
        // the corresponding value in a row object (for instance rowData.firstName).
        var emailText = fillInTemplateFromObject(emailTemplate, rowData);
      var emailSubject = "MicroStrategy Support Request # " + rowData.rowNumber;

        MailApp.sendEmail(emailAddress, emailSubject, emailText, {
            htmlBody: emailText,
            replyTo: '123@email.com',
            name:'Support Form for MSTR'

          });      
      }
      //Marks submission as sent to prevent future acknowledgements
      dataSheet.getRange(rowData.rowNumber,ACKNOWLEDGRMRNTSENT).setValue("Yes"); //12 is the column number for 'Action'
   }
}



// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = data[normalizeHeader(templateVars[i])];
    email = email.replace(templateVars[i], variableData || "");
  }

  return email;
}





//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////

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
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
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


function include(arr,obj) {
    return (arr.indexOf(obj) != -1);
}

function tz () { 

Logger.log(Session.getTimeZone());

var d = new Date()
var n = d.getTimezoneOffset();

Logger.log(d + " " +n);

}
