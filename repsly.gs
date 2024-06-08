
Conversation opened. 3 messages. 1 message unread.

Skip to content
Using Gmail with screen readers
1 of 10,772
Re: [EXTERNAL EMAIL]Fwd:
Inbox

Thomas Lloyd
2:43 PM (7 minutes ago)
if (typeof name === 'string' && isNaN(name) && name !== 'Team' && name !== 'Week' && name !== 'Trading Hours' && name !== 'ASM' && name !== 'Supervisor' && name

Thomas Lloyd
2:44 PM (6 minutes ago)
if (typeof name === 'string' && isNaN(name) && name !== 'Team' && name !== 'Week' && name !== 'Trading Hours' && name !== 'ASM' && name !== 'Supervisor' && name

Thomas Lloyd
Attachments
2:50 PM (0 minutes ago)
to me


 One attachment
  •  Scanned by Gmail
/**
Bismillah
 * Repsly Automation Project
 * Author: Joseph Hayes
 * 
 * This script automates the transformation of a rota into a structured table format
 * in Google Sheets using Google Apps Script. It processes data from an existing sheet,
 * extracts relevant information, and generates a new sheet with a standardized table.
 * 
 * Functions:
 * - generateTable(): Main function to generate the table.
 * - updateSchedule(): Converts employee names to representative IDs.
 * - getEmployeeNames(): Retrieves employee names from a specified sheet.
 * - holidays(): Placeholder for custom holiday handling logic.
 */


function generateTable() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = spreadsheet.getSheetByName("repslyWeek");

  // Specify the sheet name for the new table
  var newSheetName = "Generated Table";

  // Check if the sheet already exists
  var existingSheet = spreadsheet.getSheetByName(newSheetName);

  // If the sheet exists, delete it
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }

  // Create a new sheet for the table
  var newSheet = spreadsheet.insertSheet(newSheetName);

  // Set header row for the new sheet
  newSheet.appendRow(["Place ID", "Representative ID", "From date", "Repeat every (x) weeks", "Visit time", "Duration (in minutes)", "Task type", "Task Description"]);

  var employeeNames = getEmployeeNames();

  // Loop through the data in the original sheet
  var data = originalSheet.getRange(3, 1, originalSheet.getLastRow() - 2, originalSheet.getLastColumn()).getValues();
  var durationCol = 4;
  var visitCol = 2;
  var startCol = "C";
  var endCol = "E";

  while (durationCol < 23) {
    for (var i = 0; i < data.length; i++) {
      var rowData = data[i];

      // Extract relevant information from the original data
      var representativeID = rowData[0];

      //employees at the company can be edited and changed by changing employees
      if (employeeNames.includes(representativeID)) {
        // Call the readMergedCellData function to get the data from the merged cell
        var rowNum = 4;
        var range = startCol + rowNum + ":" + endCol + rowNum;
        var mergedRange = originalSheet.getRange(range);

        var fromDate = mergedRange.getCell(1, 1).getValue();
        var visitTime = Utilities.formatDate(new Date(rowData[visitCol]), spreadsheet.getSpreadsheetTimeZone(), "HH:mm");


        var cellBelowValue = originalSheet.getRange(rowNum + 1, visitCol -1).getValue();

        if (cellBelowValue === "hol") {
          Logger.log('Works.');
        }

        Logger.log(cellBelowValue);

       

        if (parseFloat(rowData[durationCol]) >= 7.00) {
          var duration = (parseFloat(rowData[durationCol]) + 1) * 60; // Convert hours to minutes
        } else {
          var duration = parseFloat(rowData[durationCol]) * 60;
        }


        // Add a new row to the new sheet with the extracted data
        if (duration !== 0) {
          newSheet.appendRow(["O10POS05", representativeID, fromDate, 0, visitTime, duration, "form", "12.Break Log(GB)"]);
        }
      }
    }

    // Update startCol and endCol after each iteration
    startCol = String.fromCharCode(startCol.charCodeAt(0) + 3);
    endCol = String.fromCharCode(endCol.charCodeAt(0) + 3);

    //Increasing counters to next row of data which would be for the next day
    durationCol += 3;
    visitCol += 3;
  }
  updateSchedule();
}

//converting names from Names to ID numbers
function updateSchedule() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var mappingSheet = spreadsheet.getSheetByName("EmployeeMapping");
  var scheduleSheet = spreadsheet.getSheetByName("Generated Table");

  // Get the data from the EmployeeMapping sheet
  var mappingData = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, 2).getValues();

  // Get the data from the Schedule sheet
  var scheduleData = scheduleSheet.getDataRange().getValues();

  // Create a mapping object for quick lookups
  var mappingObject = {};
  mappingData.forEach(function (row) {
    mappingObject[row[1]] = row[0];
  });

  // Update the names in the Schedule sheet based on the mapping
  for (var i = 1; i < scheduleData.length; i++) {
    var representativeName = scheduleData[i][1];
    if (mappingObject.hasOwnProperty(representativeName)) {
      scheduleData[i][1] = mappingObject[representativeName];
    }
  }

  // Update the data in the Schedule sheet
  scheduleSheet.getRange(1, 1, scheduleData.length, scheduleData[0].length).setValues(scheduleData);
}

function getEmployeeNames() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = spreadsheet.getSheetByName("repslyWeek"); // Change "test" to your sheet name
  var data = originalSheet.getRange(3, 1, originalSheet.getLastRow() - 2, 1).getValues(); // Adjust range as necessary

  var employeeNames = [];
  for (var i = 0; i < data.length; i++) {
    var name = data[i][0];
    // Check if the entry is a string and is not a header or invalid entry
    if (typeof name === 'string' && isNaN(name) && name !== 'Team' && name !== 'Week' && name !== 'Trading Hours' && name !== 'ASM' && name !== 'Supervisor' && name !== 'Retail Expert'&& name !== 'Store Manager' && name !== 'Retail Stylist'&& name !== 'Stylist'&& name !== 'Expert'&& name !== 'Daily Total' && name !== 'Week '&& name !== 'Daily Total '&& name !== 'FTE' && name !== 'FTE ') {
      employeeNames.push(name);
    }
  }
  return employeeNames;
}
TESTER.TXT
Displaying TESTER.TXT.
