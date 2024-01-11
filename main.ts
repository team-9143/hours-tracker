/*
 * Developed by Siddharth Thomas '2025
 *
 * Requires a sheet called 'Result Sheet' to store data in, with a sheet-wide filter view
 * Logs the elapsed time between check in and check out by week, timing out after 2 hours, adding notes to describe metadata
 * Takes data from a spreadsheet's form responses in the format: [timestamp, member ID, input('In' or 'Out'), metadata]
 *
 * Run triggers: onFormSubmit; updateTimeouts on each hour
 */

import Base = GoogleAppsScript.Base;
import Spreadsheet = GoogleAppsScript.Spreadsheet;

const firstDataRowIndex = 2; // Index of first row with a member address

const memberColIndex = 1; // Index of column of member addresses
const totalHoursColIndex = 2; // Index of column of total hours logged
const checkInColIndex = 3; // Index of column with check in times
const timeoutColIndex = 4; // Index of column with timeouts
const currentWeekColIndex = 5; // Index of column representing current week of logged hours

// Legible date formatter in format [Day HH:MM:SS AM/PM]
const humanDateFormatter = new Intl.DateTimeFormat('en-us', {weekday: 'short', hour: '2-digit', minute: '2-digit', second: '2-digit'});

const timeoutReturnTime: Date = new Date(1_800_000); // Time given back after a timeout, 30 minutes
const timeoutReq: Date = new Date(7_200_000); // Time until an automated timeout is performed, 2 hours

const resultSheet: Spreadsheet.Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Result Sheet') as Spreadsheet.Sheet; // Sheet we're working with

const members: string[] = []; // Array of members with indexes relative to spreadsheet rows
var numDataRows: number; // Number of rows with member data

// Updates members array and number of data rows from spreadsheet storage
function updateVars(): void {
  numDataRows = resultSheet.getLastRow() - firstDataRowIndex + 1; // Recalculate number of data rows

  // Get range of addresses from sheet and add to array
  members.length = 0; // Reset array
  resultSheet.getRange(firstDataRowIndex, memberColIndex, numDataRows).getDisplayValues().forEach(row => members.push(row[0]));
}

// Formats a date within a month of the epoch into [HH:MM:SS] with leading zeros and a minus sign if necessary
function formatElapsedTime(date: Date): string {
  // If the date is negative, treat it as a positive date and add the negative later
  let isNegative: boolean = false;
  if (date.getTime() < 0) {
    date.setTime(-date.getTime());
    isNegative = true;
  }

  // Initialize specific time variables from date
  const hours: number = (date.getUTCDate()-1)*24 + date.getUTCHours();
  const minutes: number = date.getUTCMinutes();
  const seconds: number = date.getUTCSeconds();

  // Add minus sign and leading zeroes and return
  return (isNegative ? '-' : '')
    + (hours < 10 ? '0'+hours : hours) + ':'
    + (minutes < 10 ? '0'+minutes : minutes) + ':'
    + (seconds < 10 ? '0'+seconds : seconds);
}

// Formats user-input metadata for consistency
function formatMetadata(metadata: string): string {
  metadata = metadata.trim();

  // Return 'N/A' if string is empty or equivalent to N/A
  if (metadata === '' || metadata.toLowerCase() === 'n/a') {
    return 'N/A'
  }

  // Replace line breaks with semicolons and return
  return metadata.replace(/\r?\n|\r/g, '; ');
}

// Adds elapsed time to a row with an annotation in the cell note describing time elapsed and metadata
function addHours(rowIndex: number, elapsed: Date, callStack: string, metadata: string): void {
  const logCell: Spreadsheet.Range = resultSheet.getRange(rowIndex, currentWeekColIndex);

  // Create date object from member's logged time and new elapsed time
  // Interpreting the display value here is more coherent than the literal cell value
  const [hours, minutes, seconds]: string[] = logCell.getDisplayValue().split(':');
  const time: Date = new Date(
    Number(hours) * 3_600_000 // hours to milliseconds
    + Number(minutes) * 60_000 // minutes to milliseconds
    + Number(seconds) * 1_000 // seconds to milliseconds
    + elapsed.getTime() // Add elapsed time to logged time
  );

  // Check that old and new logged times are valid
  if (time.toString() === 'Invalid Date') {throw 'Invalid logged hours';}
  if (time.getTime() < 0) {throw 'Cannot log a negative number of hours';}

  // Send date object to cell in format [HH:MM:SS]
  logCell.setValue(formatElapsedTime(time));

  // Send metadata to cell note in new line in format 'Logged [HH:MM:SS] from [callStack] for:\n[metadata]'
  logCell.setNote(
    logCell.getNote() + '\n\n'
    + `Logged ${formatElapsedTime(elapsed)} from ${callStack} for:\n`
    + formatMetadata(metadata)
  );
}

// Checks in a row with the current time
function checkIn(rowIndex: number): void {
  resultSheet.getRange(rowIndex, checkInColIndex).setValue(new Date());
}

// If a row is checked in, checks it out and logs the elapsed time and metadata
function checkOut(rowIndex: number, metadata: string): void {
  const checkInTime: any = resultSheet.getRange(rowIndex, checkInColIndex).getValue();
  // Check that the member is checked in
  if (checkInTime !== '') {
    // Add hours elapsed since checked in time
    addHours(
      rowIndex,
      new Date(Date.now() - checkInTime.getTime()),
      'checkin ' + humanDateFormatter.format(checkInTime),
      metadata
    );
    resultSheet.getRange(rowIndex, checkInColIndex).setValue(''); // Remove check in data
  }
}

// Voids a row's check-in, increments the timeout counter, and logs the tiemout return time with a note signifying a timeout
function timeout(rowIndex: number): void {
  const checkInTime: any = resultSheet.getRange(rowIndex, checkInColIndex).getValue();
  // Check that the member is checked in
  if (checkInTime.getValue() !== '') {
    // Add the time and note
    addHours(
      rowIndex,
      timeoutReturnTime,
      'checkin ' + humanDateFormatter.format(checkInTime),
      'Timeout'
    );
    // Void the check in and increment the timeout counter
    resultSheet.getRange(rowIndex, checkInColIndex).setValue('');
    resultSheet.getRange(rowIndex, timeoutColIndex).setValue(resultSheet.getRange(rowIndex, timeoutColIndex).getValue() + 1);
  }
}

// Checks all members for timeouts and returns a list of row indexes for those who have passed the required time
function timeoutCheck(): number[] {
  const checkInTimes: any[] = resultSheet.getRange(firstDataRowIndex, checkInColIndex, numDataRows).getValues().map(row => row[0]);
  const timeoutRowIndexes: number[] = [] // Array to fill with timed out members

  checkInTimes.forEach((val, i) => {
    // Check that value is not blank and has passed the timeout time, then push member to list
    if (val !== '') {
      if (Date.now() - val.getTime()  > timeoutReq.getTime()) {
        timeoutRowIndexes.push(i + firstDataRowIndex); // checkInTimes array should be relative to sheet
      }
    }
  });

  return timeoutRowIndexes;
}

// Times out members who have passed the time requirement
function updateTimeouts(): void {
  timeoutCheck().forEach(rowIndex => timeout(rowIndex));
}

// Creates a new log column for the current week
function startWeek(): void {
  // Create a date object for the start of the Monday in the current week
  const weekStart: Date = new Date();
  const dayNum: number = weekStart.getDay() || 7; // Day of the week, 1 = Monday and 7 = Sunday
  weekStart.setHours(-24 * (dayNum-1)); // Set at 0 hours and subtract 24 for each day past Monday
  weekStart.setMinutes(0, 0, 0); // Set minutes, seconds, and milliseconds to 0

  // Create a column headed by the Monday's date and filled with zero times
  resultSheet.insertColumnBefore(currentWeekColIndex); // New column 5 inherits formatting from previous column 5
  resultSheet.getRange(firstDataRowIndex-1, currentWeekColIndex).setValue(weekStart); // Header set to date of the Monday
  resultSheet.getRange(firstDataRowIndex, currentWeekColIndex, numDataRows).setValues(new Array<string[]>(numDataRows).fill(['0:0:0'])); // Set column values to 0
}

// Adds a new row for a new member, sorts the sheet, and updates variables to match
function addMember(id: string): void {
  resultSheet.insertRowBefore(firstDataRowIndex); // Create row
  // Initialize row
  resultSheet.getRange(firstDataRowIndex, memberColIndex).setValue(id);
  resultSheet.getRange(firstDataRowIndex, totalHoursColIndex).setValue(`=SUM(${String.fromCharCode(currentWeekColIndex+64)}${firstDataRowIndex}:${firstDataRowIndex})`);
  resultSheet.getRange(firstDataRowIndex, timeoutColIndex).setValue(0);
  resultSheet.getRange(firstDataRowIndex, currentWeekColIndex).setValue('0:0:0');

  resultSheet.getFilter().sort(totalHoursColIndex, false); // Sort sheet by total hours descending

  // Update variables to account for change in spreadsheet
  updateVars();
}

// Handles automated updates, runs when a connected google form is submitted
function onFormSubmit(e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  updateVars();
  const [timestamp, address, input, metadata] = e.values; // Retrieve ordered values from form

  // Add a new member if necessary and find relative row index
  let index: number = members.indexOf(address);
  if (index === -1) {
    addMember(address);
    index = members.indexOf(address);
  }

  // If inputting to the form and has been checked in, add hours and metadata and remove check-in
  checkOut(index + firstDataRowIndex, metadata);
  // If checking in, add timestamp to sheet
  if (input === 'In') {
    checkIn(index + firstDataRowIndex);
  }
}