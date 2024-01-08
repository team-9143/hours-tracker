/*
 * Developed by Siddharth Thomas '2025
 *
 * Logs the elapsed time between check in and check out by week, with a 30 minute timeout
 * Takes data from a spreadsheet's form responses in the format: [timestamp, member ID, input("In" or "Out")]
 * Creates a custom menu for users from the "admins" list to check a member in or out or modify a member's most recent week's time
 * Requires a sheet called "Result Sheet" to store data in, with a sheet-wide filter view
 *
 * Run triggers: onFormSubmit, onOpen
 */

const resultSheet = SpreadsheetApp.getActive().getSheetByName('Result Sheet');
var checkedIn = Object.create(null);
var members = [];

// Update variables from spreadsheet storage
function updateVars() {
  // Updates members array with all values in the sheet
  members = []; // Reset array
  const realMembers = resultSheet.getRange(2,1, resultSheet.getLastRow()-1).getValues(); // Get range of addresses from sheet
  for (let row of realMembers) {members.push(row[0]);} // Add addresses from sheet to array

  // Updates checkedIn object to contain members with timestamps for check-ins in the sheet
  checkedIn = Object.create(null); // Reset object
  const checkInTimes = resultSheet.getRange(2,3, resultSheet.getLastRow()-1).getValues().map(row => row[0]); // Get array of timestamps from sheet
  checkInTimes.forEach((timestamp, i) => {
    if (timestamp !== '') {
      // For all check-in timestamps
      checkedIn[members[i]] = new Date(timestamp); // Associate timestamps with members
    }
  })
}

// Calculate elapsed time since last check in and add data to the sheet, then remove check in data from sheet
function checkOut(id, elapsed, modifier) {
  // If more than a week has passed since the next column, create a new column for this week
  if (checkedIn[id] - resultSheet.getRange('E1').getValue() > 604800000) {
    // Create a date object for the Monday in the check in time's week
    let weekStart = new Date(checkedIn[id]);
    let day = weekStart.getDay() || 7;
    weekStart.setHours(-24 * (day-1));
    weekStart.setMinutes(0,0,0);

    // Create a column headed by the Monday's date and filled with empty times
    resultSheet.insertColumnBefore(5);
    resultSheet.getRange('E1').setValue(weekStart);
    for (let i = 2; i <= resultSheet.getMaxRows(); i++) {
      resultSheet.getRange(i, 5).setValue('0:0:0');
    }
  }

  // Add elapsed time to column
  // Get time from sheet
  var time = resultSheet.getRange(members.indexOf(id)+2, 5).getDisplayValue();
  // Create date object from time
  const [hours, minutes, seconds] = time.split(':');
  time = new Date(Date.UTC(1970, 0, hours/24 + 1, hours%24, minutes, seconds));

  // Add elapsed time to date object, multiplying by the modifier if given
  modifier = modifier || 1;
  // Stop if time or change is invalid
  if (time.toString() === 'Invalid Date') {throw 'Invalid starting time';}
  if (modifier < 0 && time.getTime() + elapsed*modifier < 0) {throw 'Cannot remove more hours than are logged';}
  time.setTime(time.getTime() + elapsed*modifier);

  // Send string from date object to spreadsheet
  time = `${(time.getUTCDate()-1)*24 + time.getUTCHours()}:${time.getUTCMinutes()}:${time.getUTCSeconds()}`;
  resultSheet.getRange(members.indexOf(id)+2, 5).setValue(time);

  // Remove check in data from sheet
  resultSheet.getRange(members.indexOf(id)+2, 3).setValue('');
}

// Check members out if they have been checked in for over 2 hours, and add 30 minutes to their time
function checkForTimeout() {
  const time = Date.now();
  for (const member in checkedIn) {
    if (time - checkedIn[member] > 7200000) {
      checkOut(member, 1800000);
      // Remove check in data to avoid double check-outs
      delete checkedIn[member];

      // Add to the number of timeouts
      resultSheet.getRange(members.indexOf(member)+2, 4).setValue(resultSheet.getRange(members.indexOf(member)+2, 4).getValue() + 1);
    }
  }
}

// Update the sheet to include any new member from a form response
function addNewMember(id) {
  if (!members.includes(id)) {
    resultSheet.insertRowBefore(2);
    resultSheet.getRange(2,1).setValue(id);
    resultSheet.getRange(2,2).setValue('=SUM(E2:2)');
    resultSheet.getRange(2,4).setValue(0);
    resultSheet.getRange(2, 5).setValue('0:0:0');
    resultSheet.getFilter().sort(2, false);

    // Update member list to account for change in spreadsheet
    members.push(member);
  }
}

function onFormSubmit(e) {
  var [timestamp, id, input] = e.values;
  timestamp = new Date(timestamp);

  updateVars();
  addNewMember(id);
  checkForTimeout();

  // If inputting to the form and has been checked in, check out
  if (checkedIn.hasOwnProperty(id)) {checkOut(id, timestamp - checkedIn[id]);}
  // If checking in, add timestamp to sheet
  if (input === 'In') {resultSheet.getRange(members.indexOf(id)+2, 3).setValue(timestamp);}
}

// Receive and validate a member name input
function inputMember() {
  const ui = SpreadsheetApp.getUi();

  const idResponse = ui.prompt('Select Member', 'Partial or full email', ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() != ui.Button.OK || idResponse.getResponseText() == '') {throw 'Operation canceled';}
  const id = members.find(member => member.startsWith(idResponse.getResponseText()));
  if (id === undefined) {throw `Member ${idResponse.getResponseText()} not found`;}

  return id;
}

// Modify a member's hours by a time input by a user
function modifyHours() {
  updateVars();
  const ui = SpreadsheetApp.getUi();
  const id = inputMember();

  // Recieve and validate a time input
  const timeResponse = ui.prompt('Amend Hours', id + '\nTime modifier (h:m:s)', ui.ButtonSet.OK_CANCEL);
  if (timeResponse.getSelectedButton() != ui.Button.OK) {throw 'Operation canceled';}

  // Create date and modifier from time input
  var elapsed = timeResponse.getResponseText();
  var modifier = 1;
  if (elapsed.startsWith('+')) {
    elapsed = elapsed.slice();
  } else if (elapsed.startsWith('-')) {
    modifier = -1;
    elapsed = elapsed.slice(1);
  }
  const [hours, minutes, seconds] = elapsed.split(':');
  elapsed = new Date(Date.UTC(1970, 0, 1, hours, minutes, seconds));
  // Stop at invalid time inputs
  if (elapsed.toString() === 'Invalid Date') {throw 'Invalid modifier time';}

  // Modify time as dictated, replacing check in data, and send confirmation message
  const checkInTime = resultSheet.getRange(members.indexOf(id)+2, 3).getValue();
  checkOut(id, elapsed.getTime(), modifier);
  resultSheet.getRange(members.indexOf(id)+2, 3).setValue(checkInTime);
  if (modifier === -1) {modifier = '-';} else {modifier = '+'}
  ui.alert('Confirmation', `${id} 's time modified by ${modifier}${elapsed.getUTCHours()}:${elapsed.getUTCMinutes()}:${elapsed.getUTCSeconds()}`, ui.ButtonSet.OK);
}

// Check in the inputted member
function adminCheckIn(t) {
  updateVars();
  const id = inputMember(),
    time = t || Date.now();

  if (checkedIn.hasOwnProperty(id)) {checkOut(id, time - checkedIn[id]);}
  resultSheet.getRange(members.indexOf(id)+2, 3).setValue(new Date(time));

  const ui = SpreadsheetApp.getUi();
  ui.alert('Confirmation', `${id} checked in`, ui.ButtonSet.OK);
}

// Check out the inputted member
function adminCheckOut() {
  updateVars();
  const id = inputMember();
  if (!checkedIn.hasOwnProperty(id)) {throw `${id} is not checked in`;}
  checkOut(id, Date.now() - checkedIn[id].getTime());

  const ui = SpreadsheetApp.getUi();
  ui.alert('Confirmation', `${id} checked out`, ui.ButtonSet.OK);
}

// Re-check in every member already checked in, effectively resetting the timeout counter
function timeoutReset() {
  updateVars();
  var updatedMembers = [];
  const time = Date.now();
  for (const id in checkedIn) {
    if (checkedIn.hasOwnProperty(id)) {checkOut(id, time - checkedIn[id]);}
    resultSheet.getRange(members.indexOf(id)+2, 3).setValue(new Date(time));
    updatedMembers.push(id);
  }

  const ui = SpreadsheetApp.getUi();
  const message = updatedMembers.length === 0 ? 'No checked in members' : (updatedMembers.length === 1 ? updatedMembers[0] + ' has been checked in' : updatedMembers.join(',\n') + '\nhave been checked in');
  ui.alert('Confirmation', message, ui.ButtonSet.OK);
}

// Update variables, then check for any timeouts
function adminTimeout() {
  updateVars();
  checkForTimeout();
}

// Initialize admin menu for sheet editors
function onOpen(e) {
  // Stop if user's email does not match an editor
  if (!resultSheet.getEditors().find(editor => e.user.getEmail() === editor.getEmail())) {return;}

  updateVars();

  // Create admin menu
  SpreadsheetApp.getUi().createMenu('Admin Settings')
    .addItem('Check in', 'adminCheckIn')
    .addItem('Check out', 'adminCheckOut')
    .addItem('Amend hours', 'modifyHours')
    .addItem('Re-check in', 'timeoutReset')
    .addItem('Check for timeout', 'adminTimeout')
    .addToUi();
}