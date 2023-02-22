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

const resultSheet = SpreadsheetApp.getActive().getSheetByName('Result Sheet'),
  checkedIn = new Object(),
  admins = ['edean2025@', 'ftorrano@', 'jgreenbaum2025@', 'jstevens@', 'sthomas2025@'];
var members = [];

// Update variables from storage
function updateVars() {
  // Updates members array with all values in the sheet
  members = [];
  const memberList = resultSheet.getRange(2,1, resultSheet.getLastRow()-1).getValues();
  for (let member of memberList) {members.push(member[0]);}

  // Updates checkedIn object to include all members with timestamps for check-ins in the sheet, and remove those without
  const checkInTimes = resultSheet.getRange(2,3, resultSheet.getLastRow()-1).getValues().map(item => item[0]);
  checkInTimes.forEach((timestamp, i) => {
    let member = members[i]
    if (timestamp === '') {
      delete checkedIn[member];
    } else {
      checkedIn[member] = new Date(timestamp);
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
      
      // Add to the number timeouts
      resultSheet.getRange(members.indexOf(member)+2, 4).setValue(resultSheet.getRange(members.indexOf(member)+2, 4).getValue() + 1);
      const weeklyCell = resultSheet.getRange(members.indexOf(member)+2, 5);
      if (parseInt(weeklyCell.getNote().split(' ')[1]) >= 1) {
        weeklyCell.setNote('Timeouts: ' + (parseInt(weeklyCell.getNote().split(' ')[1]) + 1));
      } else {
        weeklyCell.setNote('Timeouts: 1');
      }
    }
  }
}

// Update the sheet to include any new member from a form response
function addNewMember(id) {
  if (!members.includes(id)) {
    resultSheet.insertRowBefore(2);
    resultSheet.getRange(2,1).setValue(id);
    resultSheet.getRange(2,2).setValue('=SUM(OFFSET(A2, 0, 4, 1, COLUMNS(2:2)-2))');
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
  if (idResponse.getSelectedButton() != ui.Button.OK) {throw 'Operation canceled';}
  const id = members.find(member => member.includes(idResponse.getResponseText()));
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

// Checks in every member already checked in
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

function onOpen(e) {
  // Stop if user's email does not include a string from the admin list
  if (!admins.find(item => e.user.getEmail().includes(item))) {return;}
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
