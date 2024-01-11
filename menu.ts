/*
 * Developed by Siddharth Thomas '2025
 *
 * Creates a custom menu for users with editing permissions to check a member in or out or modify a member's logged time
 *
 * Run triggers: onOpen
 */

// Prompt and find a member address from input, returning the member address if it exists
function addressFromInput(): string {
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const input: Base.PromptResponse = ui.prompt('Select Member', 'Start of address', ui.ButtonSet.OK_CANCEL);
  const inputText: string = input.getResponseText();

  // Check for user-side cancelation
  if (input.getSelectedButton() !== ui.Button.OK || inputText === '') {
    throw 'Operation canceled';
  }

  // Search for and validate address
  const id: string | undefined = members.find(address => address.startsWith(inputText));
  if (id === undefined) {
    throw `Address '${inputText}' not found`;
  }

  // Return first hit of given address
  return id;
}

// Check in a member from admin input, using admin notees if re-checking in
function adminCheckIn(): void {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const id: string = addressFromInput();
  const rowIndex: number = members.indexOf(id) + firstDataRowIndex;

  // If the member is already checked in, add hours and admin-provided metadata
  if (!resultSheet.getRange(rowIndex, checkInColIndex).isBlank()) {
    checkOut(
      rowIndex,
      'Admin nt: ' + formatMetadata(ui.prompt('Re-check in notes', 'Projects/tasks worked on', ui.ButtonSet.OK).getResponseText())
    );
  }

  checkIn(rowIndex);

  // Confirmation message
  ui.alert('Confirmation', `${id} checked in`, ui.ButtonSet.OK);
}

// Check out a member from admin input with admin notes
function adminCheckOut(): void {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const id: string = addressFromInput();
  const rowIndex: number = members.indexOf(id) + firstDataRowIndex;

  // Check that member is checked in
  if (resultSheet.getRange(rowIndex, checkInColIndex).isBlank()) {
    throw `Member ${id} is not checked in`;
  }

  // Check out with admin-provided metadata
  checkOut(
    rowIndex,
    'Admin nt: ' + formatMetadata(ui.prompt('Check out notes', 'Projects/tasks worked on', ui.ButtonSet.OK).getResponseText())
  );

  // Confirmation message
  ui.alert('Confirmation', `${id} checked out`, ui.ButtonSet.OK);
}

// Modify a member's hours by an admin time input with admin notes
function adminModifyHours() {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const id: string = addressFromInput();
  const input: Base.PromptResponse = ui.prompt('Amend Hours', `${id}\nTime modifier [+/-H:M:S]`, ui.ButtonSet.OK_CANCEL);
  let inputText: string = input.getResponseText();

  // Check for user-side cancelation
  if (input.getSelectedButton() !== ui.Button.OK || inputText === '') {
    throw 'Operation canceled';
  }

  // Remove first non-NaN character from the input and set isNegative if input starts with a '-'
  const isNegative: boolean = inputText.charAt(0) === '-';
  if (isNaN(Number(inputText.charAt(0)))) {
    inputText = inputText.substring(1);
  }

  // Create Date object from time input
  const [hours, minutes, seconds]: string[] = inputText.split(':');
  const time: Date = new Date(
    Number(hours) * 3_600_000 // hours to milliseconds
    + Number(minutes) * 60_000 // minutes to milliseconds
    + Number(seconds) * 1_000 // seconds to milliseconds
  );

  // Check for invalid time input
  if (time.toString() === 'Invalid Date') {throw 'Invalid time input';}

  // Add hours, with negative if applicable
  addHours(
    members.indexOf(id) + firstDataRowIndex,
    isNegative ? new Date(-time.getTime()) : time,
    'admin',
    'Admin nt: ' + formatMetadata(ui.prompt('Modification notes', 'Projects/tasks worked on', ui.ButtonSet.OK).getResponseText())
  );

  // Confirmation message
  ui.alert('Confirmation', `${id} modified by ${isNegative ? '-' : '+'}${formatElapsedTime(time)}`, ui.ButtonSet.OK);
}

// Re-check in all members
function adminResetTimeouts(): void {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const checkInTimes: any[] = resultSheet.getRange(firstDataRowIndex, checkInColIndex, numDataRows).getValues().map(row => row[0]);
  const resets: string[] = [];

  checkInTimes.forEach((val, i) => {
    // For all checked in members, add their hours and then re-check them in
    if (val !== '') {
      // Add hours elapsed since first check in time
      addHours(
        i + firstDataRowIndex,
        new Date(Date.now() - val.getTime()),
        'checkin ' + humanDateFormatter.format(val),
        'Admin timeout reset'
      );

      // Re-check in
      checkIn(i + firstDataRowIndex);

      // Push address of reset member to list for confirmation message
      resets.push(members[i]);
    }
  });

  // Confirmation message
  ui.alert('Confirmation', `${resets.length} members have been re-checked in:\n${resets.join(',\n')}`, ui.ButtonSet.OK);
}

// Timeout a member from admin input
function adminTimeoutMember(): void {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const id: string = addressFromInput();
  const rowIndex: number = members.indexOf(id) + firstDataRowIndex;
  const checkInCell: Spreadsheet.Range = resultSheet.getRange(rowIndex, checkInColIndex);

  // Check that member is checked in
  if (checkInCell.isBlank()) {
    throw `Member ${id} is not checked in`;
  }

  // Add the time and note
  addHours(
    rowIndex,
    timeoutReturnTime,
    'checkin ' + humanDateFormatter.format(checkInCell.getValue()),
    'Admin timeout nt: ' + formatMetadata(ui.prompt('Timeout notes', 'Reason for timeout', ui.ButtonSet.OK).getResponseText())
  );

  // Void the check in and increment the timeout counter
  checkInCell.setValue('');
  resultSheet.getRange(rowIndex, timeoutColIndex).setValue(resultSheet.getRange(rowIndex, timeoutColIndex).getValue() + 1);

  // Confirmation message
  ui.alert('Confirmation', `${id} timed out`, ui.ButtonSet.OK);
}

// Google will prompt user for authorization before this function runs from the menu, this function provides next step instructions for the user
function authorizeUser(): void {
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const userAddr: string = Session.getActiveUser().getEmail();

  if (SpreadsheetApp.getActiveSpreadsheet().getEditors().some(editor => editor.getEmail() === userAddr)) {
    ui.alert('Success', `User ${userAddr} authorized. Please refresh`, ui.ButtonSet.OK);
  } else {
    ui.alert('Failure', `User ${userAddr} is not authorized to edit this sheet`, ui.ButtonSet.OK);
  }
}

// Creates admin menu for sheet editors, runs when the spreadsheet is opened
function onOpen(e: GoogleAppsScript.Events.SheetsOnOpen): void {
  // Create auth menu to prompt for user authorization
  SpreadsheetApp.getUi().createMenu('Admin Auth')
    .addItem('Authorize', 'authorizeUser')
    .addToUi();

  // If the script is authorized and the user is an editor, create the admin menu
  if (SpreadsheetApp.getActiveSpreadsheet().getEditors().some(editor => editor.getEmail() === Session.getActiveUser().getEmail())) {
    SpreadsheetApp.getUi().createMenu('Admin Settings')
      .addItem('Check in member', 'adminCheckIn')
      .addItem('Check out member', 'adminCheckOut')
      .addItem('Amend hours', 'adminModifyHours')
      .addItem('Reset timeouts', 'adminResetTimeouts')
      .addItem('Timeout member', 'adminTimeoutMember')
      .addToUi();
  }
}