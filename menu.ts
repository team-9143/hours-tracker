/*
 * Developed by Siddharth Thomas '2025
 *
 * Creates a custom menu for sheet editors to check a member in or out, time members out, modify a member's logged time, or exempt members from a certain week of hour requirements
 *
 * Run triggers: onOpen
 * Permissions needed: https://www.googleapis.com/auth/spreadsheets.currentonly
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

// Get the row index from the current selected cell
function rowIndexFromSelection(): number {
  let rowIndex: number = SpreadsheetApp.getActiveRange().getRow();

  // Check that selected row is valid for operations
  if (rowIndex < firstDataRowIndex || rowIndex > resultSheet.getLastRow()) {
    throw 'Invalid selection, select a row with an address';
  }

  return rowIndex;
}

// Get the week column index from the current selected cell (defaults to current week)
function colIndexFromSelection(): number {
  let colIndex: number = SpreadsheetApp.getActiveRange().getColumn();

  // Limit the selected column to valid ones
  colIndex = Math.max(colIndex, currentWeekColIndex);

  return colIndex;
}

function hoursFromInput(header: string, id: string): Date {
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const input: Base.PromptResponse = ui.prompt(header, id + '\nTime input [+/-H:M]', ui.ButtonSet.OK_CANCEL);
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
  const [hours, minutes]: string[] = inputText.split(':');
  const time: Date = new Date(
    Number(hours) * 3_600_000 // hours to milliseconds
    + Number(minutes) * 60_000 // minutes to milliseconds
  );

  // Check for invalid time input
  if (time.toString() === 'Invalid Date') {throw 'Invalid time input';}

  return isNegative ? new Date(-time.getTime()) : time;
}

// Check in a member from admin selection, using admin notes if re-checking in
function adminCheckIn(): void {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const rowIndex: number = rowIndexFromSelection();
  const id: string = resultSheet.getRange(rowIndex, addressColIndex).getDisplayValue();

  // If the member is already checked in, add hours and admin-provided metadata
  if (!resultSheet.getRange(rowIndex, checkInColIndex).isBlank()) {
    let metadata: Base.PromptResponse = ui.prompt('Re-check in notes', id + '\nProjects/tasks worked on', ui.ButtonSet.OK_CANCEL);

    // Check for user-side cancelation
    if (metadata.getSelectedButton() !== ui.Button.OK) {
      throw 'Operation canceled';
    }

    checkOut(
      rowIndex,
      undefined,
      'Admin nt: ' + formatMetadata(metadata.getResponseText())
    );
  }

  checkIn(rowIndex);

  // Confirmation message
  ui.alert('Confirmation', `${id} checked in`, ui.ButtonSet.OK);
}

// Check out a member from admin selection with admin notes
function adminCheckOut(): void {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const rowIndex: number = rowIndexFromSelection();
  const id: string = resultSheet.getRange(rowIndex, addressColIndex).getDisplayValue();

  // Check that member is checked in
  if (resultSheet.getRange(rowIndex, checkInColIndex).isBlank()) {
    throw `Member ${id} is not checked in`;
  }

  const metadata: Base.PromptResponse = ui.prompt('Check out notes', id + '\nProjects/tasks worked on', ui.ButtonSet.OK_CANCEL);

  // Check for user-side cancelation
  if (metadata.getSelectedButton() !== ui.Button.OK) {
    throw 'Operation canceled';
  }

  // Check out with admin-provided metadata
  checkOut(
    rowIndex,
    undefined,
    'Admin nt: ' + formatMetadata(metadata.getResponseText())
  );

  // Confirmation message
  ui.alert('Confirmation', `${id} checked out`, ui.ButtonSet.OK);
}

// Modify a member's hours from admin selection by an admin time input with admin notes
function adminModifyHours() {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const id: string = resultSheet.getRange(rowIndexFromSelection(), addressColIndex).getDisplayValue();
  const week: string = resultSheet.getRange(headerRowIndex, colIndexFromSelection()).getDisplayValue();
  const modifier = hoursFromInput('Amend Hours', id);

  const metadata: Base.PromptResponse = ui.prompt('Modification notes', `${id} for week of ${week}\n${formatElapsedTime(modifier)}\nProjects/tasks worked on`, ui.ButtonSet.OK_CANCEL);

  // Check for user-side cancelation
  if (metadata.getSelectedButton() !== ui.Button.OK) {
    throw 'Operation canceled';
  }

  // Add hours, with negative if applicable
  addHours(
    members.indexOf(id) + firstDataRowIndex,
    modifier,
    undefined,
    'admin',
    'Admin nt: ' + formatMetadata(metadata.getResponseText())
  );

  // Confirmation message
  let confirmation: string;
  if (modifier.getTime() < 0) {
    confirmation = '-' + formatElapsedTime(new Date(-modifier.getTime()));
  } else {
    confirmation = '+' + formatElapsedTime(modifier);
  }
  ui.alert('Confirmation', `${id} modified by ${confirmation}`, ui.ButtonSet.OK);
}

// Re-check in all members
function adminResetTimeouts(): void {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const checkInTimes: any[] = resultSheet.getRange(firstDataRowIndex, checkInColIndex, numDataRows()).getValues().map(row => row[0]);
  const resets: string[] = [];

  checkInTimes.forEach((val, i) => {
    // For all checked in members, add their hours and then re-check them in
    if (val !== '') {
      // Add hours elapsed since first check in time
      addHours(
        i + firstDataRowIndex,
        new Date(Date.now() - val.getTime()),
        undefined,
        'checkin ' + timeDateFormatter.format(val),
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

// Timeout a member from admin selection
function adminTimeoutMember(): void {
  updateVars();
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const rowIndex: number = rowIndexFromSelection();
  const id: string = resultSheet.getRange(rowIndex, addressColIndex).getDisplayValue();
  const checkInCell: Spreadsheet.Range = resultSheet.getRange(rowIndex, checkInColIndex);

  // Check that member is checked in
  if (checkInCell.isBlank()) {
    throw `Member ${id} is not checked in`;
  }

  // Add the time and note
  addHours(
    rowIndex,
    timeoutReturnTime,
    undefined,
    'checkin ' + timeDateFormatter.format(checkInCell.getValue()),
    'Admin timeout nt: ' + formatMetadata(ui.prompt('Timeout notes', id + '\nReason for timeout', ui.ButtonSet.OK).getResponseText())
  );

  // Void the check in and increment the timeout counter
  checkInCell.setValue('');
  resultSheet.getRange(rowIndex, timeoutColIndex).setValue(resultSheet.getRange(rowIndex, timeoutColIndex).getValue() + 1);

  SpreadsheetApp.flush();

  // Confirmation message
  ui.alert('Confirmation', `${id} timed out`, ui.ButtonSet.OK);
}

// Increase a member's hours to a set time for a week with a note explaining why
function adminExemptFromWeek(): void {
  const ui: Base.Ui = SpreadsheetApp.getUi();
  const logCell: Spreadsheet.Range = resultSheet.getRange(rowIndexFromSelection(), colIndexFromSelection());
  const id: string = resultSheet.getRange(logCell.getRow(), addressColIndex).getDisplayValue();
  const week: string = resultSheet.getRange(headerRowIndex, logCell.getColumn()).getDisplayValue();
  const target: Date = hoursFromInput('Target Hours (please fill out at the end of the week)', `${id} for week of ${week}`);

  // Create date object with a minimum of target time
  // Interpreting the display value here is more coherent than the literal cell value
  const [hours, minutes, seconds]: string[] = logCell.getDisplayValue().split(':');
  const time: Date = new Date(
    Number(hours) * 3_600_000 // hours to milliseconds
    + Number(minutes) * 60_000 // minutes to milliseconds
    + Number(seconds) * 1_000 // seconds to milliseconds
  );

  if (time.getTime() >= target.getTime()) {
    throw 'Member has already exceeded this hour requirement';
  }

  time.setTime(target.getTime());

  // Send metadata to cell note in new line in format 'Set to [HH:MM:SS] for exemption for:\n[metadata]'
  logCell.setNote(
    `Set ${formatElapsedTime(time)} for exemption for:\n`
    + formatMetadata(ui.prompt('Exemption notes', `${id} for week of ${week}\nReason for exemption`, ui.ButtonSet.OK).getResponseText())
    + logCell.getNote()
  );

  // Send date object to cell in format [HH:MM:SS]
  // Comes after note to avoid synchronizing issues if previous user prompt is aborted
  logCell.setValue(formatElapsedTime(time));

  SpreadsheetApp.flush();

  // Confirmation message
  ui.alert('Confirmation', `${id} exempted from week of ${week} with ${formatElapsedTime(time)}`, ui.ButtonSet.OK);
}

// Creates admin menu, runs when the spreadsheet is opened by a sheet editor
function onOpen(e: GoogleAppsScript.Events.SheetsOnOpen): void {
  // Create the admin menu
  SpreadsheetApp.getUi().createMenu('Admin Settings')
    .addItem('Check in selected row', 'adminCheckIn')
    .addItem('Check out selected row', 'adminCheckOut')
    .addItem('Amend hours for selected cell', 'adminModifyHours')
    .addItem('Reset timeouts', 'adminResetTimeouts')
    .addItem('Timeout selected row', 'adminTimeoutMember')
    .addItem('Exempt selected cell', 'adminExemptFromWeek')
    .addToUi();
}