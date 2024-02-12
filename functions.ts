/*
 * Developed by Siddharth Thomas '2025
 *
 * Implements custom functions for in-cell calculations of values.
 */

// Multiplier for missed hours into recovery time
const missedTimeMultiplier = 2;

// Calculates hours required for a member to meet active status
function GET_MISSED_HOURS(row: number): string {
  // Interpreting the display value here is more coherent than the literal cell value
  const weeklyReqMS: number = durationToDate(resultSheet.getRange(row, hourReqColIndex).getDisplayValue()).getTime();
  let requiredMS: number = 0;

  // Weeks and array of logged times to minimize number of API calls
  const weeks: number = resultSheet.getLastColumn() - currentWeekColIndex + 1;
  const loggedTime: string[] = resultSheet.getRange(row, currentWeekColIndex, 1, weeks).getDisplayValues()[0];

  // For each hour tracking column, moving from oldest (right) to newest (left), not counting current week
  for (let col = weeks - 1; col > 0; col--) {
    let deltaMS: number = durationToDate(loggedTime[col]).getTime() - weeklyReqMS

    if (deltaMS < 0) {
      // Increment requirement by difference between logged time and hour minimum with multiplier applied
      requiredMS += -deltaMS * missedTimeMultiplier;
    } else if (requiredMS > 0) {
      // Decrement requirement by additional hours over minimum requirement (block from being below 0)
      requiredMS -= Math.min(deltaMS, requiredMS);
    }
  }

  // Decrement counter by additional hours logged in the current week
  const currentLoggedAboveMin: number = durationToDate(loggedTime[0]).getTime() - weeklyReqMS;
  if (currentLoggedAboveMin > 0) {
    requiredMS -= Math.min(currentLoggedAboveMin, requiredMS);
  }

  // Return final formatted value
  return formatElapsedTime(new Date(requiredMS));
}