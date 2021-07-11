/**
* Get a base column of the Gantt timeline.
* @param   {Object} sheet A sheet object.
* @return  {Number} The starting column position of the base column in the sheet.
*/
function getBaseCol(sheet) {
  return sheet.getRange('F:F')
    .getColumn();
}


/**
 * Get a date string from the last date column of Gantt chart.
 * @param   {Object} sheet A sheet object.
 * @param   {Number} dailyCol A number represents for the first column of a daily column.
 * @return  {String} A string of the last day of the Gant chart formatted as "dddd \n DD MMM".
 */
function getDateOfColumn(sheet, dailyCol) {
  var aryValues = sheet.getRange(1, dailyCol - 2, 1, 2).getValues();
  if (aryValues[0][0] !== "") return aryValues[0][0];
  else return aryValues[0][1];
}


/**
 * Get a moment object from date string of a day column of the Gantt chart.
 * @param   {String} date A string date.
 * @return  {Object} A moment object. 
 */
function getMomentOfDate(date) {
  var aryDate = date.split('\n');
  var ddMMM = aryDate[1].split(' ');
  var date = moment();
  return date.month(ddMMM[1]).date(ddMMM[0]);
}


/**
 * Get a moment object of the next monday.
 * @param   {Object} lastColMoment A moment object.
 * @return  {Object} A new moment object of the next monday. 
 */
function getNextMonday(lastColMoment) {
  var monday = lastColMoment.isoWeekday(1).add(1, 'week');
  return monday;
}


/**
 * Parse a string date from a moment date object and format it.
 * @param   {Object} date A moment object.
 * @return  {String} A string of a date formatted as "dddd \n DD MMM".
 */
function formatDate(date) {
  var newdate = date.format('dddd DD MMM');
  var arydate = newdate.split(' ');
  newdate = arydate[0] + '\n' + arydate[1] + ' ' + arydate[2];
  return newdate;
}


/**
* Copy value and format from a range to another range.
* @param   {Object} tSheet A sheet which contains the release template.
* @param   {Object} gSheet A sheet where the release will be insert to.
*/
function createTaskFromTemplate(tSheet, gSheet) {
  var baseCol = getBaseCol(gSheet);
  var g_lastRowID = gSheet.getLastRow();
  var t_lastRowID = tSheet.getLastRow();
  var fromRange = tSheet.getRange(2, 1, t_lastRowID - 1, baseCol);
  gSheet.insertRowsAfter(g_lastRowID, tSheet.getLastRow() - 1);
  var toRange = gSheet.getRange(g_lastRowID + 1, 1, t_lastRowID - 1, baseCol);
  fromRange.copyTo(toRange);
  decorTaskHeader(g_lastRowID + 1, gSheet);
}
