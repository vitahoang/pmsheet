const { fn } = require("moment");

/**
 * Add a new week in the gantt chart timeline
 */
function addNewWeek() {
    console.log('Start addNewWeek');
    var sheet = getSheet('PM test');
    var lastDayCol = sheet.getLastColumn() - 2;
    var maxCol = sheet.getMaxColumns();
    var lastColDate = getDateOfColumn(sheet, lastDayCol);
    var lastColMoment = getMomentOfDate(lastColDate);
    var nextMoment = getNextMonday(lastColMoment);

    for (var i = 0; i < 5; i++) {
        sheet.insertColumnsAfter(maxCol, 2);
        sheet.setColumnWidths(maxCol + 1, 2, 40);
        sheet.getRange(1, maxCol + 1).setValue(formatDate(nextMoment));
        sheet.getRange(1, maxCol + 1, 1, 2).merge();
        decorLastDay(sheet);
        nextMoment.add(1, 'day');
        maxCol = maxCol + 2;
    }
    decorLastWeek(sheet);
}

/**
* Format the Gantt timeline.
*/
function formatGanttime(sheet) {
    var sheet = getSheet('PM test');
    var baseCol = getBaseCol(sheet);
    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();
    var maxWeek = (maxCol - baseCol) / 5;

    for (var i = 0; i < maxCol; i++) {
        var dailyCol = sheet.getRange(1,baseCol + 1, maxRow, 2);
        decorAday(dailyCol);
        baseCol = baseCol + 2;
    }
    var baseCol = getBaseCol(sheet);

    for (var i = 0; i < maxWeek; i++) {
        var weeklyRange = sheet.getRange(1, baseCol + 1, maxRow, 10);
        weeklyRange.setBorder = (false, true, false, true, null, null, "gray", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    }
}


/**
* Get a base column of the Gantt timeline.
* @param   {Object} sheet A sheet object.
* @return  {Integer} The starting column position of the base column in the sheet.
*/
function getBaseCol(sheet) {
    return sheet.getRange('E:E')
        .getColumn();
}



/**
 * Decorate the lastest week range.
 * @param   {Object} sheet A sheet object.
 */
function decorLastWeek(sheet) {
    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();
    sheet.getRange(1, maxCol - 9, maxRow - 1, 10)
        .setBorder(false, true, false, true, null, null, "gray", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
* Decorate a weekly column of the Gantt timeline.
* @param   {Object} weeklyrange A range object which represents for a weekly column.
*/


/**
 * Decorate the lastest day column.
 * @param   {Object} sheet A sheet object.
 */
function decorLastDay(sheet) {
    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();
    sheet.getRange(1, maxCol - 1, maxRow - 1, 2)
        .setBorder(false, true, false, true, false, false, "gray", SpreadsheetApp.BorderStyle.DOTTED);
}


/**
* Decorate a daily column of the Gantt timeline.
* @param   {Object} dailyRange A range object which represents for a daily column.
*/
function decorAday(dailyRange) {
    var rangeHeader = dailyRange.offset(0,0,1,2);
    dailyRange.setBorder(false, true, false, true, false, false, "gray", SpreadsheetApp.BorderStyle.DOTTED);
    rangeHeader.setTextStyle(dailyHeaderStyle);
}



/**
 * Get a date string from the last date column of Gantt chart.
 * @param   {Object} sheet A sheet object.
 * @param   {Integer} dailyCol A number represents for the first column of a daily column.
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
