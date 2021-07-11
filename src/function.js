const ts = require("typescript");

var ganttSheet = 'PM test';
var templateSheet = 'Template Backup';



/**
 * Add a new week in the gantt chart timeline
* @param   {String} ganttSheet Name of the sheet that stores the Gantt Chart.
 */
function addNewWeek(ganttSheet) {
    console.log('Start addNewWeek');
    var sheet = getSheet(ganttSheet);
    var lastDayCol = sheet.getLastColumn() - 2;
    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();
    var lastColDate = getDateOfColumn(sheet, lastDayCol);
    var lastColMoment = getMomentOfDate(lastColDate);
    var nextMoment = getNextMonday(lastColMoment);

    for (var i = 0; i < 5; i++) {
        sheet.insertColumnsAfter(maxCol, 2);
        sheet.setColumnWidths(maxCol + 1, 2, 40);
        sheet.getRange(1, maxCol + 1)
            .setValue(formatDate(nextMoment));
        sheet.getRange(1, maxCol + 1, 1, 2)
            .merge()
            .setTextStyle(dailyHeaderStyle);
        decorLastDay(sheet);
        nextMoment.add(1, 'day');
        maxCol = maxCol + 2;
    }
    maxcol = sheet.getMaxColumns();
    decorLastWeek(maxCol, maxRow, sheet);
}

/**
* Format the Gantt timeline.
* @param   {String} ganttSheet Name of the sheet that stores the Gantt Chart.
*/
function formatGanttime(ganttSheet) {
    console.log('Start formatGanttime');
    var sheet = getSheet(ganttSheet);
    var baseCol = getBaseCol(sheet);
    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();
    var maxDay = (maxCol - baseCol) / 2;

    for (var i = 0; i < maxDay; i++) {
        var dailyCol = sheet.getRange(1, baseCol + 1, maxRow, 2);
        decorAday(dailyCol);
        baseCol = baseCol + 2;
    }
    var baseCol = getBaseCol(sheet);

    while (maxCol > baseCol) {
        decorLastWeek(maxCol, maxRow, sheet)
        maxCol = maxCol - 10;
    }
}


/**
* Add a new release to the Gantt Chart.
* @param   {String} ganttSheet Name of the sheet that stores the Gantt Chart.
*/
function addNewRelease(ganttSheet, templateSheet) {
    console.log('Start addNewRelease');
    var gSheet = getSheet(ganttSheet);
    var tSheet = getSheet(templateSheet);
    var baseCol = getBaseCol(sheet);
    var fromRange = tSheet.getRange(2, 1, tSheet.getLastRow(), baseCol);
    gSheet.insertRows(gSheet.getLastRow, tSheet.getLastRow() - 1);
    var toRange = gSheet.getRange(gSheet.getLastRow + 1, 1, tSheet.getLastRow() - 1, baseCol);

    copyRange(fromRange, toRange);
}



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
