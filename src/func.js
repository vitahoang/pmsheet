
/**
 * Add a new week in the gantt chart timeline
 */
function addNewWeek() {
    console.log('Start addNewWeek');
    var sheet = getSheet('Project Management');
    var ss = getSpreadSheet();
    var baseCol = getNamedRangeByName(ss, 'status').getColumn();
    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();
    var lastColDate = getLastColValue(sheet);
    var lastDate = getLastMoment(lastColDate);
    var nextDate = getNextMonday(lastDate);

    for (var i = 0; i < 5; i++) {
        sheet.insertColumnsAfter(maxCol, 2);
        sheet.setColumnWidths(maxCol + 1, 2, 40);
        sheet.getRange(1, maxCol + 1).setValue(formatDate(nextDate));
        sheet.getRange(1, maxCol + 1, 1, 2).merge();
        decorLastDay(sheet);
        nextDate.add(1, 'day');
        maxCol = maxCol + 2;
    }
    decorLastWeek(sheet);
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
 * Get a date string from the last day column of Gantt chart.
 * @param   {Object} sheet A sheet object.
 * @return  {String} The last date string. 
 */
function getLastColValue(sheet) {
    var maxCol = sheet.getLastColumn();
    var aryValues = sheet.getRange(1, maxCol - 2, 1, 2).getValues();
    if (aryValues[0][0] !== "") return aryValues[0][0];
    else return aryValues[0][1];
}


/**
 * Get a moment object from the last date string of the Gantt chart.
 * @param   {String} date A string date.
 * @return  {Object} A moment object. 
 */
function getLastMoment(date) {
    var aryDate = date.split('\n');
    var ddMMM = aryDate[1].split(' ');
    var date = moment();
    return date.month(ddMMM[1]).date(ddMMM[0]);
}


/**
 * Get a moment object of the next monday.
 * @param   {Object} lastDate A moment object.
 * @return  {Object} A new moment object. 
 */
function getNextMonday(lastDate) {
    var monday = lastDate.isoWeekday(1).add(1, 'week');
    return monday;
}


/**
 * Parse a string from a moment date and format it.
 * @param   {Object} date A moment object.
 * @return  {String} A formatted string of a date.
 */
function formatDate(date) {
    var newdate = date.format('dddd DD MMM');
    var arydate = newdate.split(' ');
    newdate = arydate[0] + '\n' + arydate[1] + ' ' + arydate[2];
    return newdate;
}




