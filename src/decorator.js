/**
 * Decorate the lastest week range.
 * @param   {Object} sheet A sheet object.
 * @param   {Number} maxCol The max number of column of the sheet.
 * @param   {Number} maxRow The max number of row of the sheet.
 */
function decorLastWeek(maxCol, maxRow, sheet) {
    sheet.getRange(1, maxCol - 9, maxRow - 1, 10)
        .setBorder(false, true, false, true, null, null, "gray", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
        .setHorizontalAlignment("center");
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
* Decorate a daily column of the Gantt timeline.
* @param   {Object} dailyRange A range object which represents for a daily column.
*/
function decorAday(dailyRange) {
    var maxRow = dailyRange.getValues().length;
    var rangeHeader = dailyRange.offset(0, 0, 1, 2);
    var rangeBody = dailyRange.offset(1, 0, maxRow - 1, 2);
    rangeBody.setBorder(false, false, false, false, false, false);
    dailyRange.setBorder(false, true, false, true, false, false, "gray", SpreadsheetApp.BorderStyle.DOTTED);
    rangeHeader.setTextStyle(dailyHeaderStyle);
}