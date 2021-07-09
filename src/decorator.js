/**
 * Decorate the lastest week range.
 * @param   {Object} sheet A sheet object.
 * @param   {Number} maxCol The max number of column of the sheet.
 * @param   {Number} maxRow The max number of row of the sheet.
 */
function decorLastWeek(maxCol, maxRow, sheet) {
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
    sheet.getRange(1, maxCol + 1, 1, 2)
            .merge()
            .setTextStyle(dailyHeaderStyle);
    sheet.getRange(1, maxCol - 1, maxRow - 1, 2)
        .setBorder(false, true, false, true, false, false, "gray", SpreadsheetApp.BorderStyle.DOTTED);
}


/**
* Decorate a daily column of the Gantt timeline.
* @param   {Object} dailyRange A range object which represents for a daily column.
*/
function decorAday(dailyRange) {
    var rangeHeader = dailyRange.offset(0, 0, 1, 2);
    dailyRange.setBorder(false, true, false, true, false, false, "gray", SpreadsheetApp.BorderStyle.DOTTED);
    rangeHeader.setTextStyle(dailyHeaderStyle);
}