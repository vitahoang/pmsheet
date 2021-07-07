// Import momentjs lib
eval(UrlFetchApp.fetch("https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js").getContentText());

/**
 * Add a new week in the gantt chart timeline
 */
function addNewWeek() {
    console.log('Start addNewWeek');
    var sheet = getSheet('PM test');
    var ss = getSpreadSheet();
    var baseCol = getNamedRangeByName(ss, 'status').getColumn();
    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();
    var lastColDate = getLastColValue(sheet);
    var lastDate = parseMoment(lastColDate);
    var nextDate = getNextMonday(lastDate);

    for (var i = 0; i < 5; i++) {
        sheet.insertColumnsAfter(maxCol, 2);
        sheet.setColumnWidths(maxCol + 1, 2, 40);
        sheet.getRange(1, maxCol + 1).setValue(formatDate(nextDate));
        sheet.getRange(1, maxCol + 1, 1, 2).merge();
        decorNewDay(sheet);
        nextDate.add(1, 'day');
        maxCol = maxCol + 2;
    }
    decorNewWeek(sheet);
}


function decorNewWeek(sheet) {
    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();
    sheet.getRange(1, maxCol - 9, maxRow - 1, 10)
    .setBorder(false, true, false, true, null, null, "gray", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); 
}


function decorNewDay(sheet) {
    var maxCol = sheet.getMaxColumns();
    var maxRow = sheet.getMaxRows();
    sheet.getRange(1, maxCol - 1, maxRow - 1, 2)
        .setBorder(false, true, false, true, false, false, "gray", SpreadsheetApp.BorderStyle.DOTTED);
}


function getLastColValue (sheet) {
    var maxCol = sheet.getLastColumn();
    var aryValues = sheet.getRange(1, maxCol - 2, 1, 2).getValues();
    if (aryValues[0][0] !== "") return aryValues[0][0];
    else return aryValues[0][1];
}


function parseDayOfWeek(date) {
    var aryDate = date.split('\n');
    return aryDate[0];
}


function parseMoment(date) {
    var aryDate = date.split('\n');
    var ddMMM = aryDate[1].split(' ');
    var date = moment();
    return date.month(ddMMM[1]).date(ddMMM[0]);
}


function getNextMonday(lastDate) {
    var monday = lastDate.isoWeekday(1).add(1, 'week');
    return monday;
}


function formatDate(date) {
    var newdate = date.format('dddd DD MMM');
    var arydate = newdate.split(' ');
    newdate = arydate[0] + '\n' + arydate[1] + ' ' + arydate[2];
    return newdate;
}




