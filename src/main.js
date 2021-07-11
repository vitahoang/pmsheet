var ganttSheet = 'Project Management';
var templateSheet = 'Template Backup';


/**
 * Add a new week in the gantt chart timeline
 */
function addNewWeek() {
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
*/
function formatGanttime() {
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
*/
function addNewTask() {
  console.log('Start addNewTask');
  var gSheet = getSheet(ganttSheet);
  var tSheet = getSheet(templateSheet);
  createTaskFromTemplate(tSheet, gSheet);
  formatGanttime();
}