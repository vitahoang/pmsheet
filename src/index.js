// Import momentjs lib
eval(UrlFetchApp.fetch("https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js").getContentText());

function onInstall (e) {
  onOpen(e);
}


function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Gantt Chart");
  menu.addItem("Add a new week", "addNewWeek");
  menu.addToUi();
}