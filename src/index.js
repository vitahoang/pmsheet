function onInstall (e) {
  onOpen(e);
}


function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Gantt Chart");
  menu.addItem("Add a new week", "addNewWeek");
  menu.addToUi();
}