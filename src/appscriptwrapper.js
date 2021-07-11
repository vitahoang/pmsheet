/**
 * Get speadsheet object.
 * @return  {Object} A sheet object. 
 */
function getSpreadSheet() {
  if (getSpreadSheet.ss) { return getSpreadSheet.ss; };
  getSpreadSheet.ss = SpreadsheetApp.getActive();
  return getSpreadSheet.ss;
}


/**
 * Get a sheet object by its name.
 * @param   {String} sheetName  A sheet name.     
 * @return  {Object} A sheet object. 
 */
function getSheet(sheetName) {
  var ss = getSpreadSheet();
  return ss.getSheetByName(sheetName);
};


/**
 * Set a key-value property of a document.
 * @param   {String} key     
 * @param   {String} value   
 * @return  {String} 
 */
function setDocProperty(key, value) {
  var docProperties = PropertiesService.getDocumentProperties();
  docProperties.setProperty(key, value);
}


/**
 * Get the document property by its name.
 * @param   {String} key   A name of a named range.
 * @return  {String} 
 */
function getDocProperty(key) {
  var docProperties = PropertiesService.getDocumentProperties();
  return docProperties.getProperty(key);
}


/**
 * Search a named range by its name and return its range object.
 * @param   {Object} spreadsheet    A Google Spreadsheet file object.
 * @param   {String} name  A name of a named range.
 * @return  {Object} A range object. 
 */
function getNamedRangeByName(spreadsheet, name) {
  var namedRanges = spreadsheet.getNamedRanges();
  if (namedRanges.length > 0) {
    for (var i = 0; i < namedRanges.length; i++) {
      if (namedRanges[i].getName() == name) {
        return namedRanges[i].getRange();
      }
    }
  }
  else {
    console.log("The range " + name + " doesn't exist");
  }
}