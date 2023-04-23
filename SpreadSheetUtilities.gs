/**
 * creates new sheet ands the data to new sheet
 */
function addDataToNewSheet(data = '', sheetName = '') {
  if (data.length == 0) {
    throw ('there is no data');
    return;
  }
  var date = (new Date()).toLocaleString();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName + date);
  sheet.appendRow(data);
}


/**
 * given range in A1notation
 *  converts it to indices..
 * returns the following:
 *   var data = {
    "range": range,
    "emailCol": emailCol,
    "gradeCol": gradeCol,
    "startIndex": startIndex,
    "endIndex": endIndex
  }
 */
function convertRangeToIndices(range) {
  range = range.trim();
  var c = range.indexOf(",");
  if (c == -1) {
    c = range.indexOf(":");
    if (c == -1) {
      throw ('no colon, no comma in the range string');
    }

  }

  firstPart = range.slice(0, c).replaceAll(/[ \[\],:]/g, "");
  secondPart = range.slice(c, range.length).replaceAll(/[ \[\],:]/g, "");

  emailCol = firstPart.charCodeAt(0) - "A".charCodeAt(0);

  startIndex = (firstPart.length > 1) ?
    parseInt(firstPart.slice(1, firstPart.lenthg), 10) : 1;
  endIndex = (secondPart.length > 1) ?
    parseInt(secondPart.slice(1, secondPart.length), 10) : 1000;//Set to a Maksimum 1000.

  gradeCol = secondPart.charCodeAt(0) - "A".charCodeAt(0);
  var data = {
    "range": range,
    "emailCol": emailCol,
    "gradeCol": gradeCol,
    "startIndex": startIndex,
    "endIndex": endIndex
  }
  return data;
}
