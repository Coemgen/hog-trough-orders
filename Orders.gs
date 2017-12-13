/*jslint
browser, for, maxlen: 80, single, white
*/
/*global
SpreadsheetApp
*/
/*property
    abs, appendRow, clear, filter, forEach, getActive, getDisplayValue,
    getLastRow, getName, getRange, getSheets, getValues, getYear, insertSheet,
    length, match, push, reverse, setValue, slice, sort
*/

/**
1. update active index which is the row index (row number minus one) of the
   first row of active data in the Current sheet.
2. archive previous year, create new Current sheet, and perform roll-over
   functions at new year.
*/

var ROLL_OVER_HOURS = 6;

/**
 * @param {object} ss
 * @param {string} newYear
 * @param {object} sheetOldYr
 * @param {number} currActIdx
 * @param {number} lastRow
 * @param {object} dataArr
 * @param {number} index
 */
function startNewYear(ss,
    newYear,
    sheetOldYr,
    currActIdx,
    lastRow,
    dataArr,
    index) {
  'use strict';
  var row = 0;
  var column = 1;
  var numRows = 0;
  var numColumns = 11;
  var templ = ss.getSheetByName('Template');
  var sheetNewYr = templ.copyTo(ss);
  sheetNewYr.setName(newYear);
  if (currActIdx) {
  // remove current data from old sheet
  row = currActIdx + index;
  numRows = lastRow - index;
  sheetOldYr.getRange(row + 1, column, numRows, numColumns).clear();
  // add current data to new sheet
  dataArr.slice(index).forEach(
      function(current, idx) {
        if (idx === 0) { current[10] = 1; }
        current[9] = idx + 1;
        sheetNewYr.appendRow(current);
      }
  );
  }
  sheetOldYr.hideSheet();
  return;
}


function updateIndex() {
  'use strict';
  var ss = SpreadsheetApp.getActive();
  var sheet = {};
  var row = 2;
  var column = 11;
  var currActIdx = 0;
  var lastRow = 0;
  var dataArr = [];
  var currDate = new Date();
  var tempDate = {};
  var hours = 0;
  var sheetObjArr = ss.getSheets();
  var sheetNameObjArr = [];
  var i = 0;
  // get all sheets named for years
  sheetObjArr = sheetObjArr.filter(
      function(current) {
        return current.getName().match(/^\d{4}$/);
      });
  sheetObjArr.forEach(
      function(current) {
        sheetNameObjArr.push([current.getName(), current]);
      }
  );
  sheetNameObjArr.sort();
  sheetNameObjArr.reverse();
  sheet = sheetNameObjArr[0][1];
  currActIdx = Number(sheet.getRange(row, column).getDisplayValue());
  lastRow = sheet.getLastRow();
  if (lastRow !== currActIdx) {
    dataArr = sheet.getRange(
        currActIdx + 1,
        1,
        lastRow - currActIdx,
        11).getValues();
    // replace this code with the findIndex call
    for (i = 0; i < dataArr.length; i += 1) {
      tempDate = dataArr[i][0];
      hours = Math.abs(currDate - tempDate) / 3600000;
      if (hours < ROLL_OVER_HOURS) {
        sheet.getRange(2, 11).setValue(dataArr[i][9]);
        break;
      }
    }
    if (i === dataArr.length) {
      sheet.getRange(2, 11).setValue(lastRow);
    }
  }
  if (currDate.getYear() > Number(sheet.getName())) {
    if (i < dataArr.length) {
      startNewYear(ss,
          currDate.getYear().toString(),
          sheet,
          currActIdx,
          lastRow,
          dataArr,
          i);
    } else {
      startNewYear(ss, currDate.getYear().toString(), sheet);
    }
  }
}
