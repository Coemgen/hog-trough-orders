/*jslint
browser, maxlen: 80, single, white
*/
/*global
AdminDirectory, Logger, SpreadsheetApp
*/
/*
 * Need to uncomment lines in function buildUsersSheet
 * to update the Users sheet.
 */


/** globals **/
var userKeyObj = {};
var badListObj = {
'@meditech.com': true,
'aannantuonio@meditech.com': true,
'aayers@meditech.com': true,
'ddsouza@meditech.com': true,
'dnetto@meditech.com': true,
'eblitz@meditech.com': true,
'ejoncas@meditech.com': true,
'ftouze@meditech.com': true,
'jcook@meditech.com': true,
'jsolari@meditech.com': true,
'klennon@meditech.com': true,
'ldibona@meditech.com': true,
'lrounds@meditech.com': true,
'mbabbitt@meditech.com': true,
'mbates@meditech.com': true,
'mcymer@meditech.com': true,
'mhobin@meditech.com': true,
'sbirch@meditech.com': true,
'sleroux@meditech.com': true,
'smceachern@meditech.com': true,
'smoquin@meditech.com': true,
'vbosteels@meditech.com': true
};
var badDataArr = [];

/** functions **/


/**
 * @param {string} userKey
 * @return {string}
 */
function getNumericKey(userKey) {
  'use strict';
  var user = {};
  //
  if (badListObj[userKey]) {
    return '0';
  }
  if (userKey.match(/^\d{21}$/)) {
    return userKey;
  }
  if (userKeyObj[userKey]) {
    return userKeyObj[userKey];
  }
  try {
    //
    user = AdminDirectory.Users
        .get(
        userKey,
        {
          projection: 'basic',
          viewType: 'domain_public'
        }
        );
    //
    userKeyObj[userKey] = user.id;
    return user.id;
  }
  catch (err) {
    Logger.log('err: %s user: %s', err, userKey);
    return '0';
  }
}


function xtractUsers(sheet, usersObj) {
  'use strict';
  var userKey = '';
  var sum = 0;
  //
  sheet.getDataRange()
      .getValues().slice(1).forEach(
      //
      function(current) {
        //
        if (current[8] === 'Deleted') { return; }
        userKey = getNumericKey(current[1]);
        try {
          if (current[5].toString().match(/\d+.?\d*/) === null) {
            badDataArr.push([sheet.getName()].concat(current));
          }
        }
        catch (err) {
          badDataArr.push([err, sheet.getName()].concat(current));
        }
        sum = Number(current[5]) + 0;
        if (usersObj[userKey]) {
          usersObj[userKey].sum = usersObj[userKey].sum + sum;
        } else if (userKey.match(/^\d{21}$/)) {
          //
          usersObj[userKey] = {};
          usersObj[userKey].sum = sum;
          usersObj.usersArr.push(userKey);
        }
      }
      );
  return usersObj;
}


function buildUsersSheet() {
  'use strict';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  var usersObj = {};
  var userKey = '';
  var sum = 0;
  var sheetBadData = {};
  usersObj.usersArr = [];
  ss.getSheets().forEach(
      function(current) {
        var name = current.getName();
        if (name.match(/^\d{4}$/)) {
          usersObj = xtractUsers(current, usersObj);
        }
      }
  );
  usersSheet.clearContents();
  usersObj.usersArr.sort();
  usersObj.usersArr.forEach(
      function(current) {
        userKey = current;
        sum = usersObj[userKey].sum;
        usersSheet.appendRow([userKey, sum]);
      }
  );
  if (badDataArr.length > 0) {
    sheetBadData = ss.insertSheet();
    badDataArr.forEach(
        function(current) {
          sheetBadData.appendRow(current);
        }
    );
  }
}
