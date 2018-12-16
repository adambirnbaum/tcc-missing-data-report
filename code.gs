Logger = BetterLog.useSpreadsheet('1WkVozJVnLzQWaPzp-Owq2VzLdTKjkVBMqmozaCesaxQ');

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  /*****************************************************************************************************************
  *
  *
  *
  *****************************************************************************************************************/
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("3CC Missing Data Report for Accounting")
      .addItem("Manually Run Reports", "triggerGenerateReports")
      .addSeparator()
      .addItem("Enable Weekly Reporting", "menuItem2")
      .addItem("Disable Weekly Reporting", "menuItem3")
      .addToUi();
}

function menuItem2() {
  Logger.log("**** Running menuItem2() ****");
  // Trigger every Monday at 01:00AM CT.
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert('Please confirm',
                        'Are you sure you want to enable automatic weekly reporting?',
                         ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    // Check if any time triggers have already been created
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getTriggerSource() == ScriptApp.TriggerSource.CLOCK) {
        throw 'Error:  Weekly reporting is already enabled.  Can only enable one time.';
      }
    }
    ScriptApp.newTrigger('triggerGenerateReports')
    .timeBased()
    //.everyMinutes(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(1)
    .inTimezone("America/Chicago")
    .create();

    ui.alert('Weekly reporting has been enabled.  Reports are sent every Monday at 1:00 am CT.');
  } else {
    // User clicked "No" or X in the title bar.
  }
}

function menuItem3() {
  Logger.log("**** Running menuItem3() ****");
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert('Please confirm',
                        'Are you sure you want to disable automatic weekly reporting?',
                         ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    // Loop over all triggers.
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
      // If the current trigger is a time trigger, delete it.
      if (allTriggers[i].getTriggerSource() == ScriptApp.TriggerSource.CLOCK) {
        ScriptApp.deleteTrigger(allTriggers[i]);
      }
    }
    ui.alert('Automatic weekly reporting has been disabled.  Weekly reports will not be sent.');
  } else {
    // User clicked "No" or X in the title bar.
  }
}

function triggerGenerateReports() {
  Logger.log("##################### S T A R T Running triggerGenerateReports() #####################");
  var myActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  myActiveSpreadsheet.toast('Loading settings', 'Status');
  var settings = getSettings();

  myActiveSpreadsheet.toast('Loading data from Masters', 'Status');
  var objectRowsData = getAllDataFromMasters(settings);

  myActiveSpreadsheet.toast('Sorting data for reports', 'Status');
  var sortedKeysFromObjectRowsData = sortKeysFromObjectRowsData(Object.keys(objectRowsData), settings);

  myActiveSpreadsheet.toast('Filtering data for reports', 'Status');
  var filteredCarrierReportRowsData = filterObjectRowsData(objectRowsData, sortedKeysFromObjectRowsData, "Carrier");
  var filteredVendorReportRowsData = filterObjectRowsData(objectRowsData, sortedKeysFromObjectRowsData, "Vendor");
  var filteredCustomerReportRowsData = filterObjectRowsData(objectRowsData, sortedKeysFromObjectRowsData, "Customer");
  var filteredWeightOrDateReportRowsData = filterObjectRowsData(objectRowsData, sortedKeysFromObjectRowsData, "WeightOrDate");

  myActiveSpreadsheet.toast('Creating reports', 'Status');
  var dataTableForCarrierEmail = createReport(filteredCarrierReportRowsData, sortedKeysFromObjectRowsData, "Carrier");
  var dataTableForVendorEmail = createReport(filteredVendorReportRowsData, sortedKeysFromObjectRowsData, "Vendor");
  var dataTableForCustomerEmail = createReport(filteredCustomerReportRowsData, sortedKeysFromObjectRowsData, "Customer");
  var dataTableForWeightOrDateEmail = createReport(filteredWeightOrDateReportRowsData, sortedKeysFromObjectRowsData, "WeightOrDate");

  myActiveSpreadsheet.toast('Sending email reports', 'Status');
  sendEmailReport(dataTableForCarrierEmail, settings, "Carrier");
  sendEmailReport(dataTableForVendorEmail, settings, "Vendor");
  sendEmailReport(dataTableForCustomerEmail, settings, "Customer");
  sendEmailReport(dataTableForWeightOrDateEmail, settings, "WeightOrDate");

  Logger.log("That's the end of the show, folks!");
}

function getSettings() {
  Logger.log("**** Running getSettings() ****");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings");

  if (sheet == null) {        // if Settings sheet does not exist in destination
      throw 'Error:  The Settings sheet could not be found';
  }
  var data = sheet.getDataRange().getValues();

  var ssUrl = ss.getUrl();

  var emailsWeightWeeklyReport = data[4][4];
  var emailsCustomerWeeklyReport = data[5][4];
  var emailsCarrierWeeklyReport = data[6][4];
  var emailsVendorWeeklyReport = data[7][4];

  var totalNumTraders = data[9][4];

  var masterHeaderRowNum = data[30][4];
  var masterDataStartRowNum = data[31][4];
  var masterDataStartColNum = letterToColumn(data[32][4]);
  var masterDataEndColNum = letterToColumn(data[33][4]);

  var arrayTraderNames = data[57][4].split(',');
  var arrayMasterUrls = data[58][4].split(',');

  return {ssUrl:ssUrl,

          emailsWeightWeeklyReport:emailsWeightWeeklyReport,
          emailsCustomerWeeklyReport:emailsCustomerWeeklyReport,
          emailsCarrierWeeklyReport:emailsCarrierWeeklyReport,
          emailsVendorWeeklyReport:emailsVendorWeeklyReport,

          totalNumTraders:totalNumTraders,

          masterHeaderRowNum:masterHeaderRowNum,
          masterDataStartRowNum:masterDataStartRowNum,
          masterDataStartColNum:masterDataStartColNum,
          masterDataEndColNum:masterDataEndColNum,

          arrayTraderNames:arrayTraderNames,
          arrayMasterUrls:arrayMasterUrls};
}

function getAllDataFromMasters(settings) {
  Logger.log("**** Running getAllDataFromMasters ****");
  var arrayMasterUrls = settings.arrayMasterUrls;
  var arrayTraderNames = settings.arrayTraderNames;
  var masterHeaderRowNum = settings.masterHeaderRowNum;
  var masterDataStartRowNum = settings.masterDataStartRowNum;
  var masterDataStartColNum = settings.masterDataStartColNum;
  var masterDataTotalCols = settings.masterDataEndColNum - masterDataStartColNum + 1;

  var objectRowsData = {};

  // Loop through the master Urls
  for (var i = 0; i < arrayMasterUrls.length; i++) {
    // Get a list of sheets within the current master
    var ss = SpreadsheetApp.openByUrl(arrayMasterUrls[i])
    var sheets = ss.getSheets()
      .filter(function(x) {
        return (x.getName() != "Settings" && x.getName() != "Master Template TO DUPLICATE" && x.getName() != "Log");     // filter out Setting/template sheets
      });
    // Logger.log('Sheets = ' + sheets);
    // Loop through the sheets within each master
    for (var j = 0; j < sheets.length; j++) {
      var sheet = sheets[j];
      // >>>> Eventually will need to look into whether using LastRow is an issue here, vs stripping out blank cells at bottom
      var range = sheet.getRange(masterDataStartRowNum, masterDataStartColNum, sheet.getLastRow(), masterDataTotalCols);

      // store row data indexed by column header name, with trader name as the key
      var sheetNameAndTraderName = sheet.getName() + "-" + arrayTraderNames[i];
      objectRowsData[sheetNameAndTraderName] = getRowsData(sheet, range, masterHeaderRowNum);
      // Logger.log("objectRowsData[sheetNameAndTraderName][0] = " + objectRowsData[sheetNameAndTraderName][0]);
      // Logger.log("objectRowsData[sheetNameAndTraderName][0][carrierInvNumber] = " + objectRowsData[sheetNameAndTraderName][0]["carrierInvNumber"]);
    }
  }
  // Logger.log('objectRowsData = ' + objectRowsData);
  return objectRowsData;
}

function sortKeysFromObjectRowsData(keysFromObjectRowsData, settings) {
  Logger.log("**** Running sortKeysFromObjectRowsData() ****");
  Logger.log("keysFromObjectRowsData = " + keysFromObjectRowsData);
  var sheetAndTraderNumber = [];
  var sortedKeysFromObjectRowsData = [];
  // convert into a numerical, sortable value
  for (var i = 0; i < keysFromObjectRowsData.length; i++) {
    sheetAndTraderNumber.push(convertSheetNameAndTraderNameToNumber(keysFromObjectRowsData[i], settings));
  }
  // sort the array
  sheetAndTraderNumber.sort(function(a, b){return a-b});

  // convert back into original alphanumberical value
  for (var i = 0; i < sheetAndTraderNumber.length; i++) {
    sortedKeysFromObjectRowsData.push(convertSheetAndTraderNumberToName(sheetAndTraderNumber[i], settings));
  }
  // Logger.log('sortedKeysFromObjectRowsData = ' + sortedKeysFromObjectRowsData);
  return sortedKeysFromObjectRowsData;
}

function filterObjectRowsData(objectRowsData, sortedKeysFromObjectRowsData, reportType) {
  Logger.log("**** Running filterObjectRowsData() ****");
  Logger.log("sortedKeysFromObjectRowsData = " + sortedKeysFromObjectRowsData);
  var filteredObjectRowsData = {};
  switch(reportType) {
    case "Carrier":
      for (var i = 0; i < sortedKeysFromObjectRowsData.length; i++) {
        filteredObjectRowsData[sortedKeysFromObjectRowsData[i]] = objectRowsData[sortedKeysFromObjectRowsData[i]].filter(function(x) {
          if (!x.carrierInvNumber && x.dateLastSynced && x.status != "Canceled") {
            return true
          } else {
            return false
          }
        });
      }
      break;
    case "Customer":
      for (var i = 0; i < sortedKeysFromObjectRowsData.length; i++) {
        filteredObjectRowsData[sortedKeysFromObjectRowsData[i]] = objectRowsData[sortedKeysFromObjectRowsData[i]].filter(function(x) {
          if (!x.customerInvNumber3ccInvoice && x.dateLastSynced && x.status != "Canceled") {
            return true
          } else {
            return false
          }
        });
      }
      break;
    case "Vendor":
      for (var i = 0; i < sortedKeysFromObjectRowsData.length; i++) {
        filteredObjectRowsData[sortedKeysFromObjectRowsData[i]] = objectRowsData[sortedKeysFromObjectRowsData[i]].filter(function(x) {
          if (!x.vendorInv && x.dateLastSynced && x.status != "Canceled") {
            return true
          } else {
            return false
          }
        });
      }
      break;
    case "WeightOrDate":
      for (var i = 0; i < sortedKeysFromObjectRowsData.length; i++) {
        filteredObjectRowsData[sortedKeysFromObjectRowsData[i]] = objectRowsData[sortedKeysFromObjectRowsData[i]].filter(function(x) {
          if ((!x.actualDeliveryDate && x.dateLastSynced && x.status != "Canceled") || (!x.finalWeightDestination && !x.finalWeightOrigin && x.dateLastSynced && x.status != "Canceled")) {
            return true
          } else {
            return false
          }
        });
      }
  }
  // Logger.log("filteredObjectRowsData = " + filteredObjectRowsData);
  // Logger.log("filteredObjectRowsData[sortedKeysFromObjectRowsData[0]] = " + filteredObjectRowsData[sortedKeysFromObjectRowsData[0]]);
  return filteredObjectRowsData
}

function createReport(filteredObjectRowsData, sortedKeysFromObjectRowsData, reportType) {
  Logger.log("**** Running createReport() ****");
  var dataForEmail = [];
  var dataTableForEmail = '';
  var totalNumberOfRowsInAllMasterSheets = 0
  var currentMonthString = "";
  var lastMonthString = "";

  var carrierTdTag = '</td><td>';
  var customerTdTag = '</td><td>';
  var vendorTdTag = '</td><td>';

  switch(reportType) {
    case "Carrier":
      carrierTdTag = '</td><td style="background-color:#ffffb2">';
      break;
    case "Customer":
      customerTdTag = '</td><td style="background-color:#ffffb2">';
      break;
    case "Vendor":
      vendorTdTag = '</td><td style="background-color:#ffffb2">';
      break;
    case "WeightOrDate":
      // Don't need to do anything in this case
  }


  // Loop through all months across all masters
  for (var i = 0; i < sortedKeysFromObjectRowsData.length; i++) {
    var numberOfRowsInMasterSheet = filteredObjectRowsData[sortedKeysFromObjectRowsData[i]].length;

    if (numberOfRowsInMasterSheet === 0) {
      continue;
    }

    totalNumberOfRowsInAllMasterSheets += numberOfRowsInMasterSheet;

    // loop through the rows in sheet
    for (var j = 0; j < numberOfRowsInMasterSheet; j++) {
      /*
      // THIS SECTION FOR TESTING
      dataForEmail[j] = [];
      dataForEmail[j].push(sortedKeysFromObjectRowsData[i].substring(0, sortedKeysFromObjectRowsData[i].indexOf('-')));
      dataForEmail[j].push(filteredObjectRowsData[sortedKeysFromObjectRowsData[i]][j]["masterRecord"]);
      dataForEmail[j].push(filteredObjectRowsData[sortedKeysFromObjectRowsData[i]][j]["customer"]);
      dataForEmail[j].push(filteredObjectRowsData[sortedKeysFromObjectRowsData[i]][j]["vendor"]);
      dataForEmail[j].push(filteredObjectRowsData[sortedKeysFromObjectRowsData[i]][j]["carrierShipVia"]);*/

      currentMonthString = sortedKeysFromObjectRowsData[i].substring(0, sortedKeysFromObjectRowsData[i].indexOf('-'));

      if ((lastMonthString != currentMonthString) && (i != sortedKeysFromObjectRowsData.length - 1) && (lastMonthString != "")) {
        dataTableForEmail += '</table><br><table width="600" style="border:1px solid #333"><tr style="border-bottom: solid 1px black"><th style="background-color:#aaaaaa">Date</th><th style="background-color:#aaaaaa">Estimated Ship Date</th><th style="background-color:#aaaaaa">3CC Number</th><th style="background-color:#aaaaaa">Customer</th><th style="background-color:#aaaaaa">Vendor</th><th style="background-color:#aaaaaa">Carrier</th></tr>';
      }

      try { var estimatedShipDateString = filteredObjectRowsData[sortedKeysFromObjectRowsData[i]][j]["estimatedShipDate"]; } catch(err) { var estimatedShipDateString = 'Undefined'; }
      try { var masterRecordString = filteredObjectRowsData[sortedKeysFromObjectRowsData[i]][j]["masterRecord"]; } catch(err) { var masterRecordString = 'Undefined'; }
      try { var customerString = filteredObjectRowsData[sortedKeysFromObjectRowsData[i]][j]["customer"]; } catch(err) { var customerString = 'Undefined'; }
      try { var vendorString = filteredObjectRowsData[sortedKeysFromObjectRowsData[i]][j]["vendor"]; } catch(err) { var vendorString = 'Undefined'; }
      try { var carrierString = filteredObjectRowsData[sortedKeysFromObjectRowsData[i]][j]["carrierShipVia"]; } catch(err) { var carrierString = 'Undefined'; }

      if (Utilities.formatDate(new Date(estimatedShipDateString), "EST", "MMM d yyyy") == "Dec 31 1969") {
        estimatedShipDateStringFormatted = estimatedShipDateString;
      } else {
        estimatedShipDateStringFormatted = Utilities.formatDate(new Date(estimatedShipDateString), "EST", "MMM d yyyy");
      }

      dataTableForEmail += '<tr style="border-bottom: solid 1px black"><td>' + currentMonthString +
                           '</td><td>' + estimatedShipDateStringFormatted +
                           '</td><td style="background-color:#ffffb2">' + masterRecordString +
                           customerTdTag + customerString +
                           vendorTdTag + vendorString +
                           carrierTdTag + carrierString +
                           '</td></tr>';
      lastMonthString = currentMonthString;
    }
  }
  if (totalNumberOfRowsInAllMasterSheets === 0) {
    return null;
  }
  return dataTableForEmail;
}

function sendEmailReport(dataTableForEmail, settings, reportType) {
  Logger.log("**** Running sendEmailReport() ****");
  var d = new Date();
  var month = d.getMonth() + 1; //months from 1-12
  var day = d.getDate();
  var year = d.getFullYear();
  var formatedDate = month + "/" + day + "/" + year;
  var mailBcc = "birnbaum.adam@gmail.com"

  switch(reportType) {
    case "Carrier":
      var mailTo = settings.emailsCarrierWeeklyReport;
      var reportTypeString = "Carrier Invoice Number";
      break;
    case "Customer":
      var mailTo = settings.emailsCustomerWeeklyReport;
      var reportTypeString = "Customer Invoice Number";
      break;
    case "Vendor":
      var mailTo = settings.emailsVendorWeeklyReport;
      var reportTypeString = "Vendor Invoice Number";
      break;
    case "WeightOrDate":
      var mailTo = settings.emailsWeightWeeklyReport;
      var reportTypeString = "Final Weight and/or Actual Delivery Date";
  }
  var mailSubject = formatedDate + ' - Weekly Missing ' + reportTypeString + ' Report';
  if (dataTableForEmail == null) {
    var mailBody = '<p>As of ' + formatedDate + ', in the visible sheets there are no loads missing <strong>' + reportTypeString + '</strong>.</p>';
  } else {
    var mailBody = '<p>As of ' + formatedDate + ', here are the loads missing <strong>' + reportTypeString + '</strong>:</p>';
    mailBody += '<table width="600" style="border:1px solid #333"><tr style="border-bottom: solid 1px black"><th style="background-color:#aaaaaa">Date</th><th style="background-color:#aaaaaa">Estimated Ship Date</th><th style="background-color:#aaaaaa">3CC Number</th><th style="background-color:#aaaaaa">Customer</th><th style="background-color:#aaaaaa">Vendor</th><th style="background-color:#aaaaaa">Carrier</th></tr>';
    mailBody += dataTableForEmail;
    mailBody += '</table><p>End of Report.</p>';
  }

  // Logger.log('mailTo = ' + mailTo);
  // Logger.log('mailSubject = ' + mailSubject);
  // Logger.log('mailBody = ' + mailBody);

  MailApp.sendEmail({
    to: mailTo,
    subject: mailSubject,
    htmlBody: mailBody,
    noReply: true,
    bcc: mailBcc
  });
}

// ******************************************************************** Helper Functions *******************************************************************

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function getMonthName(monthNumber) {
  monthNumber = Number(monthNumber);
  var monthNames = {00:"January",
                    01:"February",
                    02:"March",
                    03:"April",
                    04:"May",
                    05:"June",
                    06:"July",
                    07:"August",
                    08:"September",
                    09:"October",
                    10:"November",
                    11:"December"};
  // Logger.log('monthNames[monthNumber] = ' + monthNames[monthNumber]);
  return monthNames[monthNumber];
}

function getMonthNumber(m) {
  var monthNames = {January:0,
                    February:1,
                    March:2,
                    April:3,
                    May:4,
                    June:5,
                    July:6,
                    August:7,
                    September:8,
                    October:9,
                    November:10,
                    December:11};
  return monthNames[m];
}

// ************************* Sheet processing library functions from https://developers.google.com/apps-script/articles/mail_merge#section4 *************

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

function convertSheetNameAndTraderNameToNumber(sheetNameAndTraderName, settings) {
  Logger.log("**** Running convertSheetNameAndTraderNameToNumber() ****");
  //Logger.log("sheetNameAndTraderName = " + sheetNameAndTraderName);
  var sheetName = sheetNameAndTraderName.substring(0, sheetNameAndTraderName.indexOf('-'));
  var traderName = sheetNameAndTraderName.substring(sheetNameAndTraderName.indexOf('-') + 1);
  var sheetMonthNumber = getMonthNumber(sheetName.substring(0, sheetName.indexOf(" ")));
  var sheetYearNumber = Number(sheetName.substring(sheetName.indexOf(" ") + 1));
  var traderNumber = settings.arrayTraderNames.indexOf(traderName);

  //Logger.log("traderName = " + traderName);
  //Logger.log("settings.arrayTraderNames = " + settings.arrayTraderNames);
  //Logger.log("settings.arrayTraderNames.indexOf(traderName) = " + settings.arrayTraderNames.indexOf(traderName));

  if (!sheetYearNumber) { sheetYearNumber = 9999; }
  if (!sheetMonthNumber && sheetMonthNumber != 0) { sheetMonthNumber = 99; }
  if (!traderNumber && traderNumber != 0) { traderNumber = 99; }

  sheetYearNumber = parseInt(sheetYearNumber, 10);
  sheetMonthNumber = parseInt(sheetMonthNumber, 10);
  traderNumber = parseInt(traderNumber, 10);

  var yearMonthTraderNum = sheetYearNumber * 10000 + sheetMonthNumber * 100 + traderNumber;

  //Logger.log("yearMonthTraderNum = " + yearMonthTraderNum);
  //Logger.log("sheetMonthNumber = " + sheetMonthNumber);
  //Logger.log("sheetYearNumber = " + sheetYearNumber);
  //Logger.log("traderNumber = " + traderNumber);

  return yearMonthTraderNum;
}

function convertSheetAndTraderNumberToName(sheetAndTraderNumber, settings) {
  Logger.log("**** Running convertSheetAndTraderNumberToName() ****");
  // Logger.log("sheetAndTraderNumber = " + sheetAndTraderNumber);
  var s = sheetAndTraderNumber.toString();
  var myYearString = s.substring(0,4);
  var myMonthString = s.substring(4,6);
  var myTraderString = s.substring(6);

  if (myYearString == 9999 || myMonthString == 99 || myTraderString == 99) {
    return "Incorrect date format";
  } else {
    var myYear = myYearString;
    var myMonth = getMonthName(myMonthString);
    var myTrader = settings.arrayTraderNames[Number(myTraderString)];
    return (myMonth + " " + myYear + "-" + myTrader);
  }
}

/*
- Change minute trigger to trigger once per week
- regardless all data will be captured even if sheet is named "Augfffst 2017" -- if sheet name is misspelled it will show up at bottom of email with note explaining that it wasn't found in the db of permitted names
*/
