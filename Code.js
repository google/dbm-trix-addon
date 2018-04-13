// Copyright 2018 Google Inc. All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and

'use strict';

/**
 * @fileoverview This addon syncs DoubleClick search web query reports with
 * the trix that it is installed on.
 */

/**
 * Trigger function that fires when spreadsheet loads to create the add-on menu.
 * @param {Object} e The current ScriptApp Object.
 */
function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Select DBM report', 'selectDbmReportDialog')
      .addItem('Refresh current sheet', 'refreshCurrentSheet')
      .addItem('Refresh all sheets', 'refreshAllSheets')
      .addItem('Show last sync time', 'showLastSyncDetails')
      .addItem('Schedule reports', 'scheduleReportsDialog')
      .addSeparator()
      .addItem('Reset Sheet link to DBM Report', 'resetSheetLinkage')
      .addSeparator()
      .addItem('Purge Properties', 'purgeProperties');

  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    menu.addToUi();
  } else {
    try {
      if (APP_ADMIN_EMAILS.indexOf(Session.getActiveUser().getEmail()) > -1) {
        menu.addItem('Show Debug Info', 'showDebugInfo');
      }
    } catch (err) {
      console.error(err);
    } finally {
      menu.addToUi();
    }
  }
}

/**
 * Convenience function to clear all properties set. Used only in debug mode.
 */
function purgeProperties() {
  var ui = SpreadsheetApp.getUi();
  var purgeAlertString =  'Are you sure you wish to purge? This will delete' +
          'your report linkages, permissions and schedules. ' +
          'You will need to authenticate yourself again after you purge.';
  var userResponse = ui.alert(purgeAlertString, ui.ButtonSet.OK_CANCEL);
  if (userResponse == ui.Button.OK) {
    PropertiesService.getDocumentProperties().deleteAllProperties();
    PropertiesService.getUserProperties().deleteAllProperties();

    // Also delete any triggers currently set by this user.
    var triggers =
        ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    var purgeBrowserString = 'All stored properties have been purged. ' +
        'Please authorize yourself again by going to ' +
        'Add-ons -> DBM Add-on -> Select DBM Report';
    Browser.msgBox(purgeBrowserString);
  }
}

/**
 * Resets the DBM Report to sheet linking.
 */
function resetSheetLinkage() {
  var ui = SpreadsheetApp.getUi();
  var resetSheetAlertString = 'This will remove the link. Your data will not ' +
      'be affected, but you won\'t be able to refresh data  ' +
      'until you do the linking again.';
  var response = ui.alert(
      'Resetting DBM Link',
      resetSheetAlertString,
      ui.ButtonSet.OK_CANCEL);

  if (response == ui.Button.OK) {
    var currentSheet = SpreadsheetApp.getActiveSheet();
    var currentSheetId = currentSheet.getSheetId();
    resetReportLink(currentSheetId);
  }
}

/**
 * Displays debug info available for application admin.
 */
function showDebugInfo() {
  var debugInfoTemplate = HtmlService.createTemplateFromFile('DebugInfo');
  debugInfoTemplate.currentSchedule =
      retrieveCurrentSchedule().replace(/<b>/g, '').replace(/<\/b>/g, '');
  debugInfoTemplate.debugInfo = getDebugInfo();
  var ui = debugInfoTemplate.evaluate()
               .setSandboxMode(HtmlService.SandboxMode.IFRAME)
               .setWidth(600)
               .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Debug Info');
}

/**
 * Function that fires when the add-on is installed.
 * @param {Object} e The current ScriptApp Object.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Method that shows a dialog box to choose the report.
 */
function selectDbmReportDialog() {
  var OAuthService = getOAuthService();
  OAuthService = validateOAuthService(OAuthService, false);

  if (!OAuthService.hasAccess()) {
    displayAuthorizationDialog();
  } else {
    var ui = HtmlService.createTemplateFromFile('Dialog')
                 .evaluate()
                 .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                 .setHeight(170);
    SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_SELECT_DBM_REPORT_TITLE);
  }
}

/**
 * Method to show ad dialog with last sync details.
 */
function showLastSyncDetails() {
  var OAuthService = getOAuthService();
  OAuthService = validateOAuthService(OAuthService, false);

  if (!OAuthService.hasAccess()) {
    displayAuthorizationDialog();
  } else {
    var ui = HtmlService.createTemplateFromFile('LastSyncDetails')
                 .evaluate()
                 .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                 .setWidth(400)
                 .setHeight(80);
    SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_LAST_SYNC_DETAILS_TITLE);
  }
}

/**
 * Method that fires when Refresh Current Sheet command is chosen.
 */
function refreshCurrentSheet() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  activeSpreadsheet.toast('Pulling new data', 'Status', -1);
  try {
    pullNewData(activeSpreadsheet);
  } catch (err) {
    activeSpreadsheet.toast(err, 'Status', -1);
    return;
  }
  activeSpreadsheet.toast('Refresh Complete', 'Status', 3);
}

/**
 * Function that pulls fresh data into all DBM Report Linked Sheets.
 * @param {Object} activeSpreadsheet The sheet object in which data needs to
 *     be populated.
 */
function pullNewDataAll(activeSpreadsheet) {
  var sheets = activeSpreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    pullNewData(sheets[i]);
  }
}

/**
 * Method that fires when Refresh All Sheet command is chosen.
 */
function refreshAllSheets() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  activeSpreadsheet.toast('Pulling new data', 'Status', -1);
  try {
    DBM_offlineReportSync();
  } catch (err) {
    activeSpreadsheet.toast(err, 'Status', -1);
    return;
  }
  activeSpreadsheet.toast('Refresh Complete', 'Status', 3);
}

/**
 * Function that handles response when OAuth posts back security token.
 * @param {Request} request The HTTP Request object.
 * @return {Object} HtmlOutput.
 */
function authCallback(request) {
  var OAuthService = getOAuthService();
  var isAuthorized = OAuthService.handleCallback(request);
  if (isAuthorized) {
    var authorizedHTMLString = 'Success! You can close this tab and open the ' +
        'Dialog box again.';
    return HtmlService.createHtmlOutput(authorizedHTMLString);
  } else {
    var unauthorizedHTMLString = 'Denied. You can close this tab. ' +
        'Please try opening the Dialog box again.';
    return HtmlService.createHtmlOutput(unauthorizedHTMLString);
  }
}

/**
 * Gets latest report data from the query id.
 * @param {string} queryId The query id from DBM.
 * @param {string} reportName Name of the report.
 */
function fetchLatestData(queryId, reportName) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  activeSpreadsheet.toast('Pulling DBM report...', 'Status', -1);

  var documentProperties = PropertiesService.getDocumentProperties();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var currentSheetId = currentSheet.getSheetId();

  // Delete properties used by other DDM Add-ons for this sheet, so that it
  // doesn't interfere with GCS Offline Sync.
  documentProperties.deleteProperty(currentSheetId + '_WEBQUERY_URL');
  documentProperties.deleteProperty(currentSheetId + '_PROFILE_ID');
  documentProperties.deleteProperty(currentSheetId + '_REPORT_ID');
  documentProperties.setProperty(currentSheetId + '_QUERY_ID', queryId);
  documentProperties.setProperty(
      currentSheetId + '_DBM_REPORT_NAME', reportName);
  documentProperties.setProperty(
      currentSheetId + '_DBM_REPORT_SETUP_USER',
      Session.getActiveUser().getEmail());
  pullNewData(currentSheet);
  var fetchDataString = 'DBM report data added. You can now manually refresh ' +
      'or setup a scheduled sync';
  activeSpreadsheet.toast(fetchDataString, 'Status', 5);
}

/**
 * Logs to the console log as an error message.
 * @param {string} msg The message to log.
 */
function logError(msg) {
  var spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var errorMsg = msg + ', Spreadsheet name: ' + spreadsheetName +
      ', Spreadsheet URL: ' + spreadsheetUrl +
      ', Active sheet: ' + currentSheet;
  console.error(errorMsg);
}

/**
 * Logs to the console log as an info message.
 * @param {string} msg The message to log.
 */
function logInfo(msg) {
  var spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var infoMsg = msg + ', Spreadsheet name: ' + spreadsheetName +
      ', Spreadsheet URL: ' + spreadsheetUrl +
      ', Active sheet: ' + currentSheet;
  console.info(infoMsg);
}

/**
 * Checks if the document properties has a bucket name saved and replaces
 * it with the query id. This is needed for the DBM API changes which don't
 * allow access to cloud buckets anymore.
 * @param {string} sheetId The sheet unique id in the spreadsheet.
 * @throws if the bucketName and queryId is invalid.
 * @return {string} The query id of the DBM query.
 */
function replaceBucketNameWithQueryId(sheetId) {
  var queryId;
  var urlFetchResponse;
  var responseObject;
  var allQueries;
  var isFound;

  var documentProperties = PropertiesService.getDocumentProperties();
  // Check if user is using the bucket name and migrate it to use query id.
  var bucketName = documentProperties.getProperty(sheetId + '_BUCKET_NAME');
  var queryId = documentProperties.getProperty(sheetId + '_QUERY_ID');
  var OAuthService = getOAuthService();
  OAuthService = validateOAuthService(OAuthService, true);

  if (!bucketName && !queryId) {
    throw new Error('Something is wrong. Might need to link the sheet again.');
    logError('Something is wrong. Might need to link the sheet again.');
  }
  if (!bucketName && queryId) {
    // Query id exists. Nothing to do here.
    return queryId;
  }
  if (bucketName && queryId) {
    // Both exist. Get rid of bucket name. Ideally should never come here.
    documentProperties.deleteProperty(sheetId + '_BUCKET_NAME');
    return queryId;
  }

  try {
    urlFetchResponse = UrlFetchApp.fetch(
        dbmGetReportURL,
        {headers: {Authorization: 'Bearer ' + OAuthService.getAccessToken()}});
  } catch (err) {
    logError('fetch all reports failed. Reason: ' + err);
    var authExpiredErrorString = 'Error fetching all reports. DBM API ' +
        'Credentials expired. Please click to re - authorize < a href =\'' +
        OAuthService.getAuthorizationUrl() + '\' target =\'_blank\'> here </a>';
    throw new Error(authExpiredErrorString);
    OAuthService.reset();
  }

  if (!urlFetchResponse) {
    var emptyResponseErrorString = 'Empty response from DBM API when ' +
        'querying for all reports';
    throw new Error(emptyResponseErrorString);
    logError(emptyResponseErrorString);
  }

  responseObject = JSON.parse(urlFetchResponse.getContentText());

  if (!responseObject || !responseObject['queries']) {
    var responseErrorString = 'No or Empty response from DBM API when ' +
        'querying for all reports';
    throw new Error(responseErrorString);
    logError(responseErrorString);
  }

  allQueries = responseObject['queries'];
  isFound = false;

  for (var j = 0; j < allQueries.length; j++) {
    var latestPath =
        allQueries[j]['metadata']['googleCloudStoragePathForLatestReport'];
    if (!latestPath) {
      // Ignore any queries that don't have a path yet. Refer b/64101859.
      continue;
    }
    var regExp = new RegExp('[0-1]*[^/]*_report');
    var gcsBucket = regExp.exec(latestPath)[0];
    if (gcsBucket == bucketName) {
      queryId = allQueries[j]['queryId'];
      isFound = true;
      break;
    }
  }
  if (!isFound) {
    var queryNotFoundErrorString = 'Couldn\'t find the relevant query. ' +
        'Perhaps you do not own the DBM report that the current' +
        ' sheet is linked to. Please ask the person' +
        ' who did the original linking to refresh it once. ' +
        'Alternately, please remove the linkage using the addon menu ' +
        'and link the report to the sheet again .';
    throw new Error(queryNotFoundErrorString);
    logError(queryNotFoundErrorString);
  }

  // Remove the bucket name and add query id.
  documentProperties.deleteProperty(sheetId + '_BUCKET_NAME');
  documentProperties.setProperty(sheetId + '_QUERY_ID', queryId);
  return queryId;
}

/**
 * Method that pulls the latest data using the report query ID
 * and puts in the sheet.
 * @param {Object} sheet The sheet in which data needs to be populated.
 *     If null, use current Sheet.
 * @throws If error fetching report with ID due to DBM API credentials expired.
 * @throws If empty response from DBM API call is returned.
 * @throws If none or empty response from DBM Report Query.
 */
function pullNewData(sheet) {
  var SpreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var SpreadsheetURL = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  sheet = sheet || currentSheet;
  var OAuthService = getOAuthService();
  OAuthService = validateOAuthService(OAuthService, true);
  var documentProperties = PropertiesService.getDocumentProperties();
  var sheetId = sheet.getSheetId();
  var queryId = documentProperties.getProperty(sheetId + '_QUERY_ID');
  var response;
  var responseObject;
  var v2apiBucketName = 'ddm-xbid';
  var latestFileURL;
  var fileContent = null;
  var apiVersion = 1;

  if (!queryId) {
    queryId = replaceBucketNameWithQueryId(sheetId);
  }

  try {
    response = UrlFetchApp.fetch(
        dbmFetchQueryURL + '/' + queryId,
        {headers: {Authorization: 'Bearer ' + OAuthService.getAccessToken()}});
  } catch (err) {
    console.error('fetchReport failed. Reason: ' + err);
    throw new Error(
        'Error fetching report with ID. DBM API Credentials expired. ' +
        'Please click to re-authorize <a href=\'' +
        OAuthService.getAuthorizationUrl() + '\' target=\'_blank\'> here </a>');
    OAuthService.reset();
  }

  if (!response || !response.getContentText()) {
    console.error(
        'Empty response from DBM API call, maybe report doesn\'t exist ' +
        dbmGetReportURL);
    throw new Error(
        'Empty response from DBM API call, maybe report doesn\'t exist ' +
        dbmGetReportURL);
  }

  responseObject = JSON.parse(response.getContentText());

  if (!responseObject || !responseObject['metadata']) {
    console.error(
        'No or Empty response from DBM Report Query ' + dbmGetReportURL);
    throw new Error(
        'No or Empty response from DBM Report Query ' + dbmGetReportURL);
  }

  latestFileURL =
      responseObject['metadata']['googleCloudStoragePathForLatestReport'];
  fileContent = null;

  // Checking for api version to send headers or not.
  // v1 relevant content may be removed once fully deprecated.
  if (latestFileURL.indexOf(v2apiBucketName) >
      0) {  // This is v2 report. Don't pass oAuth headers.
    fileContent = UrlFetchApp.fetch(latestFileURL);
    apiVersion = 2;
  } else {  // v1 report needs oAuth headers.
    fileContent = UrlFetchApp.fetch(
        latestFileURL,
        {headers: {Authorization: 'Bearer ' + OAuthService.getAccessToken()}});
  }

  // Populate the contents in the sheet
  populateSpreadsheet(fileContent.getContentText(), sheetId, apiVersion);
  console.info(
      'DBM - Sheet Refreshed with new data: ' + queryId + ' for ' +
      SpreadsheetName + '. Sheet Name: ' + sheet.getName() +
      '. Url: ' + SpreadsheetURL);

  var end_time = new Date();
  var latestFileUpdatedDate =
      new Date(responseObject['metadata']['latestReportRunTimeMs']);
  documentProperties.setProperty(sheetId + '_LAST_SYNC', end_time);
  documentProperties.setProperty(
      sheetId + '_DBM_REPORT_UPDATED_DATE', latestFileUpdatedDate);
}

/**
 * Helper Function to read content stream and put it in the current active
 * spreadsheet.
 * @param {string} csvContent - CSV stream.
 * @param {string} currentSheetId - Google Sheet ID.
 * @param {string} apiVersion - version of available API (v1 or v2).
 */
function populateSpreadsheet(csvContent, currentSheetId, apiVersion) {
  var rows = Utilities.parseCsv(csvContent);
  var lastRowWithData = 1;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] == '') lastRowWithData = i;
  }

  var numberOfRows = rows.length;

  // Remove apiVersion after the move to v2 is complete.
  var actualLastRow = (apiVersion == 1) ? lastRowWithData : lastRowWithData + 1;

  for (var i = actualLastRow; i <= numberOfRows; i++) {
    rows.pop();
  }

  if (rows && rows.length && rows[0] && rows[0].length) {
    // Get the existing spreadsheet using the specified filename
    var sheet = getSheetById(currentSheetId);

    // Clear the existing rows first. Only the values, retain the formatting.
    sheet.clearContents();

    // console.info("Number of rows: %s and Number of columns
    // %s",rows.length,rows[0].length);

    // Output the result rows
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

    // Truncate the sheet to number of rows with content
    if (sheet.getMaxRows() > sheet.getLastRow())
      sheet.deleteRows(
          sheet.getLastRow() + 1, sheet.getMaxRows() - sheet.getLastRow());

    // Truncate the sheet to number of columns with content
    if (sheet.getMaxColumns() > sheet.getLastColumn())
      sheet.deleteColumns(
          sheet.getLastColumn() + 1,
          sheet.getMaxColumns() - sheet.getLastColumn());
  }
}

/**
 * Method that can be used with App Script Triggers to do an offline sync.
 * @throws if access not granted or expired.
 */
function DBM_offlineReportSync() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var currentSheetId = '';
  var queryId = '';
  var bucketName = '';
  var SpreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var SpreadsheetURL = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  console.info(
      'Entered DBM offline sync method for: ' + SpreadsheetName +
      '. Url: ' + SpreadsheetURL);

  var permissions = checkPermissions();
  if (!permissions.authorized) {
    var lastAuthEmailDate = documentProperties.getProperty('lastAuthEmailDate');
    var today = new Date().toDateString();
    if (lastAuthEmailDate != today) {
      if (MailApp.getRemainingDailyQuota() > 0) {
        var html = HtmlService.createTemplateFromFile('AuthorizationEmail');
        html.url = permissions.authorizationURL;
        html.ADDON_TITLE = ADDON_TITLE;
        var message = html.evaluate();

        MailApp.sendEmail(
            SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(),
            ADDON_TITLE + ' - Authorization Required', message.getContent(),
            {name: ADDON_TITLE, htmlBody: message.getContent()});
      }
      documentProperties.setProperty('lastAuthEmailDate', today);
    }
  } else {
    console.info(
        'Started DBM offline sync for: ' + SpreadsheetName +
        '. Url: ' + SpreadsheetURL);
    for (var i = 0; i < sheets.length; i++) {
      currentSheetId = sheets[i].getSheetId();
      queryId = documentProperties.getProperty(currentSheetId + '_QUERY_ID');
      bucketName =
          documentProperties.getProperty(currentSheetId + '_BUCKET_NAME');

      // If there is sheet level query sync needed
      if (queryId || bucketName) {
        try {
          console.info(
              'Started DBM offline sync for: ' + SpreadsheetName +
              '. Sheet Name: ' + sheets[i].getName() +
              '. Url: ' + SpreadsheetURL);
          pullNewData(sheets[i]);
          console.info(
              'Finished DBM offline sync for: ' + SpreadsheetName +
              '. Sheet Name: ' + sheets[i].getName() +
              '. Url: ' + SpreadsheetURL);
        } catch (err) {
          // Credentials expired already a caught exception
          if (err.toString().indexOf('Access not granted or expired') > -1) {
            throw err;
            return;
          }

          MailApp.sendEmail({
            to: SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(),
            subject: ADDON_TITLE + ' - Offline Sync Failed',
            htmlBody:
                'Sheet URL: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl() +
                '#gid=' + currentSheetId + '<br><br>' + err
          });
          console.error(
              ADDON_TITLE + ' - Offline Sync Failed for Sheet URL: ' +
              SpreadsheetURL + '#gid=' + currentSheetId + '<br><br>' + err);
        }
      }
    }
  }
}

/**
 * Method that creates OAuth2 Object using OAuth2 Library for Appscript.
 * https://github.com/googlesamples/apps-script-oauth2
 * @return {OAuth2} OAuthService object.
 */
function getOAuthService() {
  return OAuth2
      .createService(serviceName)

      // Set the endpoint URLs, which are the same for all Google services.
      .setAuthorizationBaseUrl(authorizationUrl)
      .setTokenUrl(tokenAccessUrl)

      // Set the client ID and secret, from the Google Developers Console.
      .setClientId(clientId)
      .setClientSecret(clientSecret)

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      // Set the scopes to request (space-separated for Google services).
      .setScope(scope)

      // Below are Google-specific OAuth2 parameters.

      // Sets the login hint, which will prevent the account chooser screen
      // from being shown to users logged in with multiple accounts.
      .setParam('login_hint', Session.getActiveUser().getEmail())

      // Requests offline access.
      .setParam('access_type', 'offline')

      // Forces the approval prompt every time. This is useful for testing,
      // but not desirable in a production application.
      .setParam('approval_prompt', 'force');
}

/**
 * Clear oAuth Token, if a user manually revokes access to this app under Google
 * Security, then we need to clear the OAuth Token running this method.
 */
function clearToken() {
  var OAuthService = getOAuthService();
  OAuthService.reset();
}

/**
 * Helper Function to get Sheet from Sheet ID.
 * @param {string} sheetId Sheet ID of the Google Spreadsheet.
 * @return {Array} sheets Returns array of sheets.
 */
function getSheetById(sheetId) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length;
       i++) {  // iterate all sheets and compare ids, unfortunately there is no
               // default method available
    if (sheets[i].getSheetId() == sheetId) {
      return sheets[i];
    }
  }
}

/**
 * Function returning list of reports.
 * @return {Array} an array of report objects.
 * @throws If error fetching report with ID due to DBM API credentials expired.
 * @throws If empty response from DBM API call is returned.
 * @throws If none or empty response from DBM Report Query.
 */
function getReportList() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var sheetQueryId = documentProperties.getProperty(
      SpreadsheetApp.getActiveSheet().getSheetId() + '_QUERY_ID');

  var reports = [];
  var OAuthService = getOAuthService();
  var response;
  var responseObject;

  try {
    response = UrlFetchApp.fetch(
        dbmGetReportURL,
        {headers: {Authorization: 'Bearer ' + OAuthService.getAccessToken()}});
  } catch (err) {
    console.error('getReportList failed. Reason: ' + err);
    var authErrorString = 'Error fetching report list. DBM API  ' +
        'Credentials expired. Please click to re-authorize <a href=\'' +
        OAuthService.getAuthorizationUrl() + '\' target=\'_blank\'> here ' +
        '</a>';
    throw new Error(authErrorString);
    OAuthService.reset();
  }

  if (!response || !response.getContentText()) {
    console.error('Empty response from DBM API call ' + dbmGetReportURL);
    throw new Error('Empty response from DBM API call ' + dbmGetReportURL);
  }

  responseObject = JSON.parse(response.getContentText());

  if (!responseObject || !responseObject.queries) {
    console.error(
        'No or Empty response from DBM Report Query ' + dbmGetReportURL);
    throw new Error(
        'No or Empty response from DBM Report Query ' + dbmGetReportURL);
  }

  for (var i = 0; i < responseObject.queries.length; i++) {
    var reportMetaData = responseObject.queries[i].metadata;
    var queryId = responseObject.queries[i].queryId;
    if (!reportMetaData) continue;

    var linkedReport = false;

    if (sheetQueryId == queryId) {
      linkedReport = true;
    } else {
      linkedReport = false;
    }

    var reportObject = {
      'title': reportMetaData.title,
      'linked': linkedReport,
      'query_id': queryId
    };
    reports.push(reportObject);
  }
  return reports;
}

/**
 * Shows a dialog to schedule the report sync.
 */
function scheduleReportsDialog() {
  var OAuthService = getOAuthService();
  OAuthService = validateOAuthService(OAuthService, false);

  if (!OAuthService.hasAccess()) {
    displayAuthorizationDialog();
  } else {
    var ui = HtmlService.createTemplateFromFile('Scheduler')
                 .evaluate()
                 .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                 .setWidth(700)
                 .setHeight(150);
    SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_SCHEDULE_REPORTS_TITLE);
  }
}

/**
 * Setup a Apps Script trigger to schedule the report sync.
 * @param {boolean} enableScheduler Check if the scheduler is to be enabled
 * or not.
 * @param {string} frequency What frequency to run the report at.
 * @param {string} timer The day of week etc. to use for the scheduling.
 * @param {string} timeOfDaySelected The time of day to use.
 * @return {string} The scheduled summary.
 */
function setScheduler(enableScheduler, frequency, timer, timeOfDaySelected) {
  var documentProperties, triggerSettings;
  documentProperties = PropertiesService.getDocumentProperties();

  // Delete the trigger previously setup
  var triggers = ScriptApp.getProjectTriggers();

  for (i = 0; i < triggers.length; i++) {
    // Delete the old trigger on the document
    if (triggers[i].getUniqueId() ==
        documentProperties.getProperty('DBM_Trigger_ID')) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  if (!enableScheduler) {
    documentProperties.deleteProperty('DBM_Schedule_Frequency');
    documentProperties.deleteProperty('DBM_Schedule_Time');
    documentProperties.deleteProperty('DBM_Schedule_Time2');
    documentProperties.deleteProperty('DBM_Trigger_Created_By');
    documentProperties.deleteProperty('DBM_Trigger_ID');
    return '';
  }

  switch (frequency) {
    case 'hourly':
      var triggerID = ScriptApp.newTrigger('DBM_offlineReportSync')
                          .timeBased()
                          .everyHours(timer)
                          .create()
                          .getUniqueId();

      triggerSettings = {
        'DBM_Schedule_Frequency': frequency,
        'DBM_Schedule_Time': timer,
        'DBM_Schedule_Time2': null,
        'DBM_Trigger_Created_By': Session.getActiveUser().getEmail(),
        'DBM_Trigger_ID': triggerID
      };
      documentProperties.setProperties(triggerSettings);
      break;

    case 'daily':
      var triggerID = ScriptApp.newTrigger('DBM_offlineReportSync')
                          .timeBased()
                          .everyDays(1)
                          .atHour(timer)
                          .create()
                          .getUniqueId();

      triggerSettings = {
        'DBM_Schedule_Frequency': frequency,
        'DBM_Schedule_Time': timer,
        'DBM_Schedule_Time2': null,
        'DBM_Trigger_Created_By': Session.getActiveUser().getEmail(),
        'DBM_Trigger_ID': triggerID
      };
      documentProperties.setProperties(triggerSettings);
      break;

    case 'weekly':
      var onWeekDay = ScriptApp.WeekDay.MONDAY;
      switch (timer) {
        case 'MONDAY':
          onWeekDay = ScriptApp.WeekDay.MONDAY;
          break;
        case 'TUESDAY':
          onWeekDay = ScriptApp.WeekDay.TUESDAY;
          break;
        case 'WEDNESDAY':
          onWeekDay = ScriptApp.WeekDay.WEDNESDAY;
          break;
        case 'THURSDAY':
          onWeekDay = ScriptApp.WeekDay.THURSDAY;
          break;
        case 'FRIDAY':
          onWeekDay = ScriptApp.WeekDay.FRIDAY;
          break;
        case 'SATURDAY':
          onWeekDay = ScriptApp.WeekDay.SATURDAY;
          break;
        case 'SUNDAY':
          onWeekDay = ScriptApp.WeekDay.SUNDAY;
          break;
      }
      var triggerID = ScriptApp.newTrigger('DBM_offlineReportSync')
                          .timeBased()
                          .everyWeeks(1)
                          .onWeekDay(onWeekDay)
                          .atHour(timeOfDaySelected)
                          .create()
                          .getUniqueId();

      triggerSettings = {
        'DBM_Schedule_Frequency': frequency,
        'DBM_Schedule_Time': timer,
        'DBM_Schedule_Time2': timeOfDaySelected,
        'DBM_Trigger_Created_By': Session.getActiveUser().getEmail(),
        'DBM_Trigger_ID': triggerID
      };
      documentProperties.setProperties(triggerSettings);
      break;
  }
  return retrieveCurrentSchedule();
}

/**
 * Retrieve the current sync schedule in human readable format.
 * @return {string} Sync schedule in human readable format.
 */
function retrieveCurrentSchedule() {
  // Only returns the triggers for the current user for the DBM Add-on for this
  // document
  var triggers =
      ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());
  var documentProperties = PropertiesService.getDocumentProperties();

  // If the user who setup the Scheduler is same as the current user, check if
  // the trigger stil exists. If it doesn't exist delete the properties set.
  if (documentProperties.getProperty('DBM_Trigger_Created_By') ==
      Session.getActiveUser().getEmail()) {
    var doesTriggerExist = false;
    for (i = 0; i < triggers.length; i++) {
      if (triggers[i].getUniqueId() ==
          documentProperties.getProperty('DBM_Trigger_ID')) {
        doesTriggerExist = true;
      }
    }
    if (!doesTriggerExist) {
      documentProperties.deleteProperty('DBM_Schedule_Frequency');
      documentProperties.deleteProperty('DBM_Schedule_Time');
      documentProperties.deleteProperty('DBM_Schedule_Time2');
      documentProperties.deleteProperty('DBM_Trigger_Created_By');
      documentProperties.deleteProperty('DBM_Trigger_ID');
      return '';
    }
  }
  var currentFrequency =
      documentProperties.getProperty('DBM_Schedule_Frequency');
  var currentTimer = documentProperties.getProperty('DBM_Schedule_Time');
  var currentTimer2 = documentProperties.getProperty('DBM_Schedule_Time2');
  var createdBy = documentProperties.getProperty('DBM_Trigger_Created_By');
  var strReadableSchedule = '';

  switch (currentFrequency) {
    case 'hourly':
      strReadableSchedule +=
          '<b>Current Sync Schedule:</b> Every ' + currentTimer + ' Hours';
      break;

    case 'daily':
      strReadableSchedule += '<b>Current Sync Schedule:</b> Daily between ' +
          daily_frequency[currentTimer].text;
      break;

    case 'weekly':
      strReadableSchedule += '<b>Current Sync Schedule:</b> Weekly on ' +
          currentTimer + ' between ' + daily_frequency[currentTimer2].text;
      break;
  }
  if (createdBy && strReadableSchedule)
    strReadableSchedule += '. Created by: <b>' + createdBy + '</b>';
  return strReadableSchedule;
}

/**
 * Function unlinking DBM report for the sheet.
 */
function unlinkReport() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
      'Unlinking DBM report...', 'Status', -1);

  var currentSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  resetReportLink(currentSheetId);

  SpreadsheetApp.getActiveSpreadsheet().toast(
      'DBM report unlinked for this sheet!', 'Status', 5);
}

/**
 * Function resetting report link for a given sheet.
 * @param {string} sheetId Id of the sheet to have link reseted.
 */
function resetReportLink(sheetId) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty(sheetId + '_LAST_SYNC');
  documentProperties.deleteProperty(sheetId + '_DBM_REPORT_UPDATED_DATE');
  documentProperties.deleteProperty(sheetId + '_QUERY_ID');
  documentProperties.deleteProperty(sheetId + '_BUCKET_NAME');
  documentProperties.deleteProperty(sheetId + '_DBM_REPORT_SETUP_USER');
  documentProperties.deleteProperty(sheetId + '_DBM_REPORT_NAME');
}

/**
 * Function returning setup of the report.
 * @return {string} report setup user concatenated with report name.
 */
function getLinkedReportSetup() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var currentSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  var reportSetupUser =
      documentProperties.getProperty(currentSheetId + '_DBM_REPORT_SETUP_USER');
  var reportName =
      documentProperties.getProperty(currentSheetId + '_DBM_REPORT_NAME');
  var bucketName =
      documentProperties.getProperty(currentSheetId + '_BUCKET_NAME');
  if (bucketName) {
    replaceBucketNameWithQueryId(currentSheetId);
  }
  return {'reportSetupUser': reportSetupUser, 'reportName': reportName};
}

/**
 * Helper function returning debugging info.
 * @return {Array} debugInfo array of debug info objects.
 */
function getDebugInfo() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var documentProperties = PropertiesService.getDocumentProperties();
  var debugInfo = [];
  var debugInfoRow = {};

  for (var i = 0; i < sheets.length; i++) {
    var currentSheetId = sheets[i].getSheetId();
    var queryId = documentProperties.getProperty(currentSheetId + '_QUERY_ID');

    if (queryId == null || queryId == '') continue;
    var lastSyncDetails =
        documentProperties.getProperty(currentSheetId + '_LAST_SYNC');

    if (lastSyncDetails == null) lastSyncDetails = 'Never';

    debugInfoRow = {
      'SheetName': sheets[i].getSheetName(),
      'SheetId': currentSheetId,
      'QueryId': documentProperties.getProperty(currentSheetId + '_QUERY_ID'),
      'DBM_Report_Updated_Date': documentProperties.getProperty(
          currentSheetId + '_DBM_REPORT_UPDATED_DATE'),
      'Linked_DBM_Report_Name':
          documentProperties.getProperty(currentSheetId + '_DBM_REPORT_NAME'),
      'DBM_Report_Setup_User': documentProperties.getProperty(
          currentSheetId + '_DBM_REPORT_SETUP_USER')
    };
    debugInfo.push(debugInfoRow);
  }
  return debugInfo;
}

/**
 * Function checking authorization permission status.
 * @return {Object} authorization status object.
 */
function checkPermissions() {
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    console.warn(
        'DBM Report Add-on: Auth Required for: ' + SpreadsheetName +
        '. Url: ' + SpreadsheetURL);

    return {
      'authorized': false,
      'authorizationURL': authInfo.getAuthorizationUrl()
    };
  } else {
    return {'authorized': true, 'authorizationURL': ''};
  }
}

/**
 * Function displaying authorization dialog box.
 */
function displayAuthorizationDialog() {
  var ui = HtmlService.createTemplateFromFile('Authorization')
               .evaluate()
               .setSandboxMode(HtmlService.SandboxMode.IFRAME)
               .setWidth(400)
               .setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_AUTHORIZATION_NEEDED_TITLE);
}

/**
 * Function validating OAuthService access.
 * @param {Object} OAuthService Object for verification.
 * @param {boolean} sendEmail If true sends an email if DBM API
 *     credentials are expired.
 * @throws if email about DBM API credentials expired was sent.
 * @return {Object} authorization status object.
 */
function validateOAuthService(OAuthService, sendEmail) {
  if (!(typeof OAuthService.hasAccess == 'function' &&
        typeof OAuthService.getAccessToken == 'function')) {
    OAuthService.reset();
    if (sendEmail) {
      logError('DBM API Credentials Expired');
      // Send email to sheet owner to renew the OAuth Token
      MailApp.sendEmail({
        to: SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(),
        subject: ADDON_TITLE + ' - DBM API Credentials Expired',
        htmlBody:
            'The security token for running offline syncs for DBM report ' +
            'sync has expired. <br>To renew it please click <a href="' +
            OAuthService.getAuthorizationUrl() + '" target="_blank">here</a>'
      });
      throw (' Please check your email for instructions');
      logError(err);
    }
  }
  return OAuthService;
}

/**
 * Log data in Apps Script console.
 * @param {string} logText Text to log in the console.
 */
function doLog(logText) {
  console.info(logText);
}
