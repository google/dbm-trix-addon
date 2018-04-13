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
// limitations under the License.

'use strict';

/**
 * @fileoverview This file contains all the constants used in the addon.
 */

/**
 * Service name to use for GCS oauth service.
 * @const {string}
 */
var serviceName = 'GCSAPI';

/**
 * Scopes used by the addon.
 * @const {string}
 */
var scope = 'https://www.googleapis.com/auth/devstorage.read_only' +
    ' https://www.googleapis.com/auth/doubleclickbidmanager';

/**
 * URI to request oauth token.
 * @const {string}
 */
var requestTokenUrl = 'https://www.google.com/accounts/OAuthGetRequestToken';

/**
 * URI to access oauth token.
 * @const {string}
 */
var tokenAccessUrl = 'https://accounts.google.com/o/oauth2/token';

/**
 * Oauth authorization URL.
 * @const {string}
 */
var authorizationUrl = 'https://accounts.google.com/o/oauth2/auth';

/**
 * Name of the addon.
 * @const {string}
 */
var ADDON_TITLE = 'DBM Report - Google Sheets Addon';

/**
 * Admin users to show additional options in the menu.
 * @const {Array<string>}
 */
var APP_ADMIN_EMAILS = [];

/**
 * Service name to use for DBM oauth service.
 * @const {string}
 */
var dbmServiceName = 'DBMAPI';

/**
 * Request URI for DBM API for queries list.
 * @const {string}
 */
var dbmGetReportURL =
    'https://www.googleapis.com/doubleclickbidmanager/v1/queries';

/**
 * Request URI for DBM API for single query.
 * @const {string}
 */
var dbmFetchQueryURL =
    'https://www.googleapis.com/doubleclickbidmanager/v1/query';

/**
 * Google developer console client id.
 * @const {string}
 */
var clientId =
    'insert_client_id';

/**
 * Google developer console client secret.
 * @const {string}
 */
var clientSecret = 'insert_secret';

/**
 * OAuth2 library callback URL format.
 * @const {string}
 */
var redirectUri = 'https://script.google.com/macros/d/{SCRIPT ID}/usercallback';

/**
 * Path to the Google Cloud Storage bucket containing the latest report.
 * @const {string}
 */
var cloudPathName = 'googleCloudStoragePathForLatestReport';

/**
 * Text mapping for daily frequency values.
 * @const {Array<Object>}
 */
var daily_frequency = [
  {'value': 0, 'text': 'Midnight to 1am'},
  {'value': 1, 'text': '1am to 2am'},
  {'value': 2, 'text': '2am to 3am'},
  {'value': 3, 'text': '3am to 4am'},
  {'value': 4, 'text': '4am to 5am'},
  {'value': 5, 'text': '5am to 6am'},
  {'value': 6, 'text': '6am to 7am'},
  {'value': 7, 'text': '7am to 8am'},
  {'value': 8, 'text': '8am to 9am'},
  {'value': 9, 'text': '9am to 10am'},
  {'value': 10, 'text': '10am to 11am'},
  {'value': 11, 'text': '11am to noon'},
  {'value': 12, 'text': 'noon to 1pm'},
  {'value': 13, 'text': '1pm to 2pm'},
  {'value': 14, 'text': '2pm to 3pm'},
  {'value': 15, 'text': '3pm to 4pm'},
  {'value': 16, 'text': '4pm to 5pm'},
  {'value': 17, 'text': '5pm to 6pm'},
  {'value': 18, 'text': '6pm to 7pm'},
  {'value': 19, 'text': '7pm to 8pm'},
  {'value': 20, 'text': '8pm to 9pm'},
  {'value': 21, 'text': '9pm to 10pm'},
  {'value': 22, 'text': '10pm to 11pm'},
  {'value': 23, 'text': '11pm to midnight'}
];

/**
 * Dialog windows headers.
 * @const {string}
 */
var DIALOG_SELECT_DBM_REPORT_TITLE = 'DBM Report';
var DIALOG_AUTHORIZATION_NEEDED_TITLE = 'Authorization Needed';
var DIALOG_LAST_SYNC_DETAILS_TITLE = 'Last Sync Details';
var DIALOG_SCHEDULE_REPORTS_TITLE = 'Schedule Reports';
