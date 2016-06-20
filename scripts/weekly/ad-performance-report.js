// Copyright 2015, Google Inc. All Rights Reserved.
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

/**
 * @name Ad Performance Report
 *
 * @overview The Ad Performance Report generates a Google Spreadsheet that
 *     contains ad performance stats like Impressions, Cost, Click Through Rate,
 *     etc. as several distribution charts for an advertiser account. See
 *     https://developers.google.com/adwords/scripts/docs/solutions/ad-performance
 *     for more details.
 *
 * @author AdWords Scripts Team [adwords-scripts@googlegroups.com]
 *
 * @version 1.0.1
 *
 * @changelog
 * - version 1.0.1
 *   - Improvements to time zone handling.
 * - version 1.0
 *   - Released initial version.
 */

// Comma-separated list of recipients. Comment out to not send any emails.
var RECIPIENT_EMAIL = emailAddress;

// URL of the default spreadsheet template. This should be a copy of
// https://goo.gl/21FW5i
var SPREADSHEET_URL = adPerformanceReportSpreadsheet;

/**
 * This script computes an Ad performance report
 * and outputs it to a Google spreadsheet.
 */
function main() {
  Logger.log('Using template spreadsheet - %s.', SPREADSHEET_URL);
  var spreadsheet = copySpreadsheet(SPREADSHEET_URL);
  Logger.log('Generated new reporting spreadsheet %s based on the template ' +
      'spreadsheet. The reporting data will be populated here.',
      spreadsheet.getUrl());

  var headlineSheet = spreadsheet.getSheetByName('Headline');
  headlineSheet.getRange(1, 2, 1, 1).setValue('Date');
  headlineSheet.getRange(1, 3, 1, 1).setValue(new Date());
  var finalUrlSheet = spreadsheet.getSheetByName('Final Url');
  finalUrlSheet.getRange(1, 2, 1, 1).setValue('Date');
  finalUrlSheet.getRange(1, 3, 1, 1).setValue(new Date());
  spreadsheet.getRangeByName('account_id_headline').setValue(
      AdWordsApp.currentAccount().getCustomerId());
  spreadsheet.getRangeByName('account_id_final_url').setValue(
      AdWordsApp.currentAccount().getCustomerId());

  outputSegmentation(headlineSheet, 'Headline', function(ad) {
    return ad.getHeadline();
  });
  outputSegmentation(finalUrlSheet, 'Final Url', function(ad) {
    return ad.urls().getFinalUrl();
  });
  Logger.log('Ad performance report available at\n' + spreadsheet.getUrl());
  if (RECIPIENT_EMAIL) {
    MailApp.sendEmail(RECIPIENT_EMAIL,
      'Ad Performance Report is ready',
      spreadsheet.getUrl());
  }
}

/**
 * Retrieves the spreadsheet identified by the URL.
 * @param {string} spreadsheetUrl The URL of the spreadsheet.
 * @return {SpreadSheet} The spreadsheet.
 */
function copySpreadsheet(spreadsheetUrl) {
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl).copy(
      'Ad Performance Report - ' +
      getDateStringInTimeZone('MMM dd, yyyy HH:mm:ss z'));

  // Make sure the spreadsheet is using the account's timezone.
  spreadsheet.setSpreadsheetTimeZone(AdWordsApp.currentAccount().getTimeZone());
  return spreadsheet;
}

/**
 * Generates statistical data for this segment.
 * @param {Sheet} sheet Sheet to write to.
 * @param {string} segmentName The Name of this segment for the header row.
 * @param {function(AdWordsApp.Ad): string} segmentFunc Function that returns
 *        a string used to segment the results by.
 */
function outputSegmentation(sheet, segmentName, segmentFunc) {
  // Output header row.
  var rows = [];
  var header = [
    segmentName,
    'Num Ads',
    'Impressions',
    'Clicks',
    'CTR (%)',
    'Cost'
  ];
  rows.push(header);

  var segmentMap = {};

  // Compute data.
  var adIterator = AdWordsApp.ads()
      .forDateRange('LAST_WEEK')
      .withCondition('Impressions > 0').get();
  while (adIterator.hasNext()) {
    var ad = adIterator.next();
    var stats = ad.getStatsFor('LAST_WEEK');
    var segment = segmentFunc(ad);
    if (!segmentMap[segment]) {
      segmentMap[segment] = {
        numAds: 0,
        totalImpressions: 0,
        totalClicks: 0,
        totalCost: 0.0
      };
    }
    var data = segmentMap[segment];
    data.numAds++;
    data.totalImpressions += stats.getImpressions();
    data.totalClicks += stats.getClicks();
    data.totalCost += stats.getCost();
  }

  // Write data to our rows.
  for (var key in segmentMap) {
    if (segmentMap.hasOwnProperty(key)) {
      var ctr = 0;
      if (segmentMap[key].numAds > 0) {
        ctr = (segmentMap[key].totalClicks /
          segmentMap[key].totalImpressions) * 100;
      }
      var row = [
        key,
        segmentMap[key].numAds,
        segmentMap[key].totalImpressions,
        segmentMap[key].totalClicks,
        ctr.toFixed(2),
        segmentMap[key].totalCost];
      rows.push(row);
    }
  }
  sheet.getRange(3, 2, rows.length, 6).setValues(rows);
}

/**
 * Produces a formatted string representing a given date in a given time zone.
 *
 * @param {string} format A format specifier for the string to be produced.
 * @param {date} date A date object. Defaults to the current date.
 * @param {string} timeZone A time zone. Defaults to the account's time zone.
 * @return {string} A formatted string of the given date in the given time zone.
 */
function getDateStringInTimeZone(format, date, timeZone) {
  date = date || new Date();
  timeZone = timeZone || AdWordsApp.currentAccount().getTimeZone();
  return Utilities.formatDate(date, timeZone, format);
}
