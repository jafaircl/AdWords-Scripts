/**
 * @name Weekly Performance Report
 *
 * @overview This is a combination of Google AdWords scripts combined to create and e-mail a Google
 * 	Sheets spreadsheet. It will also run tests and apply labels for statistically significant
 * 	ad test winners and for underperforming keywords.
 * 	- Ad Performance Report
 * 		@author AdWords Scripts Team
 * 		@url https://developers.google.com/adwords/scripts/docs/solutions/account-summary
 * 	- Keyword Performance Report
 * 		@author AdWords Scripts Team
 * 		@url https://developers.google.com/adwords/scripts/docs/solutions/keyword-performance
 * 	- Declining Ad Groups
 * 		@author AdWords Scripts Team
 * 		@url https://developers.google.com/adwords/scripts/docs/solutions/declining-adgroups
 * 	- Keyword Labeler
 * 		@author AdWords Scripts Team
 * 		@url https://developers.google.com/adwords/scripts/docs/solutions/labels
 * 	- Negative Keyword Conflicts (Could also be run daily)
 * 		@author AdWords Scripts Team
 * 		@url https://developers.google.com/adwords/scripts/docs/solutions/negative-keyword-conflicts
 * 	- Display Placement Monitoring
 * 		@author Derek Martin
 * 		@url https://gist.github.com/derekmartinla/06e51b8a4298b8bbb8ff
 * 	- Search Query Mining
 * 		@author Brainlabs Digital
 * 		@url https://www.brainlabsdigital.com
 * 	- Statistically Significant Ad Creative Testing
 * 		@author Russell Savage
 * 		@url http://www.freeadwordsscripts.com/2013/12/automated-creative-testing-with.html
 *
 * @version 1.0
 */

/////////////////////////////////////////////////////////////
///                                                       ///
///                  MUST-CHANGE VARIABLES                ///
///                                                       ///
/////////////////////////////////////////////////////////////
// Recipient E-mail
var RECIPIENT_EMAIL = 'example@example.com';

// Template Spreadsheet URL
// Should be a copy of https://goo.gl/Hd4j3g
var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1MQJ1h-b2ZIYEHNREAhP7rJyrnZTtWPiGSNF_9beqHfc/edit';

/////////////////////////////////////////////////////////////
///                                                       ///
///                  CAN-CHANGE VARIABLES                 ///
///                                                       ///
/////////////////////////////////////////////////////////////

// Date Range
var dateRange = 'LAST_7_DAYS';

/**
 * Defines the rules by which keywords will be labeled.
 * The labelName field is required. Other fields may be null.
 * @type {Array.<{
 *     conditions: Array.<string>,
 *     dateRange: string,
 *     filter: function(Object): boolean,
 *     labelName: string,
 *   }>
 * }
 */
var GLOBAL_CONDITIONS = [
  'CampaignStatus = ENABLED',
  'AdGroupStatus = ENABLED',
  'Status = ENABLED',
  'Impressions > 10'
];

var RULES = [
  {
    conditions: [
      'Ctr < 0.01',
    ],
    labelName: 'Underperforming'
  }
];

// Limits on the number of keywords in an account the script can process.
var MAX_POSITIVES = 250000;
var MAX_NEGATIVES = 50000;

// Statistical Significance Ad Testing
var EXTRA_LOGS = true;
var CONFIDENCE_LEVEL = 99; // 90%, 95%, or 99% are most common
  
//If you only want to run on some campaigns, apply a label to them
//and put the name of the label here.  Leave blank to run on all campaigns.
var CAMPAIGN_LABEL = '';
  
//These two metrics are the components that make up the metric you
//want to compare. For example, this measures CTR = Clicks/Impressions
//Other examples might be:
// Cost Per Conv = Cost/Conversions
// Conversion Rate = Conversions/Clicks
// Cost Per Click = Cost/Clicks
var VISITORS_METRIC = 'Impressions';
var CONVERSIONS_METRIC = 'Clicks';
//This is the number of impressions the Ad needs to have in order
//to start measuring the results of a test.
var VISITORS_THRESHOLD = 75;
 
//Setting this to true to enable the script to check mobile ads
//against other mobile ads only. Enabling this will start new tests
//in all your AdGroups so only enable this after you have completed
//a testing cycle.
var ENABLE_MOBILE_AD_TESTING = true;
 
//Set this on the first run which should be the approximate last time
//you started a new creative test. After the first run, this setting
//will be ignored.
var OVERRIDE_LAST_TOUCHED_DATE = 'Jun 1, 2016';
  
var LOSER_LABEL = 'Loser '+CONFIDENCE_LEVEL+'% Confidence';
var CHAMPION_LABEL = 'Current Champion';
 
// Set this to true and the script will apply a label to 
// each AdGroup to let you know the date the test started
// This helps you validate the results of the script.
var APPLY_TEST_START_DATE_LABELS = false;
  
//These come from the url when you are logged into AdWords
//Set these if you want your emails to link directly to the AdGroup
var __c = '';
var __u = '';

// The currency symbol used for formatting. For example "£", "$" or "€".
var currencySymbol = "$";

// Use this if you only want to look at some campaigns
// such as campaigns with names 
var campaignNameContains = "";

// Words will be ignored if their statistics are lower than any of these thresholds
var impressionThreshold = 10;
var clickThreshold = 0;
var costThreshold = 0;
var conversionThreshold = 0;

/////////////////////////////////////////////////////////////
///                                                       ///
///                     START FUNCTIONS                   ///
///                                                       ///
/////////////////////////////////////////////////////////////
var currentAccount = AdWordsApp.currentAccount();
var accountName = currentAccount.getName();
var spreadsheet = copySpreadsheet(SPREADSHEET_URL);

function main() {
  // Generate Ad Performance Report
  adPerformance();
  
  // Generate Keyword Performance Report
  keywordPerformance();
  
  // Generate Declining Ad Groups Report
  decliningAdGroups();
  
  // Underperforming Keywords
  var results = processAccount();
  processResults(results);
  
  // Negative Keyword Conflicts
  negativeKeywordConflicts();
  
  // Display Placement Monitoring
  displayPlacementMonitoring();
  
  // Search Query Mining
  searchQueryMining();
  
  // Send the email
  if (RECIPIENT_EMAIL) {
    MailApp.sendEmail(RECIPIENT_EMAIL,
      accountName + ' Performance Report is ready',
      spreadsheet.getUrl());
  }
  
  // Statistical Significance Ad Testing
  statSigAdTesting();

}

/////////////////////////////////////////////////////////////
///                                                       ///
///                     MAIN FUNCTIONS                    ///
///                                                       ///
/////////////////////////////////////////////////////////////

/**
 * Outputs ad performance to the spreadsheet
 */
function adPerformance() {
  Logger.log('Using template spreadsheet - %s.', SPREADSHEET_URL);
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
}

/**
 * Outputs ad performance to the spreadsheet
 */
function keywordPerformance() {
  Logger.log('Using template spreadsheet - %s.', SPREADSHEET_URL);
  Logger.log('Generated new reporting spreadsheet %s based on the template ' +
      'spreadsheet. The reporting data will be populated here.',
      spreadsheet.getUrl());

  var sheet = spreadsheet.getSheetByName('KeywordReport');
  sheet.getRange(1, 2, 1, 1).setValue('Date');
  sheet.getRange(1, 3, 1, 1).setValue(new Date());
  spreadsheet.getRangeByName('account_id').setValue(AdWordsApp.currentAccount().
      getCustomerId());
  outputQualityScoreData(sheet);
  outputPositionData(sheet);
  Logger.log('Keyword performance report available at\n' +
      spreadsheet.getUrl());
}

/**
 * Declining Ad Groups function
 */
function decliningAdGroups() {
  Logger.log('Using spreadsheet - %s.', SPREADSHEET_URL);

  var sheet = spreadsheet.getSheetByName('Declining Ad Groups');
  spreadsheet.getRangeByName('account_id').setValue(
      AdWordsApp.currentAccount().getCustomerId());
  sheet.getRange(1, 2, 1, 1).setValue('Date');
  sheet.getRange(1, 3, 1, 1).setValue(new Date());
  sheet.getRange(7, 1, sheet.getMaxRows() - 7, sheet.getMaxColumns()).clear();

  var adGroupsIterator = AdWordsApp.adGroups()
      .withCondition("Status = 'ENABLED'")
      .withCondition("CampaignStatus = 'ENABLED'")
      .forDateRange(dateRange)
      .orderBy('Ctr ASC')
      .withLimit(100)
      .get();

  var today = getDateInThePast(0);
  var oneWeekAgo = getDateInThePast(7);
  var twoWeeksAgo = getDateInThePast(14);
  var threeWeeksAgo = getDateInThePast(21);

  var reportRows = [];

  while (adGroupsIterator.hasNext()) {
    var adGroup = adGroupsIterator.next();
    // Let's look at the trend of the ad group's CTR.
    var statsThreeWeeksAgo = adGroup.getStatsFor(threeWeeksAgo, twoWeeksAgo);
    var statsTwoWeeksAgo = adGroup.getStatsFor(twoWeeksAgo, oneWeekAgo);
    var statsLastWeek = adGroup.getStatsFor(oneWeekAgo, today);

    // Week over week, the ad group is declining - record that!
    if (statsLastWeek.getCtr() < statsTwoWeeksAgo.getCtr() &&
        statsTwoWeeksAgo.getCtr() < statsThreeWeeksAgo.getCtr()) {
      reportRows.push([adGroup.getCampaign().getName(), adGroup.getName(),
          statsLastWeek.getCtr(), statsLastWeek.getCost(),
          statsTwoWeeksAgo.getCtr(), statsTwoWeeksAgo.getCost(),
          statsThreeWeeksAgo.getCtr(), statsThreeWeeksAgo.getCost()]);
    }
  }
  if (reportRows.length > 0) {
    sheet.getRange(7, 2, reportRows.length, 8).setValues(reportRows);
    sheet.getRange(7, 4, reportRows.length, 1).setNumberFormat('#0.00%');
    sheet.getRange(7, 6, reportRows.length, 1).setNumberFormat('#0.00%');
    sheet.getRange(7, 8, reportRows.length, 1).setNumberFormat('#0.00%');

    sheet.getRange(7, 5, reportRows.length, 1).setNumberFormat('#,##0.00');
    sheet.getRange(7, 7, reportRows.length, 1).setNumberFormat('#,##0.00');
    sheet.getRange(7, 9, reportRows.length, 1).setNumberFormat('#,##0.00');
  }
}

/**
 * Processes the rules on the current account.
 *
 * @return {Array.<Object>} An array of changes made, each having
 *     a customerId, campaign name, ad group name, label name,
 *     and keyword text that the label was applied to.
 */
function processAccount() {
  ensureAccountLabels();
  var changes = applyLabels();

  return changes;
}

/**
 * Processes the results of the script.
 *
 * @param {Array.<Object>} changes An array of changes made, each having
 *     a customerId, campaign name, ad group name, label name,
 *     and keyword text that the label was applied to.
 */
function processResults(changes) {
  if (changes.length > 0) {
    saveToSpreadsheet(changes, RECIPIENT_EMAIL);
  } else {
    Logger.log('No labels were applied.');
  }
}

/**
 * Negative Keyword Conflicts
 */
function negativeKeywordConflicts() {
  var conflicts = findAllConflicts();

  var hasConflicts = outputConflicts(spreadsheet,
    AdWordsApp.currentAccount().getCustomerId(), conflicts);
}

/**
 * Statistical Significance Ad Testing
 */
function statSigAdTesting() {
  createLabelIfNeeded(LOSER_LABEL,"#FF00FF"); //Set the colors of the labels here
  createLabelIfNeeded(CHAMPION_LABEL,"#0000FF"); //Set the colors of the labels here
    
  //Let's find all the AdGroups that have new tests starting
  var currentAdMap = getCurrentAdsSnapshot();
  var previousAdMap = getPreviousAdsSnapshot();
  if(previousAdMap) {
    currentAdMap = updateCurrentAdMap(currentAdMap,previousAdMap);
  }
  storeAdsSnapshot(currentAdMap);
  previousAdMap = null;
    
  //Now run through the AdGroups to find tests
   var agSelector = AdWordsApp.adGroups()
    .withCondition('CampaignStatus = ENABLED')
    .withCondition('AdGroupStatus = ENABLED')
    .withCondition('Status = ENABLED');
  if(CAMPAIGN_LABEL !== '') {
    var campNames = getCampaignNames();
    agSelector = agSelector.withCondition("CampaignName IN ['"+campNames.join("','")+"']");
  }
  var agIter = agSelector.get();
  var todayDate = getDateString(new Date(),'yyyyMMdd');
  var touchedAdGroups = [];
  var finishedEarly = false;
  while(agIter.hasNext()) {
    var ag = agIter.next();
 
    var numLoops = (ENABLE_MOBILE_AD_TESTING) ? 2 : 1;
    for(var loopNum = 0; loopNum < numLoops; loopNum++) {
      var isMobile = (loopNum == 1);
      var rowKey;
      if(isMobile) {
        info('Checking Mobile Ads in AdGroup: "'+ag.getName()+'"');
        rowKey = [ag.getCampaign().getId(),ag.getId(),'Mobile'].join('-');
      } else {
        info('Checking Ads in AdGroup: "'+ag.getName()+'"');
        rowKey = [ag.getCampaign().getId(),ag.getId()].join('-');
      }
 
      if(!currentAdMap[rowKey]) {  //This shouldn't happen
        warn('Could not find AdGroup: '+ag.getName()+' in current ad map.');
        continue; 
      }
       
      if(APPLY_TEST_START_DATE_LABELS) {
        var dateLabel;
        if(isMobile) {
          dateLabel = 'Mobile Tests Started: '+getDateString(currentAdMap[rowKey].lastTouched,'yyyy-MM-dd');
        } else {
          dateLabel = 'Tests Started: '+getDateString(currentAdMap[rowKey].lastTouched,'yyyy-MM-dd');
        }
 
        createLabelIfNeeded(dateLabel,"#8A2BE2");
        //remove old start date
        var labelIter = ag.labels().withCondition("Name STARTS_WITH '"+dateLabel.split(':')[0]+"'")
                                   .withCondition("Name != '"+dateLabel+"'").get();
        while(labelIter.hasNext()) {
          var label = labelIter.next();
          ag.removeLabel(label.getName());
          if(!label.adGroups().get().hasNext()) {
            //if there are no more entities with that label, delete it.
            label.remove();
          }
        }
        applyLabel(ag,dateLabel);
      }
           
      //Here is the date range for the test metrics
      var lastTouchedDate = getDateString(currentAdMap[rowKey].lastTouched,'yyyyMMdd');
      info('Last Touched Date: '+lastTouchedDate+' Todays Date: '+ todayDate);
      if(lastTouchedDate === todayDate) {
        //Special case where the AdGroup was updated today which means a new test has started.
        //Remove the old labels, but keep the champion as the control for the next test
        info('New test is starting in AdGroup: '+ag.getName());
        removeLoserLabelsFromAds(ag,isMobile);
        continue;
      }
       
      //Is there a previous winner? if so we should use it as the control.
      var controlAd = checkForPreviousWinner(ag,isMobile);
       
      //Here we order by the Visitors metric and use that as a control if we don't have one
      var adSelector = ag.ads().withCondition('Status = ENABLED').withCondition('AdType = TEXT_AD');
      if(!AdWordsApp.getExecutionInfo().isPreview()) {
        adSelector = adSelector.withCondition("LabelNames CONTAINS_NONE ['"+[LOSER_LABEL,CHAMPION_LABEL].join("','")+"']");
      }
      var adIter = adSelector.forDateRange(lastTouchedDate, todayDate)
                             .orderBy(VISITORS_METRIC+" DESC")
                             .get();
      if( (controlAd == null && adIter.totalNumEntities() < 2) ||
          (controlAd != null && adIter.totalNumEntities() < 1) )
      { 
        info('AdGroup did not have enough eligible Ads. Had: '+adIter.totalNumEntities()+', Needed at least 2'); 
        continue; 
      }
       
      if(!controlAd) {
        info('No control set for AdGroup. Setting one.');
        while(adIter.hasNext()) {
          var ad = adIter.next();
          if(shouldSkip(isMobile,ad)) { continue; }
          controlAd = ad;
          break;
        }
        if(!controlAd) {
          continue;
        }
        applyLabel(controlAd,CHAMPION_LABEL);
      }
       
      while(adIter.hasNext()) {
        var testAd = adIter.next();
        if(shouldSkip(isMobile,testAd)) { continue; }
        //The Test object does all the heavy lifting for us.
        var test = new Test(controlAd,testAd,
                            CONFIDENCE_LEVEL,
                            lastTouchedDate,todayDate,
                            VISITORS_METRIC,CONVERSIONS_METRIC);
        info('Control - Visitors: '+test.getControlVisitors()+' Conversions: '+test.getControlConversions());
        info('Test    - Visitors: '+test.getTestVisitors()+' Conversions: '+test.getTestConversions());
        info('P-Value: '+test.getPValue());
         
        if(test.getControlVisitors() < VISITORS_THRESHOLD ||
           test.getTestVisitors() < VISITORS_THRESHOLD)
        {
          info('Not enough visitors in the control or test ad.  Skipping.');
          continue;
        }
         
        //Check for significance
        if(test.isSignificant()) {
          var loser = test.getLoser();
          removeLabel(loser,CHAMPION_LABEL); //Champion has been dethroned
          applyLabel(loser,LOSER_LABEL);
           
          //The winner is the new control. Could be the same as the old one.
          controlAd = test.getWinner();
          applyLabel(controlAd,CHAMPION_LABEL);
           
          //We store some metrics for a nice email later
          if(!ag['touchCount']) {
            ag['touchCount'] = 0;
            touchedAdGroups.push(ag);
          }
          ag['touchCount']++;
        }
      }
       
      //Let's bail if we run out of time so we can send the emails.
      if((!AdWordsApp.getExecutionInfo().isPreview() && AdWordsApp.getExecutionInfo().getRemainingTime() < 60) ||
         ( AdWordsApp.getExecutionInfo().isPreview() && AdWordsApp.getExecutionInfo().getRemainingTime() < 10) )
      {
        finishedEarly = true;
        break;
      }
    }
  }
  if(touchedAdGroups.length > 0) {
    sendMailForTouchedAdGroups(touchedAdGroups,finishedEarly);
  }
}

/**
 * Display placement monitoring
 */
// This script reviews your GDN placements for the following conditions:
// 1) Placements that are converting at less than $40
// 2) Placements that have cost more than $50 but haven't converted
// 3) Placements that have more than 1K impressions and less than .10 CTR

function displayPlacementMonitoring() {
   
  var body = "<h2>Google Display Network Alert - " + accountName + "</h2>";
  body += "<h3>Placements that are converting at less than $40:</h3> " ;
  body += "<ul>";
  
  var list = runLowCostAndConvertingReport();
  
  for (i=0; i < list.length; i++) {
    body += "<li><strong>" + list[i].placement + "</strong> - " + list[i].adgroup + ' - $' + list2[i].cost + "</li>";
    
  } 
  body += "</ul>";
  
  body += "<h3>Placements that have cost more than $50 but haven't converted:</h3> " ;
  body += "<ul>";
  
  var list2 = runHighCostNoConversionsReport();

   for (i=0; i < list2.length; i++) {
    body += "<li><strong>" + list2[i].placement + "</strong> - " + list2[i].adgroup + ' - $' + list2[i].cost + "</li>";
    
  } 
  body += "</ul>";

  body += "<h3>Placements that have more than 1K impressions and less than .10 CTR:</h3> " ;
  body += "<ul>";

  var list3 = runLowCtrAndHighImpressionsReport();
  
   for (i=0; i < list3.length; i++) {
    body += "<li><strong>" + list3[i].placement + "</strong> - " + list3[i].adgroup +" - " + parseFloat(list3[i].clicks/list3[i].impressions).toFixed(4) + "% - " + list3[i].clicks + " clicks - " + list3[i].impressions + ' impressions ' + "</li>";
    
  } 
  body += "</ul>";
  
  MailApp.sendEmail(RECIPIENT_EMAIL,'Display Network Alerts - ' + accountName, body,{htmlBody: body}); 
}



/////////////////////////////////////////////////////////////
///                                                       ///
///                   HELPER FUNCTIONS                    ///
///                                                       ///
/////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////
///                                                       ///
///               AD PERFORMANCE FUNCTIONS                ///
///                                                       ///
/////////////////////////////////////////////////////////////

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
      .forDateRange(dateRange)
      .withCondition('Impressions > 0').get();
  while (adIterator.hasNext()) {
    var ad = adIterator.next();
    var stats = ad.getStatsFor(dateRange);
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

/////////////////////////////////////////////////////////////
///                                                       ///
///                   KEYWORD FUNCTIONS                   ///
///                                                       ///
/////////////////////////////////////////////////////////////
/**
 * Outputs Quality score related data to the spreadsheet
 * @param {Sheet} sheet The sheet to output to.
 */
function outputQualityScoreData(sheet) {
  // Output header row
  var header = [
    'Quality Score',
    'Num Keywords',
    'Impressions',
    'Clicks',
    'CTR (%)',
    'Cost'
  ];
  sheet.getRange(3, 2, 1, 6).setValues([header]);

  // Initialize
  var qualityScoreMap = [];
  for (i = 1; i <= 10; i++) {
    qualityScoreMap[i] = {
      numKeywords: 0,
      totalImpressions: 0,
      totalClicks: 0,
      totalCost: 0.0
    };
  }

  // Compute data
  var keywordIterator = AdWordsApp.keywords()
      .forDateRange(dateRange)
      .withCondition('Impressions > 0')
      .get();
  while (keywordIterator.hasNext()) {
    var keyword = keywordIterator.next();
    var stats = keyword.getStatsFor(dateRange);
    var data = qualityScoreMap[keyword.getQualityScore()];
    if (data) {
      data.numKeywords++;
      data.totalImpressions += stats.getImpressions();
      data.totalClicks += stats.getClicks();
      data.totalCost += stats.getCost();
    }
  }

  // Output data to spreadsheet
  var rows = [];
  for (var key in qualityScoreMap) {
    var ctr = 0;
    var cost = 0.0;
    if (qualityScoreMap[key].numKeywords > 0) {
      ctr = (qualityScoreMap[key].totalClicks /
        qualityScoreMap[key].totalImpressions) * 100;
    }
    var row = [
      key,
      qualityScoreMap[key].numKeywords,
      qualityScoreMap[key].totalImpressions,
      qualityScoreMap[key].totalClicks,
      ctr.toFixed(2),
      qualityScoreMap[key].totalCost];
    rows.push(row);
  }
  sheet.getRange(4, 2, rows.length, 6).setValues(rows);
}

/**
 * Outputs average position related data to the spreadsheet.
 * @param {Sheet} sheet The sheet to output to.
 */
function outputPositionData(sheet) {
  // Output header row
  headerRow = [];
  var header = [
    'Avg Position',
    'Num Keywords',
    'Impressions',
    'Clicks',
    'CTR (%)',
    'Cost'
  ];
  headerRow.push(header);
  sheet.getRange(16, 2, 1, 6).setValues(headerRow);

  // Initialize
  var positionMap = [];
  for (i = 1; i <= 12; i++) {
    positionMap[i] = {
      numKeywords: 0,
      totalImpressions: 0,
      totalClicks: 0,
      totalCost: 0.0
    };
  }

  // Compute data
  var keywordIterator = AdWordsApp.keywords()
      .forDateRange(dateRange)
      .withCondition('Impressions > 0')
      .get();
  while (keywordIterator.hasNext()) {
    var keyword = keywordIterator.next();
    var stats = keyword.getStatsFor(dateRange);
    if (stats.getAveragePosition() <= 11) {
      var data = positionMap[Math.ceil(stats.getAveragePosition())];
    } else {
      // All positions greater than 11
      var data = positionMap[12];
    }
    data.numKeywords++;
    data.totalImpressions += stats.getImpressions();
    data.totalClicks += stats.getClicks();
    data.totalCost += stats.getCost();
  }

  // Output data to spreadsheet
  var rows = [];
  for (var key in positionMap) {
    var ctr = 0;
    var cost = 0.0;
    if (positionMap[key].numKeywords > 0) {
      ctr = (positionMap[key].totalClicks /
        positionMap[key].totalImpressions) * 100;
    }
    var row = [
      key <= 11 ? key - 1 + ' to ' + key : '>11',
      positionMap[key].numKeywords,
      positionMap[key].totalImpressions,
      positionMap[key].totalClicks,
      ctr.toFixed(2),
      positionMap[key].totalCost
    ];
    rows.push(row);
  }
  sheet.getRange(17, 2, rows.length, 6).setValues(rows);
}

/////////////////////////////////////////////////////////////
///                                                       ///
///          UNDERPERFORMING KEYWORD FUNCTIONS            ///
///                                                       ///
/////////////////////////////////////////////////////////////

/**
 * Retrieves the names of all labels in the account.
 *
 * @return {Array.<string>} An array of label names.
 */
function getAccountLabelNames() {
  var labelNames = [];
  var iterator = AdWordsApp.labels().get();

  while (iterator.hasNext()) {
    labelNames.push(iterator.next().getName());
  }

  return labelNames;
}

/**
 * Checks that the account has a label for each rule and
 * creates the rule's label if it does not already exist.
 * Throws an exception if a rule does not have a labelName.
 */
function ensureAccountLabels() {
  var labelNames = getAccountLabelNames();

  for (var i = 0; i < RULES.length; i++) {
    var labelName = RULES[i].labelName;

    if (!labelName) {
      throw 'Missing labelName for rule #' + i;
    }

    if (labelNames.indexOf(labelName) == -1) {
      AdWordsApp.createLabel(labelName);
      labelNames.push(labelName);
    }
  }
}

/**
 * Retrieves the keywords in an account satisfying a rule
 * and that do not already have the rule's label.
 *
 * @param {Object} rule An element of the RULES array.
 * @return {Array.<Object>} An array of keywords.
 */
function getKeywordsForRule(rule) {
  var selector = AdWordsApp.keywords();

  // Add global conditions.
  for (var i = 0; i < GLOBAL_CONDITIONS.length; i++) {
    selector = selector.withCondition(GLOBAL_CONDITIONS[i]);
  }

  // Add selector conditions for this rule.
  if (rule.conditions) {
    for (var i = 0; i < rule.conditions.length; i++) {
      selector = selector.withCondition(rule.conditions[i]);
    }
  }

  // Exclude keywords that already have the label.
  selector.withCondition('LabelNames CONTAINS_NONE ["' + rule.labelName + '"]');

  // Add a date range.
  selector = selector.forDateRange(dateRange);

  // Get the keywords.
  var iterator = selector.get();
  var keywords = [];

  // Check filter conditions for this rule.
  while (iterator.hasNext()) {
    var keyword = iterator.next();

    if (!rule.filter || rule.filter(keyword)) {
      keywords.push(keyword);
    }
  }

  return keywords;
}

/**
 * For each rule, determines the keywords matching the rule and which
 * need to have a label newly applied, and applies it.
 *
 * @return {Array.<Object>} An array of changes made, each having
 *     a customerId, campaign name, ad group name, label name,
 *     and keyword text that the label was applied to.
 */
function applyLabels() {
  var changes = [];
  var customerId = AdWordsApp.currentAccount().getCustomerId();

  for (var i = 0; i < RULES.length; i++) {
    var rule = RULES[i];
    var keywords = getKeywordsForRule(rule);
    var labelName = rule.labelName;

    for (var j = 0; j < keywords.length; j++) {
      var keyword = keywords[j];

      keyword.applyLabel(labelName);

      changes.push({
        customerId: customerId,
        campaignName: keyword.getCampaign().getName(),
        adGroupName: keyword.getAdGroup().getName(),
        labelName: labelName,
        keywordText: keyword.getText(),
      });
    }
  }

  return changes;
}

/**
 * Outputs a list of applied labels to a new spreadsheet and gives editor access
 * to a list of provided emails.
 *
 * @param {Array.<Object>} changes An array of changes made, each having
 *     a customerId, campaign name, ad group name, label name,
 *     and keyword text that the label was applied to.
 * @param {Array.<Object>} emails An array of email addresses.
 * @return {string} The URL of the spreadsheet.
 */
function saveToSpreadsheet(changes, emails) {
  var sheet = spreadsheet.getSheetByName('Underperforming Keywords');

  Logger.log('Saving changes to spreadsheet at ' + spreadsheet.getUrl());

  var headers = spreadsheet.getRangeByName('Headers');
  var outputRange = headers.offset(1, 0, changes.length);

  var outputValues = [];
  for (var i = 0; i < changes.length; i++) {
    var change = changes[i];
    outputValues.push([
      change.customerId,
      change.campaignName,
      change.adGroupName,
      change.keywordText,
      change.labelName
    ]);
  }
  outputRange.setValues(outputValues);

  spreadsheet.getRangeByName('RunDate').setValue(new Date());

  return spreadsheet.getUrl();
}

/////////////////////////////////////////////////////////////
///                                                       ///
///          NEGATIVE KEYWORD CONFLICT FUNCTIONS          ///
///                                                       ///
/////////////////////////////////////////////////////////////

/**
 * Finds all negative keyword conflicts in an account.
 *
 * @return {Array.<Object>} An array of conflicts.
 */
function findAllConflicts() {
  var campaignIds;
    campaignIds = getAllCampaignIds();

  var campaignCondition = '';
  if (campaignIds.length > 0) {
    campaignCondition = 'AND CampaignId IN [' + campaignIds.join(',') + ']';
  }

  Logger.log('Downloading keywords performance report');
  var query =
    'SELECT CampaignId, CampaignName, AdGroupId, AdGroupName, ' +
    '       Criteria, KeywordMatchType, IsNegative ' +
    'FROM KEYWORDS_PERFORMANCE_REPORT ' +
    'WHERE CampaignStatus = "ENABLED" AND AdGroupStatus = "ENABLED" AND ' +
    '      Status = "ENABLED" AND IsNegative IN [true, false] ' +
    '      ' + campaignCondition + ' ' +
    'DURING YESTERDAY';
  var report = AdWordsApp.report(query);

  Logger.log('Building cache and populating with keywords');
  var cache = {};
  var numPositives = 0;
  var numNegatives = 0;

  var rows = report.rows();
  while (rows.hasNext()) {
    var row = rows.next();

    var campaignId = row['CampaignId'];
    var campaignName = row['CampaignName'];
    var adGroupId = row['AdGroupId'];
    var adGroupName = row['AdGroupName'];
    var keywordText = row['Criteria'];
    var keywordMatchType = row['KeywordMatchType'];
    var isNegative = row['IsNegative'];

    if (!cache[campaignId]) {
      cache[campaignId] = {
        campaignName: campaignName,
        adGroups: {},
        negatives: [],
        negativesFromLists: [],
      };
    }

    if (!cache[campaignId].adGroups[adGroupId]) {
      cache[campaignId].adGroups[adGroupId] = {
        adGroupName: adGroupName,
        positives: [],
        negatives: [],
      };
    }

    if (isNegative == 'true') {
      cache[campaignId].adGroups[adGroupId].negatives
        .push(normalizeKeyword(keywordText, keywordMatchType));
      numNegatives++;
    } else {
      cache[campaignId].adGroups[adGroupId].positives
        .push(normalizeKeyword(keywordText, keywordMatchType));
      numPositives++;
    }

    if (numPositives > MAX_POSITIVES ||
        numNegatives > MAX_NEGATIVES) {
      throw 'Trying to process too many keywords. Please restrict the ' +
            'script to a smaller subset of campaigns.';
    }
  }

  Logger.log('Downloading campaign negatives report');
  var query =
    'SELECT CampaignId, Criteria, KeywordMatchType ' +
    'FROM CAMPAIGN_NEGATIVE_KEYWORDS_PERFORMANCE_REPORT ' +
    'WHERE CampaignStatus = "ENABLED" ' +
    '      ' + campaignCondition;
  var report = AdWordsApp.report(query);

  var rows = report.rows();
  while (rows.hasNext()) {
    var row = rows.next();

    var campaignId = row['CampaignId'];
    var keywordText = row['Criteria'];
    var keywordMatchType = row['KeywordMatchType'];

    if (cache[campaignId]) {
      cache[campaignId].negatives
        .push(normalizeKeyword(keywordText, keywordMatchType));
    }
  }

  Logger.log('Populating cache with negative keyword lists');
  var negativeKeywordLists =
    AdWordsApp.negativeKeywordLists().withCondition('Status = ACTIVE').get();

  while (negativeKeywordLists.hasNext()) {
    var negativeKeywordList = negativeKeywordLists.next();

    var negatives = [];
    var negativeKeywords = negativeKeywordList.negativeKeywords().get();

    while (negativeKeywords.hasNext()) {
      var negative = negativeKeywords.next();
      negatives.push(normalizeKeyword(negative.getText(),
                                      negative.getMatchType()));
    }

    var campaigns = negativeKeywordList.campaigns()
        .withCondition('Status = ENABLED').get();

    while (campaigns.hasNext()) {
      var campaign = campaigns.next();
      var campaignId = campaign.getId();

      if (cache[campaignId]) {
        cache[campaignId].negativesFromLists = negatives;
      }
    }
  }

  Logger.log('Finding negative conflicts');
  var conflicts = [];

  // Adds context about the conflict.
  var enrichConflict = function(conflict, campaignId, adGroupId, level) {
    conflict.campaignId = campaignId;
    conflict.adGroupId = adGroupId;
    conflict.campaignName = cache[campaignId].campaignName;
    conflict.adGroupName = cache[campaignId].adGroups[adGroupId].adGroupName;
    conflict.level = level;
  };

  for (var campaignId in cache) {
    for (var adGroupId in cache[campaignId].adGroups) {
      var positives = cache[campaignId].adGroups[adGroupId].positives;

      var negativeLevels = {
        'Campaign': cache[campaignId].negatives,
        'Ad Group': cache[campaignId].adGroups[adGroupId].negatives,
        'Negative list': cache[campaignId].negativesFromLists
      };

      for (var level in negativeLevels) {
        var newConflicts =
          checkForConflicts(negativeLevels[level], positives);

        for (var i = 0; i < newConflicts.length; i++) {
          enrichConflict(newConflicts[i], campaignId, adGroupId, level);
        }
        conflicts = conflicts.concat(newConflicts);
      }
    }
  }

  return conflicts;
}

/**
 * Saves conflicts to a spreadsheet if present.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 * @param {string} customerId The account the conflicts are for.
 * @param {Array.<Object>} conflicts A list of conflicts.
 * @return {boolean} True if there were conflicts and false otherwise.
 */
function outputConflicts(spreadsheet, customerId, conflicts) {
  if (conflicts.length > 0) {
    saveConflictsToSpreadsheet(spreadsheet, customerId, conflicts);
    Logger.log('Conflicts were found for ' + customerId +
               '. See ' + spreadsheet.getUrl());
    return true;
  } else {
    Logger.log('No conflicts were found for ' + customerId + '.');
    return false;
  }
}

/**
 * Sets up the spreadsheet to receive output.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 */
function initializeSpreadsheet(spreadsheet) {
  // Make sure the spreadsheet is using the account's timezone.
  spreadsheet.setSpreadsheetTimeZone(AdWordsApp.currentAccount().getTimeZone());

  // Clear the last run date on the spreadsheet.
  spreadsheet.getRangeByName('RunDate').clearContent();

  // Clear all rows in the spreadsheet below the header row.
  var outputRange = spreadsheet.getRangeByName('Headers')
    .offset(1, 0, spreadsheet.getSheetByName('Conflicts')
        .getDataRange().getLastRow())
    .clearContent();
}

/**
 * Saves conflicts for a particular account to the spreadsheet starting at the
 * first unused row.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 * @param {string} customerId The account that the conflicts are for.
 * @param {Array.<Object>} conflicts A list of conflicts.
 */
function saveConflictsToSpreadsheet(spreadsheet, customerId, conflicts) {
  // Find the first open row on the Report tab below the headers and create a
  // range large enough to hold all of the failures, one per row.
  var lastRow = spreadsheet.getSheetByName('Negative Keyword Conflicts')
    .getDataRange().getLastRow();
  var headers = spreadsheet.getRangeByName('NegativeKeywordHeaders');
  var outputRange = headers
    .offset(lastRow - headers.getRow() + 1, 0, conflicts.length);

  // Build each row of output values in the order of the columns.
  var outputValues = [];
  for (var i = 0; i < conflicts.length; i++) {
    var conflict = conflicts[i];
    outputValues.push([
      customerId,
      conflict.negative,
      conflict.level,
      conflict.positives.join(', '),
      conflict.campaignName,
      conflict.adGroupName
    ]);
  }
  outputRange.setValues(outputValues);

  spreadsheet.getRangeByName('NegativeKeywordRunDate').setValue(new Date());
}

/**
 * Retrieves the campaign IDs of a campaign iterator.
 *
 * @param {Object} campaigns A CampaignIterator object.
 * @return {Array.<Integer>} An array of campaign IDs.
 */
function getCampaignIds(campaigns) {
  var campaignIds = [];
  while (campaigns.hasNext()) {
    campaignIds.push(campaigns.next().getId());
  }

  return campaignIds;
}

/**
 * Retrieves all campaign IDs in an account.
 *
 * @return {Array.<Integer>} An array of campaign IDs.
 */
function getAllCampaignIds() {
  return getCampaignIds(AdWordsApp.campaigns().get());
}

/**
 * Retrieves the campaign IDs with a given label.
 *
 * @param {string} labelText The text of the label.
 * @return {Array.<Integer>} An array of campaign IDs, or null if the
 *     label was not found.
 */
function getCampaignIdsWithLabel(labelText) {
  var labels = AdWordsApp.labels()
    .withCondition('Name = "' + labelText + '"')
    .get();

  if (!labels.hasNext()) {
    return null;
  }
  var label = labels.next();

  return getCampaignIds(label.campaigns().get());
}

/**
 * Compares a set of negative keywords and positive keywords to identify
 * conflicts where a negative keyword blocks a positive keyword.
 *
 * @param {Array.<Object>} negatives A list of objects with fields
 *     display, raw, and matchType.
 * @param {Array.<Object>} positives A list of objects with fields
 *     display, raw, and matchType.
 * @return {Array.<Object>} An array of conflicts, each an object with
 *     the negative keyword display text causing the conflict and an array
 *     of blocked positive keyword display texts.
 */
function checkForConflicts(negatives, positives) {
  var conflicts = [];

  for (var i = 0; i < negatives.length; i++) {
    var negative = negatives[i];
    var anyBlock = false;
    var blockedPositives = [];

    for (var j = 0; j < positives.length; j++) {
      var positive = positives[j];

      if (negativeBlocksPositive(negative, positive)) {
        anyBlock = true;
        blockedPositives.push(positive.display);
      }
    }

    if (anyBlock) {
      conflicts.push({
        negative: negative.display,
        positives: blockedPositives
      });
    }
  }

  return conflicts;
}

/**
 * Removes leading and trailing match type punctuation from the first and
 * last character of a keyword's text, if any.
 *
 * @param {string} text A keyword's text to remove punctuation from.
 * @param {string} open The character that may be the first character.
 * @param {string} close The character that may be the last character.
 * @return {Object} The same text, trimmed of open and close if present.
 */
function trimKeyword(text, open, close) {
  if (text.substring(0, 1) == open &&
      text.substring(text.length - 1) == close) {
    return text.substring(1, text.length - 1);
  }

  return text;
}

/**
 * Normalizes a keyword by returning a raw and display version and consistent
 * match type. The raw version has no leading and trailing punctuation for
 * phrase and exact match keywords, no consecutive whitespace, is all
 * lowercase, and removes broad match qualifiers. The display version has no
 * consecutive whitespace and is all lowercase. The match type is uppercase.
 *
 * @param {string} text A keyword's text that should be normalized.
 * @param {string} matchType The keyword's match type.
 * @return {Object} An object with fields display, raw, and matchType.
 */
function normalizeKeyword(text, matchType) {
  var display;
  var raw = text;
  matchType = matchType.toUpperCase();

  // Replace leading and trailing "" for phrase match keywords and [] for
  // exact match keywords, if it is there.
  if (matchType == 'PHRASE') {
    raw = trimKeyword(raw, '"', '"');
  } else if (matchType == 'EXACT') {
    raw = trimKeyword(raw, '[', ']');
  }

  // Collapse any runs of whitespace into single spaces.
  raw = raw.replace(new RegExp('\\s+', 'g'), ' ');

  // Keywords are not case sensitive.
  raw = raw.toLowerCase();

  // Set display version.
  display = raw;
  if (matchType == 'PHRASE') {
    display = '"' + display + '"';
  } else if (matchType == 'EXACT') {
    display = '[' + display + ']';
  }

  // Remove broad match modifier '+' sign.
  raw = raw.replace(new RegExp('\\s\\+', 'g'), ' ');

  return {display: display, raw: raw, matchType: matchType};
}

/**
 * Tests whether all of the tokens in one keyword's raw text appear in
 * the tokens of a second keyword's text.
 *
 * @param {string} keywordText1 the raw keyword text whose tokens may
 *     appear in the other keyword text.
 * @param {string} keywordText2 the raw keyword text which may contain
 *     the tokens of the other keyword.
 * @return {boolean} Whether all tokens in keywordText1 appear among
 *     the tokens of keywordText2.
 */
function hasAllTokens(keywordText1, keywordText2) {
  var keywordTokens1 = keywordText1.split(' ');
  var keywordTokens2 = keywordText2.split(' ');

  for (var i = 0; i < keywordTokens1.length; i++) {
    if (keywordTokens2.indexOf(keywordTokens1[i]) == -1) {
      return false;
    }
  }

  return true;
}

/**
 * Tests whether all of the tokens in one keyword's raw text appear in
 * order in the tokens of a second keyword's text.
 *
 * @param {string} keywordText1 the raw keyword text whose tokens may
 *     appear in the other keyword text.
 * @param {string} keywordText2 the raw keyword text which may contain
 *     the tokens of the other keyword in order.
 * @return {boolean} Whether all tokens in keywordText1 appear in order
 *     among the tokens of keywordText2.
 */
function isSubsequence(keywordText1, keywordText2) {
  return (' ' + keywordText2 + ' ').indexOf(' ' + keywordText1 + ' ') >= 0;
}

/**
 * Tests whether a negative keyword blocks a positive keyword, taking into
 * account their match types.
 *
 * @param {Object} negative An object with fields raw and matchType.
 * @param {Object} positive An object with fields raw and matchType.
 * @return {boolean} Whether the negative keyword blocks the positive keyword.
 */
function negativeBlocksPositive(negative, positive) {
  var isNegativeStricter;

  switch (positive.matchType) {
    case 'BROAD':
      isNegativeStricter = negative.matchType != 'BROAD';
      break;

    case 'PHRASE':
      isNegativeStricter = negative.matchType == 'EXACT';
      break;

    case 'EXACT':
      isNegativeStricter = false;
      break;
  }

  if (isNegativeStricter) {
    return false;
  }

  switch (negative.matchType) {
    case 'BROAD':
      return hasAllTokens(negative.raw, positive.raw);
      break;

    case 'PHRASE':
      return isSubsequence(negative.raw, positive.raw);
      break;

    case 'EXACT':
      return positive.raw === negative.raw;
      break;
  }
}

/////////////////////////////////////////////////////////////
///                                                       ///
///     STATISTICAL SIGNIFICANCE AD TESTING FUNCTIONS     ///
///                                                       ///
/////////////////////////////////////////////////////////////
 
// A helper function to return the list of campaign ids with a label for filtering 
function getCampaignNames() {
  var campNames = [];
  var labelIter = AdWordsApp.labels().withCondition("Name = '"+CAMPAIGN_LABEL+"'").get();
  if(labelIter.hasNext()) {
    var label = labelIter.next();
    var campIter = label.campaigns().get();
    while(campIter.hasNext()) {
      campNames.push(campIter.next().getName()); 
    }
  }
  return campNames;
}
  
function applyLabel(entity,label) {
  if(!AdWordsApp.getExecutionInfo().isPreview()) {
    entity.applyLabel(label);
  } else {
    var adText = (entity.getEntityType() === 'Ad') ? [entity.getHeadline(),entity.getDescription1(),
                                                      entity.getDescription2(),entity.getDisplayUrl()].join(' ') 
                                                   : entity.getName();
    Logger.log('PREVIEW: Would have applied label: '+label+' to Entity: '+ adText);
  }
}
  
function removeLabel(ad,label) {
  if(!AdWordsApp.getExecutionInfo().isPreview()) {
    ad.removeLabel(label);
  } else {
    var adText = [ad.getHeadline(),ad.getDescription1(),ad.getDescription2(),ad.getDisplayUrl()].join(' ');
    Logger.log('PREVIEW: Would have removed label: '+label+' from Ad: '+ adText);
  }
}
  
// This function checks if the AdGroup has an Ad with a Champion Label
// If so, the new test should use that as the control.
function checkForPreviousWinner(ag,isMobile) {
  var adSelector = ag.ads().withCondition('Status = ENABLED')
                           .withCondition('AdType = TEXT_AD');
  if(!AdWordsApp.getExecutionInfo().isPreview()) {
    adSelector = adSelector.withCondition("LabelNames CONTAINS_ANY ['"+CHAMPION_LABEL+"']");
  }
  var adIter = adSelector.get();
  while(adIter.hasNext()) {
    var ad = adIter.next();
    if(shouldSkip(isMobile,ad)) { continue; }
    info('Found a previous winner. Using it as the control.');
    return ad;
  }
  return null;
}
 
function shouldSkip(isMobile,ad) {
  if(isMobile) {
    if(!ad.isMobilePreferred()) {
      return true;
    }
  } else {
    if(ad.isMobilePreferred()) {
      return true;
    }
  }
  return false;
}
  
// This function sends the email to the people in the TO array.
// If the script finishes early, it adds a notice to the email.
function sendMailForTouchedAdGroups(ags,finishedEarly) {
  var htmlBody = '<html><head></head><body>';
  if(finishedEarly) {
    htmlBody += 'The script was not able to check all AdGroups. ' +
                'It will check additional AdGroups on the next run.<br / >' ;
  }
  htmlBody += 'The following AdGroups have one or more creative tests that have finished.' ;
  htmlBody += buildHtmlTable(ags);
  htmlBody += '<p><small>Generated by <a href="http://www.freeadwordsscripts.com">FreeAdWordsScripts.com</a></small></p>' ;
  htmlBody += '</body></html>';
  var options = { 
    htmlBody : htmlBody,
  };
  var subject = ags.length + ' Creative Test(s) Completed - ' + 
    Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
  MailApp.sendEmail(RECIPIENT_EMAIL, subject, ags.length+' AdGroup(s) have creative tests that have finished.', options);
}
 
// This function uses my HTMLTable object to build the styled html table for the email.
function buildHtmlTable(ags) {
  var table = new HTMLTable();
  //CSS from: http://coding.smashingmagazine.com/2008/08/13/top-10-css-table-designs/
  //Inlined using: http://inlinestyler.torchboxapps.com/
  table.setTableStyle(['font-family: "Lucida Sans Unicode","Lucida Grande",Sans-Serif;',
                       'font-size: 12px;',
                       'background: #fff;',
                       'margin: 45px;',
                       'width: 480px;',
                       'border-collapse: collapse;',
                       'text-align: left'].join(''));
  table.setHeaderStyle(['font-size: 14px;',
                        'font-weight: normal;',
                        'color: #039;',
                        'padding: 10px 8px;',
                        'border-bottom: 2px solid #6678b1'].join(''));
  table.setCellStyle(['border-bottom: 1px solid #ccc;',
                      'padding: 4px 6px'].join(''));
  table.addHeaderColumn('#');
  table.addHeaderColumn('Campaign Name');
  table.addHeaderColumn('AdGroup Name');
  table.addHeaderColumn('Tests Completed');
  for(var i in ags) {
    table.newRow();
    table.addCell(table.getRowCount());
    var campName = ags[i].getCampaign().getName();
    var name = ags[i].getName();
    var touchCount = ags[i]['touchCount'];
    var campLink, agLink;
    if(__c !== '' && __u !== '') { // You should really set these.
      campLink = getUrl(ags[i].getCampaign(),'Ad groups');
      agLink = getUrl(ags[i],'Ads');
      table.addCell(a(campLink,campName));
      table.addCell(a(agLink,name));
    } else {
      table.addCell(campName);
      table.addCell(name);
    }
    table.addCell(touchCount,'text-align: right');
  }
  return table.toString();
}
 
// Just a helper to build the html for a link.
function a(link,val) {
  return '<a href="'+link+'">'+val+'</a>';
}
  
// This function finds all the previous losers and removes their label.
// It is used when the script detects a change in the AdGroup and needs to 
// start a new test.
function removeLoserLabelsFromAds(ag,isMobile) {
  var adSelector = ag.ads().withCondition('Status = ENABLED');
  if(!AdWordsApp.getExecutionInfo().isPreview()) {
    adSelector = adSelector.withCondition("LabelNames CONTAINS_ANY ['"+LOSER_LABEL+"']");
  }
  var adIter = adSelector.get();
  while(adIter.hasNext()) {
    var ad = adIter.next();
    if(shouldSkip(isMobile,ad)) { continue; }
    removeLabel(ad,LOSER_LABEL);
  }
}
  
// A helper function to create a new label if it doesn't exist in the account.
function createLabelIfNeeded(name,color) {
  if(!AdWordsApp.labels().withCondition("Name = '"+name+"'").get().hasNext()) {
    info('Creating label: "'+name+'"');
    AdWordsApp.createLabel(name,"",color);
  } else {
    info('Label: "'+name+'" already exists.');
  }
}
  
// This function compares the previous and current Ad maps and
// updates the current map with the date that the AdGroup was last touched.
// If OVERRIDE_LAST_TOUCHED_DATE is set and there is no previous data for the 
// AdGroup, it uses that as the last touched date.
function updateCurrentAdMap(current,previous) {
  info('Updating the current Ads map using historical snapshot.');
  for(var rowKey in current) {
    var currentAds = current[rowKey].adIds;
    var previousAds = (previous[rowKey]) ? previous[rowKey].adIds : [];
    if(currentAds.join('-') === previousAds.join('-')) {
      current[rowKey].lastTouched = previous[rowKey].lastTouched;
    }
    if(previousAds.length === 0 && OVERRIDE_LAST_TOUCHED_DATE !== '') {
      current[rowKey].lastTouched = new Date(OVERRIDE_LAST_TOUCHED_DATE);
    }
    //if we make it here without going into the above if statements
    //then the adgroup has changed and we should keep the new date
  }
  info('Finished updating the current Ad map.');
  return current;
}
  
// This stores the Ad map snapshot to a file so it can be used for the next run.
// The data is stored as a JSON string for easy reading later.
function storeAdsSnapshot(data) {
  info('Storing the Ads snapshot to Google Drive.');
  var fileName = getSnapshotFilename();
  var file = DriveApp.getFilesByName(fileName).next();
  file.setContent(Utilities.jsonStringify(data));
  info('Finished.');
}
  
// This reads the JSON formatted previous snapshot from a file on GDrive
// If the file doesn't exist, it creates a new one and returns an empty map.
function getPreviousAdsSnapshot() {
  info('Loading the previous Ads snapshot from Google Drive.');
  var fileName = getSnapshotFilename();
  var fileIter = DriveApp.getFilesByName(fileName);
  if(fileIter.hasNext()) {
    return Utilities.jsonParse(fileIter.next().getBlob().getDataAsString());
  } else {
    DriveApp.createFile(fileName, '');
    return {};
  }
}
  
// A helper function to build the filename for the snapshot.
function getSnapshotFilename() {
  var accountId = AdWordsApp.currentAccount().getCustomerId();
  return (accountId + ' Ad Testing Script Snapshot.json');
}
  
// This function pulls the Ad Performance Report which is the fastest
// way to build a snapshot of the current ads in the account.
// This only pulls in active text ads.
function getCurrentAdsSnapshot() {
  info('Running Ad Performance Report to get current Ads snapshot.');
  var OPTIONS = { includeZeroImpressions : true };
  var cols = ['CampaignId','AdGroupId','Id','DevicePreference','Impressions'];
  var report = 'AD_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where AdType = TEXT_AD',
               'and AdNetworkType1 = SEARCH',
               'and CampaignStatus = ENABLED',
               'and AdGroupStatus = ENABLED',
               'and Status = ENABLED',
               'during','TODAY'].join(' ');
  var results = {}; // { campId-agId : row, ... }
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var rowKey = [row.CampaignId,row.AdGroupId].join('-');
    if(ENABLE_MOBILE_AD_TESTING && row.DevicePreference == 30001) {
      rowKey += '-Mobile';
    }
    if(!results[rowKey]) {
      results[rowKey] = { adIds : [], lastTouched : new Date() };
    }
    results[rowKey].adIds.push(row.Id);
  }
  for(var i in results) {
    results[i].adIds.sort();
  }
  info('Finished building the current Ad map.');
  return results;
}
  
//Helper function to format the date
function getDateString(date,format) {
  return Utilities.formatDate(new Date(date),AdWordsApp.currentAccount().getTimeZone(),format); 
}
  
// Function to build out the urls for deeplinking into the AdWords account.
// For this to work, you need to have __c and __u filled in.
// Taken from: http://www.freeadwordsscripts.com/2013/11/building-entity-deep-links-with-adwords.html
function getUrl(entity,tab) {
  var customerId = __c;
  var effectiveUserId = __u;
  var decodedTab = getTab(tab);  
     
  var base = 'https://adwords.google.com/cm/CampaignMgmt?';
  var url = base+'__c='+customerId+'&__u='+effectiveUserId+'#';
    
  if(typeof entity['getEntityType'] === 'undefined') {
    return url+'r.ONLINE.di&app=cm';
  }
    
  var type = entity.getEntityType()
  if(type === 'Campaign') {
    return url+'c.'+entity.getId()+'.'+decodedTab+'&app=cm';
  }
  if(type === 'AdGroup') {
    return url+'a.'+entity.getId()+'_'+entity.getCampaign().getId()+'.'+decodedTab+'&app=cm';
  }
  if(type === 'Keyword') {
    return url+'a.'+entity.getAdGroup().getId()+'_'+entity.getCampaign().getId()+'.key&app=cm';
  }
  if(type === 'Ad') {
    return url+'a.'+entity.getAdGroup().getId()+'_'+entity.getCampaign().getId()+'.create&app=cm';
  }
  return url+'r.ONLINE.di&app=cm';
     
  function getTab(tab) {
    var mapping = {
      'Ad groups':'ag','Settings:All settings':'st_sum',
      'Settings:Locations':'st_loc','Settings:Ad schedule':'st_as',
      'Settings:Devices':'st_p','Ads':'create',
      'Keywords':'key','Audiences':'au','Ad extensions':'ae',
      'Auto targets':'at','Dimensions' : 'di'
    };
    if(mapping[tab]) { return mapping[tab]; }
    return 'key'; //default to keyword tab
  }
}
  
// Helper function to print info logs
function info(msg) {
  if(EXTRA_LOGS) {
    Logger.log('INFO: '+msg);
  }
}
  
// Helper function to print more serious warnings
function warn(msg) {
  Logger.log('WARNING: '+msg);
}
  
/*********************************************
* Test: A class for runnning A/B Tests for Ads
* Version 1.0
* Based on VisualWebsiteOptimizer logic: http://goo.gl/jiImn
* Russ Savage
* FreeAdWordsScripts.com
**********************************************/
// A description of the parmeters:
// control - the control Ad, test - the test Ad
// startDate, endDate - the start and end dates for the test
// visitorMetric, conversionMetric - the components of the metric to use for the test
function Test(control,test,desiredConf,startDate,endDate,visitorMetric,conversionMetric) {
  this.desiredConfidence = desiredConf/100;
  this.verMetric = visitorMetric;
  this.conMetric = conversionMetric;
  this.startDate = startDate;
  this.endDate = endDate;
  this.winner;
    
  this.controlAd = control;
  this.controlStats = (this.controlAd['stats']) ? this.controlAd['stats'] : this.controlAd.getStatsFor(this.startDate, this.endDate);
  this.controlAd['stats'] = this.controlStats;
  this.controlVisitors = this.controlStats['get'+this.verMetric]();
  this.controlConversions = this.controlStats['get'+this.conMetric]();
  this.controlCR = getConversionRate(this.controlVisitors,this.controlConversions);
    
  this.testAd = test;
  this.testStats = (this.testAd['stats']) ? this.testAd['stats'] : this.testAd.getStatsFor(this.startDate, this.endDate);
  this.testAd['stats'] = this.testStats;
  this.testVisitors = this.testStats['get'+this.verMetric]();
  this.testConversions = this.testStats['get'+this.conMetric]();
  this.testCR = getConversionRate(this.testVisitors,this.testConversions);
    
  this.pValue;
    
  this.getControlVisitors = function() { return this.controlVisitors; }
  this.getControlConversions = function() { return this.controlConversions; }
  this.getTestVisitors = function() { return this.testVisitors; }
  this.getTestConversions = function() { return this.testConversions; }
    
  // Returns the P-Value for the two Ads
  this.getPValue = function() {
    if(!this.pValue) {
      this.pValue = calculatePValue(this);
    }
    return this.pValue;
  };
    
  // Determines if the test has hit significance
  this.isSignificant = function() {
    var pValue = this.getPValue();
    if(pValue && pValue !== 'N/A' && (pValue >= this.desiredConfidence || pValue <= (1 - this.desiredConfidence))) {
      return true;
    }
    return false;
  }
    
  // Returns the winning Ad
  this.getWinner = function() {
    if(this.decideWinner() === 'control') {
      return this.controlAd;
    }
    if(this.decideWinner() === 'challenger') {
      return this.testAd;
    }
    return null;
  };
    
  // Returns the losing Ad
  this.getLoser = function() {
    if(this.decideWinner() === 'control') {
      return this.testAd;
    }
    if(this.decideWinner() === 'challenger') {
      return this.controlAd;
    }
    return null;
  };
    
  // Determines if the control or the challenger won
  this.decideWinner = function () {
    if(this.winner) {
      return this.winner;
    }
    if(this.isSignificant()) {
      if(this.controlCR >= this.testCR) {
        this.winner = 'control';
      } else {
        this.winner = 'challenger';
      }
    } else {
      this.winner = 'no winner';
    }
    return this.winner;
  }
    
  // This function returns the confidence level for the test
  function calculatePValue(instance) {
    var control = { 
      visitors: instance.controlVisitors, 
      conversions: instance.controlConversions,
      cr: instance.controlCR
    };
    var challenger = { 
      visitors: instance.testVisitors, 
      conversions: instance.testConversions,
      cr: instance.testCR
    };
    var z = getZScore(control,challenger);
    if(z == -1) { return 'N/A'; }
    var norm = normSDist(z);
    return norm;
  }
    
  // A helper function to make rounding a little easier
  function round(value) {
    var decimals = Math.pow(10,5);
    return Math.round(value*decimals)/decimals;
  }
    
  // Return the conversion rate for the test
  function getConversionRate(visitors,conversions) {
    if(visitors == 0) {
      return -1;
    }
    return conversions/visitors;
  }
    
  function getStandardError(cr,visitors) {
    if(visitors == 0) {
      throw 'Visitors cannot be 0.';
    }
    return Math.sqrt((cr*(1-cr)/visitors));
  }
    
  function getZScore(c,t) {
    try {
      if(!c['se']) { c['se'] = getStandardError(c.cr,c.visitors); }
      if(!t['se']) { t['se'] = getStandardError(t.cr,t.visitors); }
    } catch(e) {
      Logger.log(e);
      return -1;
    }
      
    if((Math.sqrt(Math.pow(c.se,2)+Math.pow(t.se,2))) == 0) { 
      Logger.log('WARNING: Somehow the denominator in the Z-Score calulator was 0.');
      return -1;
    }
    return ((c.cr-t.cr)/Math.sqrt(Math.pow(c.se,2)+Math.pow(t.se,2)));
  }
    
  //From: http://www.codeproject.com/Articles/408214/Excel-Function-NORMSDIST-z
  function normSDist(z) {
    var sign = 1.0;
    if (z < 0) { sign = -1; }
    return round(0.5 * (1.0 + sign * erf(Math.abs(z)/Math.sqrt(2))));
  }
    
  // From: http://picomath.org/javascript/erf.js.html
  function erf(x) {
    // constants
    var a1 =  0.254829592;
    var a2 = -0.284496736;
    var a3 =  1.421413741;
    var a4 = -1.453152027;
    var a5 =  1.061405429;
    var p  =  0.3275911;
      
    // Save the sign of x
    var sign = 1;
    if (x < 0) {
      sign = -1;
    }
    x = Math.abs(x);
      
    // A&S formula 7.1.26
    var t = 1.0/(1.0 + p*x);
    var y = 1.0 - (((((a5*t + a4)*t) + a3)*t + a2)*t + a1)*t*Math.exp(-x*x);
      
    return sign*y;
  }
}
  
/*********************************************
* HTMLTable: A class for building HTML Tables
* Version 1.0
* Russ Savage
* FreeAdWordsScripts.com
**********************************************/
function HTMLTable() {
  this.headers = [];
  this.columnStyle = {};
  this.body = [];
  this.currentRow = 0;
  this.tableStyle;
  this.headerStyle;
  this.cellStyle;
   
  this.addHeaderColumn = function(text) {
    this.headers.push(text);
  };
   
  this.addCell = function(text,style) {
    if(!this.body[this.currentRow]) {
      this.body[this.currentRow] = [];
    }
    this.body[this.currentRow].push({ val:text, style:(style) ? style : '' });
  };
   
  this.newRow = function() {
    if(this.body != []) {
      this.currentRow++;
    }
  };
   
  this.getRowCount = function() {
    return this.currentRow;
  };
   
  this.setTableStyle = function(css) {
    this.tableStyle = css;
  };
   
  this.setHeaderStyle = function(css) {
    this.headerStyle = css; 
  };
   
  this.setCellStyle = function(css) {
    this.cellStyle = css;
    if(css[css.length-1] !== ';') {
      this.cellStyle += ';';
    }
  };
   
  this.toString = function() {
    var retVal = '<table ';
    if(this.tableStyle) {
      retVal += 'style="'+this.tableStyle+'"';
    }
    retVal += '>'+_getTableHead(this)+_getTableBody(this)+'</table>';
    return retVal;
  };
   
  function _getTableHead(instance) {
    var headerRow = '';
    for(var i in instance.headers) {
      headerRow += _th(instance,instance.headers[i]);
    }
    return '<thead><tr>'+headerRow+'</tr></thead>';
  };
   
  function _getTableBody(instance) {
    var retVal = '<tbody>';
    for(var r in instance.body) {
      var rowHtml = '<tr>';
      for(var c in instance.body[r]) {
        rowHtml += _td(instance,instance.body[r][c]);
      }
      rowHtml += '</tr>';
      retVal += rowHtml;
    }
    retVal += '</tbody>';
    return retVal;
  };
   
  function _th(instance,val) {
    var retVal = '<th scope="col" ';
    if(instance.headerStyle) {
      retVal += 'style="'+instance.headerStyle+'"';
    }
    retVal += '>'+val+'</th>';
    return retVal;
  };
   
  function _td(instance,cell) {
    var retVal = '<td ';
    if(instance.cellStyle || cell.style) {
      retVal += 'style="';
      if(instance.cellStyle) {
        retVal += instance.cellStyle;
      }
      if(cell.style) {
        retVal += cell.style;
      }
      retVal += '"';
    }
    retVal += '>'+cell.val+'</td>';
    return retVal;
  };
}

/////////////////////////////////////////////////////////////
///                                                       ///
///             DISPLAY PLACEMENT MONITORING              ///
///                                                       ///
/////////////////////////////////////////////////////////////

function runLowCostAndConvertingReport()
{
	list = [];
	
	// Any placement detail (individual page) that has converted 2 or more times with CPA below $40
	var report = AdWordsApp.report(
     'SELECT Url, CampaignName, AdGroupName, Clicks, Impressions, Conversions, Cost ' +
     'FROM URL_PERFORMANCE_REPORT ' +
     'WHERE Cost < 40000000 ' +
     'AND Conversions > 2 ' + 
     'DURING LAST_30_DAYS');

  var rows = report.rows();

	 while (rows.hasNext()) {
           
       var row = rows.next();      
       
      var anonymous = row['Url'].match(/anonymous\.google/g);
       if (anonymous == null) { 
	 	   var placement = row['Url'];

       var campaign = row['CampaignName'];
		   var adgroup = row['AdGroupName'];
		   var clicks = row['Clicks'];
		   var impressions = row['Impressions'];
		   var conversions = row['Conversions'];
		   var cost = row['Cost'];
		   
		   var placementDetail = new placementObject(placement, campaign, adgroup, clicks, impressions, conversions, cost);
           
		   list.push(placementDetail);
       } 
	 }
	 return list;
}
function runLowCtrAndHighImpressionsReport()
{
	list = [];
	
	// Any placement detail (individual page) that has converted 2 or more times with CPA below $40
	var report = AdWordsApp.report(
     'SELECT Url, CampaignName, AdGroupName, Clicks, Impressions, Conversions, Cost ' +
     'FROM URL_PERFORMANCE_REPORT ' +
     'WHERE Impressions > 1000 ' +
     'AND Ctr < 0.1 ' + 
     'DURING LAST_30_DAYS');

  var rows = report.rows();

	 while (rows.hasNext()) {
           
       var row = rows.next();      
       
      var anonymous = row['Url'].match(/anonymous\.google/g);
       if (anonymous == null) { 
	 	   var placement = row['Url'];

       var campaign = row['CampaignName'];
		   var adgroup = row['AdGroupName'];
		   var clicks = row['Clicks'];
		   var impressions = row['Impressions'];
		   var conversions = row['Conversions'];
		   var cost = row['Cost'];
		   
		   var placementDetail = new placementObject(placement, campaign, adgroup, clicks, impressions, conversions, cost);
           
		   list.push(placementDetail);
       } 
	 }
	 return list;
}

function runHighCostNoConversionsReport()
{
	list = [];
	
	// Any placement detail (individual page) that has converted 2 or more times with CPA below $40
	var report = AdWordsApp.report(
     'SELECT Url, CampaignName, AdGroupName, Clicks, Impressions, Conversions, Cost ' +
     'FROM URL_PERFORMANCE_REPORT ' +
     'WHERE Cost > 50000000 ' +
     'AND Conversions < 1 ' + 
     'DURING LAST_30_DAYS');

  var rows = report.rows();

	 while (rows.hasNext()) {
           
       var row = rows.next();      
       
      var anonymous = row['Url'].match(/anonymous\.google/g);
       if (anonymous == null) { 
	 	   var placement = row['Url'];

       var campaign = row['CampaignName'];
		   var adgroup = row['AdGroupName'];
		   var clicks = row['Clicks'];
		   var impressions = row['Impressions'];
		   var conversions = row['Conversions'];
		   var cost = row['Cost'];
		   
		   var placementDetail = new placementObject(placement, campaign, adgroup, clicks, impressions, conversions, cost);
           
		   list.push(placementDetail);
       } 
	 }
	 return list;
}

function placementObject(placement, campaign, adgroup, clicks, impressions, conversions, cost) {
  this.placement = placement;
  this.campaign = campaign;
  this.adgroup = adgroup;
  this.clicks = clicks;
  this.impressions = impressions;
  this.conversions = conversions;
  this.cost = cost;
   
}

// Helpers 
function warn(msg) {
  Logger.log('WARNING: '+msg);
}
 
function info(msg) {
  Logger.log(msg);
}

//////////////////////////////////////////////////////////////////////////////
// Find the negative keywords
function searchQueryMining(){
var negativesByGroup = [];
var negativesByCampaign = [];
var sharedSetData = [];
var sharedSetNames = [];
var sharedSetCampaigns = [];
var activeCampaignIds = [];

// Gather ad group level negative keywords

var keywordReport = AdWordsApp.report(
"SELECT CampaignId, AdGroupId, Criteria, KeywordMatchType " +
"FROM   KEYWORDS_PERFORMANCE_REPORT " +
"WHERE CampaignStatus = ENABLED AND AdGroupStatus = ENABLED AND Status = ENABLED AND IsNegative = TRUE " +
"AND CampaignName CONTAINS_IGNORE_CASE '" + campaignNameContains + "' " +
"DURING " + dateRange);

var keywordRows = keywordReport.rows();
while (keywordRows.hasNext()) {
var keywordRow = keywordRows.next();

if (negativesByGroup[keywordRow["AdGroupId"]] == undefined) {
negativesByGroup[keywordRow["AdGroupId"]] = 
[[keywordRow["Criteria"].toLowerCase(),keywordRow["KeywordMatchType"].toLowerCase()]];
} else {

negativesByGroup[keywordRow["AdGroupId"]].push([keywordRow["Criteria"].toLowerCase(),keywordRow["KeywordMatchType"].toLowerCase()]);
}

if (activeCampaignIds.indexOf(keywordRow["CampaignId"]) < 0) {
activeCampaignIds.push(keywordRow["CampaignId"]);
}
}//end while

// Gather campaign level negative keywords

var campaignNegReport = AdWordsApp.report(
"SELECT CampaignId, Criteria, KeywordMatchType " +
"FROM   CAMPAIGN_NEGATIVE_KEYWORDS_PERFORMANCE_REPORT " +
"WHERE  IsNegative = TRUE "
);
var campaignNegativeRows = campaignNegReport.rows();
while (campaignNegativeRows.hasNext()) {
var campaignNegativeRow = campaignNegativeRows.next();

if (negativesByCampaign[campaignNegativeRow["CampaignId"]] == undefined) {
negativesByCampaign[campaignNegativeRow["CampaignId"]] = [[campaignNegativeRow["Criteria"].toLowerCase(),campaignNegativeRow["KeywordMatchType"].toLowerCase()]];
} else {

negativesByCampaign[campaignNegativeRow["CampaignId"]].push([campaignNegativeRow["Criteria"].toLowerCase(),campaignNegativeRow["KeywordMatchType"].toLowerCase()]);
}
}//end while

// Find which campaigns use shared negative keyword sets

var campaignSharedReport = AdWordsApp.report(
"SELECT CampaignName, CampaignId, SharedSetName, SharedSetType, Status " +
"FROM   CAMPAIGN_SHARED_SET_REPORT " +
"WHERE SharedSetType = NEGATIVE_KEYWORDS " +
"AND CampaignName CONTAINS_IGNORE_CASE '" + campaignNameContains + "'");
var campaignSharedRows = campaignSharedReport.rows();
while (campaignSharedRows.hasNext()) {
var campaignSharedRow = campaignSharedRows.next();

if (sharedSetCampaigns[campaignSharedRow["SharedSetName"]] == undefined) {
sharedSetCampaigns[campaignSharedRow["SharedSetName"]] = [campaignSharedRow["CampaignId"]];
} else {

sharedSetCampaigns[campaignSharedRow["SharedSetName"]].push(campaignSharedRow["CampaignId"]);
}
}//end while

// Map the shared sets' IDs (used in the criteria report below)
// to their names (used in the campaign report above)

var sharedSetReport = AdWordsApp.report(
"SELECT Name, SharedSetId, MemberCount, ReferenceCount, Type " +
"FROM   SHARED_SET_REPORT " +
"WHERE ReferenceCount > 0 AND Type = NEGATIVE_KEYWORDS ");
var sharedSetRows = sharedSetReport.rows();
while (sharedSetRows.hasNext()) {
var sharedSetRow = sharedSetRows.next();
sharedSetNames[sharedSetRow["SharedSetId"]] = sharedSetRow["Name"];
}//end while

// Collect the negative keyword text from the sets,
// and record it as a campaign level negative in the campaigns that use the set

var sharedSetReport = AdWordsApp.report(
"SELECT SharedSetId, KeywordMatchType, Criteria " +
"FROM   SHARED_SET_CRITERIA_REPORT ");
var sharedSetRows = sharedSetReport.rows();
while (sharedSetRows.hasNext()) {
var sharedSetRow = sharedSetRows.next();
var setName = sharedSetNames[sharedSetRow["SharedSetId"]];
if (sharedSetCampaigns[setName] !== undefined) {
for (var i=0; i<sharedSetCampaigns[setName].length; i++) {
var campaignId = sharedSetCampaigns[setName][i];
if (negativesByCampaign[campaignId] == undefined) {
negativesByCampaign[campaignId] = 
[[sharedSetRow["Criteria"].toLowerCase(),sharedSetRow["KeywordMatchType"].toLowerCase()]];
} else {

negativesByCampaign[campaignId].push([sharedSetRow["Criteria"].toLowerCase(),sharedSetRow["KeywordMatchType"].toLowerCase()]);
}
}
}
}//end while

Logger.log("Finished negative keyword lists.");

//////////////////////////////////////////////////////////////////////////////
// Defines the statistics to download or calculate, and their formatting

var statColumns = ["Clicks", "Impressions", "Cost", "ConvertedClicks", "ConversionValue"];
var calculatedStats = [["CTR","Clicks","Impressions"],
["CPC","Cost","Clicks"],
["Conv. Rate","ConvertedClicks","Clicks"],
["Cost / conv.","Cost","ConvertedClicks"],
["Conv. value/cost","ConversionValue","Cost"]]
var currencyFormat = currencySymbol + "#,##0.00";
var formatting = ["#,##0", "#,##0", currencyFormat, "#,##0", currencyFormat,"0.00%",currencyFormat,"0.00%",currencyFormat,"0.00%"];


//////////////////////////////////////////////////////////////////////////////
// Go through the search query report, remove searches already excluded by negatives
// record the performance of each word in each remaining query

var queryReport = AdWordsApp.report(
"SELECT CampaignName, CampaignId, AdGroupId, AdGroupName, Query, " + statColumns.join(", ") + " " +
"FROM SEARCH_QUERY_PERFORMANCE_REPORT " +
"WHERE CampaignStatus = ENABLED AND AdGroupStatus = ENABLED " +
"AND CampaignName CONTAINS_IGNORE_CASE '" + campaignNameContains + "' " +
"DURING " + dateRange);

var campaignSearchWords = [];
var totalSearchWords = [];
var totalSearchWordsKeys = [];
var numberOfWords = [];

var queryRows = queryReport.rows();
while (queryRows.hasNext()) {
var queryRow = queryRows.next();
var searchIsExcluded = false;

// Checks if the query is excluded by an ad group level negative

if (negativesByGroup[queryRow["AdGroupId"]] !== undefined) {
for (var i=0; i<negativesByGroup[queryRow["AdGroupId"]].length; i++) {
if ( (negativesByGroup[queryRow["AdGroupId"]][i][1] == "exact" &&
queryRow["Query"] == negativesByGroup[queryRow["AdGroupId"]][i][0]) ||
(negativesByGroup[queryRow["AdGroupId"]][i][1] != "exact" &&
(" "+queryRow["Query"]+" ").indexOf(" "+negativesByGroup[queryRow["AdGroupId"]][i][0]+" ") > -1 )){
searchIsExcluded = true;
break;
}
}
}

// Checks if the query is excluded by a campaign level negative

if (!searchIsExcluded && negativesByCampaign[queryRow["CampaignId"]] !== undefined) {
for (var i=0; i<negativesByCampaign[queryRow["CampaignId"]].length; i++) {
if ( (negativesByCampaign[queryRow["CampaignId"]][i][1] == "exact" &&
queryRow["Query"] == negativesByCampaign[queryRow["CampaignId"]][i][0]) ||
(negativesByCampaign[queryRow["CampaignId"]][i][1]!= "exact" &&
(" "+queryRow["Query"]+" ").indexOf(" "+negativesByCampaign[queryRow["CampaignId"]][i][0]+" ") > -1 )){
searchIsExcluded = true;
break;
}
}
}

if (searchIsExcluded) {continue;}
// if the search is already excluded by the current negatives,
// we ignore it and go on to the next query

var currentWords = queryRow["Query"].split(" ");
var doneWords = [];

if (campaignSearchWords[queryRow["CampaignName"]] == undefined) {
campaignSearchWords[queryRow["CampaignName"]] = [];
}

var wordLength = currentWords.length;
if (wordLength > 6) {
wordLength = "7+";
}
if (numberOfWords[wordLength] == undefined) {
numberOfWords[wordLength] = [];
}
for (var i=0; i<statColumns.length; i++) {
if (numberOfWords[wordLength][statColumns[i]] > 0) {
numberOfWords[wordLength][statColumns[i]] += parseFloat(queryRow[statColumns[i]].replace(/,/g, ""));
} else {
numberOfWords[wordLength][statColumns[i]] = parseFloat(queryRow[statColumns[i]].replace(/,/g, ""));
}
}


// Splits the query into words and records the stats for each

for (var w=0;w<currentWords.length;w++) {
if (doneWords.indexOf(currentWords[w]) < 0) { //if this word hasn't been in the query yet

if (campaignSearchWords[queryRow["CampaignName"]][currentWords[w]] == undefined) {
campaignSearchWords[queryRow["CampaignName"]][currentWords[w]] = [];
}
if (totalSearchWords[currentWords[w]] == undefined) {
totalSearchWords[currentWords[w]] = [];
totalSearchWordsKeys.push(currentWords[w]);
}

for (var i=0; i<statColumns.length; i++) {
var stat = parseFloat(queryRow[statColumns[i]].replace(/,/g, ""));
if (campaignSearchWords[queryRow["CampaignName"]][currentWords[w]][statColumns[i]] > 0) {
campaignSearchWords[queryRow["CampaignName"]][currentWords[w]][statColumns[i]] += stat;
} else {
campaignSearchWords[queryRow["CampaignName"]][currentWords[w]][statColumns[i]] = stat;
}
if (totalSearchWords[currentWords[w]][statColumns[i]] > 0) {
totalSearchWords[currentWords[w]][statColumns[i]] += stat;
} else {
totalSearchWords[currentWords[w]][statColumns[i]] = stat;
}
}

doneWords.push(currentWords[w]);
}//end if
}//end for
}//end while

Logger.log("Finished analysing queries.");


//////////////////////////////////////////////////////////////////////////////
// Output the data into the spreadsheet

var campaignSearchWordsOutput = [];
var campaignSearchWordsFormat = [];
var totalSearchWordsOutput = [];
var totalSearchWordsFormat = [];
var wordLengthOutput = [];
var wordLengthFormat = [];

// Add headers

var calcStatNames = [];
for (var s=0; s<calculatedStats.length; s++) {
calcStatNames.push(calculatedStats[s][0]);
}
var statNames = statColumns.concat(calcStatNames);
campaignSearchWordsOutput.push(["Campaign","Word"].concat(statNames));
totalSearchWordsOutput.push(["Word"].concat(statNames));
wordLengthOutput.push(["Word count"].concat(statNames));

// Output the campaign level stats

for (var campaign in campaignSearchWords) {
for (var word in campaignSearchWords[campaign]) {

if (campaignSearchWords[campaign][word]["Impressions"] < impressionThreshold) {continue;}
if (campaignSearchWords[campaign][word]["Clicks"] < clickThreshold) {continue;}
if (campaignSearchWords[campaign][word]["Cost"] < costThreshold) {continue;}
if (campaignSearchWords[campaign][word]["ConvertedClicks"] < conversionThreshold) {continue;}

// skips words under the thresholds

var printline = [campaign, word];

for (var s=0; s<statColumns.length; s++) {
printline.push(campaignSearchWords[campaign][word][statColumns[s]]);
}

for (var s=0; s<calculatedStats.length; s++) {
var multiplier = calculatedStats[s][1];
var divisor = calculatedStats[s][2];
if (campaignSearchWords[campaign][word][divisor] > 0) {
printline.push(campaignSearchWords[campaign][word][multiplier] / campaignSearchWords[campaign][word][divisor]);
} else {
printline.push("-");
}
}

campaignSearchWordsOutput.push(printline);
campaignSearchWordsFormat.push(formatting);
}
} // end for


totalSearchWordsKeys.sort(function(a,b) {return totalSearchWords[b]["Cost"] - totalSearchWords[a]["Cost"];});

for (var i = 0; i<totalSearchWordsKeys.length; i++) {
var word = totalSearchWordsKeys[i];

if (totalSearchWords[word]["Impressions"] < impressionThreshold) {continue;}
if (totalSearchWords[word]["Clicks"] < clickThreshold) {continue;}
if (totalSearchWords[word]["Cost"] < costThreshold) {continue;}
if (totalSearchWords[word]["ConvertedClicks"] < conversionThreshold) {continue;}

// skips words under the thresholds

var printline = [word];

for (var s=0; s<statColumns.length; s++) {
printline.push(totalSearchWords[word][statColumns[s]]);
}

for (var s=0; s<calculatedStats.length; s++) {
var multiplier = calculatedStats[s][1];
var divisor = calculatedStats[s][2];
if (totalSearchWords[word][divisor] > 0) {
printline.push(totalSearchWords[word][multiplier] / totalSearchWords[word][divisor]);
} else {
printline.push("-");
}
}

totalSearchWordsOutput.push(printline);
totalSearchWordsFormat.push(formatting);
} // end for

for (var i = 1; i<8; i++) {
if (i < 7) {
var wordLength = i;
} else {
var wordLength = "7+";
}

var printline = [wordLength];

if (numberOfWords[wordLength] == undefined) {
printline.push([0,0,0,0,"-","-","-","-"]);
} else {
for (var s=0; s<statColumns.length; s++) {
printline.push(numberOfWords[wordLength][statColumns[s]]);
}

for (var s=0; s<calculatedStats.length; s++) {
var multiplier = calculatedStats[s][1];
var divisor = calculatedStats[s][2];
if (numberOfWords[wordLength][divisor] > 0) {
printline.push(numberOfWords[wordLength][multiplier] / numberOfWords[wordLength][divisor]);
} else {
printline.push("-");
}
}
}

wordLengthOutput.push(printline);
wordLengthFormat.push(formatting);
} // end for

// Finds available names for the new sheets

var campaignWordName = "Campaign Word Analysis";
var totalWordName = "Total Word Analysis";
var wordCountName = "Word Count Analysis";
var campaignWordSheet = spreadsheet.getSheetByName(campaignWordName);
var totalWordSheet = spreadsheet.getSheetByName(totalWordName);
var wordCountSheet = spreadsheet.getSheetByName(wordCountName);
var i = 1;
while (campaignWordSheet != null || wordCountSheet != null || totalWordSheet != null) {
campaignWordName = "Campaign Word Analysis " + i;
totalWordName = "Total Word Analysis " + i;
wordCountName = "Word Count Analysis " + i;
campaignWordSheet = spreadsheet.getSheetByName(campaignWordName);
totalWordSheet = spreadsheet.getSheetByName(totalWordName);
wordCountSheet = spreadsheet.getSheetByName(wordCountName);
i++;
}
campaignWordSheet = spreadsheet.insertSheet(campaignWordName);
totalWordSheet = spreadsheet.insertSheet(totalWordName);
wordCountSheet = spreadsheet.insertSheet(wordCountName);

campaignWordSheet.getRange("R1C1").setValue("Analysis of Words in Search Query Report, By Campaign");
wordCountSheet.getRange("R1C1").setValue("Analysis of Search Query Performance by Words Count");

if (campaignNameContains == "") {
totalWordSheet.getRange("R1C1").setValue("Analysis of Words in Search Query Report, By Account");
} else {
totalWordSheet.getRange("R1C1").setValue("Analysis of Words in Search Query Report, Over All Campaigns Containing '" + campaignNameContains + "'");
}

campaignWordSheet.getRange("R2C1:R" + (campaignSearchWordsOutput.length+1) + "C" + campaignSearchWordsOutput[0].length).setValues(campaignSearchWordsOutput);
campaignWordSheet.getRange("R3C3:R" + (campaignSearchWordsOutput.length+1) + "C" + (formatting.length+2)).setNumberFormats(campaignSearchWordsFormat);
totalWordSheet.getRange("R2C1:R" + (totalSearchWordsOutput.length+1) + "C" + totalSearchWordsOutput[0].length).setValues(totalSearchWordsOutput);
totalWordSheet.getRange("R3C2:R" + (totalSearchWordsOutput.length+1) + "C" + (formatting.length+1)).setNumberFormats(totalSearchWordsFormat);
wordCountSheet.getRange("R2C1:R" + (wordLengthOutput.length+1) + "C" + wordLengthOutput[0].length).setValues(wordLengthOutput);
wordCountSheet.getRange("R3C2:R" + (wordLengthOutput.length+1) + "C" + (formatting.length+1)).setNumberFormats(wordLengthFormat);

Logger.log("Finished writing to spreadsheet.");
}

/////////////////////////////////////////////////////////////
///                                                       ///
///                  UNIVERSAL FUNCTIONS                  ///
///                                                       ///
/////////////////////////////////////////////////////////////

/**
 * Retrieves the spreadsheet identified by the URL.
 * @param {string} spreadsheetUrl The URL of the spreadsheet.
 * @return {SpreadSheet} The spreadsheet.
 */
function copySpreadsheet(spreadsheetUrl) {
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl).copy(
      accountName + ' - Performance Report - ' +
      getDateStringInTimeZone('MMM dd, yyyy HH:mm:ss z'));

  // Make sure the spreadsheet is using the account's timezone.
  spreadsheet.setSpreadsheetTimeZone(currentAccount.getTimeZone());
  return spreadsheet;
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

// Format CTR
function ctr(number) {
  return parseInt(number * 10000) / 10000 + '%';
}
// Returns YYYYMMDD-formatted date.
function getDateInThePast(numDays) {
  var today = new Date();
  var timeZone = currentAccount.getTimeZone();
  today.setDate(today.getDate() - numDays);
  return Utilities.formatDate(today, timeZone, 'yyyyMMdd');
}
