const ACCURANKER_API_KEY = // GET FROM DEV

const SIX_STAR_DOMAIN_ADDRESS = // GET FROM DEV
const SIX_STAR_DOMAIN_ID = // GET FROM DEV

const SLACK_WEBHOOK_URL = // GET FROM DEV

/**
 * Fetches the list of keywords and their respective rank from Accuranker.
 * 
 * @returns Array[Object]
 */
function getAccuRankerKeywords () {
  var url = `https://app.accuranker.com/api/v4/domains/${ SIX_STAR_DOMAIN_ID }/keywords?fields=keyword,ranks.rank`;
  
  var options = {
    headers: {
      'Authorization': 'Token ' + ACCURANKER_API_KEY
    }
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var rankings = JSON.parse(response.getContentText());

  var keywords = [];
  rankings.forEach(ranking => {
    keywords.push({
      'keyword': ranking.keyword,
      'rank': ranking.ranks.length > 0 ? ranking.ranks[0].rank : null
    });
  });

  return keywords;
}

/**
 * Get the list of keywords from the spreadsheet as an array of strings.
 * 
 * @returns Array[String]
 */
function getSpreadsheetKeywords () {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  if (sheet !== null) {
    var numRows = sheet.getLastRow();
    var keywordObjects = sheet.getRange(`A2:A${ numRows }`).getValues();

    var keywords = keywordObjects.map(function(row) {
      return row[0].toString();
    });

    return keywords;
  }
}

/**
 * Returns the best rank for a provided keyword.
 * 
 * @returns Int - the lowest rank
 */
function getRankForKeyword (keyword, keywords) {
  var keyword = keyword.trim();

  var ranks = [];
  keywords.forEach(keywordObject => {
    if (keywordObject.keyword.trim() === keyword)
      ranks.push(keywordObject.rank);
  });

  var filteredRanks = ranks.filter(function(value) {
    return value !== null && typeof value !== 'undefined';
  });

  // Return 101 if no valid rank is available, this allows for the
  // keyword to still be filtered in the Data Studio console.
  if (filteredRanks.length == 0)
    return 101;

  return Math.min(...filteredRanks);
}

/**
 * Returns the letter for a given column index.
 * 
 * @returns String
 */
function getColumnLetter (columnIndex) {
  var columnLetter = "";
  while (columnIndex > 0) {
    var remainder = (columnIndex - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnIndex = Math.floor((columnIndex - 1) / 26);
  }
  
  return columnLetter;
}

/**
 * Returns the titles of the columns of the spreadsheet.
 * 
 * @returns Array[String]
 */
function getColumnTitles () {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet !== null) {
    var range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    var firstRowValues = range.getValues()[0];

    return firstRowValues;
  }

  return [];
}

/**
 * Returns the column index for the new rank column.
 * 
 * @returns Integer
 */
function getRankColumnIndex () {
  var columns = getColumnTitles();

  var columnIndex = -1;
  columns.forEach((column, index) => {
    if (column.toLowerCase().includes('rank') && columnIndex === -1)
      columnIndex = index;
  });

  return columnIndex + 1;
}

/**
 * Returns the column index for the new weighted rank column.
 * 
 * @returns Integer
 */
function getWeightedRankColumnIndex () {
  var columns = getColumnTitles();

  return columns.indexOf('Search volume') + 2;
}

/**
 * Fetches the keyword list from the spreadsheet and cross-compares with the keyword
 * data returned by AccuRanker. Any data that has multiple ranks is compared so that
 * the best rank is returned. Returned data is added as a new column into the spreadsheet.
 */
function getRankings () {
  var accurankerKeywords = getAccuRankerKeywords();
  var spreadsheetKeywords = getSpreadsheetKeywords();

  var completeKeywordList = [];
  spreadsheetKeywords.forEach(keyword => {
    completeKeywordList.push({
      'keyword': keyword,
      'rank': getRankForKeyword(keyword, accurankerKeywords)
    });
  });

  var rankColumnIndex = getRankColumnIndex();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet !== null) {
    // -1 to account for the title cell
    var numRows = sheet.getLastRow() - 1;

    if (numRows !== completeKeywordList.length) {
      // Log the error
      Logger.log('The number of rows does not match up for the spreadsheet and the keywords.');

      // Construct and send a Slack notification
      var payload = constructSlackNotification(
        'Couldn\'t update the spreadsheet', 'The number of rows does not match up for the spreadsheet and the keywords.', false);
      sendSlackAlert(payload);
    }

    sheet.insertColumnBefore(rankColumnIndex);

    // Get the current date
    var currentDate = new Date();
    // Get the previous month's date
    var previousMonthDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1);
    // Format the previous month's date
    var formattedDate = Utilities.formatDate(previousMonthDate, Session.getScriptTimeZone(), 'MMMyy');

    // Generate a title cell
    var titleCell = sheet.getRange(1, rankColumnIndex);
    titleCell.setValue(`${ formattedDate } Rank`);

    // Generate the spreadsheet rank rows by looping through the keywords array
    var rankRows = [];
    completeKeywordList.forEach(keywordObject => {
      rankRows.push([ keywordObject.rank ]);
    });

    var range = sheet.getRange(2, rankColumnIndex, numRows);
    range.setValues(rankRows);

    // Called here to account for the addition of the earlier column for rank
    var weightedColumnIndex = getWeightedRankColumnIndex();

    sheet.insertColumnBefore(weightedColumnIndex);

    // Generate a title cell
    var titleCell = sheet.getRange(1, weightedColumnIndex);
    titleCell.setValue(`${ formattedDate } Weighted`);

    for (var row = 2; row < completeKeywordList.length; row++) {
      var cell = sheet.getRange(row, weightedColumnIndex);

      // Get column letter for Search Volume (column before Weighted)
      var searchVolumeColumnLetter = getColumnLetter(weightedColumnIndex - 1);
      // Get column letter for corresponding rank (to the weighted rank)
      var rankColumnLetter = getColumnLetter(rankColumnIndex);

      cell.setFormula(`=SUM(${ searchVolumeColumnLetter }${ row }*${ rankColumnLetter }${ row })`);
    }

    // Construct and send a Slack notification
    var payload = constructSlackNotification(
      'Successfully updated the spreadsheet', `The spreadsheet was successfully updated to include ranking data for ${ formattedDate }.`);
    sendSlackAlert(payload);
  }
}

/**
 * Constructs and returns the body of a Slack notification.
 * 
 * @returns Array[String]
 */
function constructSlackNotification (title, message, success = true) {
  var emoji = ':white_check_mark:';
  if (!success)
    emoji = ':warning:';

  return {
    'blocks': [
      {
        'type': 'section',
        'text': {
          'type': 'mrkdwn',
          'text': `*${ title }* ${ emoji }`
        }
      },
      {
        'type': 'divider'
      },
      {
        'type': 'section',
        'text': {
          'type': 'mrkdwn',
          'text': `${ message }`
        }
      }
    ]
  };
}

/**
 * Sends a Slack alert to the webhook.
 */
function sendSlackAlert (payload) {
  const webhook = SLACK_WEBHOOK_URL;
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'muteHttpExceptions': true,
    'payload': JSON.stringify(payload)
  };
  
  try {
    UrlFetchApp.fetch(webhook, options);
  } catch(e) {
    Logger.log(e);
  }
}
