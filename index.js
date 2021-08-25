// To learn how to run this script in Google Sheets go to https://keywordsinsheets.com/title-checker/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Title Checker')
    .addItem('Check Titles', 'titlechecker')
    .addToUi();
}

function titlechecker() {

  // Add Moment.js library
  eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js').getContentText());

  // add API key from https://rapidapi.com/apigeek/api/google-search3
  const apiKey = "YOUR-API-HERE";

  // get active spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // function to get last row of single column
  const getLastRowCol = (range) => {
    let rowNum = 0;
    let blank = false;
    for (let row = 0; row < range.length; row++) {

      if (range[row][0] === "" && !blank) {
        rowNum = row;
        blank = true;

      } else if (range[row][0] !== "") {
        blank = false;
      };
    };
    return rowNum;
  }

  // range variables to make sure data can always be appended
  const titleRange = ss.getRange('B:B').getValues();
  const titleLastRow = getLastRowCol(titleRange) + 1;
  const lastRowDiff = ss.getLastRow() - titleLastRow + 1;
  const queryRange = ss.getRange(titleLastRow, 1, lastRowDiff).getValues();

  // domain and language variables
  const domain = ss.getRange('G2').getValue().replace(/^(?:https?:\/\/)?(?:www\.)?/i, "").split('/')[0];
  const country = ss.getRange('H2').getValue();
  const language = ss.getRange('I2').getValue();

  // function to fetch SERPs from rapidapi
  const fetchSerps = (keyword, language, country) => {
    try {
      const options = {
        'method': 'GET',
        'contentType': 'application/json',
        'headers': {
          'x-rapidapi-key': apiKey,
          'x-rapidapi-host': 'google-search3.p.rapidapi.com'
        }
      };

      const serpResponse = UrlFetchApp.fetch(`https://google-search3.p.rapidapi.com/api/v1/search/q=${keyword}&gl=${country}&hl=lang_${language}&num=100`, options);

      const content = JSON.parse(serpResponse.getContentText());

      const organicResults = content.results;

      let row = [],
        data;
      for (i = 0; i < organicResults.length; i++) {

        let data = organicResults[i];

        const link = data.link;

        // if any of the top 100 ranking URLs include the domain, return data
        if (link.includes(domain)) {

          // scrape the ranking URL
          const urlResponse = UrlFetchApp.fetch(data.link).getContentText();

          const $ = Cheerio.load(urlResponse);

          // extract the title
          const title = $('title').first().text().trim();

          // check whether ranking title is different to page title
          const changed = title !== data.title ? "Changed" : "Unchanged";

          row.push(title, data.title, changed);

          return row;

        }
      }


    } catch (e) {
      return false;
    }
  }

  // loop over remaining URLs and set values while the loop is running
  queryRange.forEach(function (row, i) {
    row.forEach(function (col) {

      // visually display a fetch status. Script inspired by https://script.gs visually-display-status-when-looping-through-google-sheets-data/
      ss.getRange(titleLastRow + i, 5).setValue("Loading...");

      SpreadsheetApp.flush();

      const check = fetchSerps(col, language, country);

      // if the SERPs check is successful, return row 
      check ? ss.getRange(titleLastRow + i, 2, 1, 3).setValues([check]) : ss.getRange(titleLastRow + i, 2).setValue("No data");
      ss.getRange(titleLastRow + i, 5).setValue("Done");

      // Add timestamp of extraction
      ss.getRange(titleLastRow + i, 6).setValue(moment().format('llll'))

      SpreadsheetApp.flush();

    });
  });

}