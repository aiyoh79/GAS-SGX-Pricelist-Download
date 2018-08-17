// Project: Script to download SGX price list for personal portfolio price update
// Author: cwtan
// Date created: 14 Aug 2018

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Import SGX data...', functionName: 'importSGX'},
  ];
  spreadsheet.addMenu('Import', menuItems);
}

function importSGX() {
  var spreadsheet = SpreadsheetApp.getActive();
  var dataSheet = spreadsheet.getSheetByName('Data');
  dataSheet.activate();
   
  // Since the data is only available after market close
  // codes below will make sure when it is run before market close, pick yesterday date
  // Current assumption is data is available after 6pm
  var todayDateTime = new Date();
  var currentHour = todayDateTime.getHours();
  
  if (currentHour >= 18) {
    var dataDate = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    Logger.log(dataDate);
  }else{
    var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    var now = new Date();
    var dataDate = Utilities.formatDate(new Date(now.getTime() - MILLIS_PER_DAY), "GMT+8", "yyyy-MM-dd");
    Logger.log(dataDate);
  }

  var importString = "http://infopub.sgx.com/Apps?A=COW_Prices_Content&B=SecuritiesHistoricalPrice&F=5254&G=SESprice.dat&H=" + dataDate;
  
  // import SGX price list data into datasheet
  var csvUrl = importString;
  var csvContent = UrlFetchApp.fetch(csvUrl).getContentText();
  var csvData = Utilities.parseCsv(csvContent,';');
  
  dataSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  
  // Check if data is correct, if there is valid data, cell(1,1) won't be blank.
  // If there is valid data, split the data using ; delimiter.
  var checkValidData = dataSheet.getRange(1,1);
  if (!checkValidData.isBlank()){
    
    var startRow = 1;
    var endRow = dataSheet.getLastRow();
   
    // Get the range value for stock codes, convert all the values to string and trim it.
    // Withut trim, matching could be troublesome because you need to
    // manually add space so to match
    var range = dataSheet.getRange("O1:O");
    var rangeValues = range.getValues();
    range.setNumberFormat("@");
    
    for (var i = startRow; i <= endRow; i++) {
       // accessing the array data, array start from 0, that's why need to put i-1
       var data = rangeValues[i-1][0];
       dataSheet.getRange(i,15).setValue(String(data).trim());
      }
   }
}