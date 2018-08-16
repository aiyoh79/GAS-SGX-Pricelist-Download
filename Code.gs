// Project: Script to download SGX price list for personal portfolio price update
// Author: cwtan
// Date created: 14 Aug 2018
// Revision History:
// 14 Aug 2018 - version 0.1 - Code creation

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Import SGX data...', functionName: 'importSGX_'},
  ];
  spreadsheet.addMenu('Import', menuItems);
}

function importSGX_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var dataSheet = spreadsheet.getSheetByName('Data');
  dataSheet.activate();
   
  // Since the data is only available after market close
  // codes below will make sure when it is run before market close, pick yesterday date
  // Current assumption is data is available after 6pm
  var todayDateTime = new Date();
  var currentHour = todayDateTime.getHours();

  if (currentHour > 18) {
    var dataDate = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    Logger.log(dataDate);
  }else{
    var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    var now = new Date();
    var dataDate = Utilities.formatDate(new Date(now.getTime() - MILLIS_PER_DAY), "GMT+8", "yyyy-MM-dd");
    Logger.log(dataDate);
  }

  var importString = "http://infopub.sgx.com/Apps?A=COW_Prices_Content&B=SecuritiesHistoricalPrice&F=5254&G=SESprice.dat&H=" + dataDate;
  //var importString = "http://infopub.sgx.com/Apps?A=COW_Prices_Content&B=SecuritiesHistoricalPrice&F=5254&G=SESprice.dat&H=2018-08-14";
  
  // import SGX price list data into datasheet
  dataSheet.getRange('A1').setFormula('=IMPORTDATA("' + importString + '")')
  
  // Check if data is correct, if there is valid data, cell(1,1) won't be blank.
  // If there is valid data, split the data using ; delimiter.
  var checkValidData = dataSheet.getRange(1,1);
  if (!checkValidData.isBlank()){
    
    var startRow = 1;
    var endRow = dataSheet.getLastRow();
    
    // Get the range value for the imported data, which is column 1. 
    // This is faster than use getValue which will retrieve value one by one from server
    // getRange(row, column, numRows, numColumns)
    var rangeValues = dataSheet.getRange(1,1,endRow,1).getValues();
    
    for (var i = startRow; i <= endRow; i++) {
       // accessing the array data, array start from 0, that's why need to put i-1
       var data = rangeValues[i-1][0];
       var splitData = data.split(";");
       dataSheet.getRange(i,3,1,splitData.length).setValues([splitData]); 
       }
   
    // Get the range value for stock codes, convert all the values to string and trim it.
    // Withut trim, matching could be troublesome because you need to
    // manually add space so to match
    var rangeValues = dataSheet.getRange("Q1:Q").getValues();
    for (var i = startRow; i <= endRow; i++) {
       // accessing the array data, array start from 0, that's why need to put i-1
       var data = rangeValues[i-1][0];
       dataSheet.getRange(i,17).setValue(String(data).trim());
      }
   }
}