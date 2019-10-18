// Project: Script to download SGX price list for personal portfolio price update
// Author: cwtan

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
   
  // Getting date and the counter from 'Data' excel sheet 
  var prevDate = new Date(dataSheet.getRange(1, 25).getValue());
  var counter = dataSheet.getRange(1, 26).getValue();
  // current date
  var todayDate = new Date();

  // if prevDate is already indicate today's date, no need to run script as data already in  
  if (sameDay(prevDate,todayDate)){
    Logger.log("same date, no need to calculate, exit the script")
    return 0;
  }
    
  if (todayDate.getHours() < 17) {
    // set to yesterday date
    todayDate.setDate(todayDate.getDate() - 1)
    if (sameDay(prevDate,todayDate)){
      Logger.log("script run before 7pm where today's file is not available for download, and yesterday date file is already processed, so no need to run script, exit the script.")
      return 0;
    }
    
    Logger.log("Script run before 7pm where today's file is not available for download, let the script run with yesterday date to retrieve yesterday file for processing.")
  }
    
  // calculate the number of days between today's date and prevDate (date recorded in the excel sheet)
  // if prevDate is 14 Oct 2019 and today's date is 16 Oct 2019, the number of days should be 2
  const oneDay = 24 * 60 * 60 * 1000; // hours*minutes*seconds*milliseconds
  var nDays = Math.floor(Math.abs((todayDate - prevDate) / oneDay));
  dataSheet.getRange(1, 1).setValue(nDays);
  
  // loop through number of days to see what is the correct counter value to use
  var loopDate = prevDate;
  for (var i = 0; i < nDays; i++) {
    loopDate.setDate(loopDate.getDate() + 1);
    // day => Sunday = 0 ... Saturday = 6
    var day = loopDate.getDay();
    // only increase the counter is it is Monday to Friday (1 to 5) 
    if (day != 0 && day != 6){
      // on increase the counter if it is not holiday
      if (holiday(loopDate) == 1) {
        Logger.log("is holiday, no need to increase counter")
      }
      else{
        counter = counter + 1;
      }
    }
  }
  Logger.log(todayDate + " + " + counter);
    
  // build the import url string
  // 5566 = 14 Oct 2019
  var importString = "https://links.sgx.com/1.0.0/securities-historical/" + counter + "/SESprice.dat";
    
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
    
    dataSheet.getRange(1, 20).setValue(endRow);
    
    // Update the "data" sheet for latest date and counter
    dataSheet.getRange(1, 25).setValue(todayDate); // date
    dataSheet.getRange(1, 26).setValue(counter); // counter  
   }
}

function holiday(dateToCheck) {
  // This function can help check whether the date you enter is public holiday
  // This is based on Google holiday calander, however, Google does not follow Minstry of Manpower (MOM) published public holiday, 
  // therefore an exclusion list is created to filter out those holiday not listed in MOM list.
  // special holiday should not be included, example, polling day
  
  //var cal = CalendarApp.getCalendarById("en.singapore#holiday@group.v.calendar.google.com");
  var cal = CalendarApp.getCalendarsByName("Holidays in Singapore");
  // exclusion list of holidays which are not part of MOM published public holiday
  const nonHolidayArray = ["Christmas Eve","New Year's Eve","Children's Day","Easter Saturday","Easter Sunday"];
  //var holidays = cal.getEventsForDay(dateToCheck);
  var holidays = cal[0].getEventsForDay(dateToCheck);
  
  var isHoliday = 0;
  // if there is only one holiday per day  
  if (holidays.length >= 1){
    for (var i = 0; i < holidays.length; i++){
      var eventTitle = holidays[i].getTitle();
      // if the title is not in the exclusion list, then is holiday 
      if ( nonHolidayArray.indexOf(eventTitle) == -1 ){
        isHoliday = 1;
      }
    }
  }
  return isHoliday;
}

function sameDay(date1, date2) {
  if (date1.getFullYear() == date2.getFullYear() && date1.getMonth() == date2.getMonth() && date1.getDate() == date2.getDate()) {
    return 1;
  }else{
    return 0;
  }
}