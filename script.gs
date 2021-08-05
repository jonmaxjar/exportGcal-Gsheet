// create Menu

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('CalTool')
      .addItem('Calendar Loader', 'showUserForm')
      .addToUi();
}

//Use HTML form to get the Data

function showUserForm () {

  var template = HtmlService.createTemplateFromFile("userform");

  var html = template.evaluate();

  html.setTitle("Calendar Loader");

  SpreadsheetApp.getUi().showSidebar(html);

}

// Use the data to get all events betwen

function loadCalendar (data){  

  var mycal = [data.name];
  var cal = CalendarApp.getCalendarById(mycal)
  var calName = cal.getName();
  
  //validates is sheet exist, if not it create a sheet with the ID name
  if(!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(calName))){
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(String(calName));
  }
  var sheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(calName));
  var fromDate = [data.fromc];
  var toDate = [data.toc];
  
  //Get all events in the calendar at the selected time
  var cal = CalendarApp.getCalendarById(mycal);
  var events = cal.getEvents(new Date(fromDate + "T00:00:01"), new Date(toDate+ "T23:00:01"));
  sheet.clearContents();
  //put headers
  var header = [["Calendar Address", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event"]]
  var range = sheet.getRange(1,1,1,14);
  range.setValues(header);

//Loop to add all events in the calendar
  for (var i=0;i<events.length;i++) {
var row=i+2;
var myformula_placeholder = '';
// Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
// NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
var details=[[mycal,events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent()]];
var range=sheet.getRange(row,1,1,14);
range.setValues(details);

// Writing formulas from scripts requires that you write the formulas separate from non-formulas
// Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns F-E for time calc
var cell=sheet.getRange(row,7);
cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))');
cell.setNumberFormat('.00');

}
}