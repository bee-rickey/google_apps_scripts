/*
  Step 1: Get list of all calendars one has access to. 
  Step 2: For the first calendar, go to the SHEET_NAME sheet. Sort by date column and delete all events that are greater than today - DAYS_BEFORE. This will allow for any future events to be updated in the SHEET_NAME sheet. DAYS_BEFORE is a small buffer given to accomodate for any events older than two days being updated retrospectively.
  Step 3: Pull all events starting from today - DAYS_BEFORE to today + NUMBER_OF_MONTHS for the current calendar.
  Step 4: For each of the event, get the required values and populate in SHEET_NAME sheet.
  Step 5: Repeate Step 3 and Step 4 for the rest of the calendars one has access to.

  Additional optimisation: In step 2, while deleting, if you hit 10 consecutive events that are having a timestamp lesser than today - DAYS_BEFORE, then stop deleting rows and move to adding new events (this is to avoid unwanted deletion of whole sheet). 

*/

/* Sheet into which the calendar events will be added. Make sure this sheet name matches the sheet where you want entries to be made */
const SHEET_NAME = 'HNI Calendar';

/* The start date from which calendar pull will be done. Any event before this will be ignored */
const START_DATE = '2021-04-01 00:00:00';

/* Number of months of calendar events to be pulled from the current date */
const NUMBER_OF_MONTHS = 2;

/* 
  This is used to decide which all rows of the sheet need to be cleared during each run.
  Example, if DAYS_BEFORE = 2, then while populating the sheet, all rows having timestamp greater than today - 2 days will be deleted and repopulated.
*/
const DAYS_BEFORE = 2;

/* This function takes a calendar id and a date to fetch from (optional) and fetches all events for the calendar for next 'x' months */
function listUpcomingEvents(calendarId, dateToFetchFrom) {
 
  if (calendarId == undefined) {
    calendarId = "<>";
    console.log("Error: Calendar Id not passed");
    //return [];
  } 

  /* If no starting date is given, then the current date is assumed to be the starting date */
  if (dateToFetchFrom == undefined) {
    dateToFetchFrom = new Date(START_DATE);
  }

  let today = new Date();
  let tillDate = new Date(today.getFullYear(), today.getMonth() + NUMBER_OF_MONTHS, today.getDate());

  
  let optionalArgs = {
    timeMin: dateToFetchFrom.toISOString(),
    timeMax: tillDate.toISOString(),
    showDeleted: false,
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 2500
  };
  let response = Calendar.Events.list(calendarId, optionalArgs);
  let events = response.items;
  
  let eventsArray = [];
  
  if (events != undefined && events.length > 0) {
    for (let i = 0; i < events.length; i++) {
      let event = events[i];
      let attendees = event.attendees;
      let eventDetails = {};

      eventDetails['recurring'] = "No";
      if(event.recurringEventId != undefined)
        eventDetails['recurring'] = "Yes";  
      
      let when = event.start.dateTime;
      if (!when) {
        when = event.start.date;
      }

      eventDetails['name'] = event.creator.email;
      eventDetails['summary'] = event.summary;
      if(event.getDescription() == undefined)
        eventDetails['description'] = '';
      else
        eventDetails['description'] = event.getDescription();
      eventDetails['id'] = event.id;

      if (event.start.dateTime != undefined)
        eventDetails['date'] = new Date(event.start.dateTime);
      else
        eventDetails['date'] = new Date(event.start.date);
      if(event.start.dateTime != undefined)
        eventDetails['startDate'] = new Date(event.start.dateTime);
      else 
        eventDetails['startDate'] = new Date(event.start.date);
      if(event.end.dateTime != undefined)
        eventDetails['endDate'] = new Date(event.end.dateTime);
      else {
        eventDetails['endDate'] = new Date(event.end.date);
      }
      eventDetails['attendees'] = "";
      if (attendees != undefined) {
        for(let i = 0; i < attendees.length; i++) {
          //if(attendees[i].responseStatus != "declined")
          eventDetails['attendees'] += ";" + attendees[i].email;
        }
      }

      eventDetails['duration'] = ((eventDetails['endDate'] - eventDetails['startDate'])/(1000*60*60));

      eventsArray.push(eventDetails);
    }
    
    return eventsArray;
  } else {
    //Logger.log('No upcoming events found.');
    return eventsArray;
  }
}

function getSearchList(employeeSheet) {
  let employeeRange = employeeSheet.getRange(1, 1, employeeSheet.getLastRow() - 1, employeeSheet.getLastColumn());
  let employeeValues = employeeRange.getValues();
  let searchList = [];

  employeeValues.forEach(function(employee){
    searchList.push(employee[1]);
  });

  return searchList;
}

/* This function reads the sheet, removes all future events from the sheet and repopulates them. This is required to allow for modification or new events being accounted for */
function addEventsToSheet(calendarId, calendarName, deleteRows) {

  if(calendarId == undefined) {
    return [];
  }

  //Get all events for the given calendarId.
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadSheet.getSheetByName(SHEET_NAME);
  sheet.setFrozenRows(1);

  let employeeSheet = spreadSheet.getSheetByName('Employees');
  let searchList = getSearchList(employeeSheet);
  //console.log(searchList);

  /* Delete rows from the sheet where the date column has values greater than today - daysBefore) */
  let currentDate = new Date();

  /* Get the date which is today - DAYS_BEFORE. This will be used as the starting point for repopulating events */
  let daysBefore = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - DAYS_BEFORE);
  let breakCount = 0 ;
  let startInsertRow = 1;

  if(deleteRows){
    /* Sort the sheet by timestamp (field 2 in this case)*/
    sheet.sort(2)    
    let range = sheet.getDataRange();
    let rangeVals = range.getValues();

    /* this loop is to find out the rows till which deletion should happen */
    let startDeleteRow = rangeVals.length ;
    for(let i = rangeVals.length - 1; i > 0; i--){
      if(rangeVals[i][1] != undefined){
        /* Get the date column from the sheet - the second column in this case */
        let rowDate = new Date(rangeVals[i][1]);
        if(rowDate >= daysBefore) {
          /* Get the row number from which we need to repopulate events */
          startDeleteRow -= 1;
          breakCount = 0 ;
        } else {
          breakCount++;
        }

        /* 
          Stop counting rows if more than 10 rows are found to be having lesser timestamp than "daysBefore" timestamp.
          This is to prevent unwanted long runs of the code. Alternative solution is to move data from active sheet to archive sheet every few months.
        */
        if(breakCount > 10) {
          break;
        }
      }
      
    }
    
    if(startDeleteRow != NaN && rangeVals.length > 1) {
      console.log("Deleting from " + startDeleteRow + " : " + rangeVals.length); 
      let rangeToClear = sheet.getRange(startDeleteRow + 1, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      rangeToClear.clearContent();
      startInsertRow = startDeleteRow;
    }
  }

  let outputArray = [];
  let eventsArray = listUpcomingEvents(calendarId, daysBefore);
  if(eventsArray == undefined || eventsArray.length == 0){
    return;
  }
  
  eventsArray.forEach(function(eventDetails){
    let row = [];
    
    let patientIdsInSummary = eventDetails['summary'].match(/HNI[0-9]+/g);
    let patientIdsInDescription = eventDetails['description'].match(/HNI[0-9]+/g);
    let patientIds = [];
    
    if(patientIdsInSummary != null && patientIdsInSummary != undefined)
      patientIds = (patientIdsInDescription != null && patientIdsInDescription != undefined) ? patientIdsInSummary.concat(patientIdsInDescription) : patientIdsInSummary;
      
    if(patientIds != null && patientIds != undefined && patientIds.length != 0)
      patientIds = patientIds.join(',');
    else
      patientIds = "Not Applicable";
    
    eventDetails = businessLogic(eventDetails, searchList);
    //Add the calendar events to the sheet.
    row = 
      [ eventDetails['name'], eventDetails['date'], 
        eventDetails['summary'], eventDetails['duration'], calendarName, 
        eventDetails['attendees_description'], eventDetails['recurring'], patientIds,
        eventDetails['attended']
      ];
    outputArray.push(row);
  })
    
    /*sheet.appendRow([eventsArray[i]['name'], eventsArray[i]['date'], 
    eventsArray[i]['summary'], duration, calendarName, eventsArray[i]['attendees'], eventsArray[i]['recurring'], patientIds]);*/
  
  if(outputArray.length > 0 && outputArray[0].length > 0) {
    if(!deleteRows)
      startInsertRow = sheet.getLastRow() ;
      
    sheet.getRange(startInsertRow + 1, 1, outputArray.length, outputArray[0].length).setValues(outputArray);
    console.log('Adding from ' + (startInsertRow + 1)  + ' for ' + calendarName + ' row count: ' + outputArray.length);
  }
  
}

/* All transformations on event object should happen in this function */
function businessLogic(eventDetails, searchList) {
  eventDetails['attendees_description'] = "";
  searchList.forEach(function(name){
    let matchedName = eventDetails['description'].match("\\b" + name + "\\b");
    if(matchedName != null || matchedName != undefined){
      eventDetails['attendees_description'] += "," + matchedName ;
    }
  })

  eventDetails['attended'] = "Yes";
  if(eventDetails['summary'].match(/\bCAN\b/) != null ||
    eventDetails['summary'].match(/\bMIS\b/) != null)
    eventDetails['attended'] = "No";

  return eventDetails;
}

/*This function acts as a pseudo starting point*/
function calendarEvents() {

  /* Get all calendars one has access to */
  //var calendars = CalendarApp.getAllCalendars();

  let calendars = [
    {'id': '<>','name': 'Personal Calendar'},
    {'id': '<>', 'name': 'HNI Calendar'},
  ];
  
  
  /* 
    If you have multiple calendars, then cleanup of future events should be done only for the first calendar run. Cleanup is required to account for events that get modified or those that get newly added. 
    For subsequent calendars, we can ignore cleanup of rows 
  */
  var deleteRows = true;

  for(let i = 0; i < calendars.length; i++){
    if(i > 0)
      deleteRows = false;
    addEventsToSheet(calendars[i]['id'], calendars[i]['name'], deleteRows);
  }

  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadSheet.getSheetByName(SHEET_NAME);
  

  /* Freeze the first row of the spreadsheet since this is the header row */
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  /* Sort the sheet based on the calendar name and timestamp */
  let range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  range.sort([{column: 5, ascending: true}, {column: 2, ascending: true}]);

  /* Set the timestamp format for the event timestamp column */
  sheet.getRange(2, 2, sheet.getLastRow() - 1).setNumberFormat("dd/mm/yyyy hh:mm am/pm");

}
