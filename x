diff --git a/calendarPull.js b/calendarPull.js
index 6899fad..619789a 100644
--- a/calendarPull.js
+++ b/calendarPull.js
@@ -1,25 +1,22 @@
 /*
   Step 1: Get list of all calendars one has access to. 
-  Step 2: For the first calendar, go to the SHEET_NAME sheet. Sort by date column and delete all events that are greater than today - DAYS_BEFORE.
-    This will allow for any future events to be updated in the SHEET_NAME sheet. 
-    DAYS_BEFORE is a small buffer given to accomodate for any events older than two days being updated retrospectively.
+  Step 2: For the first calendar, go to the SHEET_NAME sheet. Sort by date column and delete all events that are greater than today - DAYS_BEFORE. This will allow for any future events to be updated in the SHEET_NAME sheet. DAYS_BEFORE is a small buffer given to accomodate for any events older than two days being updated retrospectively.
   Step 3: Pull all events starting from today - DAYS_BEFORE to today + NUMBER_OF_MONTHS for the current calendar.
   Step 4: For each of the event, get the required values and populate in SHEET_NAME sheet.
   Step 5: Repeate Step 3 and Step 4 for the rest of the calendars one has access to.
 
-  Additional optimisation: In step 2, while deleting, if you hit 10 consecutive events that are having a timestamp lesser than today - DAYS_BEFORE, 
-  then stop deleting rows and move to adding new events (this is to avoid unwanted deletion of whole sheet). 
+  Additional optimisation: In step 2, while deleting, if you hit 10 consecutive events that are having a timestamp lesser than today - DAYS_BEFORE, then stop deleting rows and move to adding new events (this is to avoid unwanted deletion of whole sheet). 
 
 */
 
 /* Sheet into which the calendar events will be added. Make sure this sheet name matches the sheet where you want entries to be made */
-const SHEET_NAME = 'Test Calendar Sheet';
+const SHEET_NAME = 'HNI Calendar';
 
 /* The start date from which calendar pull will be done. Any event before this will be ignored */
-const START_DATE = '2021-10-01 00:00:00';
+const START_DATE = '2021-04-01 00:00:00';
 
 /* Number of months of calendar events to be pulled from the current date */
-const NUMBER_OF_MONTHS = 3;
+const NUMBER_OF_MONTHS = 2;
 
 /* 
   This is used to decide which all rows of the sheet need to be cleared during each run.
@@ -31,8 +28,9 @@ const DAYS_BEFORE = 2;
 function listUpcomingEvents(calendarId, dateToFetchFrom) {
  
   if (calendarId == undefined) {
+    calendarId = "26v87sgot9k66p6l6ooti9i1mo@group.calendar.google.com";
     console.log("Error: Calendar Id not passed");
-    return [];
+    //return [];
   } 
 
   /* If no starting date is given, then the current date is assumed to be the starting date */
@@ -54,20 +52,30 @@ function listUpcomingEvents(calendarId, dateToFetchFrom) {
   };
   let response = Calendar.Events.list(calendarId, optionalArgs);
   let events = response.items;
+  
   let eventsArray = [];
   
   if (events != undefined && events.length > 0) {
     for (let i = 0; i < events.length; i++) {
       let event = events[i];
       let attendees = event.attendees;
-      let eventDetails = {}
+      let eventDetails = {};
+
+      eventDetails['recurring'] = "No";
+      if(event.recurringEventId != undefined)
+        eventDetails['recurring'] = "Yes";  
       
       let when = event.start.dateTime;
       if (!when) {
         when = event.start.date;
       }
+
       eventDetails['name'] = event.creator.email;
       eventDetails['summary'] = event.summary;
+      if(event.getDescription() == undefined)
+        eventDetails['description'] = '';
+      else
+        eventDetails['description'] = event.getDescription();
       eventDetails['id'] = event.id;
 
       if (event.start.dateTime != undefined)
@@ -86,12 +94,13 @@ function listUpcomingEvents(calendarId, dateToFetchFrom) {
       eventDetails['attendees'] = "";
       if (attendees != undefined) {
         for(let i = 0; i < attendees.length; i++) {
-          if(attendees[i].responseStatus != "declined")
-            eventDetails['attendees'] += ";" + attendees[i].email;
+          //if(attendees[i].responseStatus != "declined")
+          eventDetails['attendees'] += ";" + attendees[i].email;
         }
       }
 
       eventDetails['duration'] = ((eventDetails['endDate'] - eventDetails['startDate'])/(1000*60*60));
+
       eventsArray.push(eventDetails);
     }
     
@@ -102,6 +111,18 @@ function listUpcomingEvents(calendarId, dateToFetchFrom) {
   }
 }
 
+function getSearchList(employeeSheet) {
+  let employeeRange = employeeSheet.getRange(1, 1, employeeSheet.getLastRow() - 1, employeeSheet.getLastColumn());
+  let employeeValues = employeeRange.getValues();
+  let searchList = [];
+
+  employeeValues.forEach(function(employee){
+    searchList.push(employee[1]);
+  });
+
+  return searchList;
+}
+
 /* This function reads the sheet, removes all future events from the sheet and repopulates them. This is required to allow for modification or new events being accounted for */
 function addEventsToSheet(calendarId, calendarName, deleteRows) {
 
@@ -113,7 +134,10 @@ function addEventsToSheet(calendarId, calendarName, deleteRows) {
   let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
   let sheet = spreadSheet.getSheetByName(SHEET_NAME);
   sheet.setFrozenRows(1);
-  
+
+  let employeeSheet = spreadSheet.getSheetByName('Employees');
+  let searchList = getSearchList(employeeSheet);
+  //console.log(searchList);
 
   /* Delete rows from the sheet where the date column has values greater than today - daysBefore) */
   let currentDate = new Date();
@@ -121,6 +145,7 @@ function addEventsToSheet(calendarId, calendarName, deleteRows) {
   /* Get the date which is today - DAYS_BEFORE. This will be used as the starting point for repopulating events */
   let daysBefore = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - DAYS_BEFORE);
   let breakCount = 0 ;
+  let startInsertRow = 1;
 
   if(deleteRows){
     /* Sort the sheet by timestamp (field 2 in this case)*/
@@ -157,30 +182,84 @@ function addEventsToSheet(calendarId, calendarName, deleteRows) {
       console.log("Deleting from " + startDeleteRow + " : " + rangeVals.length); 
       let rangeToClear = sheet.getRange(startDeleteRow + 1, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
       rangeToClear.clearContent();
+      startInsertRow = startDeleteRow;
     }
   }
 
+  let outputArray = [];
   let eventsArray = listUpcomingEvents(calendarId, daysBefore);
-  for (let i=0; i < eventsArray.length; i++) {
-
-    var duration = eventsArray[i]['duration'];
-    businessLogic(eventsArray[i]);
+  if(eventsArray == undefined || eventsArray.length == 0){
+    return;
+  }
+  
+  eventsArray.forEach(function(eventDetails){
+    let row = [];
+    
+    let patientIdsInSummary = eventDetails['summary'].match(/HNI[0-9]+/g);
+    let patientIdsInDescription = eventDetails['description'].match(/HNI[0-9]+/g);
+    let patientIds = [];
+    
+    if(patientIdsInSummary != null && patientIdsInSummary != undefined)
+      patientIds = (patientIdsInDescription != null && patientIdsInDescription != undefined) ? patientIdsInSummary.concat(patientIdsInDescription) : patientIdsInSummary;
+      
+    if(patientIds != null && patientIds != undefined && patientIds.length != 0)
+      patientIds = patientIds.join(',');
+    else
+      patientIds = "Not Applicable";
+    
+    eventDetails = businessLogic(eventDetails, searchList);
     //Add the calendar events to the sheet.
-    sheet.appendRow([eventsArray[i]['name'], eventsArray[i]['date'], eventsArray[i]['summary'], duration, calendarName, eventsArray[i]['attendees']]);
+    row = 
+      [ eventDetails['name'], eventDetails['date'], 
+        eventDetails['summary'], eventDetails['duration'], calendarName, 
+        eventDetails['attendees_description'], eventDetails['recurring'], patientIds,
+        eventDetails['attended']
+      ];
+    outputArray.push(row);
+  })
+    
+    /*sheet.appendRow([eventsArray[i]['name'], eventsArray[i]['date'], 
+    eventsArray[i]['summary'], duration, calendarName, eventsArray[i]['attendees'], eventsArray[i]['recurring'], patientIds]);*/
+  
+  if(outputArray.length > 0 && outputArray[0].length > 0) {
+    if(!deleteRows)
+      startInsertRow = sheet.getLastRow() ;
+      
+    sheet.getRange(startInsertRow + 1, 1, outputArray.length, outputArray[0].length).setValues(outputArray);
+    console.log('Adding from ' + (startInsertRow + 1)  + ' for ' + calendarName + ' row count: ' + outputArray.length);
   }
   
 }
 
 /* All transformations on event object should happen in this function */
-function businessLogic(eventDetails) {
-  return true;
+function businessLogic(eventDetails, searchList) {
+  eventDetails['attendees_description'] = "";
+  searchList.forEach(function(name){
+    let matchedName = eventDetails['description'].match("\\b" + name + "\\b");
+    if(matchedName != null || matchedName != undefined){
+      eventDetails['attendees_description'] += "," + matchedName ;
+    }
+  })
+
+  eventDetails['attended'] = "Yes";
+  if(eventDetails['summary'].match(/\bCAN\b/) != null ||
+    eventDetails['summary'].match(/\bMIS\b/) != null)
+    eventDetails['attended'] = "No";
+
+  return eventDetails;
 }
 
 /*This function acts as a pseudo starting point*/
 function calendarEvents() {
 
   /* Get all calendars one has access to */
-  var calendars = CalendarApp.getAllCalendars();
+  //var calendars = CalendarApp.getAllCalendars();
+
+  let calendars = [
+    {'id': 'bharath.k.hegde@gmail.com','name': 'Personal Calendar'},
+    {'id': '26v87sgot9k66p6l6ooti9i1mo@group.calendar.google.com', 'name': 'HNI Calendar'},
+  ];
+  
   
   /* 
     If you have multiple calendars, then cleanup of future events should be done only for the first calendar run. Cleanup is required to account for events that get modified or those that get newly added. 
@@ -191,12 +270,13 @@ function calendarEvents() {
   for(let i = 0; i < calendars.length; i++){
     if(i > 0)
       deleteRows = false;
-    addEventsToSheet(calendars[i].getId(), calendars[i].getName(), deleteRows);
+    addEventsToSheet(calendars[i]['id'], calendars[i]['name'], deleteRows);
   }
 
   let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
   let sheet = spreadSheet.getSheetByName(SHEET_NAME);
   
+
   /* Freeze the first row of the spreadsheet since this is the header row */
   sheet.setFrozenRows(1);
   SpreadsheetApp.flush();
