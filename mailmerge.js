function addMailMergeDetails() {
  let today = new Date();
  let currentMonth = new Date(today.getFullYear(), today.getMonth() - 1, today.getDate()).toLocaleString('default', { month: 'long' });
  let currentYear = today.getFullYear();
  
  let emailList = {};
  let emailSheet = SpreadsheetApp.openById('<>').getSheetByName('Ref Emails');
  let emailRange = emailSheet.getRange(1, 1, emailSheet.getLastRow() - 1, emailSheet.getLastColumn());
  let emailValues = emailRange.getValues();


  emailValues.forEach(function(email){
    let key = email[0].trim().toLowerCase();

    if(key === "-" || key.length === 0){
      key = email[1].trim().toLowerCase() + "_" + email[2].trim().toLowerCase();
    }
    
    if(emailList[key] === undefined)
      emailList[key] = email[3];
  })

  let mailMergeSheet = SpreadsheetApp.openById('<>').getSheetByName('New Mail Merge');
  let mailMergeValues = mailMergeSheet.getRange(2, 1, mailMergeSheet.getLastRow() - 1, mailMergeSheet.getLastColumn()).getValues();

  let skipRow = {}
  let rowsToDelete = [];
  mailMergeValues.forEach(function(mailMergeEntry, index) {
    if(mailMergeEntry[10].match(/\bMAIL SENT\b/) === null){ //&& mailMergeEntry[4] === currentMonth){
      rowsToDelete.push(index + 2);
    } else {
      skipRow[mailMergeEntry[0]] = true;
    }
  })
  
  for(let i = rowsToDelete.length - 1; i >= 0; i--){
    console.log("Deleting row id: " + rowsToDelete[i]);
    mailMergeSheet.deleteRow(rowsToDelete[i]);
  }

  let sessionManagementSpreadSheet = SpreadsheetApp.openById('<>');
  let yearlySheet = sessionManagementSpreadSheet.getSheetByName('Yearly Dues');
  let yearlyRange = yearlySheet.getRange(2, 1, yearlySheet.getLastRow() - 1, yearlySheet.getLastColumn());

  let sessionManagementSheet = sessionManagementSpreadSheet.getSheetByName('Session Management');
  let sessionManagementRange = sessionManagementSheet.getRange(2, 1, sessionManagementSheet.getLastRow() - 1, sessionManagementSheet.getLastColumn());
  
  let currentMonthPatients = {};
  
  sessionManagementRange.getValues().forEach(function(sessionDetails){
    if(sessionDetails[0] === currentYear && sessionDetails[1] === currentMonth){
      let chargedSessions = 0;

      if(Number.isInteger(sessionDetails[14]))
        chargedSessions = sessionDetails[14];
      
      currentMonthPatients[sessionDetails[2]] = {
        'chargedSessions': chargedSessions
      };
    }
  })

  
  let outputArray = [];
  yearlyRange.getValues().forEach(function(yearlySheetEntry){
    let currentMonthSessions = 0;

    /*if(skipRow[yearlySheetEntry[0]] != undefined){
      console.log("Skipping: " + yearlySheetEntry[0]);
      return;
    }*/

    if(yearlySheetEntry[7] === 0)
      return;

    console.log("Proceeding with: " + yearlySheetEntry[0]);

    if(currentMonthPatients[yearlySheetEntry[0]] != undefined)
      currentMonthSessions = currentMonthPatients[yearlySheetEntry[0]]['chargedSessions'];


    let emailLookupKey = yearlySheetEntry[0].trim().toLowerCase();
    let emailId = "-";

    if(emailList[emailLookupKey] != undefined)
      emailId = emailList[emailLookupKey];
    else {
      emailLookupKey = yearlySheetEntry[1].trim().toLowerCase() + "_" + yearlySheetEntry[2].trim().toLowerCase();
      if(emailList[emailLookupKey] != undefined)
        emailId = emailList[emailLookupKey];
    }

    let currentRow = [
      yearlySheetEntry[0], yearlySheetEntry[1], yearlySheetEntry[2], 
      currentMonthSessions, currentMonth, yearlySheetEntry[7],
      yearlySheetEntry[yearlySheetEntry.length - 1], emailId
    ];
    outputArray.push(currentRow);
    
  })
/*
  let mailMergeRange = mailMergeSheet.getRange(2, 1, mailMergeSheet.getLastRow() - 1, mailMergeSheet.getLastColumn());
  mailMergeRange.clearContent();
  
  mailMergeSheet.getRange(2, 1, historicData.length, historicData[0].length).setValues(historicData);
  mailMergeSheet.getRange(2 + historicData.length, 1, outputArray.length, outputArray[0].length).setValues(outputArray);
*/

  
  outputArray.forEach(function(row){
    mailMergeSheet.appendRow(row);
  })
  

}

