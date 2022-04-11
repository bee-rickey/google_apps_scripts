const PAYMENT_SHEET_ID = '';
const SESSION_SHEET_ID = '';
const PAYMENT_SHEET_URL = 'https://docs.google.com/spreadsheets/d/' + PAYMENT_SHEET_ID + '/edit#gid=0&range=' ;

const START_MONTH = "January";
const CALENDAR_MONTH_ARRAY = [
  "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
];

function getPaymentDetails() {
  let paymentsSpreadsheet = SpreadsheetApp.openById(PAYMENT_SHEET_ID);
  let paymentsSheet = paymentsSpreadsheet.getSheetByName('FY2021-22');
  let paymentsRange = paymentsSheet.getRange(1, 1, paymentsSheet.getLastRow() - 1, paymentsSheet.getLastColumn());
  let paymentsValues = paymentsRange.getValues();//.filter(x => x[8].toLowerCase() == "therapy fee");
  let paymentDetails = {};

  paymentsValues.forEach(function(payment, index){
    try{
      let amount = 0;
      if(payment[5] != undefined && payment[5] != "")
        amount = payment[5]; 

      if(payment[7].trim().toLowerCase() != "therapy fee")
        return;
      //let name = payment[6].trim().toLowerCase().concat("_" + payment[7].trim().toLowerCase());
      let name = payment[6].trim().toLocaleLowerCase().replace(/ +/g, '_');
      let monthYear = payment[0].toLocaleString('default', {month: 'long'}) + "_" + payment[0].getFullYear();
      if(paymentDetails[name] == undefined) {
        paymentDetails[name] = {};
        paymentDetails[name]['cumulativePaid'] = 0;
      }
      paymentDetails[name]['cumulativePaid'] += amount;
      
      if(paymentDetails[name][monthYear] == undefined)
        paymentDetails[name][monthYear] = [];

      let transactionInfo = {};
      transactionInfo["amount"] = amount;
      transactionInfo["category"] = payment[7];
      transactionInfo["row"] = index + 1;
      
      paymentDetails[name][monthYear].push(transactionInfo);
      
    }catch(Error){
      console.log('ISSUE WITH:' + payment)
    }
  });

  return paymentDetails;
}
/*
function paymentPerMonth() {
  let sessionsSpreadsheet = SpreadsheetApp.openById(SESSION_SHEET_ID);
  let paymentsSpreadsheet = SpreadsheetApp.openById(PAYMENT_SHEET_ID);
  let sessionsSheet = sessionsSpreadsheet.getSheetByName('Session Management');
  let paymentsSheet = paymentsSpreadsheet.getSheetByName('FY2021-22');

  let sessionsRange = sessionsSheet.getRange(1, 1, sessionsSheet.getLastRow() - 1, sessionsSheet.getLastColumn());
  let sessionsValues = sessionsRange.getValues();
  let paymentsRange = paymentsSheet.getRange(1, 1, paymentsSheet.getLastRow() - 1, paymentsSheet.getLastColumn());
  let paymentsValues = paymentsRange.getValues(); //.filter(x => x[8].toLowerCase() == "therapy fee");

  let paymentDetails = {};

  paymentsValues.forEach(function(payment, index){
    try{
      if(payment[8].toLowerCase() != "therapy fee")
        return;
      let name = payment[6].trim().toLowerCase().concat("_" + payment[7].trim().toLowerCase());
      let monthYear = payment[0].toLocaleString('default', {month: 'long'}) + "_" + payment[0].getFullYear();

      if(paymentDetails[name] == undefined)
        paymentDetails[name] = {};
      
      if(paymentDetails[name][monthYear] == undefined)
        paymentDetails[name][monthYear] = [];

      let transactionInfo = {};
      transactionInfo["amount"] = payment[5];
      transactionInfo["category"] = payment[8];
      transactionInfo["row"] = index + 1;
      
      paymentDetails[name][monthYear].push(transactionInfo);
    }catch(Error){

    }
  });

  let today = new Date();
  let previousMonthYear = 
        new Date(today.getFullYear(), today.getMonth() - 2, today.getDate()).toLocaleString('default', { month: 'long' }).concat("_" + today.getFullYear());

  sessionsValues.forEach(function(sessionDetail, index){
    try{
      if(index == 0){
        return;
      }
      let name = sessionDetail[2].trim().toLowerCase().concat("_", sessionDetail[3].trim().toLowerCase());


      if(paymentDetails[name] != undefined){
        if(paymentDetails[name][previousMonthYear] != undefined && 
        previousMonthYear.split('_')[0].trim().toLowerCase() == sessionDetail[0].trim().toLowerCase()){
          let expectedPayment = sessionDetail[15];
          let paidAmount = 0 ;
          let paid = false;

          paymentDetails[name][previousMonthYear].forEach(function(item) {
            paidAmount += Number(item['amount']);
          });

          //sessionsSheet.getRange(index + 1, 19).setValue(paidAmount);
          if(Number(expectedPayment) <= paidAmount){
            paid = true;
          }
          sessionsSheet.getRange(index + 1, 19).setValue(paidAmount);
          sessionsSheet.getRange(index + 1, 21).setValue('Paid: ' + paidAmount + ' Actual: ' + expectedPayment);
          
          console.log(paymentDetails[name][previousMonthYear])
          sessionsSheet.getRange(index + 1, 22).setValue(PAYMENT_SHEET_URL + 'C' + paymentDetails[name][previousMonthYear][0]['row']);
  

          if(paid){
            sessionsSheet.getRange(index + 1, 18).setBackground('green');
            sessionsSheet.getRange(index + 1, 18).setValue('Paid');
          }else{
            sessionsSheet.getRange(index + 1, 18).setBackground('yellow');
            sessionsSheet.getRange(index + 1, 18).setValue('Not Paid');
          }

        }
      }
      
    }catch(Error){
      console.log(sessionDetail)
    }
    
  });

  sessionsSheet.autoResizeColumn(22);
}

function calendarCumulative() {
  let sessionsSpreadsheet = SpreadsheetApp.openById(SESSION_SHEET_ID);
  let sessionsSheet = sessionsSpreadsheet.getSheetByName('Session Management');

  let sessionsRange = sessionsSheet.getRange(1, 1, sessionsSheet.getLastRow() - 1, sessionsSheet.getLastColumn());
  let sessionsValues = sessionsRange.getValues();

  let paymentDetails = getPaymentDetails();

  let today = new Date();
  let previousMonth = 
        new Date(today.getFullYear(), today.getMonth() - 2, today.getDate()).toLocaleString('default', { month: 'long' });

  let indexOfPreviousMonth = MONTH_ARRAY.indexOf(previousMonth);
  
  let patientSessionDetails = {};

  sessionsValues.forEach(function(sessionDetails, index){
    let month = sessionDetails[0];
    if(MONTH_ARRAY.indexOf(month) > indexOfPreviousMonth)
      return;

    let name = sessionDetails[2].trim().toLowerCase().concat("_", sessionDetails[3].trim().toLowerCase());
    if(patientSessionDetails[name] === undefined){
      patientSessionDetails[name] = {};
      patientSessionDetails[name]['cumulativeDues'] = 0;
    }
    
    let due = 0;
    if(sessionDetails[15] != undefined && sessionDetails[15].length != 0)
      due = sessionDetails[15];
    patientSessionDetails[name]['cumulativeDues'] += due;

    let transferredAmount = "No data found";
    if(paymentDetails[name] != undefined)
      transferredAmount = paymentDetails[name]['cumulativePaid'];

    if(month === previousMonth){
      console.log('month: ' + month + ' name: ' + name + ' Due: ' + patientSessionDetails[name]['cumulativeDues'] + ' Paid: ' + transferredAmount);
      sessionsSheet.getRange(index + 1, 22).setValue('Due: ' + patientSessionDetails[name]['cumulativeDues'] + ' Paid: ' + transferredAmount);
    }
    
  })
}
*/

function getSessionData(sheetId){
  let sessionsSpreadsheet = SpreadsheetApp.openById(sheetId);
  let sessionsSheet = sessionsSpreadsheet.getSheetByName('Session Management');
  
  let sessionsRange = sessionsSheet.getRange(2, 1, sessionsSheet.getLastRow() - 1, sessionsSheet.getLastColumn());
  let sessionValues = sessionsRange.getValues().filter(x => (x[18] === 'Not Paid' || x[18] === ""));

  let terminationsSheet = sessionsSpreadsheet.getSheetByName('Terminations');
  let terminationsValues = terminationsSheet.getRange(2, 1, terminationsSheet.getLastRow() - 1, terminationsSheet.getLastColumn()).getValues();  

  let terminatedPatients = {};
  terminationsValues.forEach(function(terminationRow){
    if(terminationRow[1] === undefined || terminationRow[1].trim().length === 0){
      console.log(terminationRow[2] + " " + terminationRow[3] + " does not have HNI patient ID mapped in terminations sheet");
    }
    terminatedPatients[terminationRow[1].trim()] = terminationRow[8];
  })

  let years = new Set();
  let patientsMap = {};
  sessionValues.forEach(function(sessionDetails){
    let patientId = sessionDetails[2].trim();
    let month = sessionDetails[1];
    let year = sessionDetails[0];

    if(patientsMap[patientId] === undefined){
      patientsMap[patientId] = {};
      patientsMap[patientId]['fn'] = sessionDetails[3];
      patientsMap[patientId]['ln'] = sessionDetails[4];
      patientsMap[patientId]['therapist'] = sessionDetails[11];
      patientsMap[patientId]['chargedSessions'] = sessionDetails[14];
      patientsMap[patientId]['totalDues'] = 0;
      patientsMap[patientId]['monthEnded'] = "";
      patientsMap[patientId]['status'] = "Current";
      if(terminatedPatients[patientId] != undefined) {
        patientsMap[patientId]['status'] = "Past";
        patientsMap[patientId]['monthEnded'] = terminatedPatients[patientId];
      }
    } else {
      patientsMap[patientId]['therapist'] = sessionDetails[11];
    }

    if(patientsMap[patientId][year] === undefined)
      patientsMap[patientId][year] = {}

    if(patientsMap[patientId][year][month] === undefined){
      patientsMap[patientId][year][month] = {};
      patientsMap[patientId][year][month]['due'] = 0;
      patientsMap[patientId][year][month]['details'] = [];
    }

    let due = 0;
    
    if(sessionDetails[16] != undefined && sessionDetails[16].length != 0)
      due = sessionDetails[16];

    let comment = "";
    
    /*
    if(sessionDetails[13] === 1 && sessionDetails[12].trim().toLowerCase() === "membership fee"){
      due = 700;
      comment = "Single session, membership fee. Defaulting to Rs 700";
    }
    */

    if(sessionDetails[14] === 0){
      due = 0;
      comment = "Charged sessions = " + sessionDetails[14] + ". Defaulting to Rs 0";
    }

    if(patientsMap[patientId][year]['services'] === undefined)
      patientsMap[patientId][year]['services'] = {}

    if(patientsMap[patientId][year]['services'][sessionDetails[5]] === undefined){
      patientsMap[patientId][year]['services'][sessionDetails[5]] = "\n\t" + month + " : " + due;
    } else {
      patientsMap[patientId][year]['services'][sessionDetails[5]] += "\n\t" + month + " : " + due;
    }

    let rowDetails = {
      'due': due,
      'month': sessionDetails[1],
      'year': sessionDetails[0],
      'service': sessionDetails[5],
      'comment': comment
    }
    if(sessionDetails[0] != undefined && sessionDetails[0].length != 0 && !isNaN(parseInt(sessionDetails[0])))
      years.add(sessionDetails[0]);
    patientsMap[patientId][year][month]['due'] += due;
    patientsMap[patientId]['totalDues'] += due;
    patientsMap[patientId][year][month]['details'].push(rowDetails);
    patientsMap[patientId][year][month]['comment'] = comment;
  })
}

function yearlyDues(){
  let sessionsSpreadsheet = SpreadsheetApp.openById(SESSION_SHEET_ID);
  let sessionsSheet = sessionsSpreadsheet.getSheetByName('Session Management');
  
  let yearlySpreadsheet = SpreadsheetApp.openById('');
  let yearlySheet = yearlySpreadsheet.getSheetByName('Yearly Dues');
  let sessionsRange = sessionsSheet.getRange(2, 1, sessionsSheet.getLastRow() - 1, sessionsSheet.getLastColumn());
  let sessionValues = sessionsRange.getValues().filter(x => (x[18] === 'Not Paid' || x[18] === ""));

  let terminationsSheet = sessionsSpreadsheet.getSheetByName('Terminations');
  let terminationsValues = terminationsSheet.getRange(2, 1, terminationsSheet.getLastRow() - 1, terminationsSheet.getLastColumn()).getValues();  

  let terminatedPatients = {};
  terminationsValues.forEach(function(terminationRow){
    if(terminationRow[1] === undefined || terminationRow[1].trim().length === 0){
      console.log(terminationRow[2] + " " + terminationRow[3] + " does not have HNI patient ID mapped in terminations sheet");
    }
    terminatedPatients[terminationRow[1].trim()] = terminationRow[8];
  })

  let years = new Set();
  let patientsMap = {};
  sessionValues.forEach(function(sessionDetails){
    let patientId = sessionDetails[2].trim();
    let month = sessionDetails[1];
    let year = sessionDetails[0];

    if(patientsMap[patientId] === undefined){
      patientsMap[patientId] = {};
      patientsMap[patientId]['fn'] = sessionDetails[3];
      patientsMap[patientId]['ln'] = sessionDetails[4];
      patientsMap[patientId]['therapist'] = sessionDetails[11];
      patientsMap[patientId]['chargedSessions'] = sessionDetails[14];
      patientsMap[patientId]['totalDues'] = 0;
      patientsMap[patientId]['monthEnded'] = "";
      patientsMap[patientId]['status'] = "Current";
      if(terminatedPatients[patientId] != undefined) {
        patientsMap[patientId]['status'] = "Past";
        patientsMap[patientId]['monthEnded'] = terminatedPatients[patientId];
      }
    } else {
      patientsMap[patientId]['therapist'] = sessionDetails[11];
    }

    if(patientsMap[patientId][year] === undefined)
      patientsMap[patientId][year] = {}

    if(patientsMap[patientId][year][month] === undefined){
      patientsMap[patientId][year][month] = {};
      patientsMap[patientId][year][month]['due'] = 0;
      patientsMap[patientId][year][month]['details'] = [];
    }

    let due = 0;
    
    if(sessionDetails[16] != undefined && sessionDetails[16].length != 0)
      due = sessionDetails[16];

    let comment = "";
    
    /*
    if(sessionDetails[13] === 1 && sessionDetails[12].trim().toLowerCase() === "membership fee"){
      due = 700;
      comment = "Single session, membership fee. Defaulting to Rs 700";
    }
    */

    if(sessionDetails[14] === 0){
      due = 0;
      comment = "Charged sessions = " + sessionDetails[14] + ". Defaulting to Rs 0";
    }

    if(patientsMap[patientId][year]['services'] === undefined)
      patientsMap[patientId][year]['services'] = {}

    if(patientsMap[patientId][year]['services'][sessionDetails[5]] === undefined){
      patientsMap[patientId][year]['services'][sessionDetails[5]] = "\n\t" + month + " : " + due;
    } else {
      patientsMap[patientId][year]['services'][sessionDetails[5]] += "\n\t" + month + " : " + due;
    }

    let rowDetails = {
      'due': due,
      'month': sessionDetails[1],
      'year': sessionDetails[0],
      'service': sessionDetails[5],
      'comment': comment
    }
    if(sessionDetails[0] != undefined && sessionDetails[0].length != 0 && !isNaN(parseInt(sessionDetails[0])))
      years.add(sessionDetails[0]);
    patientsMap[patientId][year][month]['due'] += due;
    patientsMap[patientId]['totalDues'] += due;
    patientsMap[patientId][year][month]['details'].push(rowDetails);
    patientsMap[patientId][year][month]['comment'] = comment;
  })

  
  //Start reading yearly sheet
  let yearlyDuesValues = yearlySheet.getRange(1, 1, yearlySheet.getLastRow(), yearlySheet.getLastColumn()).getValues();
  let notes = yearlySheet.getRange(1, 1, yearlySheet.getLastRow(), yearlySheet.getLastColumn()).getNotes();
  let existingYearlyDuesRows = {};
  yearlyDuesValues.forEach(function(yearlyDuesRow){

    if(yearlyDuesRow[0] === undefined){
      console.log("Skipping " + yearlyDuesRow);
      return;
    }

    existingYearlyDuesRows[yearlyDuesRow[0].trim()] = [yearlyDuesRow[3]];

    /*
    let extraColumns = yearlyDuesRow.slice(23, yearlyDuesRow.length);
    
    if(extraColumns.length < 10){
      extraColumns = extraColumns.concat(Array(10 - extraColumns.length).fill(''));
    }

    existingYearlyDuesRows[yearlyDuesRow[0].trim()] = existingYearlyDuesRows[yearlyDuesRow[0].trim()].concat(extraColumns);
    */
  })
  
  //let today = new Date();
  //let currentMonth = new Date(today.getFullYear(), today.getMonth() - 1, today.getDate()).toLocaleString('default', { month: 'long' });
  let commentsRange = [];
  //let rowId = 1 ;
  let outputArray = [];

  for(var patientId in patientsMap){
    if(patientId === undefined || patientId === "")
      continue;

    let phoneNumber = "";
    if(existingYearlyDuesRows[patientId] != undefined)
      phoneNumber = existingYearlyDuesRows[patientId][0];

    let rowValues = [patientId, patientsMap[patientId]['fn'], patientsMap[patientId]['ln'], 
                    phoneNumber, patientsMap[patientId]['therapist'], 
                    patientsMap[patientId]['status'], patientsMap[patientId]['monthEnded'], patientsMap[patientId]['totalDues']];

    let services = "";
    //let name = patientsMap[patientId]['fn'].trim().toLowerCase() + "_" + patientsMap[patientId]['ln'].trim().toLowerCase();

    //let currentMonthDues = 0;
    let monthArray = [];

    years.forEach(function(year){
      CALENDAR_MONTH_ARRAY.forEach(function(month){
        monthArray.push(year + " " + month)
      })
    })
    

    monthArray.forEach(function(yearMonth){
      let yearName = yearMonth.split(' ')[0];
      let monthName = yearMonth.split(' ')[1];

      if(patientsMap[patientId][yearName] === undefined || 
        patientsMap[patientId][yearName][monthName] === undefined){
        rowValues.push(' - ');
      } else {
        rowValues.push(patientsMap[patientId][yearName][monthName]['due']);
      }
    })
    
    /*
    CALENDAR_MONTH_ARRAY.forEach(function(monthName, index){
      if(patientsMap[patientId][monthName] === undefined){
        rowValues.push(' - ');
      } else {
        if(currentMonth === monthName){
          rowValues.push(patientsMap[patientId][monthName]['due']);
          currentMonthDues = patientsMap[patientId][monthName]['due'];
        }
        else {
          rowValues.push(patientsMap[patientId][monthName]['due']);
        }

        if(patientsMap[patientId][monthName]['comment'].trim() != '')
          commentsRange.push([rowId, index, patientsMap[patientId][monthName]['comment'], name]);     
      }
    })
    */
    
    years.forEach(function(year){
        
        if(patientsMap[patientId][year] === undefined)
          return
        
        if(services.length === 0)
          services = year + ":";
        else
          services += "\n" + year + ":";
        
        for (var service in patientsMap[patientId][year]['services']) {
          services += "\n" + service + patientsMap[patientId][year]['services'][service] ;
        }
      
    })
    

    /*
    for (var service in patientsMap[patientId]['services']) {
      if (services.length == 0){
        services = service + patientsMap[patientId]['services'][service] ;
      } else {
        services += "\n" + service + patientsMap[patientId]['services'][service] ;
      }
    }
    */
    //rowValues = rowValues.concat([patientsMap[patientId]['totalDues'], services]);
    rowValues.push(services);
    
    /*
    if(existingYearlyDuesRows[patientId] != undefined && existingYearlyDuesRows[patientId].length > 1)
      rowValues = rowValues.concat(existingYearlyDuesRows[patientId].splice(1,existingYearlyDuesRows[patientId].length));
    else 
      rowValues = rowValues.concat(Array(10).fill(''));
    */

    outputArray.push(rowValues);
    //rowId++;
  }

  console.log(outputArray[0]);
  
  let maxRows = yearlySheet.getMaxRows();
  let startRow = 2;
  if(maxRows === 1)
    startRow = 1;

  let range = yearlySheet.getRange(startRow, 1, yearlySheet.getLastRow(), yearlySheet.getLastColumn());
  let formulas = range.getFormulas();

  if(range.getNumRows() > 1)
    range.clearContent();
  range.clearNote();

  yearlySheet.getRange(yearlySheet.getLastRow()+1, 1, outputArray.length, outputArray[0].length).setValues(outputArray);

  /*
  let formulasBackup = [];
  for(let row in formulas)
    for(let column in formulas[row])
      if(formulas[row][column] != ''){
        //range = yearlySheet.getRange(row + 1, column);
        //range.setFormula(formulas[row][column]);
        formulasBackup.push([row, column, formulas[row][column]]);   
      }
  
  commentsRange.forEach(function(comment){
    range=yearlySheet.getRange(comment[0] + 1, comment[1] + 1 + 7).setNote(comment[2]);
  }) 
*/
  range = yearlySheet.getRange(2, 1, yearlySheet.getLastRow() - 1, yearlySheet.getLastColumn());
  range.setHorizontalAlignment("left");
  range.sort([{column: 5, ascending: true}]);
  //yearlySheet.setRowHeights(2, yearlySheet.getMaxRows()-1, 50);
  /*
  formulasBackup.forEach(function(formula){
    let row = parseInt(formula[0]) + 2;
    let column = parseInt(formula[1]) + 1;
    range = yearlySheet.getRange(row, column);
    range.setFormula(formula[2]);
    console.log('Setting formula: ' + formula[2] + ' r: ' + row + ' c ' + column);
  })  

  //yearlySheet.setFrozenColumns(23);
  //yearlySheet.setFrozenro(23);
  */
}

function doGet() {
  return HtmlService.createTemplateFromFile("chart").evaluate();
}

function getValues(){
  var chartSheet = SpreadsheetApp.openById('')
      .getSheetByName('Yearly Dues')

  var chartSheetValues = chartSheet.getRange(1, 1, chartSheet.getLastRow(), chartSheet.getLastColumn())
      .getDisplayValues();

  var data = {};
  var lineChartData = [];
  var today = new Date() ;
  var currentMonth = ("0" + (today.getMonth() + 1)).slice(-2) + "-" + today.getFullYear();
  var currentMonthIndex = 0;

  let monthlyData = {};

  chartSheetValues.forEach(function(entry, index){
    if(index === 0){
      entry.forEach(function(header, headerIndex){
        if(headerIndex < 8 || headerIndex === entry.length - 1)
          return;
        monthlyData[headerIndex] = {"sum": 0, "name": header};
      })
    } else {
      entry.forEach(function(row, columnIndex){
        if(columnIndex < 8 || columnIndex === entry.length - 1)
          return;

        let amount = 0 ;
        if(!isNaN(parseInt(row)))
          amount = parseInt(row);
        
        monthlyData[columnIndex]['sum'] += amount;
      })
      
    }
  });

  let totalDues = 0;
  lineChartData[0] = ["Month", "Dues", {role: 'style'}];
  for(key in monthlyData){
    if(monthlyData[key]['sum'] === 0)
      continue;
    lineChartData.push([monthlyData[key]['name'], monthlyData[key]['sum'], 'color: #76A7FA']);
    totalDues += monthlyData[key]['sum'];
  }
  



  /*
  var pieChartData = [['Category', 'Expenses']];

  chartSheetValues[0].forEach(function(header, headerIndex){
    let row = [];
    if(headerIndex === 0)
      return;
    row.push(header);
    row.push(parseInt(chartSheetValues[currentMonthIndex][headerIndex]));

    pieChartData.push(row);
  })
  */
  data['lineChartData'] = lineChartData;
  data['totalDues'] = totalDues;
  //data['pieChartData'] = pieChartData;
  return data;
}

