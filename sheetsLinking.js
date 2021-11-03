

function myFunction() {
  let sessionsSpreadsheet = SpreadsheetApp.openById('<sheetId>');
  let paymentsSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sessionsSheet = sessionsSpreadsheet.getSheetByName('Session Management');
  let paymentsSheet = paymentsSpreadsheet.getSheetByName('FY2021-22');

  let sessionsRange = sessionsSheet.getRange(1, 1, sessionsSheet.getLastRow() - 1, sessionsSheet.getLastColumn());
  let sessionsValues = sessionsRange.getValues();
  let paymentsRange = paymentsSheet.getRange(1, 1, paymentsSheet.getLastRow() - 1, paymentsSheet.getLastColumn());
  let paymentsValues = paymentsRange.getValues().filter(x => x[8].toLowerCase() == "therapy fee");

  let paymentDetails = {};

  paymentsValues.forEach(function(payment){
    try{

      let name = payment[6].trim().toLowerCase().concat("_" + payment[7].trim().toLowerCase());
      let monthYear = payment[0].toLocaleString('default', {month: 'long'}) + "_" + payment[0].getFullYear();

      if(paymentDetails[name] == undefined)
        paymentDetails[name] = {};
      
      if(paymentDetails[name][monthYear] == undefined)
        paymentDetails[name][monthYear] = [];

      let transactionInfo = {};
      transactionInfo["amount"] = payment[5];
      transactionInfo["category"] = payment[8];

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
      let paid = false;
      let paidAmount = 0 ;
      let expectedPayment = 0;

      if(paymentDetails[name] != undefined){
        if(paymentDetails[name][previousMonthYear] != undefined && 
        previousMonthYear.split('_')[0].trim().toLowerCase() == sessionDetail[0].trim().toLowerCase()){
          expectedPayment = sessionDetail[15];
          paidAmount = 0 ;
          
          paymentDetails[name][previousMonthYear].forEach(function(item) {
            paidAmount += Number(item['amount']);
          });

          //sessionsSheet.getRange(index + 1, 19).setValue(paidAmount);
          if(Number(expectedPayment) <= paidAmount){
            paid = true;
          }
        }

        sessionsSheet.getRange(index + 1, 19).setValue(paidAmount);
        sessionsSheet.getRange(index + 1, 21).setValue('Paid: ' + paidAmount + ' Actual: ' + expectedPayment);

        if(paid){
          sessionsSheet.getRange(index + 1, 18).setBackground('green');
          sessionsSheet.getRange(index + 1, 18).setValue('Paid');
        }else{
          sessionsSheet.getRange(index + 1, 18).setBackground('yellow');
          sessionsSheet.getRange(index + 1, 21).setValue('Paid: ' + paidAmount + ' Actual: ' + expectedPayment);
        }

      }
      
    }catch(Error){
      console.log(sessionDetail)
    }
    
  });
}
