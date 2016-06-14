/*
 * Find PNR status of Indian Railway ticket and mail if the status changed.
 * Developed by : Shuvankar Sarkar
 */

function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var pnr = "<Enter PNR Number>";
  var apikey = "<Enter API Key>"; //Find API Key from http://www.railwayapi.com/
  var url = "http://api.railwayapi.com/pnr_status/pnr/" + pnr + "/apikey/" + apikey + "/";
  var response = UrlFetchApp.fetch(url);
  var responseString = response.getContentText();
  var data = JSON.parse(responseString);
  
  //If response code is 200 then populate the values. 
  if (data.response_code == "200") {
    //PNR no
    sheet.getRange("B1").setValue(data.pnr);
    //From
    sheet.getRange("B2").setValue(data.from_station.name);
    //To
    sheet.getRange("B3").setValue(data.to_station.name);
    //Train Name
    sheet.getRange("B4").setValue(data.train_name);
    //Date Of Journey
    sheet.getRange("B5").setValue(data.doj);
    //Booking Status
    sheet.getRange("B6").setValue(data.passengers[0].booking_status);
    //Previous Status: Set value of Current Status
    sheet.getRange("B7").setValue(sheet.getRange("B8").getValue());
    //Current Status
    sheet.getRange("B8").setValue(data.passengers[0].current_status);
    //Last Updated
    sheet.getRange("B9").setValue(Date());
  }
  //Last Response Status 
  sheet.getRange("B10").setValue(data.response_code);
  //Last Run Date
  sheet.getRange("B11").setValue(Date());
  //If Current Status and Previous Status is not same then send Email
  if(sheet.getRange("B7").getValue() != sheet.getRange("B8").getValue()) {
    sendEmails();
  }
}

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 2;   // Number of rows to process
  var dataRange = sheet.getRange(startRow, 5, numRows, 1); //5 = E Column
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
    var emailAddress = row[0];
    var subject = "PNR:"+ sheet.getRange("B1").getValue()+" | Current Status: "+sheet.getRange("B8").getValue();
    var message = "Booking Status:" + sheet.getRange("B6").getValue() + " | \
Current Status:" + sheet.getRange("B8").getValue()+ " | \
Train Name:" + sheet.getRange("B4").getValue()+ " | \
From:" + sheet.getRange("B2").getValue()+ " | \
To:" + sheet.getRange("B3").getValue();
    //TODO: Fix formatting and add in body: Date of Journey:" + sheet.getRange("B5").getValue()+ " | \    

    MailApp.sendEmail(emailAddress, subject, message);
  }
}