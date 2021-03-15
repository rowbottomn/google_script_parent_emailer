/**
 * STEAM Parent Contacter Script
 * Nathan Rowbottom March 7, 2021
 * Deployed with a google sheet to email both 
 * generic and individual comments to
 * parents and possibly students
 * 
 * TODO
 * XSend Emails out based on sheet
 * Properly send out message composition
 */

/**
 * Nathan Rowbottom March 7, 2021
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet for SNP STEAM Academy Functions.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  
  var menuItems = [
    {name: 'Email Contacting', functionName: 'sendEmails'},
    {name: 'WIP-Populate Contacts', functionName: 'addContacts'},
    {name: 'WIP-Attendance Emailer', functionName: 'processAttendance'},
  ];
  spreadsheet.addMenu('SNP STEAM Academy', menuItems);
}

function addContacts(){
  var ui = SpreadsheetApp.getUi();
  ui.alert("Sorry! This feature is not ready yet!");
}


/**
 * This is the main function that does all the work.
 * Nathan Rowbottom March 7, 2021
 */

function processAttendance() {

  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var startRow = 2; // First row of data to process
 // var numRows = sheet.getLastRow(); // Number of rows to process, set to 100 to essentially do all. Inside the loop, it breaks at first empty name
  // Fetch the range of cells
  var dataRange = sheet.getRange(startRow, 1, sheet.getLastRow(), 20);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[0] === ''){
      console.log("sent emails for "+i+ " students.");
      ui.alert("sent emails for "+i+ " students.");
      return;
    }
    else if (row[19] === ""){
      var firstName = row[0]; 
      var lastName = row[1]; 
      var emailAddresses = [];

      console.log(row[5]);
      if (row[3]){
        emailAddresses.push(row[2]);//student email will be 0 of the email array
      }
      if (row[5]){
        emailAddresses.push(row[4]);//parent email will be 1 of the email array
      }
      if (row[7]){
        emailAddresses.push(row[6]);//admin email will be 2 of the email array
      }
      if (emailAddresses.length == 0){
        console.log("NO Emails selected to send!");
        continue;
      }
      //construct the message body
      var message = ""; 
      if (row[14]){
        message += row[13]+"\n ";
      }
      if (row[12]){
        message += row[11]+"\n ";
      }
      if (row[16]){
        message += row[15]+"\n ";
      }
      if (row[18]){
        message += row[17];
      }
      if (message.length == 0){
        console.log("No messages selected to send!");
        
      }
      var subject = row[8]+" "+row[9];
      var emailsSent = 0;
      for (var ii = 0; ii < emailAddresses.length; ii++){
        MailApp.sendEmail(emailAddresses[ii], subject, message);
        emailsSent++;
      }
      console.log("emailsSent = "+ emailsSent);
      sheet.getRange(startRow + i, 20).setValue(emailsSent); //change the status to processed
      SpreadsheetApp.flush(); //update the spreadsheet cell immediately

    }
    
  }
}


/**
 * This is the main function that does all the work.
 * Nathan Rowbottom March 7, 2021
 */

function sendEmails() {

  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var startRow = 2; // First row of data to process
 // var numRows = sheet.getLastRow(); // Number of rows to process, set to 100 to essentially do all. Inside the loop, it breaks at first empty name
  // Fetch the range of cells
  var dataRange = sheet.getRange(startRow, 1, sheet.getLastRow(), 20);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[0] === ''){
      console.log("sent emails for "+i+ " students.");
      ui.alert("sent emails for "+i+ " students.");
      return;
    }
    else if (row[19] === ""){
      var firstName = row[0]; 
      var lastName = row[1]; 
      var emailAddresses = [];

      console.log(row[5]);
      if (row[3]){
        emailAddresses.push(row[2]);//student email will be 0 of the email array
      }
      if (row[5]){
        emailAddresses.push(row[4]);//parent email will be 1 of the email array
      }
      if (row[7]){
        emailAddresses.push(row[6]);//admin email will be 2 of the email array
      }
      if (emailAddresses.length == 0){
        console.log("NO Emails selected to send!");
        continue;
      }
      //construct the message body
      var message = ""; 
      if (row[14]){
        message += row[13]+"\n ";
      }
      if (row[12]){
        message += row[11]+"\n ";
      }
      if (row[16]){
        message += row[15]+"\n ";
      }
      if (row[18]){
        message += row[17];
      }
      if (message.length == 0){
        console.log("No messages selected to send!");
        
      }
      var subject = row[8]+" "+row[9];
      var emailsSent = 0;
      for (var ii = 0; ii < emailAddresses.length; ii++){
        MailApp.sendEmail(emailAddresses[ii], subject, message);
        emailsSent++;
      }
      console.log("emailsSent = "+ emailsSent);
      sheet.getRange(startRow + i, 20).setValue(emailsSent); //change the status to processed
      SpreadsheetApp.flush(); //update the spreadsheet cell immediately

    }
    
  }
}
