/*
  This script is used in conjuntion with a spreadsheet for makerspace staff to log hours weekly.
  You can view a template version here
  https://docs.google.com/spreadsheets/d/1Ok_pe8HCVap0tukDjueqT7Bi0jCCpACA4G03kVObh8s/edit?usp=sharing
  Look at the formulas in the cells. To-do: write an auto-populate function for this script that can make a blank template
  
  Using an event-based timer set by the user in the Google Apps Script dashboard,
  sendReminder() emails staff to let them know to remember to finalize logged hours
  staff emails are in the spreadsheet; see the template

  And later in the day another trigger runs sendEmail()
  which emails the administrator in charge of processing timesheets a pdf copy of the weekly hours
  Afterwards it runs recordHistory, which goes into each user sheet and moves the current data into the log
  Then it clears the current data and updates the date headings (updateDates()) including the date headings for the main sheet
  
  Note that the administrator email to be notified, along with any addresses you want cc'd,
  are hardcoded in both sendEmail and sendReminder, so make sure to read and edit below.

  Note that you can edit permissions sheet by sheet so that staff can only edit their active edit-able hours in yellow
  
*/

//Record the history for a sheet by copying it into the bottom of this sheet. Then reset the data entry fields.
function recordHistoryforSheet(this_sheet_name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(this_sheet_name);
  var source = sheet.getRange("A1:I1");
  var dates = source.getDisplayValues();
  dates[0][0]='';
  sheet.appendRow(dates[0]); //Append the dates to the log
  var source = sheet.getRange("A2:I2");
  var hours = source.getDisplayValues();
  hours[0][0] = 'Hours';
  sheet.appendRow(hours[0]); //Append the hours worked on those dates to the log
  var cleararea = sheet.getRange("B2:H2"); 
  if (this_sheet_name != "All Staff") cleararea.clearContent(); //clear the entry fields
};

//Record the history for all sheets
function recordHistory() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) if (sheets[i].getName() != "All Staff") recordHistoryforSheet(sheets[i].getName() ); //Record history for sheet
  for (var i=0; i<sheets.length ; i++) updateDates(sheets[i].getName()); //Update date ranges on sheet
};

//Update the dates fields in the sheet
function updateDates(this_sheet_name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(this_sheet_name);
  var a1notation = ["H1","G1","F1","E1","D1","C1","B1"];
  var range;
  for (var i=0; i<7; i++){
    range =  sheet.getRange(a1notation[i]);
    range.setFormula("=TODAY()+7-WEEKDAY(TODAY(),1)+"+String(6-i+1)); //fixes Sunday problem??? Sundays were landing on wrong day of week
    //range.setFormula("=DATE(YEAR(TODAY()),1,1)+((WEEKNUM(TODAY()))*7)+7-WEEKDAY(DATE(YEAR(TODAY()),1,"+String(i+1)+"),1)"); //use equation to get date
    range.setValue(range.getDisplayValue()); //set it to the display value
  }
}

//Send the email to the Administrator and then record the history/update the sheets
function sendEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("All Staff");
  var id = SpreadsheetApp.getActiveSpreadsheet().getId();
  var spreadsheetFile = DriveApp.getFileById(id);
  
  //make temp copy
  var today = new Date();
  var tmpname = (today.getMonth()+1)+'-'+today.getDate()+'-'+today.getFullYear();
  var folder = DriveApp.createFolder(tmpname);
  var copyofFile = SpreadsheetApp.open(DriveApp.getFileById(spreadsheetFile.getId()).makeCopy(tmpname,folder));  

  var range = copyofFile.getSheetByName("All Staff").getRange("A1:K14");     
  range.copyTo(range, {contentsOnly:true});

  //delete redundant sheets
  var sheets = copyofFile.getSheets();
  for (i = 0; i < sheets.length; i++) if (sheets[i].getSheetName() != "All Staff") copyofFile.deleteSheet(sheets[i]);

  var blob = DriveApp.getFileById(copyofFile.getId()).getAs('application/pdf').setName(tmpname+".pdf");
  MailApp.sendEmail('recipientemail', 'Weekly Staff Hours '+tmpname, 'Hi Recipient. Attached is the weekly hours spreadsheet. Thanks! -Shira', {attachments:[blob],cc:"email addresses to be cc'd"});
  DriveApp.getFileById(copyofFile.getId()).setTrashed(true);
  folder.setTrashed(true);
  
  //Update everything now
  recordHistory();
};

//Send reminders to the staff to fill out the timesheet
function sendReminder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("All Staff");
  var emails_column = sheet.getRange("J2:J").getValues();
  var last = emails_column.filter(String).length; 
  var emails = sheet.getRange("J2:J"+last+2).getValues();
  var sendto = "";
  for (var i=0; i<last ; i++) { sendto = sendto+emails[i].toString()+",";}  
  MailApp.sendEmail(sendto, 'Friday reminder: Finalize timesheet for this week', 'Hi staff,\nPlease finalize the timesheet for this week before the cutoff today.\n\nThank you.\n-Shira', {cc:"email addresses to be cc'd"});
};