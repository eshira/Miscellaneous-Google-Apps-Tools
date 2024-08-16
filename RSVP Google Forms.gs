/* User defined variables section. Fill this section out, then run the Initalization function to create this event.
   DON'T edit the form manually! This tool will overwrite those fields.
*/

//TODO make a GUI for all this

//Event Title
EVENT_TITLE = "3D Printing and OpenSCAD";

//Event Time FOLLOW THIS FORMAT EXACTLY
EVENT_TIME = "2018-04-04 15:00"; //Later this gets converted to both a human readable and a machine readable format.

//This event will take place in...
EVENT_LOCATION = "M5, in the Good Room";

//Event duration
EVENT_DURATION = "1.5 hours"; //Any human readable string will do

//Number of seats open
EVENT_SEATS = 15; //Enter an integer representing the max number of seats for this event

//Describe the event briefly. Make sure to mention who is running this event.
EVENT_DESCRIPTION = "Learn and make your own parametric 3D models using OpenSCAD, an open source tool that uses code to generate 3D shapes. Learn about 3D printers and how to design parts around the limitations of FDM (fused deposition modeling). After the workshop concludes, the part you designed will be 3D printed for you, and you will be notified when it is ready for pick up. Please bring a laptop to this class. Hosted by Shira Epstein.";

//Is this event open to members only? If so, add a blurb about that.
MEMBERS_ONLY = true; 

//How many hours prior to the event start time would you like to remind the reservation holders about the event?
REMINDER_HRS_EARLY = 6;

//END OF USER DEFINED VARIABLES. Run the init script when you are ready!

//Code

//Initialize the form. Run this script when you are ready to publish.
function createForm() {
  //Get the current form
  var form = FormApp.getActiveForm();
  //Create an associated spreadsheet
  var ss = SpreadsheetApp.create(EVENT_TITLE+" responses");
  
  //Set the spreadsheet as the response destinatio
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  //Put this spreadsheet in the M5 Events Folder. TODO this should be less hard-coded
  DriveApp.getFolderById('0B9W3B9eMiUtYanFZcVJtQWxZa28').addFile(DriveApp.getFileById(ss.getId()));
  
  // remove document from the root folder
  var thisfile = DriveApp.getFileById(ss.getId());
  DriveApp.getRootFolder().removeFile(thisfile);
  
  respSheet = ss.getSheets()[0];  // Forms typically go into sheet 0.
  respSheet.getRange("D1").setFormula("=SUM(D2:D)");
  respSheet.getRange("E1").setValue(0);
  
  //Create the confirmation string
  var confirmation = "Thank you for your response.\n\nIf you have reserved a spot, you will receive a reminder email about this event.\n\nIf you are on the waitlist, you will be notified if you are assigned a reservation for this event.\n\nIf you have opted to rescind your spot on the reservation or waitlist, you are now removed from this event.";
  
  //Set the title of the event, description, and confirmation message 
  form.setTitle(EVENT_TITLE).setDescription(descriptionstring()).setConfirmationMessage(confirmation);

  //List all the form items
  var allItems = form.getItems();
  //Craete some variables for later
  var i=0,thisItem,thisItemType,myheader,mytext;

  //For every item in the form
  for (i=0;i<allItems.length;i+=1) {
    thisItem = allItems[i];
    thisItemType = thisItem.getType();
    if (thisItemType==FormApp.ItemType.SECTION_HEADER){ //if we found the section header
      myheader = thisItem.asSectionHeaderItem();
      myheader.setTitle("Reserve a seat for this event!");
      myheader.setHelpText("");
    }
  }

  //Open the form to accept responses
  form.setAcceptingResponses(true);
  
  /* this section taken in part from web tutorial: http://labnol.org/?p=20707  */
  deleteTriggers_(); //delete existing triggers
  
  //Trigger to send out reminder email to those with reservations
  if (EVENT_TIME !== "") {
    var remindtime = new Date(parseDate_(EVENT_TIME).getTime());
    remindtime.setHours(remindtime.getHours()-REMINDER_HRS_EARLY);
    ScriptApp.newTrigger("sendReminder")
     .timeBased()
     .at(remindtime)
     .create();
  }
  
  //Trigger to close this form once the event takes place
  if (EVENT_TIME !== "") { 
    ScriptApp.newTrigger("closeForm")
    .timeBased()
    .at(parseDate_(EVENT_TIME))
    .create(); 
  }
  
  //Trigger on form submission to handle reservation/waitlist upkeep
  if (EVENT_SEATS !== "") { 
    ScriptApp.newTrigger("checkLimit")
    .forForm(FormApp.getActiveForm())
    .onFormSubmit()
    .create();
  }
}


function sendReminder()
{
  //Send reminders to the reservees
  var form = FormApp.getActiveForm();
  //Get the associated spreadsheet
  var destId = FormApp.getActiveForm().getDestinationId();
  var ss = SpreadsheetApp.openById(destId);
  var respSheet = ss.getSheets()[0];  // Forms typically go into sheet 0.
  rsvps = respSheet.getRange("D2:D").getDisplayValues();
  emails = respSheet.getRange("B2:B").getDisplayValues();
  maxval = emails.filter(String).length;
  var formResponses = form.getResponses();
  for (var i=0; i<maxval; i++){ //iterate over spreadsheet entries
    if (rsvps[i]==1){ //RSVP holder
      //Iterate over responses to get their RSVP URL
      for (var j = 0; j < formResponses.length; j++) {
        identity = formResponses[j].getRespondentEmail();
        if (identity == emails[i]) {
          var time = Utilities.formatDate(new Date(parseDate_(EVENT_TIME).getTime()),'America/New_York', 'EEEE, MMMM dd, yyyy @ h:mm a');
          var subject = "Reminder: '"+EVENT_TITLE+"'";
          var content = "Reminder: You have reserved a spot for '"+ EVENT_TITLE + "'" + " which is taking place on " + time + " in " + EVENT_LOCATION + ". If you need to cancel your reservation, use the following link:\n\n"+formResponses[j].getEditResponseUrl();
          MailApp.sendEmail(emails[i], subject+" "+ time, content);
        }
      }
    }
  }
}

function test()
{
  var test1 = new Date(parseDate_(EVENT_TIME).getTime());
  test1.setHours(test1.getHours()-3);
  var timestring = Utilities.formatDate(test1,'America/New_York', 'EEEE, MMMM dd, yyyy @ h:mm a'); //format the user specified event time
  Logger.log(parseDate_(EVENT_TIME));
}

//Supporting functions

//Format and return the description string
function descriptionstring () {
  //Get the current form
  var form = FormApp.getActiveForm();
  //Get the associated spreadsheet
  var destId = FormApp.getActiveForm().getDestinationId();
  var ss = SpreadsheetApp.openById(destId);
  var respSheet = ss.getSheets()[0];  // Forms typically go into sheet 0.
  
  var timestring = Utilities.formatDate(new Date(parseDate_(EVENT_TIME).getTime()),'America/New_York', 'EEEE, MMMM dd, yyyy @ h:mm a'); //format the user specified event time
  if ((EVENT_SEATS - respSheet.getRange("D1").getDisplayValue() ) > 0) var seatstring = (EVENT_SEATS - respSheet.getRange("D1").getDisplayValue() )+" out of "+EVENT_SEATS+" seats available\n\n";
  else var seatstring = "0 out of "+EVENT_SEATS +" seats available\n\n";
  var durationstring = "Duration: "+EVENT_DURATION+".\n\n";
  var membersonlystring = "\n\nThis event is open to M5 members only. ECE students can register for membership at https://sites.google.com/site/m5makerspace/";
  var locationstring = "This event takes place in "+EVENT_LOCATION+".\n\n";
  var description = timestring+". " + durationstring +locationstring+seatstring + EVENT_DESCRIPTION;
  if (MEMBERS_ONLY) description = description + membersonlystring;
  return description;
}

/* Delete all existing Script Triggers */
function deleteTriggers_() {  
  var triggers = ScriptApp.getProjectTriggers();  
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

/* Close the Google Form, Stop Accepting Reponses */
function closeForm() {  
  var form = FormApp.getActiveForm();
  form.setAcceptingResponses(false);
  deleteTriggers_();
}

/* Upkeeping of form, triggered on submit */
function checkLimit() {
  //Get the current form
  var form = FormApp.getActiveForm();
  //Get the associated spreadsheet
  var destId = FormApp.getActiveForm().getDestinationId();
  var ss = SpreadsheetApp.openById(destId);
  var respSheet = ss.getSheets()[0];  // Forms typically go into sheet 0.
  
  var formResponses = form.getResponses();
  //Iterate over responses
  for (var i = 0; i < formResponses.length; i++) {
    var rescinding = false; //Is this person rescinding their reserve/waitlist?
    var identity;
    var formResponse = formResponses[i];
    var itemResponses = formResponse.getItemResponses();
    //Get the UMail identity
    identity = formResponse.getRespondentEmail();
    //Iterate over items in the response
    for (var j = 0; j < itemResponses.length; j++) {
      var itemResponse = itemResponses[j];
      //See if they have rescinded their reservation/waitlist spot
      if (String(itemResponse.getItem().getTitle()).length == 0) {
        //Logger.log(itemResponse.getResponse());
        if (itemResponse.getResponse() == "Remove me from this event") {
          rescinding = true;
        }
      }
    }

    range = respSheet.getRange("B2:B");
    status = respSheet.getRange("C2:C");
    //Find the last non-empty entry, rather than the last row in the sheet
    maxval = range.getValues().filter(String).length;
    identities = range.getDisplayValues();
    //Iterate through the destination spreadsheet
    for (var j=0; j < maxval; j++){
      //If this entry refers to the current response we are evaluating
      if (String(identities[j]) == String(identity)){

        //If the person is listed as rescinding, clear their priority and status
        if (rescinding) {
          if (respSheet.getRange("D"+String(j+2)).getValue()==1) { //if they are rescinding and giving up their reservation now
            Logger.log(String(identity)+" has opted out of a reservation.");
            var subject = "You have cancelled your reservation for '"+EVENT_TITLE+"'";
            var content = subject+" scheduled for "+ Utilities.formatDate(new Date(parseDate_(EVENT_TIME).getTime()),'America/New_York', 'EEEE, MMMM dd, yyyy @ h:mm a')+". If you change your mind, you can use the following URL to register again:\n\n"+formResponse.getEditResponseUrl()+"\n\nPlease note that your place in the queue will not be held.";
            MailApp.sendEmail(identity, subject, content);
          }
          else if ((respSheet.getRange("D"+String(j+2)).getValue()==0 ) && (String(respSheet.getRange("D"+String(j+2)).getValue()) != "" ) ) { //if they are rescinding and giving up their waitlist
            Logger.log(String(identity)+" has opted out of the waitlist.");
            var subject = "You have opted out of the waitlist for '"+EVENT_TITLE+"'";
            var content = subject+" scheduled for "+ Utilities.formatDate(new Date(parseDate_(EVENT_TIME).getTime()),'America/New_York', 'EEEE, MMMM dd, yyyy @ h:mm a')+". If you change your mind, you can use the following URL to register again:\n\n"+formResponse.getEditResponseUrl()+"\n\nPlease note that your place in the queue will not be held.";
            MailApp.sendEmail(identity, subject, content);
          }
          respSheet.getRange("E"+String(j+2)).clear();
          respSheet.getRange("D"+String(j+2)).clear();
          
        }
        //If the person is NOT rescinding
        else {
          //If the person had rescinded previously
          //if (String(status.getDisplayValues()[j]).length>0) {
          //  respSheet.getRange("C"+String(j+2)).clear(); //Clear their rescind field
          //}
          //If the person didn't already have a priority number, assign them one (new person or opting back in person)
          if (String(respSheet.getRange("E"+String(j+2)).getDisplayValue()).length==0) {
            var lastpriority = respSheet.getRange("E1").getValue();
            respSheet.getRange("E"+String(j+2)).setValue(lastpriority+1);
            respSheet.getRange("E1").setValue(lastpriority+1);
            //And figure out if they get a waitlist or reserve
            if (respSheet.getRange("D1").getDisplayValue() < EVENT_SEATS) { //spots are available
              respSheet.getRange("D"+String(j+2)).setValue(1);
              Logger.log(String(identity)+ " has been assigned a reservation!");
              var subject = "You have reserved a spot for '"+EVENT_TITLE+"'";
              var content = subject+" scheduled for "+ Utilities.formatDate(new Date(parseDate_(EVENT_TIME).getTime()),'America/New_York', 'EEEE, MMMM dd, yyyy @ h:mm a')+" and taking place in "+EVENT_LOCATION+". If you change your mind, you can use the following URL to cancel your reservation:\n\n"+formResponse.getEditResponseUrl();
              MailApp.sendEmail(identity, subject, content);
            }
            else { //no spots; person is waitlisted
              respSheet.getRange("D"+String(j+2)).setValue(0);
              Logger.log(String(identity)+ " has been placed on the waitlist.");
              var subject = "You are on the waitlist for '"+EVENT_TITLE+"'";
              var content = "There are no spots currently available for the following event: '"+ EVENT_TITLE+ "' taking place on ";
              content = content+Utilities.formatDate(new Date(parseDate_(EVENT_TIME).getTime()),'America/New_York', 'EEEE, MMMM dd, yyyy @ h:mm a');
              content = content+". You have been placed on the waitlist. You will be notified and automatically assigned a reserved spot if it becomes available. If you have changed your mind, use the following link to leave the waitlist for this event:\n\n"+formResponse.getEditResponseUrl();
              MailApp.sendEmail(identity, subject, content);
            }
          }
        }
      }
    }
  }
  
  //Find the next person on the waitlist and assign them a reservation
  var diff = EVENT_SEATS - respSheet.getRange("D1").getDisplayValue();
  for (var i=0; i < diff; i++){ //for the number of opened up slots
    respSheet.sort(5); //sort the spreadsheet by priority
    respSheet.getRange("D1").setFormula("=SUM(D2:D)"); //i think maybe it was doing an autocorrection to the sum upon sorting...
    for (var k=0; k < maxval; k++){ //iterate through all the people in the list
      if ((respSheet.getRange("D"+String(k+2)).getValue() == 0) && (String(respSheet.getRange("D"+String(k+2)).getValue()) != "") ){ //we found the first waitlisted person
        if (respSheet.getRange("D1").getDisplayValue() < EVENT_SEATS) { //extra safety to avoid giving away too many seats
          respSheet.getRange("D"+String(k+2)).setValue(1); //give them a spot
          Logger.log(String(respSheet.getRange("B"+String(k+2)).getValue())+" has been bumped from waitlist to reservation!");
          for (var i = 0; i < formResponses.length; i++) { //iterate over reponses to find the one where row B col k+2 matches the email
            var identity;
            var formResponse = formResponses[i];
            //Get the UMail identity
            identity = formResponse.getRespondentEmail();
            if (identity == String(respSheet.getRange("B"+String(k+2)).getValue())) {
              var subject = "You have been assigned a reservation for '"+EVENT_TITLE+"'";
              var content = subject+" scheduled for "+ Utilities.formatDate(new Date(parseDate_(EVENT_TIME).getTime()),'America/New_York', 'EEEE, MMMM dd, yyyy @ h:mm a')+" and taking place in "+EVENT_LOCATION+". If you change your mind, you can use the following URL to cancel your reservation:\n\n"+formResponse.getEditResponseUrl();
              MailApp.sendEmail(identity, subject, content);
            }
          }
          break; //add only the first person on the waitlist for this round
        }
      }
    }
  }
  
  //if the event is full change some of the language
  if (respSheet.getRange("D1").getDisplayValue() >= EVENT_SEATS ) {
    //List all the form items
    var allItems = form.getItems();
    //Create some variables for later
    var i=0,thisItem,thisItemType,myheader,mytext;
    for (i=0;i<allItems.length;i+=1) {
      thisItem = allItems[i];
      thisItemType = thisItem.getType();
      if (thisItemType==FormApp.ItemType.SECTION_HEADER){
        myheader = thisItem.asSectionHeaderItem();
        myheader.setTitle("Sign up for the waitlist!");
        myheader.setHelpText("This event is full. Submit this form to sign up for the waitlist.");
      }
    }
  }
  
  else { //event is not full, return to old language
   //List all the form items
    var allItems = form.getItems();
    //Create some variables for later
    var i=0,thisItem,thisItemType,myheader,mytext;
    for (i=0;i<allItems.length;i+=1) {
      thisItem = allItems[i];
      thisItemType = thisItem.getType();
      if (thisItemType==FormApp.ItemType.SECTION_HEADER){
        myheader = thisItem.asSectionHeaderItem();
        myheader.setTitle("Reserve a seat for this event!");
        myheader.setHelpText("");
      }
    }
  }
  //Set the description 
  form.setDescription(descriptionstring());  
}

/* Parse the Date for creating Time-Based Triggers */
function parseDate_(d) {
  return new Date(d.substr(0,4), d.substr(5,2)-1, 
                  d.substr(8,2), d.substr(11,2), d.substr(14,2));
}