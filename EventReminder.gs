/*
@author: Daniela Kepper
Date: 25.01.2018

Function: The aim of this script is to use the input in the Google Sheet "Team Events" and messages the owner of the monthly team event
at the beginning of the respective month to remind him/her about the organizational duties.
*/

function sendEventEmail() {

  // function variables
  var newDay = new Date();
  var thisDay = newDay.toString().replace(/(\d\d:\d\d:\d\d)/, "00:00:00");
  thisDay = new Date(thisDay);

  var destinationFile = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team Events");
  var values = destinationFile.getRange(2,1,destinationFile.getLastRow(),4).getValues();

  // loop trough elements
  for (var i = 0; i < values.length; i++){
    if(values[i][0] == ""){
      break;
    }

    var eventMonthString = values[i][2];
    var dateParts = eventMonthString.split(/\./);
    var eventMonth = new Date(dateParts[1] +"/"+ dateParts[0] + "/" + thisDay.getFullYear() + " GMT +00");

    var monthOwner = values[i][0];
    var firstName = monthOwner.split(" ")[0];
//    var firstName = nameSplit[0];
    var ownerEmail = values[i][1];
    var relevantMonthString = values[i][3];

  // email loops regular & reminder

    if (eventMonth.getTime() == thisDay.getTime()) {
      if(firstName.slice(-1) != "s") {
        MailApp.sendEmail(ownerEmail, firstName + "'s Monthly Team Event Organization [Automated email]", relevantMonthString + " is coming up in a few days and it's your turn to organize this month's team event.", {cc: "name@gmail"});
      } else {
         MailApp.sendEmail(ownerEmail, firstName + "' Monthly Team Event Organization [Automated email]", relevantMonthString + " is coming up in a few days and it's your turn to organize this month's team event.", {cc: "name@gmail"});
      }
    }

    // Logger.log((( eventMonth.getTime() - thisDay.getTime()) / (24 * 60 * 60 * 1000)) == 7);
    if ((( eventMonth.getTime() + thisDay.getTime()) / (24 * 60 * 60 * 1000)) == 10){
      if(firstName.slice(-1) != "s") {
      MailApp.sendEmail(ownerEmail, firstName + "'s Monthly Team Event Organization Reminder [Automated email]", "*If you have already organized the team event, you can ignore this e-mail* \n \n " + relevantMonthString + "'s final Team Event organization Reminder - enjoy!", {cc: "name@gmail.com"});
      } else {
       MailApp.sendEmail(ownerEmail, firstName + "' Monthly Team Event Organization Reminder [Automated email]", "*If you have already organized the team event, you can ignore this e-mail* \n \n " + relevantMonthString + "'s final Team Event organization Reminder - enjoy!", {cc: "name@gmail.com"});
      }

  }
}
}
