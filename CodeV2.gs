function myFunction() {
  
  var spreadsheetname = "nameofsheet";
  var email = "email";
  
  //Spread related
  var ss = SpreadsheetApp.getActiveSpreadsheet();// Spreadsheet
  var sheet = ss.getSheetByName(spreadsheetname);// Spreadsheet name
  var last_row = sheet.getLastRow();
  var data = sheet.getRange("A1:I" + last_row).getValues();
  var cal = CalendarApp.getCalendarById(email); 

  var bodytxt = [{}];

  //creating a new sheet to store data
  newSheet();
  
  for(var i = 0;i<data.length;i++){ // outputting the list excel sheet
  var eventname = "[" + data[i][0] + "] " + data[i][1];
    if (data[i][8] == "Daily"){
        var endofstartdate = textModifier(data[i][2],5) + " " + textModifier(data[i][3],2) ;
        
        var color = forColor(data[i][7]);
        var description = data[i][6];
        var enddate = textModifier(data[i][4],1) + " " + textModifier(data[i][5],2);
        var startdate = textModifier(data[i][2],1) + " " + textModifier(data[i][3],2);
        var enddatetext = textModifier(data[i][4],5) + " " + textModifier(data[i][5],2);
        var startmonthvalue = textModifier(data[i][2],4);
        var endmonthvalue = textModifier(data[i][4],4);
        var firstvalue = textModifier(data[i][2],3);
        var endvalue = textModifier(data[i][4],3);

        bodytxt.push([eventname,startdate,enddate,description]);

        addingDataToSheet(eventname,description,color,startmonthvalue,endmonthvalue,endofstartdate,enddatetext,firstvalue,endvalue);
    }
    else{
      CalendarApp.createAllDayEvent(eventname,data[i][2]).removeAllReminders;
    }}

  createEvents();
  emailSend(sendoutEmail(bodytxt,bodytxt.length),email);
}

function addingDataToSheet(eventname,description,color,startmonthvalue,endmonthvalue,startendtext,enddatetext,firstdatevalue,enddatevalue){
  
  // Looking for spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("forduplicate");
  var countnum = 0;
  var newcounter = 0;

  if (startmonthvalue == endmonthvalue){ // checking for the count
    countnum = enddatevalue - firstdatevalue;}
  else{
    countnum = calculatorDateLimit(startmonthvalue) + enddatevalue -firstdatevalue ;}
  
  for(var i =0;i<countnum+1; i++){ //duplicating the month
    
    if (firstdatevalue > calculatorDateLimit(startmonthvalue)){
      newcounter++
      sheet.appendRow([eventname,endmonthvalue+"/"+newcounter+startendtext,endmonthvalue+"/"+newcounter+enddatetext,description,color]);
    }else{
      sheet.appendRow([eventname,startmonthvalue+"/"+firstdatevalue+startendtext,startmonthvalue+"/"+firstdatevalue +enddatetext,description,color]);
    }
    firstdatevalue++;
  }
}

function calculatorDateLimit(month){ // date limit for each month

  var leap = new Date().getFullYear; //checking if leap year
  var counter = 0;
  if ((0 == leap % 4) && (0 != leap % 100) || (0 == leap % 400)) {
    counter = 1;
  }
  if (counter == 1 && month == 2 ){
      month = 13;
  }

  switch(month){
    case 1:
      return 31;
    case 2:
      return 28;
    case 3:
      return 31;
    case 4:
      return 30;
    case 5:
      return 31;
    case 6:
      return 30;
    case 7:
      return 31;
    case 8:
      return 31;
    case 9:
      return 30;
    case 10:
      return 31;
    case 11:
      return 30;
    case 12:
      return 30;
    case 13:
      return 29;
  }
}

function createEvents(){ //creating the events
var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var duplicateSheet = activeSpreadsheet.getSheetByName("forduplicate");
var last_row = duplicateSheet.getLastRow();
var data = duplicateSheet.getRange("A1:I" + last_row).getValues();

for(var i = 0; i < data.length;i++){
  CalendarApp.createEvent(data[i][0],data[i][1],data[i][2]).setColor(data[i][4]).setDescription(data[i][3]).addPopupReminder;
}}

function newSheet(){ // new sheet to store the missing the gaps of dates
var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var duplicateSheet = activeSpreadsheet.getSheetByName("forduplicate");

if(duplicateSheet != null){
    activeSpreadsheet.deleteSheet(duplicateSheet);
}
duplicateSheet = activeSpreadsheet.insertSheet();
duplicateSheet.setName("forduplicate")}

function textModifier(text,counter){
  var textsplit = [{}];
  textsplit = String(text).split(" "); // to prevent repeating code.
  var newvalue;

switch(counter){
  case 1: // Date
    newvalue = monthConverter(textsplit[1])+"/"+textsplit[2]+"/"+textsplit[3];
    return newvalue;
  
  case 2: // Time
    newvalue = textsplit[4];
    return newvalue;

  case 3: // date number
    return parseInt(textsplit[2]);

  case 4: // month 
    return monthConverter(textsplit[1]);
    
  case 5: // text for end date
    newvalue = "/"+ textsplit[3];
    return newvalue;
}}

// converting text to integer
function monthConverter(date){
  switch(date){
    case "Jan":
      date = 01;
      return date;
    case "Feb":
      date = 02;
      return date;
    case "Mar":
      date = 03;
      return date;
    case "Apr":
      date = 04;
      return date;
    case "May":
      date = 05;
      return date;
    case "Jun":
      date = 06;
      return date;
    case "Jul":
      date = 07;
      return date;
    case "Aug":
      date = 08;
      return date;
    case "Sep":
      date = 09;
      return date;
    case "Oct":
      date = 10;
      return date;
    case "Nov":
      date = 11;
      return date;
    case "Dec":
      date = 12;
      return date;
  }}

function forColor(color){// converting calendar color
   switch(color){
   case "PALE_BLUE":
      return 1;
    case "PALE_GREEN":
      return 2;
    case "MAUVE":
      return 3;
    case "PALE_RED":
      return 4;
    case "YELLOW":
      return 5;
    case "ORANGE":
      return 6;
    case "CYAN":
      return 7;
    case "GRAY":
      return 8;
    case "BLUE":
      return 9;
}}

function sendoutEmail(data,length){ //sending out automated email
  var events = [{}];

  for(var i = 0; i<length;i++){
    if(data[i][0] == null ){
    }else{
    var banner = "+------------------------------------+" + "\n"
    var eventname = "| Name : " + data[i][0] +"\n";
    var startevent = "| Start date : " + data[i][1]  + "\n";
    var endevent = "| End   date: " + data[i][2] + "\n" ;
    var emaildesc = "| Description  : " + data[i][3] +"\n";
    var endemail  = "+------------------------------------+" + "\n";
    
    var output_final = banner +eventname+startevent+endevent+emaildesc +endemail;
    events.push(output_final);
    }}
  return events;
  }

function emailSend(data,email){
  var emailtxt = "Hello," + "\n" + "This just a reminder that the script has been activated and this is rerecorded data. " + "\n" + "The recorded data is only the daily resources:" + "\n" + data +"\n" + "Regards" + "\n" + "Auto Scripted | G SCRIPT PROGRAM";
  MailApp.sendEmail(email,"Calendar Script Activated", emailtxt)
}


