function createEvents() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();// Spreadsheet
    var sheet = ss.getSheetByName("Events");// Spreadsheet name
    var last_row = sheet.getLastRow();
    var data = sheet.getRange("A1:G" + last_row).getValues();
    var cal = CalendarApp.getCalendarById(""); // Calendar event
    
    //checking if there a duplicate sheet
    createNewSheet();
  
    for(var i = 0;i< data.length;i++){
      
      if (data[i][4] == "Reminder"){
      var event = CalendarApp.getDefaultCalendar().createAllDayEvent(data[i][0],data[i][1]);
      
      }
      else if (data[i][4] == "Daily"){
      
      var startday = dateInfo(data[i][1],3);
      var endday = dateInfo(data[i][2],3);
      
      var startmonthdate = dateInfo(data[i][1],2);
      var endmonthdate = dateInfo(data[i][2],2);
  
      starttime =dateInfo(data[i][1],5);
      endtime = dateInfo(data[i][2],5);
      
      dateSeparator(startmonthdate,endmonthdate,startday,endday,data[i][0],starttime,endtime,data[i][3],data[i][5]);
      }
      else{
        var event = CalendarApp.getDefaultCalendar().createEvent(data[i][0],new Date(data[i][1]),new Date(data[i][2])).setColor("1").removeAllReminders().setDescription(data[i][3]);
      }
    }
  
    makingNewEvents()
    sendingEmail();
  }
  
  function convertDate(monthtxt){
    var date = monthtxt;
    
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
      case "Dev":
        date = 12;
        return date;
    }
  }
  
  //Created a Sheet to Separate the dates then to create the events
  function createNewSheet(){
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var duplicateSheet = activeSpreadsheet.getSheetByName("Duplicate");
  
    if(duplicateSheet != null){
       activeSpreadsheet.deleteSheet(duplicateSheet);
    }
    duplicateSheet = activeSpreadsheet.insertSheet();
    duplicateSheet.setName("Duplicate")
  }
  

  // Breaking the dates down to months 
  function monthDateLimit(numberMonth,year){
    month = numberMonth;
    yearValue = year;
  
    if(yearValue == true && month == 02){
      month = month + 11;
    }
    
    switch(month){
      case 01:
        datelimit = 31 ;
        return datelimit;
      case 02:
        datelimit = 28 ;
        return datelimit;
      case 03:
        datelimit = 31 ;
        return datelimit;
      case 04:
        datelimit = 30 ;
        return datelimit;
      case 05:
        datelimit = 31 ;
        return datelimit;
      case 06:
        datelimit = 30 ;
        return datelimit;
      case 07:
        datelimit = 31 ;
        return datelimit;
      case 08:
        datelimit = 31 ;
        return datelimit;
      case 09:
        datelimit = 30 ;
        return datelimit;
      case 10:
        datelimit = 31 ;
        return datelimit;
     case 11:
        datelimit = 30 ;
        return datelimit;
     case 12:
        datelimit = 31 ;
        return datelimit;
      case 13:
        datelimit = 29;
        return datelimit;
    }
  }
  
  // Created a counter based on the date situation.
  function dateInfo(txt,counter){
    var array = [{}];
    var string = txt;
    var counter = counter;
    var newdatevalue;
  
    switch(counter){
      case 1: // Data Split
        array = String(string).split("/");
        newdatevalue = array[0] +"/"+array[1];
        Logger.log(newdatevalue);
        return newdatevalue;
      case 2: // month Data
        array = String(string).split(" ");
        newdatevalue = array[1] ;
        newdatevalue = convertDate(newdatevalue)
        Logger.log("Testing - " +newdatevalue);
        return newdatevalue;
      case 3: // Day data
        array = String(string).split(" ");
        newdatevalue = parseInt([array[2]]) ;
        //Logger.log("Day Funtion - " + newdatevalue);
        return newdatevalue;
      case 4: //
        array = String(string).split(" ");
        monthdate = convertDate(array[1]);
        var newstartdate = monthdate+ "/"+array[2]+"/"+array[3] + " " + array[4];
        return String(newstartdate);
      case 5:
        array = String(string).split(" ");
        timevalue = array[4];
        return timevalue;
    }
  
  }
  
  function leapYear(){
  // If a year is multiple of 400,
  // then it is a leap year
    var year = year;
    if (year % 400 == 0)
      return true;
    
    // Else If a year is multiple of 100,
    // then it is not a leap year
    if (year % 100 == 0)
      return false;
    
    // Else If a year is multiple of 4,
    // then it is a leap year
    if (year % 4 == 0)
      return true;
    
    return false;
  }
  
  function makingNewEvents(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();// Spreadsheet
    var sheet = ss.getSheetByName("Duplicate");// Spreadsheet name
    var last_row = sheet.getLastRow();
    var data = sheet.getRange("A1:G" + last_row).getValues();
    var cal = CalendarApp.getCalendarById(); // PUT EMAIL HERE
  
    for(var i = 0;i< data.length;i++){
        event = CalendarApp.getDefaultCalendar().createEvent(data[i][0],new Date(data[i][1]),new Date(data[i][2])).setColor(data[i][4]).removeAllReminders().setDescription(data[i][3]);
  
    }
  
  }
  
  function dateSeparator(startmonthvalue,endmonthvalue,startmonthdayvalue,lastmonthdayvalue,title,starttime,endtime,description,color){
    
    counter = startmonthdayvalue;
    var newcounter = 0;
    var monthmaxvalue = lastmonthdayvalue;
    var maxvalue = (lastmonthdayvalue -startmonthdayvalue);
    var yearvalue = new Date();
    var countervalue = monthDateLimit(startmonthvalue,yearvalue)
    var datearray = [];
    var startdate = startmonthvalue + "/" + startmonthdayvalue + "/" + yearvalue.getUTCFullYear()+ " " + starttime;
    var enddate = startmonthvalue + "/" + startmonthdayvalue + "/" + yearvalue.getUTCFullYear()+ " " + endtime;
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();// Spreadsheet
    var duplicatesheet = ss.getSheetByName("Duplicate");
    
    duplicatesheet.appendRow([title,startdate,enddate,description,color]);
  
    if (endmonthvalue > startmonthvalue){
      maxvalue = countervalue - startmonthdayvalue + lastmonthdayvalue;
    }
    
    for (var i = 0; i < maxvalue ; i++){
        counter = counter + 1;
        
        if(counter > countervalue){
          newcounter = newcounter+1;
          datearray.push(newcounter);
          var startdate = endmonthvalue + "/" +newcounter + "/" + yearvalue.getUTCFullYear()+ " " + starttime;
          var enddate = endmonthvalue + "/" +newcounter + "/" + yearvalue.getUTCFullYear()+ " " + endtime;
          duplicatesheet.appendRow([title,startdate,enddate,description,color]);
          Logger.log("Normal - " +newcounter);
        }
        if (counter <= countervalue){
          datearray.push(counter);
          var startdate = startmonthvalue + "/" +counter + "/" + yearvalue.getUTCFullYear()+ " " + starttime;
          var enddate = startmonthvalue + "/" +counter + "/" + yearvalue.getUTCFullYear()+ " " + endtime;
          duplicatesheet.appendRow([title,startdate,enddate,description,color]);
          Logger.log("Dup" + counter);
        }
    }
  }
  
  // Email sent out to show tell myself
  function sendingEmail(){
    
    replyemail = "Please reply with the Data that was the created."
  
    var startingtxt = "Hi There,\n" + "This is a email to indicate that the it has worked. \n" + "The following data that has been entered : \n" + replyemail;
  
    var mail = MailApp.sendEmail("","CALENDAR SCRIPT- ACTIVATED",startingtxt);
  }