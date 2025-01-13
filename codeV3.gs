//   ____    _    _     _____ _   _ ____    _    ____  
//  / ___|  / \  | |   | ____| \ | |  _ \  / \  |  _ \ 
// | |     / _ \ | |   |  _| |  \| | | | |/ _ \ | |_) |
// | |___ / ___ \| |___| |___| |\  | |_| / ___ \|  _ < 
//  \____/_/   \_\_____|_____|_| \_|____/_/   \_\_| \_\

// Getting the date of the week
function getDayofWeek() {
  var date = new Date();
  var datetospring = date.toString();
  var datedata = datetospring.split(" ");
  return datedata;
}

// how to use the method - endOfMonth(new Date())
function endOfMonth(date) {
  var lastdayofthemonth = new Date(date.getFullYear(), date.getMonth() + 1, 0);
  var splitvalue = lastdayofthemonth.toString().split(" ")

  return splitvalue[2]; // limit value
}

function wholeWeek() {
  var today = getDayofWeek();
  var thismonth = today[0];
  var nextmonth = endOfMonth(new Date());

  switch (thismonth) {
    case 'Mon':
      var lastoftheweek = parseInt(today[2]) + 4;
      var firstoftheweek = parseInt(today[2]);
      var requireddata = [firstoftheweek, lastoftheweek, lastoftheweek - firstoftheweek];
      var days = dayCount(firstoftheweek, lastoftheweek);
      return days;
      break;

    case 'Tue':
      var lastoftheweek = parseInt(today[2]) + 3;
      var firstoftheweek = parseInt(today[2]) - 1;
      var requireddata = [firstoftheweek, lastoftheweek, lastoftheweek - firstoftheweek];
      var days = dayCount(firstoftheweek, lastoftheweek);
      return days;
      break;

    case 'Wed':
      var lastoftheweek = parseInt(today[2]) + 2;
      var firstoftheweek = parseInt(today[2]) - 2;
      var requireddata = [firstoftheweek, lastoftheweek, lastoftheweek - firstoftheweek];
      dayCount(firstoftheweek, lastoftheweek);
      return requireddata;
      break;

    case 'Thu':
      var lastoftheweek = parseInt(today[2]) + 1;
      var firstoftheweek = parseInt(today[2]) - 3;
      var requireddata = [firstoftheweek, lastoftheweek, lastoftheweek - firstoftheweek];
      dayCount(firstoftheweek, lastoftheweek);
      return requireddata;
      break;

    case 'Fri':
      var lastoftheweek = parseInt(today[2]);
      var firstoftheweek = parseInt(today[2]) - 4;
      var requireddata = [firstoftheweek, lastoftheweek, lastoftheweek - firstoftheweek];
      var days = dayCount(firstoftheweek, lastoftheweek);
      return days;
      break;

    case 'Sat':
      var lastoftheweek = parseInt(today[2]) - 1;
      var firstoftheweek = parseInt(today[2]) - 5;
      var requireddata = [firstoftheweek, lastoftheweek, lastoftheweek - firstoftheweek];
      var days = dayCount(firstoftheweek, lastoftheweek);
      return days;
      break;

    case 'Sun':
      var lastoftheweek = parseInt(today[2]) - 2;
      var firstoftheweek = parseInt(today[2]) - 6;
      var requireddata = [firstoftheweek, lastoftheweek, lastoftheweek - firstoftheweek];
      dayCount(firstoftheweek, lastoftheweek);
      return requireddata;
      break;
  }

}

// This to cound all the days of the week
function dayCount(firstdate, enddate) {
  var lastvalue = endOfMonth(new Date());
  var lastmonthenddate = endOfMonth(new Date("01 " + String(monthManpulation(getDayofWeek()[1], 2)) + " 2024"));

  var newvalue = 0;
  var backvalue = 0;
  var alldays = [];

  for (var i = parseInt(firstdate); i <= parseInt(enddate); i++) {

    if (i <= parseInt(lastvalue) && i > 0) { // count normally
      var currentdate = i + " " + getDayofWeek()[1];
      alldays.push(currentdate);

    } else if (i <= 0) { // if the month value is behind this one

      backvalue = parseInt(lastmonthenddate) + i
      finalbackvalue = backvalue + " " + monthManpulation(getDayofWeek()[1], 2);
      alldays.push(finalbackvalue);
    }
    else { // if the value overs the month
      newvalue++;
      var nextmonth = newvalue + " " + monthManpulation(getDayofWeek()[1], 1);
      alldays.push(nextmonth);
    }
  }
  return alldays;
}

function monthManpulation(input, counter) {
  var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var counter;

  switch (counter) {
    case 1: // Forward
      var nextmonth;
      for (var i = 0; i < months.length; i++) {

        if (input == months[i]) {
          nextmonth = months[i + 1];
        }
      }
      return nextmonth;

    case 2: // backwards
      var lastmonth;
      months.reverse();

      for (var i = 0; i < months.length; i++) {

        if (input == months[i]) {
          lastmonth = months[i + 1];

        }
      }
      return lastmonth;

  }
}

function createEvents() { // Read sheet then create event
  var spreadsheetname = "Templates";
  var ss = SpreadsheetApp.getActiveSpreadsheet();// Spreadsheet
  var sheet = ss.getSheetByName(spreadsheetname);// Spreadsheet name
  var last_row = sheet.getLastRow();
  var data = sheet.getRange("A2:I" + last_row).getValues();

  var year = new Date();

  for (var i = 0; i < data.length; i++) {
    var eventname = "[" + data[i][0] + "] " + data[i][1];
    var description = data[i][2];

    var eventcategory = data[i][5];
    var color = forColor(data[i][6]);
    var eventstatus = data[i][8];

    var days = wholeWeek();

    if (eventcategory == "Daily" && eventstatus == "On") {
      for (var j = 0; j < days.length; j++) {
        var startdate = days[j] + ", " + year.getFullYear() + " " + splitdate(data[i][3]);
        var enddate = days[j] + ", " + year.getFullYear() + " " + splitdate(data[i][4]);
        CalendarApp.createEvent(eventname, new Date(startdate), new Date(enddate)).setDescription(description).setColor(color).removeAllReminders();
      }
    } else if (eventcategory == "Weekly" && eventstatus== "On") {

      var weekday = data[i][7];
      var weekdayvalues = weekdayChoseData(weekday, days);
      var startdate = weekdayvalues + ", " + year.getFullYear() + " " + splitdate(data[i][3]);
      var enddate = weekdayvalues + ", " + year.getFullYear() + " " + splitdate(data[i][4]);

      CalendarApp.createEvent(eventname, new Date(startdate), new Date(enddate)).setDescription(description).setColor(color).removeAllReminders();

    }
  }
}

function splitdate(date) {
  var datetospring = date.toString();
  var datedata = datetospring.split(" ");
  return datedata[4];
}


function forColor(color) {// converting calendar color
  switch (color) {
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
  }
}

function weekdayChoseData(dates, input) { // checking which day it is
  var datevalues = ["Mon", "Tue", "Wed", "Thu", "Fri"];
  var finaloutput;

  for (var i = 0; i < datevalues.length; i++) {
    if (datevalues[i] == dates) {
      finaloutput = input[i];
    }
  }

  return finaloutput;
}

// Adding Menu Item 
// TODO CODE - I may need to create a set up button for this
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Timetable Support Tool')
      .addItem('About Tool', 'about')
      .addSeparator()
      .addItem('Run Script', 'eventsrun')
      .addItem('Set Up Excel', 'nothing')
      .addToUi();
}

function about() {
  SpreadsheetApp.getUi()
     .alert('This is the About tool');
}

function eventsrun() {
  SpreadsheetApp.getUi() 
     .alert('Creating Calendar Events');
  createEvents()
}
