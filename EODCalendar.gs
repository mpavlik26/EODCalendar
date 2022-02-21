var ss = SpreadsheetApp.getActiveSpreadsheet();
var s = ss.getActiveSheet(); 

var CALENDAR_PAGERDUTY_ID = "zonky.cz_3439363838313737383333@resource.calendar.google.com";
var CALENDAR_AGILE_MEETINGS_ID = "zonky.cz_2d323130323133313933@resource.calendar.google.com";
var CALENDAR_SEND_INVITES = true;

var CALENDAR_TAKEOVER_MEETING_DEFAULT_GUESTS = ["iva.balhar@zonky.cz", "martin.pavlik@zonky.cz", "katerina.jenikova@zonky.cz", "katerina.matkova@zonky.cz", "petr.pokorny@airbank.cz"];
var CALENDAR_REGULAR_MEETING_DEFAULT_GUESTS = ["iva.balhar@zonky.cz", "katerina.jenikova@zonky.cz", "katerina.matkova@zonky.cz", "zuzana.tokolyova@zonky.cz", "petr.pokorny@airbank.cz"];

var CALENDAR_TAKEOVER_MEETING_TIME = (8 * 60 + 40) * 60 * 1000;
var CALENDAR_REGULAR_MEETING_TIME = (8 * 60 + 50) * 60 * 1000;
var CALENDAR_TAKEOVER_MEETING_DURATION = 20 * 60 * 1000;
var CALENDAR_REGULAR_MEETING_DURATION = 10 * 60 * 1000;

var CALENDAR_TAKEOVER_MEETING_NAME = "EOD / PD switch sync";
var CALENDAR_REGULAR_MEETING_NAME = "EOD / PD daily sync";

var COLUMN_WEEK_START_DATE = 2;
var COLUMN_BE_EOD = 4;
var COLUMN_BE_PD  = 8;
var COLUMN_FE_EOD = 12;
var COLUMN_FE_PD  = 16;
var COLUMN_PO     = 20;
var COLUMN_SRE    = 24;

var COLUMN_EMAIL_ADDRESSES = 27;

var EMAIL_SUFFIX = "@zonky.cz";
var EMAIL_FIRST_NAME_LAST_NAME_SEPARATOR = ".";
var EMAILS_SEPARATOR = ",";

var STRANGE_EMAILS_MAP = [//lower case with no diacritics
  {firstName: "karel", lastName: "zelnicek", email: "karel.zelnicek2@airbank.cz"} 
];


var DIACRITICS_MAP = [
  {input: "á", output: "a"},
  {input: "č", output: "c"},
  {input: "ď", output: "d"},
  {input: "é", output: "e"},
  {input: "ě", output: "e"},
  {input: "í", output: "i"},
  {input: "ľ", output: "l"},
  {input: "ň", output: "n"},
  {input: "ó", output: "o"},
  {input: "ř", output: "r"},
  {input: "š", output: "s"},
  {input: "ť", output: "t"},
  {input: "ú", output: "u"},
  {input: "ů", output: "u"},
  {input: "ý", output: "y"},
  {input: "ž", output: "z"}
  ];

var NATIONAL_HOLIDAYS = [
    "01-01",
    "05-01",
    "05-08",
    "07-05",
    "07-06",
    "09-28",
    "10-28",
    "11-17",
    "12-24",
    "12-25",
    "12-26"
  ];

var EASTER_MONDAYS = [
    "2020-04-13",
    "2021-04-05",
    "2022-04-18",
    "2023-04-10",
    "2024-04-01",
    "2025-04-21",
    "2026-04-06",
    "2027-03-29",
    "2028-04-17",
    "2029-04-02"
  ];



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("EOD");
  
  menu.addItem("Generovat vše", "generateEverything");
  menu.addSeparator();
  menu.addItem("Generovat e-mailové adresy", "generateEmailAddresses");
  menu.addItem("Generovat události do PD kalendáře", "generatePagerDutyCalendarEvents");
  menu.addItem("Generovat status meetingy", "generateStatusMeetings");
  menu.addToUi();
}


function addDays(d, days){//Date d
  return addMiliseconds(d, days * 86400 * 1000);
}


function addMiliseconds(d, ms){//Date d
  return eliminateDST(new Date(d.getTime() + ms));
}


function addPopupReminders(event){
  event.addPopupReminder(1440 * 2.5);
  event.addPopupReminder(1440 * 4.5);
  event.addPopupReminder(1440 * 6.5); 
}


function arrayUnion(a1, a2){
  for(var i = 0; i < a2.length; i++)
    a1.push(a2[i]);
  
  return a1;
}


function eliminateDST(d){//Date d
  var hours = d.getHours();
  
  if(hours > 12)
    return new Date(d.getTime() + 3600 * (24 - hours) * 1000);
  
  if(!hours)
    return new Date(d.getTime() - 3600 * hours * 1000);
  
  return d;
}


function generateEmailAddresses(){
  var shift = getActiveShift();
  
  shift.generateEmailAddresses();
}


function generateEverything(){
  generateEmailAddresses();
  generatePagerDutyCalendarEvents();
  generateStatusMeetings();
}


function generatePagerDutyCalendarEvents(){
  var shift = getActiveShift();
  
  shift.generatePagerDutyCalendarEvents();
}


function generateStatusMeetings(){
  var shift = getActiveShift();
  
  shift.generateStatusMeetings();
}


function getActiveShift(){
  return new Shift(s.getActiveCell().getRow());
}


function getFirstWorkingDateStartingEarliestOn(d){//Date d
  for(currentDate = d; !isWorkingDay(currentDate); currentDate = addDays(currentDate, 1));
   
  return currentDate;
}


function getLongDateString(d){//Date d
  return d.getFullYear() + "-" + getShortDateString(d);
}


function getSafeDate(d){//Date d
  var tzOffset = d.getTimezoneOffset();
  
  return eliminateDST(new Date(d.getTime() + tzOffset * 60 * 1000));
}


function getShortDateString(d){//Date d
  var dd = d.getDate();
  var mm = d.getMonth() + 1;
  
  return getStringWithLeadingZero(mm) + "-" + getStringWithLeadingZero(dd);
}


function getStringWithLeadingZero(n){//int n
  return ((n < 10) ? "0" : "") + n;
}


function getTimeAtDayBySafeDate(safeDate, safeDateDiff, timeOffset){//safeDateDiff (miliseconds between safeDate and midnight in current TZ) //timeOffset (miliseconds between point in time and midnight in current TZ)
  return addMiliseconds(safeDate, timeOffset - safeDateDiff);
}


function implode(inputArray, separator, skipEmptyStrings = true){
  var result = "";
  var item = "";
  
  for(var i = 0; i < inputArray.length; i++){
    item = inputArray[i];
    
    if(item == "" && skipEmptyStrings)
        continue;
      
    if(result != "")
      result += separator;
    
    result += item;
  }
  
  return result;
}


function isHoliday(d){//Date d
  var sd = getShortDateString(d);
  var ld = getLongDateString(d);
  var ldf = getLongDateString(addDays(d, 3)); //easter friday to easter monday
  
  return (NATIONAL_HOLIDAYS.includes(sd) || EASTER_MONDAYS.includes(ld) || EASTER_MONDAYS.includes(ldf))
}


function isWeekendDay(d){//Date d
  var dow = d.getDay();
  
  return (!dow || dow == 6);
}


function isWorkingDay(d){//Date d
  return (!isWeekendDay(d) && !isHoliday(d));
}


function removeDiacritics(input){//string at lowercase is expected as $input param
  var output = "";
  var replaced = false;
  
  for(var i = 0; i < input.length; i++){
    var ch = input[i];
    if(ch >= 'a' && ch <= 'z'){
      output += ch; 
    }
    else{
      replaced = false;
      
      for(var j = 0; j < DIACRITICS_MAP.length; j++){
        if(ch == DIACRITICS_MAP[j].input){
          output += DIACRITICS_MAP[j].output;
          replaced = true;
          break;
        }
      }
      
      if(!replaced)
        output += ch;
    }
  }
  
  return output;
}


/*******************
* class Person *
*******************/
class Person{
  constructor(shift, firstName, lastName, type){
    this.shift = shift;
    this.firstName = firstName.trim();
    this.lastName = lastName.trim();
    this.type = type;
  }
  
  
  generatePagerDutyCalendarEvent(calendar){
    if(this.isNull())
      return;
    
    var shiftDateIntervals = this.generateShiftDateIntervals();
    var event;

    for(var i = 0; i < shiftDateIntervals.length; i++){
      event = calendar.createAllDayEvent(this.getCalendarEventName(), shiftDateIntervals[i].startDate, shiftDateIntervals[i].endDate, {guests: this.getEmailAddress(), sendInvites: CALENDAR_SEND_INVITES});
    
      addPopupReminders(event);
    }
  }
  
  
  generateShiftDateIntervals(){
    var shiftStartDate = this.getShiftStartDate();
    var shiftEndDate = this.getShiftEndDate();
    var shiftDateIntervals = new Array();
    
    var intervalStartDate;
    var intervalEndDate;
    var intervalStartDateSet = false;
    
    for(var currentDate = shiftStartDate; currentDate < shiftEndDate; currentDate = addDays(currentDate, 1)){
      if(this.isEOD() && isHoliday(currentDate)){
        if(intervalStartDateSet){
          intervalEndDate = currentDate;
          shiftDateIntervals.push({startDate: intervalStartDate, endDate: intervalEndDate});
          intervalStartDateSet = false;
        }
      }
      else{
        if(!intervalStartDateSet){
          intervalStartDate = currentDate;
          intervalStartDateSet = true;
        }
      }
    }
    
    if(intervalStartDateSet){
      intervalEndDate = currentDate;
      shiftDateIntervals.push({startDate: intervalStartDate, endDate: intervalEndDate});
    }
    
    return shiftDateIntervals;
  }
  
  
  getCalendarEventName(){
    return this.type + " - " + this.getName();
  }
  
  
  getEmailAddress(){
    if(this.isNull())
      return "";

    var strangeEmailAddress = this.getStrangeEmailAddress();
  
    return (strangeEmailAddress != "") ? strangeEmailAddress : (this.getFirstNameLowerCasedWithNoDiacritics() + EMAIL_FIRST_NAME_LAST_NAME_SEPARATOR + this.getLastNameLowerCasedWithNoDiacritics() + EMAIL_SUFFIX);
  }
  

  getFirstNameLowerCasedWithNoDiacritics(){
    return (this.isNull()) ? "" : removeDiacritics(this.firstName.toLowerCase());
  }
  
  
  getLastNameLowerCasedWithNoDiacritics(){
    return (this.isNull()) ? "" : removeDiacritics(this.lastName.toLowerCase());
  }


  getName(){
    return this.firstName + " " + this.lastName; 
  }
  

  getShiftDuration(){//returns length of PD / EOD shift in days (PD is 7, EOD is 5)
    return (this.isEOD()) ? 5 : 7;
  }
  
  
  getShiftEndDate(){
    var shiftEndDate;
   
    for(shiftEndDate = addDays(this.shift.weekStartSafeDate, this.getShiftDuration()); !(this.isEOD() || isWorkingDay(shiftEndDate)); shiftEndDate = addDays(shiftEndDate, 1));
    
    return shiftEndDate;
  }

  
  getShiftStartDate(){
    var shiftStartDate;
    
    for(shiftStartDate = this.shift.weekStartSafeDate; !isWorkingDay(shiftStartDate); shiftStartDate = addDays(shiftStartDate, 1));
    
    return shiftStartDate;
  }
  

  getStrangeEmailAddress(){//if person has a strange e-mail address, it's returned. "" is returned otherwise
    var firstNameLowerCasedWithNoDiacritics = this.getFirstNameLowerCasedWithNoDiacritics();
    var lastNameLowerCasedWithNoDiacritics = this.getLastNameLowerCasedWithNoDiacritics();
    
    for(var i = 0; i < STRANGE_EMAILS_MAP.length; i++){
      if(STRANGE_EMAILS_MAP[i].firstName == firstNameLowerCasedWithNoDiacritics && STRANGE_EMAILS_MAP[i].lastName == lastNameLowerCasedWithNoDiacritics)
        return STRANGE_EMAILS_MAP[i].email;
    }

    return "";
  }  
  

  isEOD(){
    return this.type.includes("EOD"); 
  }
  
  
  isNull(){
    return (this.lastName == ""); 
  }
}



/*******************
* class Shift *
*******************/
class Shift{
  constructor(row){
    this.row = row;
    this.weekStartDate = s.getRange(this.row, COLUMN_WEEK_START_DATE).getValue();
    this.weekStartSafeDate = getSafeDate(this.weekStartDate);
    this.persons = [
      new Person(this, s.getRange(this.row, COLUMN_BE_EOD).getValue(), s.getRange(this.row, COLUMN_BE_EOD + 1).getValue(), "EOD - BE"),
      new Person(this, s.getRange(this.row, COLUMN_BE_PD).getValue(), s.getRange(this.row, COLUMN_BE_PD + 1).getValue(), "PagerDuty - BE"),
      new Person(this, s.getRange(this.row, COLUMN_FE_EOD).getValue(), s.getRange(this.row, COLUMN_FE_EOD + 1).getValue(), "EOD - FE"),
      new Person(this, s.getRange(this.row, COLUMN_FE_PD).getValue(), s.getRange(this.row, COLUMN_FE_PD + 1).getValue(), "PagerDuty - FE"),
      new Person(this, s.getRange(this.row, COLUMN_PO).getValue(), s.getRange(this.row, COLUMN_PO + 1).getValue(), "PagerDuty - PO"),
      new Person(this, s.getRange(this.row, COLUMN_SRE).getValue(), s.getRange(this.row, COLUMN_SRE + 1).getValue(), "PagerDuty - SRE")
    ];
  }
  
  
  generateEmailAddresses(){
    s.getRange(this.row, COLUMN_EMAIL_ADDRESSES).setValue(this.getEmailAddressesCommaSeparatedList(CALENDAR_REGULAR_MEETING_DEFAULT_GUESTS));
  }

  
  generatePagerDutyCalendarEvents(){
    var calendar = CalendarApp.getCalendarById(CALENDAR_PAGERDUTY_ID);
    
    for(var i = 0; i < this.persons.length; i++)
      this.persons[i].generatePagerDutyCalendarEvent(calendar);
  }
  
  
  generateRegularStatusMeetings(calendar){
    var regularStatusMeetingsDates = this.generateRegularStatusMeetingsDates();

    for(var i = 0; i < regularStatusMeetingsDates.length; i++){
      var start = getTimeAtDayBySafeDate(regularStatusMeetingsDates[i], this.getSafeDateDiff(), CALENDAR_REGULAR_MEETING_TIME);
      var end = addMiliseconds(start, CALENDAR_REGULAR_MEETING_DURATION);

      var event = calendar.createEvent(
        CALENDAR_REGULAR_MEETING_NAME,
        start,
        end,{
          guests: this.getEmailAddressesCommaSeparatedList(CALENDAR_REGULAR_MEETING_DEFAULT_GUESTS),
          sendInvites: CALENDAR_SEND_INVITES
        }
      );
    }
  }
  
  
  generateRegularStatusMeetingsDates(){
    var regularMeetingsDates = new Array();
    
    for(var currentDate = getFirstWorkingDateStartingEarliestOn(addDays(getFirstWorkingDateStartingEarliestOn(this.weekStartSafeDate), 1)); !isWeekendDay(currentDate); currentDate = addDays(currentDate, 1))
      if(isWorkingDay(currentDate))
        regularMeetingsDates.push(currentDate);
    
    return regularMeetingsDates;
  }
  
  
  generateStatusMeetings(){
    var calendar = CalendarApp.getCalendarById(CALENDAR_AGILE_MEETINGS_ID);
    
    this.generateTakeoverMeeting(calendar);
    this.generateRegularStatusMeetings(calendar);
  }
  
  
  generateTakeoverMeeting(calendar){
    var meetingInterval = this.getTakeoverMeetingInterval();
    
    var event = calendar.createEvent(
      CALENDAR_TAKEOVER_MEETING_NAME,
      meetingInterval.start,
      meetingInterval.end,{
        guests: this.getEmailAddressesCommaSeparatedList(arrayUnion(CALENDAR_TAKEOVER_MEETING_DEFAULT_GUESTS, this.getEmailAddressesOfPreviousShift())),
        sendInvites: CALENDAR_SEND_INVITES
      }
    );
  }
  
  
  getEmailAddresses(){
    var emailAddresses = new Array();
    
    for(var i = 0; i < this.persons.length; i++)
      emailAddresses.push(this.persons[i].getEmailAddress());
    
    return emailAddresses;
  }
  
  
  getEmailAddressesOfPreviousShift(){
    return new Shift(this.row - 1).getEmailAddresses();
  }
  

  getEmailAddressesCommaSeparatedList(defaultAddresses){
    return implode(arrayUnion(this.getEmailAddresses(), defaultAddresses), ",");
  }

  
  getSafeDateDiff(){
    return this.weekStartSafeDate - this.weekStartDate;
  }
  

  getTakeoverMeetingInterval(){
    var firstWorkingDate = getFirstWorkingDateStartingEarliestOn(this.weekStartSafeDate);
    var start = getTimeAtDayBySafeDate(firstWorkingDate, this.getSafeDateDiff(), CALENDAR_TAKEOVER_MEETING_TIME);
    var end = addMiliseconds(start, CALENDAR_TAKEOVER_MEETING_DURATION);
    
    return {start, end};
  }

}
