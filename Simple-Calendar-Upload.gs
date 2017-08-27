var SS = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
var UI = SpreadsheetApp.getUi();

function onInstall(e) {
  onOpen(e);
  UI.alert("Welcome to Simple Calendar Upload!\n\n",
           "There are sample lines on this sheet already. Use them as examples for different types of GCal events. "+
           "Once an event is processed successfully, the \"Event Created\" column will be populated. "+
           "Events will only be processed if this column is blank.\n\n"+
           "To select which calendar you will export events to, go to the Addons menu, Simple Calendar Upload, and click Load Calendars.\n\n"+
           "To export all new events, go to the Addons menu, Simple Calendar Upload, and click Export.",
          UI.ButtonSet.OK);
}

function onOpen(e) {
  UI.createAddonMenu()
       .addItem("Load Calendars", "getCalendars")
       .addItem("Export", "createEvents")
       .addItem("Help", "onHelp")
       .addToUi();
}

function onHelp() {
  UI.alert("There are sample lines on this sheet already. Use them as examples for different types of GCal events. "+
           "Once an event is processed successfully, the \"Event Created\" column will be populated. "+
           "Events will only be processed if this column is blank.\n\n"+
           "To select which calendar you will export events to, go to the Addons menu, Simple Calendar Upload, and click Load Calendars.\n\n"+
           "To export all new events, go to the Addons menu, Simple Calendar Upload, and click Export.");
}

function CalendarEvent(name, start, end, notes) {
  this.name = name;
  this.start = start;
  this.end = end;
  this.notes = notes;
  if (this.end == "") {
    this.allDay = 1;
  } else {
    this.allDay = 0;
  }
}

function getCalendars() {
  var cals = CalendarApp.getAllOwnedCalendars();
  var calList = [];
  for (var i in cals) {
    calList.push(cals[i].getName()+","+cals[i].getId());
  }
  var rule = SpreadsheetApp.newDataValidation()
                             .requireValueInList(calList).setAllowInvalid(false).build();
  SS.getRange(1, 2).setDataValidation(rule).setValue(calList[0]);
}

function createEvents() {
  var calSelect = SS.getRange(1, 2).getValue().split(",")[1];
  var cal = CalendarApp.getCalendarById(calSelect);
  var range = SS.getRange(4, 1, SS.getDataRange().getNumRows()-3,5).getValues();
  var allRows = [];
  for (var row in range) {
    if (range[row][4] == "") {
      allRows.push(new CalendarEvent(range[row][0], range[row][1], range[row][2], range[row][3]));
    }
  }
  for (var i = 0; i < allRows.length; i++) {
    try {
      if (allRows[i].allDay == 1) {
        cal.createAllDayEvent(allRows[i].name, allRows[i].start, {"description":allRows[i].notes});
      } else {
        cal.createEvent(allRows[i].name, allRows[i].start, allRows[i].end, {"description":allRows[i].notes});
      }
      SS.getRange(i+4, 5).setValue("Added " + new Date());
    }
    catch (err) {
      SS.getRange(i+4, 5).setValue("ERROR: "+err+"\nCheck data, delete this message, then try again.");
    }
  }
}
