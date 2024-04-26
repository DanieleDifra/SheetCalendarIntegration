const calendarName = CalendarApp.getCalendarById(calendarId).getName();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Scripts')
      .addItem('Add To Calendar', 'menuItem1')
      .addItem('Delete From Calendar', 'menuItem2')
      .addToUi();
}

function menuItem1() {
  createCalendarEvent();
  SpreadsheetApp.getUi()
     .alert('Added events in calendar!');
}

function menuItem2() {
  deleteAllEvents();
  SpreadsheetApp.getUi()
     .alert('Deleted all events from calendar');
}