const calendarName = CalendarApp.getCalendarById(calendarId).getName();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Scripts')
      .addItem('Add To Calendar', 'addToCalendar')
      .addSubMenu(ui.createMenu('Delete From Calendar')
        .addItem('7 days', 'delete7days')
        .addItem('15 days', 'delete15days')
        .addItem('30 days', 'delete30days')
        .addItem('All', 'deleteFromCalendar'))
      .addToUi();
}

function addToCalendar() {
  createCalendarEvent();
  SpreadsheetApp.getUi()
     .alert('Added events in calendar!');
}

function delete7days() {
  deleteDaysEvents(7);
  SpreadsheetApp.getUi()
     .alert('Deleted 7 days of events from calendar');
}

function delete15days() {
  deleteDaysEvents(15);
  SpreadsheetApp.getUi()
     .alert('Deleted all events from calendar');
}

function delete30days() {
  deleteDaysEvents(30);
  SpreadsheetApp.getUi()
     .alert('Deleted all events from calendar');
}

function deleteFromCalendar() {
  deleteAllEvents();
  SpreadsheetApp.getUi()
     .alert('Deleted all events from calendar');
}
