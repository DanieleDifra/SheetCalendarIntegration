//Calendar Integration ADV for StreetMarket
//Istruzioni: sullo spreadsheet evidenziare un range che comprenda almeno tutte le celle che andranno lette/scritte e salvarlo come "calendar"
//            La logica con la quale lo script funziona è basata sull seguente ordine delle colonne (e[0] = A, e[1] = B...)
//            e[1] : Data inizio
//            e[2] : Data fine
//            e[3] : EventId (Colonna D)
//            e[4] : Status
//            e[5] : Titolo
//            e[7] : Descrizione
//2024 Daniele Di Francesco

const calendarId = "de360f6f630ade439750b9513a84284b0862240c835b65dfee3afcc69db761a7@group.calendar.google.com"; //test calendar
//const calendarId = "cf031a28b14ffdb39f9b75a677fdc521621b5f86b17ca0ee899d51fa51545111@group.calendar.google.com "; //SM calendar
const ss = SpreadsheetApp.getActiveSpreadsheet();

function createCalendarEvents() {
	let eventsRange = ss.getRangeByName("calendar");
  let events = eventsRange.getValues();
  let ttn = Date.now();
  let today = new Date(ttn);
  let count = 2;

  events.forEach(function(e){
    var title = "ADV - " + e[5];
    var startDate = new Date(e[1]);
    var endDate = new Date(e[2]);
    var status = e[4];
    var description = e[7];

    var eventCellRange = "D" + count;
    var eventCell = ss.getRange(eventCellRange);
    var statusCellRange = "E" + count;
    
    if(status != "Saltato"){
      var statusCell = ss.getRange(statusCellRange);
      if(startDate.getTime() < today.getTime() && today.getTime() < endDate.getTime()){
        statusCell.setValue("Live");
      } else if(today.getTime() > endDate.getTime()) {
        statusCell.setValue("Concluso");
      }

      if(new Date(e[1]) != "Invalid Date" && e[3] == ""){
        var event = CalendarApp.getCalendarById(calendarId).createAllDayEvent(title,startDate,endDate, {description: description});
        eventId = event.getId();
        eventCell.setValue(eventId); //Writing the event ID
      }
    }
    count++;
  }) 
}

function deleteEvents(startDate, endDate){
  var replace_with = ""; //Leave blank to delete Text
  var lastRow = ss.getLastRow();
  var ranges = ['D2:D' + lastRow];

  //Cycle through all events and delete them and delete the eventId in the sheet
  var allEvents = CalendarApp.getCalendarById(calendarId).getEvents(startDate,endDate);
  for (var e=0;e<allEvents.length;e++){
    var to_replace = allEvents[e].getId();
    allEvents[e].deleteEvent();
    ss.getRangeList(ranges).getRanges().forEach(r => r.createTextFinder(to_replace).matchEntireCell(true).replaceAllWith(replace_with));
  } 
  //Logger.log('Deleted ' + allEvents.length + ' events');
}

function deleteDaysEvents(days){
  var ttn = Date.now();
  var startDate = new Date(ttn);
  var daysInMillis = days * 24 * 60 * 60 * 1000;
  var result = startDate.getTime() + daysInMillis;
  var endDate = new Date(result);
  //Logger.log("startDate: " + startDate);
  //Logger.log("endDate: " + endDate);
  deleteEvents(startDate,endDate);
}

function deleteAllEvents(){
    //Setting a very wide range of dates to delete all
    var startDate = new Date ('2024-01-01');
    var endDate = new Date ('2044-01-01');
    //Logger.log('Deleting events starting from: ' + startDate + ' up to: ' + endDate);
    deleteEvents(startDate,endDate);
}