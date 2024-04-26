//Calendar Integration for StreetMarket
//Istruzioni: sullo spreadsheet evidenziare un range che comprenda almeno tutte le celle che andranno lette/scritte e salvarlo come "calendar"
//            La logica con la quale lo script funziona è basata sull seguente ordine delle colonne (e[0] = A, e[1] = B...)
//            e[1] : Data
//            e[2] : Ora
//            e[3] : EventId (Colonna D)
//            e[4] : Titolo
//            e[10]: Descrizione
//2024 Daniele Di Francesco

const calendarId = "de360f6f630ade439750b9513a84284b0862240c835b65dfee3afcc69db761a7@group.calendar.google.com"
const ss = SpreadsheetApp.getActiveSpreadsheet();

function createCalendarEvent() {
	let eventsRange = ss.getRangeByName("calendar");
  let events = eventsRange.getValues();
  let count = 2;
	
  // Creates an event for each item in events array   
	events.forEach(function(e){


      if(new Date(e[1]) != "Invalid Date" && e[3] == ""){
        var title = e[4]
        var description = e[10];
        
        //Apparently complex date parsing and concat
        var startDateRaw = new Date (e[1]);
        var timeDate = new Date (e[2]);
        if(timeDate == "Invalid Date"){ //If there is no time set I set a default of 10:00 AM instead of error
          var hours = 10;
          var minutes = 0;
        } else { 
          var hours = timeDate.getHours();
          var minutes = timeDate.getMinutes() - 50; //Weird fix because of weird time getting from Google Sheets
        }
        var timeString = hours + ':' + minutes + ':00';
        var year = startDateRaw.getFullYear();
        var month = startDateRaw.getMonth() + 1; // Jan is 0, dec is 11
        var day = startDateRaw.getDate();
        var dateString = '' + year + '-' + month + '-' + day;
        var startDate = new Date(dateString + ' ' + timeString);

        //Setting the end of the event one hour after the beginning
        var numberOfMillis = startDate.getTime();
        var addMillis = 60 * 60 * 1000;
        var endDate = new Date(numberOfMillis + addMillis)

        Logger.log("Creating event named " + title + " from " + startDate + " until " + endDate);
    	  var event = CalendarApp.getCalendarById(calendarId).createEvent(title,startDate,endDate, {description: description});
        Logger.log('Event successfully created with Id: ' + event.getId());

        //Writing the event ID
        eventCellRange = "D"+count;
        eventCell = ss.getRange(eventCellRange);
        eventCell.setValue(event.getId());
      }
    count++;
  })
}

function deleteAllEvents(){
    //Setting a very wide range of dates to delete all
    var startDate = new Date ('2024-01-01');
    var endDate = new Date ('2044-01-01');
    Logger.log('Deleting events starting from: ' + startDate + ' up to: ' + endDate);

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
    Logger.log('Deleted ' + allEvents.length + ' events');
}

