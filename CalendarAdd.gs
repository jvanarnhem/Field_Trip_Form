function addToCalendar(form) {
  var eventDate = form.tripdate;
  var eventTitle = form.destination;
  var eventDetails = "Application Number: " + form.appnum + "\nAdult in Charge: " + form.adultincharge + "\nComments: " + form.comments;
  
  //Get the calendar
  var cal = CalendarApp.getCalendarsByName('OFCS Field Trip Calendar')[0];//Change the calendar name
  var eventStartTime = new Date(eventDate+","+form.start);
  //End time is calculated by adding an hour in the event start time 
  var eventEndTime = new Date(eventDate+","+form.end);
  //Create the events
  Logger.log(eventTitle + " - " + eventStartTime + " - " + eventEndTime);
  cal.createEvent(eventTitle, eventStartTime,eventEndTime ,{description:eventDetails});
}
