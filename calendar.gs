function syncCalendar() {
  // Replace with your actual calendar ID
  const calendarId = "youremail@gmail.com";
  
  // Replace with the sheet name and data range
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getRange(4, 1, sheet.getLastRow() - 1, 5); // Skip header row (row 1)
  //const dataValues = dataRange.getValues();
  
  //var cal = CalendarApp.getCalendarById("trudsdata@gmail.com");
  //var events = cal.getEvents(new Date(sht.getRange("B4").getValues()) , new Date(sht.getRange("B5").getValues()));


  // Get Calendar events
  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(new Date(sheet.getRange("G4").getValues()) , new Date(sheet.getRange("G5").getValues()));
  
  // Clear existing sheet data (optional)
  dataRange.clearContent();
  
  // Sync Calendar events to Sheet
  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    dataRange.getCell(i + 1, 5).setValue(event.getTitle()); // Adjust column index based on your data
    dataRange.getCell(i + 1, 4).setValue(event.getDescription()); // Adjust column index based on your data
    dataRange.getCell(i + 1, 2).setValue(event.getEndTime().toString()); // Adjust column index based on your data
    dataRange.getCell(i + 1, 1).setValue(event.getStartTime().toString()); // Adjust column index based on your data

    //dataRange.getCell(i + 1, 3).setValue(event.getEnd().toString()); // Adjust column index based on your data
    // Add more data columns as needed (Description, etc.)

   // var eventTitle = events[eventCtr].getTitle();
    //var eventDesc = events[eventCtr].getDescription(); 
  } 
}
