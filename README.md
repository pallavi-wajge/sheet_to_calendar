# sheet_to_calendar
Trace scheduling of studios and accounts
function onEdit(e) {
  if (!e) {
    Logger.log("Event object is undefined.");
    return;
  }

  // Specify the sheet name
  const sheetName = 'cal';
  
  // Get the active sheet
  const sheet = e.source.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet ${sheetName} not found.`);
    return;
  }
  
  // Get the edited range
  const range = e.range;
  const row = range.getRow();
  const column = range.getColumn();
  
  // Log the details of the edit event
  Logger.log(`Edit event detected. Sheet: ${sheet.getName()}, Row: ${row}, Column: ${column}`);
  
  // Assuming date is in column 1 (A) and event description is in column 2 (B)
  if (column === 1 || column === 2) {
    eventToCalendar(sheet, row);
  }
}

function eventToCalendar(sheet, row) {
  const dateCell = sheet.getRange(row, 1);
  const eventTitleCell = sheet.getRange(row, 2);

  if (!dateCell || !eventTitleCell) {
    Logger.log("Date or event title cell is undefined.");
    return;
  }

  const date = dateCell.getValue();
  const eventTitle = eventTitleCell.getValue();

  Logger.log(`Processing event for date: ${date}, title: ${eventTitle}`);
  
  if (date && eventTitle) {
    const calendarId = 'c_d624cc7da9b8c0761768d0bbfe4bf707961fa11b15b8fb05f7947d4212accf7f@group.calendar.google.com';
    const calendar = CalendarApp.getCalendarById(calendarId);

    if (!calendar) {
      Logger.log('Calendar not found.');
      return;
    }

    // Delete existing all-day events on the same date
    const events = calendar.getEventsForDay(date);
    events.forEach(event => {
      if (event.isAllDayEvent()) {
        event.deleteEvent();
      }
    });

    // Create new all-day event
    const event = calendar.createAllDayEvent(eventTitle, date);
    Logger.log(`All-day event created: ${event.getTitle()} on ${event.getAllDayStartDate()}`);
  }
}

function exportToSpecificCalendar() {
  const sheetName = 'cal';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet ${sheetName} not found.`);
    return;
  }
  
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  const calendarId = 'c_d624cc7da9b8c0761768d0bbfe4bf707961fa11b15b8fb05f7947d4212accf7f@group.calendar.google.com';
  const calendar = CalendarApp.getCalendarById(calendarId);
  
  if (!calendar) {
    Logger.log('Calendar not found.');
    return;
  }
  
  for (let i = 1; i < values.length; i++) {
    handleEvent(sheet, i + 1); // i+1 to map to the correct row in the sheet
  }
}

function handleEvent(sheet, row) {
  const dateCell = sheet.getRange(row, 1);
  const eventTitleCell = sheet.getRange(row, 2);

  if (!dateCell || !eventTitleCell) {
    Logger.log("Date or event title cell is undefined.");
    return;
  }

  const date = dateCell.getValue();
  const eventTitle = eventTitleCell.getValue();

  Logger.log(`Processing event for date: ${date}, title: ${eventTitle}`);
  
  if (date && eventTitle) {
    const calendarId = 'c_d624cc7da9b8c0761768d0bbfe4bf707961fa11b15b8fb05f7947d4212accf7f@group.calendar.google.com';
    const calendar = CalendarApp.getCalendarById(calendarId);

    if (!calendar) {
      Logger.log('Calendar not found.');
      return;
    }

    // Delete existing all-day events on the same date
    const events = calendar.getEventsForDay(date);
    events.forEach(event => {
      if (event.isAllDayEvent()) {
        event.deleteEvent();
      }
    });

    // Create new all-day event
    const event = calendar.createAllDayEvent(eventTitle, date);
    Logger.log(`All-day event created: ${event.getTitle()} on ${event.getAllDayStartDate()}`);
  }
}
