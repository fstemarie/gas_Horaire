function test_764() {
  Logger.clear()
  preparePhase2()
  fillCalendars()
}

function fillCalendars() {
  Logger.log('-- fillCalendars()')
  fillCalendar(myCal, 'Moi')
  fillCalendar(matesCal, 'Collegues')
}

function fillCalendar(cal, sheetName) {
  Logger.log('-- fillCalendar()')
  const MAXLIMIT = 100
  let regRange, regData, row, event, limit = 0
  let employee, eventName, eventStart, eventEnd, eventId
  regRange = SpreadsheetApp.open(DriveApp.getFileById(registryId)).
    getSheetByName(sheetName).getDataRange().offset(1, 0)
  regData = regRange.getValues()
  for (row = 1; row < regData.length; row++) {
    if (regData[row][4] == 1) continue; // Si deja traite, on le saute.
    [employee, eventName, eventId, eventStart, eventEnd,] = regData[row]
    event = cal.createEvent(eventName, eventStart.toDate(),
      eventEnd.toDate(), {'description': eventId})
    switch (eventName[0]) {
      case 'w':
        if (employee = 'Ste-Marie, François') {
          event.addPopupReminder(15)
          event.addPopupReminder(5)
        }
        event.setColor(CalendarApp.EventColor.ORANGE)
        break

      case 'l':
        if (employee = 'Ste-Marie, François') {
          event.addPopupReminder(15)
          event.addPopupReminder(5)
        }
        event.setColor(CalendarApp.EventColor.MAUVE)
        break
    }
    regData[row][4] = 1 // Signale comme quoi il a ete traite dans le registre
    if (++limit == MAXLIMIT) break
    Utilities.sleep(500)
  }
  regRange.setValues(regData)
}
