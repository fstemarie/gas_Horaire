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
  let regRange, regData, event
  let employee, eventName, eventStart, eventEnd, eventId
  regRange = registry.getSheetByName(sheetName).sort(4)
  regRange = regRange.getDataRange().offset(1, 0)
  regData = regRange.getValues()
  for (let row_i = 0; row_i < regData.length; row_i++) {
    const row = regData[row_i]
    if (moment().diff(startTs, 'seconds') >= 350) break
    if (row[5] == 1) {
      Logger.log('Rangee deja traitee. On la saute')
      continue; // Si deja traite, on le saute.
    }
    if (row[0].length == 0) {
      Logger.log('Rangee vide. On la saute')
      continue // Si rangee vide, on le saute
    }
    if (row_i % 20 == 0) {
      regRange.setValues(regData)
      Logger.log('Mise a jour du registre...')
      // Utilities.sleep(1000) // Prend une pause
    }
    [employee, eventName, eventId, eventStart, eventEnd,] = row
    eventStart = moment(eventStart).toDate()
    eventEnd = moment(eventEnd).toDate()
    // eventStart = moment(eventStart).utcOffset(-5).utc().toDate()
    // eventEnd = moment(eventEnd).utcOffset(-5).utc().toDate()
    Logger.log('Creation evenement: ' + eventId)
    event = cal.createEvent(eventName, eventStart,
      eventEnd, {description: eventId})
    switch (eventId[0]) {
      case 'w':
        if (employee == 'Ste-Marie, François') {
          event.addPopupReminder(15)
          event.addPopupReminder(5)
          event.setColor(CalendarApp.EventColor.ORANGE)
        }
        else {
          event.setColor(CalendarApp.EventColor.PALE_BLUE)
        }
        break

      case 'l':
        if (employee == 'Ste-Marie, François') {
          event.addPopupReminder(15)
          event.addPopupReminder(5)
          event.setColor(CalendarApp.EventColor.MAUVE)
        }
        else {
          event.setColor(CalendarApp.EventColor.PALE_RED)
        }
        break
    }
    row[5] = 1 // Signale comme quoi il a ete traite dans le registre
    Logger.log('Pause 500ms...')
    Utilities.sleep(500)
  }
  regRange.setValues(regData)
}
