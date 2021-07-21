function fillCalendars() {
  Logger.log("-- fillCalendars()")
  for (const sheet of registry.getSheets()) {
    if (moment().diff(startTs, "seconds") >= 350) {
      break
    }
    let sheetName = sheet.getName()
    if (sheetName === "TEMPLATE") continue
    let cal = CalendarApp.getCalendarsByName(sheetName)[0]
    if (!cal) cal = CalendarApp.createCalendar(sheetName)
    fillCalendar(cal, sheet.sort(4))
  }
}

function fillCalendar(cal, sheet) {
  Logger.log("-- fillCalendar()")
  let regRange, regEvents, event, employee
  let eventName, eventStart, eventEnd, eventId
  regRange = sheet.getDataRange().offset(1, 0)
  regEvents = regRange.getValues()
  for (let row_i = 0; row_i < regEvents.length; row_i++) {
    if (moment().diff(startTs, "seconds") >= 350) {
      throw "End of allotted time reached. Exiting..."
    }
    const row = regEvents[row_i]
    if (row[5] == 1) {
      Logger.log("Rangee deja traitee. On la saute")
      continue; // Si deja traite, on le saute.
    }
    if (row[0].length == 0) {
      Logger.log("Rangee vide. On la saute")
      continue // Si rangee vide, on le saute
    }
    [employee, eventName, eventId, eventStart, eventEnd,] = row
    eventStart = moment(eventStart).toDate()
    eventEnd = moment(eventEnd).toDate()
    Logger.log("Creation evenement: " + eventId)
    event = cal.createEvent(eventName, eventStart,
      eventEnd, {description: eventId})
    switch (eventId[0]) {
      case "w":
        if (employee == "Ste-Marie, François") {
          event.addPopupReminder(15)
          event.addPopupReminder(5)
          event.setColor(CalendarApp.EventColor.ORANGE)
        }
        else {
          event.setColor(CalendarApp.EventColor.PALE_BLUE)
        }
        break

      case "l":
        if (employee == "Ste-Marie, François") {
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
    Logger.log("Mise a jour du registre...")
    regRange.setValues(regEvents)
    Logger.log("Pause 500ms...")
    Utilities.sleep(500)
  }
  regRange.setValues(regEvents)
}
