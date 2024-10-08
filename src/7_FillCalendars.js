function fillCalendar() {
  // Logger.log("-- fillCalendars()")
  // row[0] = employee, row[1] = summary, row[2] = start, row[3] = end
  // row[4] = errored, row[5] = workDay, row[6] = processed, row[7] = Id
  // row[8] = timestamp
  const spreadsheets = unprocFolder.getFilesByType(MimeType.GOOGLE_SHEETS)
  while (spreadsheets.hasNext()) {
    const file = spreadsheets.next()
    const ss = SpreadsheetApp.open(file)
    const registry = ss.getSheetByName(REGISTRYNAME)
    if (registry) {
      let regRange, regEvents, calEvent
      regRange = registry.getDataRange().offset(1, 0)
      regEvents = regRange.getValues()
      for (let iRow = 0; iRow < regEvents.length; iRow++) {
        const row = regEvents[iRow]
        if (moment().diff(startTs, "seconds") >= 350) {
          throw "End of allotted time reached. Exiting..."
        }
        if (!row[0] || row[4] == 1 || row[6] == 1) continue;
        calEvent = createCalEvent(row)
        row[6] = 1 // Signale comme quoi il a ete traite dans le registre
        row[7] = calEvent.getId()
        row[8] = moment().format()
        Logger.log("Update of registry. Pause 500ms...")
        regRange.setValues(regEvents)
        Utilities.sleep(500)
      }
    }
    registry.autoResizeColumns(1, 9)
    file.moveTo(procFolder)
  }
}

function createCalEvent(row) {
  // Logger.log("-- createCalEvent()")
  let cal, calEvent
  let [employee, summary, start, end, , , id] = row
  start = moment(start)
  end = moment(end)
  cal = getCalendar(employee)
  Logger.log(`>> ${summary} | ${start.format()} | ${end.format()}`)
  calEvent = cal.createEvent(summary,
    start.toDate(), end.toDate(), { description: id })
  if (summary.startsWith("<>")) {
    calEvent.setColor(CalendarApp.EventColor.BLUE)
  } else {
    calEvent.setColor(CalendarApp.EventColor.BLUE)
  }
  if (employee == "Ste-Marie, François") {
    calEvent.addPopupReminder(15)
    calEvent.addPopupReminder(5)
    if (summary.startsWith("<>")) {
      calEvent.setColor(CalendarApp.EventColor.ORANGE)
    } else {
      calEvent.setColor(CalendarApp.EventColor.ORANGE)
    }
  }
  return calEvent
}

function getCalendar(employee) {
  // Logger.log("-- getCalendar()")
  let cal, calName, color
  if (employee == "Ste-Marie, François") {
    calName = "gs-ste-marie-francois"
    color = CalendarApp.Color.ORANGE
  } else {
    calName = "gs-collegues"
    color = CalendarApp.Color.BLUE
  }
  cal = CalendarApp.getCalendarsByName(calName)[0]
  if (!cal) {
    cal = CalendarApp.createCalendar(calName)
    cal.setColor(color)
    Logger.log("Calendar creation. Pause 1s...")
    Utilities.sleep(1000)
  }
  return cal
}
