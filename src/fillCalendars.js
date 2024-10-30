  // row[0] = employee
  // row[1] = summary
  // row[2] = start
  // row[3] = end
  // row[4] = errored
  // row[5] = workDay
  // row[6] = processed
  // row[7] = Id
  // row[8] = timestamp

function fillCalendars() {
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
        if (!row[0] || row[4] == 1 || row[6] == 1) continue;
        calEvent = createCalEvent(row)
        row[6] = 1 // Signale comme quoi il a ete traite dans le registre
        row[7] = calEvent.getId()
        row[8] = moment().format()
        regRange.setValues(regEvents)
        Logger.log("Update of registry. Pause 500ms...")
        Utilities.sleep(500)
      }
    }
    registry.autoResizeColumns(1, 9)
    file.moveTo(procFolder)
    if (moment().diff(startTs, "seconds") >= 290) {
      throw "End of allotted time reached. Exiting..."
    }
  }
}

function createCalEvent(row) {
  let cal, calEvent
  let [employee, summary, start, end, , , id] = row
  start = moment(start)
  end = moment(end)
  cal = getCalendar(employee)
  Logger.log(`>> ${summary} | ${start.format()} | ${end.format()}`)
  calEvent = cal.createEvent(summary,
    start.toDate(), end.toDate(), { description: id })
  // if (employee == "Ste-Marie, François") {
  //   calEvent.addPopupReminder(15)
  //   calEvent.addPopupReminder(5)
  // }
  return calEvent
}

function getCalendar(employee) {
  let calName = slugify(employee)
  if (employee == "Ste-Marie, François") {
    calName = "gs-ste-marie-francois"
  } else {
    calName = "gs-collegues"
  }
  let cal = CalendarApp.getCalendarsByName(calName)[0]
  if (!cal) {
    cal = CalendarApp.createCalendar(calName)
    Logger.log("Calendar creation. Pause 1s...")
    Utilities.sleep(1000)
  }
  return cal
}
