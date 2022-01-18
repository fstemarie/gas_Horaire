function updateCalendars_tt() {
  globalThis.startTs = moment()
  Logger.clear()
  Logger.log("-- updateCalendars_tt() " + startTs)
  extractNewSchedules()
  processSpreadsheets()
  try {
    fillCalendar()
  } catch (e) {
    Logger.log(e)
  }
  globalThis.endTs = moment()
  Logger.log("Execution Time: " + endTs.diff(startTs, "seconds") + "sec")
}
