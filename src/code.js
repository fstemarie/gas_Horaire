function findNewSchedules_tt() {
  globalThis.startTs = moment()
  Logger.clear()
  try {
    findNewSchedules()
  } catch(e) {
    Logger.log(e)
  }
  globalThis.endTs = moment()
  Logger.log("Execution Time: " + endTs.diff(startTs, "seconds") + "sec")
}

function processSpreadsheets_tt() {
  globalThis.startTs = moment()
  Logger.clear()
  try {
    processSpreadsheets()
  } catch(e) {
    Logger.log(e)
  }
  globalThis.endTs = moment()
  Logger.log("Execution Time: " + endTs.diff(startTs, "seconds") + "sec")
}

function fillCalendars_tt() {
  globalThis.startTs = moment()
  Logger.clear()
  try {
  fillCalendars()
  } catch(e) {
    Logger.log(e)
  }
  globalThis.endTs = moment()
  Logger.log("Execution Time: " + endTs.diff(startTs, "seconds") + "sec")
}
