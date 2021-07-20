function updateCalendars_tt() {
  globalThis.startTs = moment()
  Logger.clear()
  Logger.log("-- updateCalendars_tt() " + startTs)
  let msgInfos = prepareSheets()
  processSheets(msgInfos)
  fillCalendars()
  globalThis.endTs = moment()
  Logger.log("Execution Time: " + endTs.diff(startTs, "seconds") + "sec")
}
