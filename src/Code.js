const LABEL_UNPROCESSED = "Horaire",
  LABEL_PROCESSED = "Processed",
  FOLDER_UNPROCESSED = "Unprocessed",
  FOLDER_PROCESSED = "Processed",
  FOLDER_ROOT = "Horaire"
  MYCALNAME = "Best buy - Francois",
  MATESCALNAME = " Best buy - Collegues",
  REGISTRYNAME = "Horaire - Registry"

var startTs, endTs, duration, myCal, matesCal,
  registry, rootFolder, unprocFolder, procFolder

function updateCalendars_tt() {
  startTs = moment()
  Logger.clear()
  Logger.log("-- prepareSheets_tt() " + startTs)
  prepare()
  let msgInfos = prepareSheets()
  processSheets(msgInfos)
  fillCalendars()
  endTs = moment()
  Logger.log("Execution Time: " + endTs.diff(startTs, "seconds") + "sec")
}
