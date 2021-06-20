const LABEL_UNPROCESSED = 'Horaire',
  LABEL_PROCESSED = 'Processed',
  FOLDER_UNPROCESSED = 'Unprocessed',
  FOLDER_PROCESSED = 'Processed',
  FOLDER_ROOT = 'Horaire'
  MYCALNAME = 'Best buy - Francois',
  MATESCALNAME = ' Best buy - Collegues',
  REGISTRYNAME = 'Horaire - Registry',
  moment = Momentjs.moment

var startTs, endTs, duration, myCal, matesCal,
  registry, rootFolder, unprocFolder, procFolder

function updateCalendars_tt() {
  startTs = moment()
  Logger.clear()
  Logger.log('-- prepareSheets_tt() ' + startTs)
  prepare()
  let msgInfos = prepareSheets()
  processSheets(msgInfos)
  fillCalendars()
  endTs = moment()
  Logger.log('Execution Time: ' + endTs.diff(startTs, 'seconds') + 'sec')
}

function updateCalendars_more_tt() {
  startTs = moment()
  Logger.clear()
  Logger.log('-- fillCalendars_tt()')
  prepare()
  fillCalendars()
  endTs = moment()
  Logger.log('Execution Time: ' + endTs.diff(startTs, 'seconds') + 'sec')
}

function cleanup() {
  Logger.log('-- cleanup()')
  prepare()

  Logger.log('Removing Processed label from threads')
  const msg1730Id = '176ee4f6c6b3bab1'
  const msg1016Id = '176c9a220894c367'
  let lbl = GmailApp.getUserLabelByName(LABEL_PROCESSED)
  let msg1730 = GmailApp.getThreadById(msg1730Id)
  msg1730.removeLabel(lbl)
  let msg1016 = GmailApp.getThreadById(msg1016Id)
  msg1016.removeLabel(lbl)

  Logger.log('Cleaning registry')
  let sheets = registry.getSheets()
  for (let sheet of sheets) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 10)
    if (sheet.getMaxRows() > 2) {
      sheet.deleteRows(2, sheet.getMaxRows() - 2)
    }
  }

  Logger.log('Cleaning up folders')
  let files = procFolder.getFiles()
  while (files.hasNext()) {
    let file = files.next()
    Drive.Files.remove(file.getId())
  }
  files = unprocFolder.getFiles()
  while (files.hasNext()) {
    let file = files.next()
    Drive.Files.remove(file.getId())
  }

  Logger.log('Removing calendars')
  let cal = CalendarApp.getCalendarsByName(MYCALNAME)[0]
  if (cal != null) {
    Calendar.Calendars.remove(cal.getId())
  }
  cal = CalendarApp.getCalendarsByName(MATESCALNAME)[0]
  if (cal != null) {
    Calendar.Calendars.remove(cal.getId())
  }
}