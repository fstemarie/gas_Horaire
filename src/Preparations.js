function prepare() {
  moment.tz.setDefault("America/Vancouver");
  prepareFolders()
  prepareRegistry()
  prepareCalendars()
}

function prepareFolders() {
  Logger.log("-- prepareFolders()")
  let folders = DriveApp.getFoldersByName(FOLDER_ROOT)
  if (folders.hasNext()) {
    rootFolder = folders.next()
  }
  else {
    rootFolder = DriveApp.createFolder(FOLDER_ROOT)
  }

  folders = rootFolder.getFoldersByName(FOLDER_PROCESSED)
  if (folders.hasNext()) {
    procFolder = folders.next()
  }
  else {
    procFolder = rootFolder.createFolder(FOLDER_PROCESSED)
  }

  folders = rootFolder.getFoldersByName(FOLDER_UNPROCESSED)
  if (folders.hasNext()) {
    unprocFolder = folders.next()
  }
  else {
    unprocFolder = rootFolder.createFolder(FOLDER_UNPROCESSED)
  }
}

function prepareRegistry() {
  Logger.log("-- prepareRegistry()")
  let files = rootFolder.getFilesByName(REGISTRYNAME)
  if (files.hasNext()){
    registry = SpreadsheetApp.open(files.next())
  }
  else {
    registry = createRegistry()
  }

  function createRegistry() {
    Logger.log("-- createRegistry()")
    let spreadsheet = SpreadsheetApp.create(REGISTRYNAME)
    DriveApp.getFileById(spreadsheet.getId()).moveTo(rootFolder)
    let sheet = spreadsheet.getSheets()[0]
    sheet.setName("Moi")
    sheet.appendRow(["Employé", "Évenement", "Id", "Date début",
      "Date fin", "Traité"])
    sheet.getRange("A1:F1").setFontSize(14).setFontWeight("bold")
      .setHorizontalAlignment("center").setBorder(null, null, true,
        null, null, null)
    sheet.setFrozenRows(1)
    sheet = spreadsheet.insertSheet("Collegues")
    sheet.appendRow(["Employé", "Évenement", "Id", "Date début",
      "Date fin", "Traité"])
    sheet.getRange("A1:F1").setFontSize(14).setFontWeight("bold")
      .setHorizontalAlignment("center").setBorder(null, null, true,
        null, null, null)
    sheet.setFrozenRows(1)
    return spreadsheet
  }
}

function prepareCalendars() {
  Logger.log("-- prepareCalendars()")
  myCal = CalendarApp.getOwnedCalendarsByName(MYCALNAME)[0]
  if (myCal == null) {
    myCal = CalendarApp.createCalendar(MYCALNAME)
    myCal.setColor(CalendarApp.Color.PURPLE)
    Utilities.sleep(250)
  }
  matesCal = CalendarApp.getOwnedCalendarsByName(MATESCALNAME)[0]
  if (matesCal == null) {
    matesCal = CalendarApp.createCalendar(MATESCALNAME)
    Utilities.sleep(250)
  }
}
