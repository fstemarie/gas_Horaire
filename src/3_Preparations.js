const LABEL_UNPROCESSED = "Horaire",
  LABEL_PROCESSED = "Processed",
  FOLDER_UNPROCESSED = "Unprocessed",
  FOLDER_PROCESSED = "Processed",
  FOLDER_ROOT = "Horaire"
  MYCALNAME = "Best buy - Francois",
  MATESCALNAME = " Best buy - Collegues",
  REGISTRYNAME = "Horaire - Registry"

! function() {
  moment.tz.setDefault("America/Vancouver");
  prepareFolders()
  prepareRegistry()
}()

function prepareFolders() {
  Logger.log("-- prepareFolders()")
  let folders = DriveApp.getFoldersByName(FOLDER_ROOT)
  if (folders.hasNext()) {
    globalThis.rootFolder = folders.next()
  } else {
    globalThis.rootFolder = DriveApp.createFolder(FOLDER_ROOT)
  }

  folders = rootFolder.getFoldersByName(FOLDER_PROCESSED)
  if (folders.hasNext()) {
    globalThis.procFolder = folders.next()
  } else {
    globalThis.procFolder = rootFolder.createFolder(FOLDER_PROCESSED)
  }

  folders = rootFolder.getFoldersByName(FOLDER_UNPROCESSED)
  if (folders.hasNext()) {
    globalThis.unprocFolder = folders.next()
  } else {
    globalThis.unprocFolder = rootFolder.createFolder(FOLDER_UNPROCESSED)
  }
}

function createRegistry() {
  Logger.log("-- createRegistry()")
  let spreadsheet = SpreadsheetApp.create(REGISTRYNAME)
  DriveApp.getFileById(spreadsheet.getId()).moveTo(rootFolder)
  let sheet = spreadsheet.getSheets()[0]
  sheet.setName("TEMPLATE")
  sheet.appendRow(["Employé", "Évenement", "Id", "Date début",
    "Date fin", "Traité", "Horodate Execution"])
  sheet.getRange("A1:F1").setFontSize(14).setFontWeight("bold")
    .setHorizontalAlignment("center").setBorder(null, null, true,
      null, null, null)
  sheet.setFrozenRows(1)
  return spreadsheet
}

function prepareRegistry() {
  Logger.log("-- prepareRegistry()")
  let files = rootFolder.getFilesByName(REGISTRYNAME)
  let registry
  if (files.hasNext()){
    registry = SpreadsheetApp.open(files.next())
  } else {
    registry = createRegistry()
  }
  globalThis.registry = registry
}
