const REGISTRYNAME = "Registre"

// Name Gmail label names
const LABEL_UNPROCESSED = "Horaire"
const LABEL_PROCESSED = "Processed"

// Names of Google Drive folders
const FOLDER_ROOT = "Horaire"
const FOLDER_NEW = "1- New"
const FOLDER_PREPARED = "2- Prepared"
const FOLDER_PROCESSED = "3- Processed"

// Names of Google Calendars
const CAL_ME = "gs-francois"
const CAL_MATES = "gs-collegues"

! function() {
  moment.tz.setDefault("America/Vancouver");
  prepareFolders()
}()

function prepareFolders() {
  Logger.log("-- prepareFolders()")
  let folders = DriveApp.getFoldersByName(FOLDER_ROOT)
  if (folders.hasNext()) {
    globalThis.rootFolder = folders.next()
  } else {
    globalThis.rootFolder = DriveApp.createFolder(FOLDER_ROOT)
  }

  folders = rootFolder.getFoldersByName(FOLDER_NEW)
  if (folders.hasNext()) {
    globalThis.newFolder = folders.next()
  } else {
    globalThis.newFolder = rootFolder.createFolder(FOLDER_NEW)
  }

  folders = rootFolder.getFoldersByName(FOLDER_PREPARED)
  if (folders.hasNext()) {
    globalThis.unprocFolder = folders.next()
  } else {
    globalThis.unprocFolder = rootFolder.createFolder(FOLDER_PREPARED)
  }

  folders = rootFolder.getFoldersByName(FOLDER_PROCESSED)
  if (folders.hasNext()) {
    globalThis.procFolder = folders.next()
  } else {
    globalThis.procFolder = rootFolder.createFolder(FOLDER_PROCESSED)
  }
}

// function createRegistry() {
//   Logger.log("-- createRegistry()")
//   let spreadsheet = SpreadsheetApp.create(REGISTRYNAME)
//   DriveApp.getFileById(spreadsheet.getId()).moveTo(rootFolder)
//   let sheet = spreadsheet.getSheets()[0]
//   sheet.setName("Registre")
//   sheet.appendRow([
//     "Employé",
//     "Sommaire",
//     "Début",
//     "Fin",
//     "Erreur",
//     "Original",
//     "Traité",
//     "Id",
//     "Horodate Execution"
//   ])
//   sheet.getRange("A1:I1").setFontSize(14).setFontWeight("bold")
//     .setHorizontalAlignment("center").setBorder(null, null, true,
//       null, null, null)
//   sheet.setFrozenRows(1)
//   return spreadsheet
// }

// function prepareRegistry() {
//   Logger.log("-- prepareRegistry()")
//   let files = rootFolder.getFilesByName(REGISTRYNAME)
//   let registry
//   if (files.hasNext()){
//     registry = SpreadsheetApp.open(files.next())
//   } else {
//     registry = createRegistry()
//   }
//   globalThis.registry = registry
// }
