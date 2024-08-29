const REGISTRYNAME = "Registre"

// Name Gmail label names
const LABEL_SCHEDULE = "Horaire"
const LABEL_PROCESSED = "Processed"
const GMAIL_QUERY = `label:${LABEL_SCHEDULE} AND NOT label:${LABEL_PROCESSED}`

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
  // Logger.log("-- prepareFolders()")
  let folders = DriveApp.getFoldersByName(FOLDER_ROOT)
  if (folders.hasNext()) {
    globalThis.rootFolder = folders.next()
  } else {
    Logger.log("Creating root folder")
    globalThis.rootFolder = DriveApp.createFolder(FOLDER_ROOT)
  }

  folders = rootFolder.getFoldersByName(FOLDER_NEW)
  if (folders.hasNext()) {
    globalThis.newFolder = folders.next()
  } else {
    Logger.log(`Creating \'${FOLDER_NEW}\' folder`)
    globalThis.newFolder = rootFolder.createFolder(FOLDER_NEW)
  }

  folders = rootFolder.getFoldersByName(FOLDER_PREPARED)
  if (folders.hasNext()) {
    globalThis.unprocFolder = folders.next()
  } else {
    Logger.log(`Creating \'${FOLDER_PREPARED}\' folder`)
    globalThis.unprocFolder = rootFolder.createFolder(FOLDER_PREPARED)
  }

  folders = rootFolder.getFoldersByName(FOLDER_PROCESSED)
  if (folders.hasNext()) {
    globalThis.procFolder = folders.next()
  } else {
    Logger.log(`Creating \'${FOLDER_PROCESSED}\' folder`)
    globalThis.procFolder = rootFolder.createFolder(FOLDER_PROCESSED)
  }
}
