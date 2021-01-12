function preparePhase1() {
  props = PropertiesService.getScriptProperties()
  cache = CacheService.getUserCache()
  prepareMoment()
  prepareFolders()
  createRegistry()
}

function preparePhase2() {
  props = PropertiesService.getScriptProperties()
  cache = CacheService.getUserCache()
  prepareMoment()
  prepareCalendars()
  fillCalendars()
}

function prepareMoment() {
  Logger.log('-- prepareMoment()')
  let momentJS = cache.get('momentJS')
  if (momentJS == null) {
    momentJS = UrlFetchApp.fetch(MOMENTURL).getContentText()
    cache.put('momentJS', momentJS, 7200)
  }
  eval(momentJS)
  moment.defaultFormat = "YYYY-MM-DD HH:mm"
}

function prepareFolders() {
  Logger.log('-- prepareFolders()')
  unprocFolderId = props.getProperty('unprocFolderId')
  if (unprocFolderId == null) {
    unprocFolderId = DriveApp.createFolder(FOLDER_UNPROCESSED).getId()
    props.setProperty('unprocFolderId', unprocFolderId)
  }
  procFolderId = props.getProperty('procFolderId')
  if (procFolderId == null) {
    procFolderId = DriveApp.createFolder(FOLDER_PROCESSED).getId()
    props.setProperty('procFolderId', procFolderId)
  }
}

function prepareRegistry() {
  Logger.log('-- prepareRegistry()')
  registryId = props.getProperty('registryId')
  if (registryId == null) {
    registryId = createRegistry()
    props.setProperty('registryId', registryId)
  }
  else {
    try {
      let registry = DriveApp.getFileById(registryId)
      Logger.log(registry.getName())
    }
    catch (err) {
      registryId = createRegistry()
      props.setProperty('registryId', registryId)
    }
  }
  function createRegistry() {
    Logger.log('-- createRegistry()')
    let registry, sheet
    registry = SpreadsheetApp.create(REGISTRYNAME)
    sheet = registry.getSheets()[0]
    sheet.setName('Moi')
    sheet.appendRow(['Employé', 'Évenement', 'Id', 'Date début',
      'Date fin', 'Traité'])
    sheet.getRange('A1:F1').setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('center').setBorder(null, null, true,
        null, null, null)
    sheet.setFrozenRows(1)
    sheet = registry.insertSheet('Collegues')
    sheet.appendRow(['Employé', 'Évenement', 'Id', 'Date début',
      'Date fin', 'Traité'])
    sheet.getRange('A1:F1').setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('center').setBorder(null, null, true,
        null, null, null)
    sheet.setFrozenRows(1)
    return registry.getId()
  }  
}

function prepareCalendars() {
  Logger.log('-- prepareCalendars()')
  myCal = CalendarApp.getOwnedCalendarsByName(MYCALNAME)[0]
  if (myCal == null) {
    myCal = CalendarApp.createCalendar(MYCALNAME)
    myCal.setTimeZone('America/Vancouver')
    Utilities.sleep(500)
  }
  matesCal = CalendarApp.getOwnedCalendarsByName(MATESCALNAME)[0]
  if (matesCal == null) {
    matesCal = CalendarApp.createCalendar(MYCALNAME)
    matesCal.setTimeZone('America/Vancouver')
    Utilities.sleep(500)
  }
}
