const MOMENTURL = 'https://momentjs.com/downloads/moment.min.js',
  LABEL_UNPROCESSED = 'Horaire',
  LABEL_PROCESSED = 'Processed',
  FOLDER_UNPROCESSED = 'Unprocessed',
  FOLDER_PROCESSED = 'Processed',
  MYCALNAME = 'Best buy - Francois - Test',
  MATESCALNAME = ' Best buy - Collegues - Test',
  REGISTRYNAME = 'Horaire - Registre'

var cache, props, registryId, unprocFolderId,
  procFolderId, myCal, matesCal


function prepareSheets_tt() {
  Logger.clear();
  Logger.log('-- prepareSheets_tt()')
  preparePhase1()
  let msgInfos = prepareSheets()
  processSheets(msgInfos)
}

function fillCalendars_tt() {
  Logger.clear();
  Logger.log('-- fillCalendars_tt()')
  preparePhase2()
}
