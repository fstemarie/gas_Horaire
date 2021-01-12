function install() {
  Logger.log('-- install()')
  ScriptApp.newTrigger('prepareSheets_tt').timeBased()
    .atHour(17).everyDays(1).create()
  ScriptApp.newTrigger('fillCalendars_tt').timeBased()
    .atHour(17).nearMinute(05).everyDays(1).create()
  ScriptApp.newTrigger('fillCalendars_tt').timeBased()
    .atHour(17).nearMinute(15).everyDays(1).create()
  ScriptApp.newTrigger('fillCalendars_tt').timeBased()
    .atHour(17).nearMinute(25).everyDays(1).create()
  ScriptApp.newTrigger('fillCalendars_tt').timeBased()
    .atHour(17).nearMinute(35).everyDays(1).create()
  ScriptApp.newTrigger('fillCalendars_tt').timeBased()
    .atHour(17).nearMinute(45).everyDays(1).create()
}

function unInstall() {
  Logger.log('-- unInstall()')
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    ScriptApp.deleteTrigger(trigger)
  })
}
