function install() {
  Logger.log('-- install()')
  ScriptApp.newTrigger('updateCalendars_tt').timeBased().everyHours(1).create()
  // ScriptApp.newTrigger('updateCalendars_tt').timeBased()
  //   .atHour(17).nearMinute(1).everyDays(1).create()
  // ScriptApp.newTrigger('updateCalendars_more_tt').timeBased()
  //   .atHour(17).nearMinute(05).everyDays(1).create()
  // ScriptApp.newTrigger('updateCalendars_more_tt').timeBased()
  //   .atHour(17).nearMinute(15).everyDays(1).create()
}

function unInstall() {
  Logger.log('-- unInstall()')
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    ScriptApp.deleteTrigger(trigger)
  })
}
