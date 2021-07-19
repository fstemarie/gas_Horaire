function install() {
  Logger.log("-- install()")
  ScriptApp.newTrigger("updateCalendars_tt").timeBased().everyHours(1).create()
}

function unInstall() {
  Logger.log("-- unInstall()")
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    ScriptApp.deleteTrigger(trigger)
  })
}
