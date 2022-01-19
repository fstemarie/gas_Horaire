function processSpreadsheets() {
  // Logger.log("-- processSpreadsheets()")
  const spreadsheets = newFolder.getFilesByType(MimeType.GOOGLE_SHEETS)
  while (spreadsheets.hasNext()) {
    const file = spreadsheets.next()
    const ss = SpreadsheetApp.open(file)
    processSpreadsheet(ss)
    file.moveTo(unprocFolder)
  }
}

function processSpreadsheet(ss) {
  // Logger.log("-- processSpreadsheet()")
  const sheet = ss.getSheets()[0]
  const schedule = sheet.getDataRange().getDisplayValues()
  const msgDate = moment(ss.createDeveloperMetadataFinder()
    .withKey("gas_Horaire.msgDate").find()[0].getValue())
  let dates, regEvents = [], registry = ss.getSheetByName(REGISTRYNAME)

  if (registry != null) ss.deleteSheet(registry)
  registry = createRegistry(ss);
  // Pour chaque rangee/employe/date de la cedule...
  // Chaque rangee est soit la cedule d'un employe ou les dates de la semaine
  for (let iRow = 0; iRow < schedule.length; iRow++) {
    let row = schedule[iRow]
    // Si la rangee est vide
    if (row[1].length == 0) continue
    if (row[0].length == 0) {
      // La rangee se trouve a etre les dates de la semaine
      row.shift() // On enleve la cellule vide du debut
      // On garde en memoire les dates transformees en Moment
      dates = fixIncompleteDates(row, msgDate)
      continue
    }
    if (!dates) continue
    // La rangee se trouve a etre la cedule d'un employe
    let employee = row.shift()
    Logger.log(`Employee: ${employee} ****************************************`)
    regEvents = regEvents.concat(transformWeekSchedule(employee, row, dates))
    if (regEvents.length > 0) {
      registry.getRange(2, 1, regEvents.length, 9).setValues(regEvents)
      registry.sort(3).sort(1, false).autoResizeColumns(1, 9)
    }
  }
}

function transformWeekSchedule(employee, row, dates) {
  // Logger.log("-- transformWeekSchedule()")
  let regEvents = []
  // Pour chaque journee de la semaine/colonne...
  for (let iCol = 0; iCol < row.length; iCol++) {
    let start, end, lunch, workStart, workFinish, lunchStart, lunchEnd,
      summary, workDay = row[iCol], errored = 0
    // Si la cellule est vide, on saute
    if (workDay.length == 0) continue
    // Si la cellule est mal formattee, on saute
    if (/[a-z]/i.test(workDay[0])) continue
    workDay = applyReplacements(workDay)
      // De cette facon, on se retrouve avec 3 tokens
      ;[start, end, lunch] = workDay.split("|")

    //----------------------- Work ---------------------------------
    workStart = dates[iCol] + " " + start
    if (moment(workStart, "YYYY-MM-DD hh:mmA").isValid()) {
      workStart = moment.tz(workStart, "YYYY-MM-DD hh:mmA",
        "America/Vancouver").utc()
    } else errored = 1

    workFinish = dates[iCol] + " " + end
    if (moment(workFinish, "YYYY-MM-DD hh:mmA").isValid()) {
      workFinish = moment.tz(workFinish, "YYYY-MM-DD hh:mmA",
        "America/Vancouver").utc()
      if (workFinish.isBefore(workStart)) {
        workFinish = workFinish.add(1, "day")
      }
    } else errored = 1

    summary = "<> " + employee
    if (errored == 0) {
      regEvents.push([employee, summary, workStart.format(),
        workFinish.format(), 0, workDay, 0, null, null])
    } else {
      regEvents.push([employee, summary, null, null,
        1, workDay, 0, null, null])
    }

    //---------------------- Lunch ---------------------------------
    errored = 0
    if (typeof lunch != "undefined") {
      lunchStart = dates[iCol] + " " + lunch
      if (moment(lunchStart, "YYYY-MM-DD hh:mmA").isValid()) {
        lunchStart = moment.tz(lunchStart, "YYYY-MM-DD hh:mmA",
          "America/Vancouver").utc()
        if (lunchStart.isBefore(workStart)) lunchStart.add(1, "day")
        lunchEnd = lunchStart.clone().add(30, "minutes")
      } else {
        errored = 1
      }

      summary = "-- " + employee
      if (errored == 0) {
        regEvents.push([employee, summary, lunchStart.format(),
          lunchEnd.format(), 0, workDay, 0, null, null])
      } else {
        regEvents.push([employee, summary, null, null,
          1, workDay, 0, null, null])
      }
    }
  }
  return regEvents
}

function fixIncompleteDates(incompleteDates, msgDate) {
  // Logger.log("-- fixIncompleteDates()")
  const dates = incompleteDates.map((d) => {
    let year = msgDate.year()
    let [, month, day] = d.split(" ")
    month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
      'Sep', 'Oct', 'Nov', 'Dec'].indexOf(month) + 1
    day = parseInt(day)
    let newDate = [year, month, day]
    if (moment(newDate) < msgDate) {
      newDate = [year + 1, month, day]
    }
    return newDate.join("-")
  })
  return dates
}

function applyReplacements(workDay) {
  // Logger.log("-- applyReplacements()")
  let newWorkDay = workDay.replace(/\r\n/, " ")
    .replace(" PST", "").replace(/NO\s*LUNCH/i, "").replace(" - ", "|")
    .replace(/LUNCH\s*:/i, "|").replace(/\s*/g, "").replace(/\|$/, "")
  Logger.log(`workDay before/after replacements: ${workDay} -> ${newWorkDay}`)
  return newWorkDay
}

function createRegistry(ss) {
  // Logger.log("-- createRegistry()")
  Logger.log("Creating new registry")
  let sheet = ss.insertSheet(REGISTRYNAME)
  sheet.appendRow([
    "Employé", "Sommaire", "Début", "Fin", "Erreur",
    "Original", "Traité", "Id", "Horodate Execution"
  ])
  sheet.getRange("A1:I1").setFontSize(14).setFontWeight("bold")
    .setHorizontalAlignment("center").setBorder(null, null, true,
      null, null, null)
  sheet.setFrozenRows(1)
  return sheet
}
