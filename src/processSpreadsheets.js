function processSpreadsheets() {
  const spreadsheets = newFolder.getFilesByType(MimeType.GOOGLE_SHEETS)
  while (spreadsheets.hasNext()) {
    const file = spreadsheets.next()
    const ss = SpreadsheetApp.open(file)
    processSpreadsheet(ss)
    file.moveTo(unprocFolder)
    if (moment().diff(startTs, "seconds") >= 290) {
      throw "End of allotted time reached. Exiting..."
    }
  }
}

function processSpreadsheet(ss) {
  const sheet = ss.getSheets()[0]
  const schedule = sheet.getDataRange().getDisplayValues()
  let fromDate, dates, regEvents = [], registry = ss.getSheetByName(REGISTRYNAME)

  fromDate = ss.getName().split(' ')[1]
  fromDate = moment.tz(fromDate, "America/Vancouver")
  if (registry != null) ss.deleteSheet(registry)
  registry = createRegistry(ss);
  // Pour chaque rangee/employe/date de la cedule...
  // Chaque rangee est soit la cedule d'un employe ou les dates de la semaine
  for (let iRow = 0; iRow < schedule.length; iRow++) {
    let row = schedule[iRow]
    if (row[1].length == 0) continue // Si la rangee est vide
    if (row[0].length == 0) {
      // La rangee se trouve a etre les dates de la semaine
      dates = [moment(fromDate).format('YYYY-MM-DD')]
      for (let i = 1; i <= 6; i++) {
        fromDate.add(1, 'd')
        dates.push(moment(fromDate).format('YYYY-MM-DD'))
      }
      fromDate.add(1, 'd')
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
      workFinish = moment.tz(workFinish, "YYYY-MM-DD hh:mmA", "America/Vancouver").utc()
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
        lunchStart = moment.tz(lunchStart, "YYYY-MM-DD hh:mmA", "America/Vancouver").utc()
        if (lunchStart.isBefore(workStart)) lunchStart.add(1, "day")
        if (lunchStart.isAfter(workFinish)) lunchStart.subtract(1, "day")
        if (!lunchStart.isBetween(workStart, workFinish)) errored = 1
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

function applyReplacements(workDay) {
  let newWorkDay = workDay.replace(/\r\n/, " ")
    .replace(" PST", "").replace(/NO\s*LUNCH/i, "").replace(" - ", "|")
    .replace(/LUNCH\s*:/i, "|").replace(/\s*/g, "").replace(/\|$/, "")
  Logger.log(`workDay before/after replacements: ${workDay} -> ${newWorkDay}`)
  return newWorkDay
}

function createRegistry(ss) {
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
