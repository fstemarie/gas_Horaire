function processSheets(msgInfos) {
  Logger.log("-- processSheets()")
  Logger.log(msgInfos)
  for (let msgInfo of msgInfos) {
    let regEvents = []
    let { msgDate, fileId } = msgInfo
    msgDate = moment(msgDate)
    const file = DriveApp.getFileById(fileId)
    const sheet = SpreadsheetApp.open(file).getSheets()[0]
    // ... ouvre la spreadsheet et ajoute les infos dans regData
    regEvents = regEvents.concat(processSheet(sheet, msgDate))
    file.moveTo(procFolder)
    // Va ajouter les donnees dans le registre
    addEventsToRegistry(regEvents)
  }
}

function processSheet(sheet, msgDate) {
  Logger.log("-- processSheet()")
  // Dans cette section, on transforme les rangees
  // de date dans un format utilisable.
  let regEvents = [], dates
  let schedule = sheet.getDataRange().getDisplayValues()
  // Pour chaque rangee de la cedule...
  for (let iRow = 0; iRow < schedule.length; iRow++) {
    let employee, row = schedule[iRow]
    // Si la rangee est vide
    if (row[1].length == 0) {
      Logger.log("Skipped empty row")
      continue
    } else if (row[0].length == 0) {
      // La rangee se trouve a etre les dates de la semaine
      row.shift() // On enleve la cellule vide du debut
      // On garde en memoire les dates transformees en Moment
      dates = fixIncompleteDates(row, msgDate)
    } else {
      // La rangee se trouve a etre la cedule d'un employe
      employee = row.shift()
      Logger.log("************************************************************")
      Logger.log("Employee: " + employee)
      // Pour chaque journee de la semaine (ou chaque colonne)...
      for (let iCol = 0; iCol < row.length; iCol++) {
        let start, end, lunch, workStart, workFinish, summary,
          lunchStart, lunchEnd, workDay = row[iCol], errored = 0
        // Si la cellule est vide, on saute
        if (workDay.length == 0) continue
        // Si la cellule est mal formattee, on saute
        if (/[a-z]/i.test(workDay[0])) continue
        workDay = applyReplacements(workDay)
          // De cette facon, on se retrouve avec 3 tokens
          ;[start, end, lunch] = workDay.split("|")

        //----------------------- Work ---------------------------------
        workStart = dates[iCol].format("YYYY-MM-DD ") + start
        Logger.log("Raw workStart: " + workStart)
        if (moment(workStart, "YYYY-MM-DD hh:mmA").isValid()) {
          workStart = moment.tz(workStart, "YYYY-MM-DD hh:mmA",
            "America/Vancouver").utc()
        } else {
          errored = 1
          Logger.log("Invalid Date - row=" + iRow +
            "; col=" + iCol + "; emp=" + employee)
        }

        workFinish = dates[iCol].format("YYYY-MM-DD ") + end
        Logger.log("Raw workEnd: [" + workFinish + "]")
        if (moment(workFinish, "YYYY-MM-DD hh:mmA").isValid()) {
          workFinish = moment.tz(workFinish, "YYYY-MM-DD hh:mmA",
            "America/Vancouver").utc()
          if (workFinish.isBefore(workStart)) {
            workFinish = workFinish.add(1, "day")
          }
        } else {
          errored = 1
          Logger.log("Invalid Date - row=" + iRow +
            "; col=" + iCol + "; emp=" + employee)
        }

        summary = "<> " + employee
        if (errored == 0) {
          regEvents.push([employee, summary, workStart.format(),
            workFinish.format(), 0, workDay])
        } else {
          regEvents.push([employee, summary,,, 1, workDay])
        }

        //---------------------- Lunch ---------------------------------
        errored = 0
        if (typeof lunch != "undefined") {
          lunchStart = dates[iCol].format("YYYY-MM-DD ") + lunch
          Logger.log("Raw lunch: " + lunchStart)
          if (moment(lunchStart, "YYYY-MM-DD hh:mmA").isValid()) {
            lunchStart = moment.tz(lunchStart, "YYYY-MM-DD hh:mmA",
              "America/Vancouver").utc()
            if (lunchStart.isBefore(workStart)) lunchStart.add(1, "day")
            lunchEnd = lunchStart.clone().add(30, "minutes")
          } else {
            errored = 1
            Logger.log("Invalid Date - row=" + iRow +
              "; col=" + iCol + "; emp=" + employee)
          }

          summary = "-- " + employee
          if (errored == 0) {
            regEvents.push([employee, summary, lunchStart.format(),
              lunchEnd.format(), 0, workDay])
            } else {
            regEvents.push([employee, summary,,, 1, workDay])
          }

        }
      }
    }
  }
  return regEvents
}

function applyReplacements(workDay) {
  Logger.log("workDay before replacements: " + workDay)
  workDay = workDay.replace(/\r\n/, " ")
  workDay = workDay.replace(" PST", "")
  workDay = workDay.replace(/NO\s*LUNCH/i, "")
  workDay = workDay.replace(" - ", "|")
  workDay = workDay.replace(/LUNCH\s*:/i, "|")
  workDay = workDay.replace(/\s*/g, "")
  workDay = workDay.replace(/\|$/, "")
  Logger.log("workDay after replacements: " + workDay)
  return workDay
}

// Create an object where each property is an employee (slug) which contains
// an Array with its events
function separateEventsByEmployees(regEvents) {
  Logger.log("-- separateEventsByEmp()")
  const regEventsByEmp = {}
  regEvents.forEach(row => {
    const employeeSlug = slugify(row[0])
    if (!regEventsByEmp[employeeSlug]) regEventsByEmp[employeeSlug] = []
    regEventsByEmp[employeeSlug].push(row)
  });
  return regEventsByEmp
}

function addEventsToRegistry(regEvents) {
  Logger.log("-- addEventsToRegistry()")
  if (regEvents.length == 0) return
  let regEventsByEmp = separateEventsByEmployees(regEvents)
  for (const employeeSlug in regEventsByEmp) {
    let sheet = registry.getSheetByName(employeeSlug)
    if (!sheet) {
      let template = registry.getSheetByName("TEMPLATE")
      sheet = registry.insertSheet(employeeSlug,
        Option = { "template": template })
    }
    const regEvents = regEventsByEmp[employeeSlug]
    const range = sheet.getRange(sheet.getLastRow() + 1, 1, regEvents.length, 6)
    range.setValues(regEvents)
    sheet.sort(4).autoResizeColumns(1, 5)
  }
}

function fixIncompleteDates(incompleteDates, msgDate) {
  Logger.log("-- fixIncompleteDates()")
  dates = incompleteDates.map((d) => {
    d = moment(d, "ddd, MMM D").year(msgDate.year())
    if (d < msgDate) {
      d.add(1, "year")
    }
    return d
  })
  return dates
}
