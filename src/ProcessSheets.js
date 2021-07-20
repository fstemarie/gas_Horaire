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
  // Pour chaque employe (ou dates)...
  for (let iRow = 0; iRow < schedule.length; iRow++) {
    let employee, row = schedule[iRow]
    if (row[1].length == 0) {
      Logger.log("Skipped empty row")
      continue
    }
    else if (row[0].length == 0) {
      // La rangee se trouve a etre les dates de la semaine
      row.shift() // On enleve la cellule vide du debut
      // On garde en memoire les dates transformees en Moment
      dates = fixDates(row, msgDate)
    }
    else {
      // La rangee se trouve a etre la cedule d'un employe
      employee = row.shift()
      Logger.log("Employee: " + employee)
      // Pour chaque journee de la semaine...
      for (let iCol = 0; iCol < row.length; iCol++) {
        let start, end, lunch, workStart, workEnd, wuid, summary,
          lunchStart, lunchEnd, workDay = row[iCol]
        // Si la cellule est vide, on saute
        if (workDay.length == 0) continue
        // Si la cellule est mal formattee, on saute
        if (/[a-z]/i.test(workDay[0])) continue
        // Debarrasse du texte inutile
        workDay = workDay.replace(" PST", "")
        workDay = workDay.replace(/NO LUNCH/, "")
        workDay = workDay.replace(" - ", "|")
        workDay = workDay.replace(/LUNCH :/, "|")
        workDay = workDay.replace(/\s*/g, "")
        workDay = workDay.replace(/\|$/, "")
        // De cette facon, on se retrouve avec 3 tokens
        ;[start, end, lunch] = workDay.split("|")

        //----------------------- Work ---------------------------------
        workStart = dates[iCol].format("YYYY-MM-DD ") + start
        switch (start.length) {
          case 3: case 4:
            workStart = moment.tz(workStart, "YYYY-MM-DD hhA",
              "America/Vancouver").utc()
            break

          case 6: case 7:
            workStart = moment.tz(workStart, "YYYY-MM-DD hh:mmA",
              "America/Vancouver").utc()
            break
        }
        if (workStart.isValid()) {
          Logger.log("workStart is Valid: " + workStart.format())
          wuid = "w/" + workStart.format("X") + "/" + slugify(employee)
        }
        else {
          Logger.log("Invalid Date - row=" + iRow +
            "; col=" + iCol + "; emp=" + employee)
        }

        workEnd = dates[iCol].format("YYYY-MM-DD ") + end
        switch (end.length) {
          case 3: case 4:
            workEnd = moment.tz(workEnd, "YYYY-MM-DD hhA",
              "America/Vancouver").utc()
            break

          case 6: case 7:
            workEnd = moment.tz(workEnd, "YYYY-MM-DD hh:mmA",
              "America/Vancouver").utc()
            break

        }
        if (workEnd.isValid()) {
          Logger.log("workEnd is Valid: " + workEnd.format())
          if (workEnd.isBefore(workStart)) workEnd = workEnd.add(1, "day")
        }
        else {
          Logger.log("Invalid Date - row=" + iRow +
            "; col=" + iCol + "; emp=" + employee)
        }

        summary = " <> " + employee
        regEvents.push([employee, summary, wuid, workStart.format(),
          workEnd.format(), 0])

        //---------------------- Lunch ---------------------------------
        if (typeof lunch != "undefined") {
          lunchStart = dates[iCol].format("YYYY-MM-DD ") + lunch
          Logger.log("lunch: " + lunchStart)
          switch (lunch.length) {
            case 3: case 4:
              lunchStart = moment.tz(lunchStart, "YYYY-MM-DD hhA",
                "America/Vancouver").utc()
              break

            case 6: case 7:
              lunchStart = moment.tz(lunchStart, "YYYY-MM-DD hh:mmA",
                "America/Vancouver").utc()
              break
            }
          if (lunchStart.isValid()) {
            Logger.log("lunchStart is Valid: " + lunchStart.format())
            if (lunchStart.isBefore(workStart)) lunchStart.add(1, "day")
            lunchEnd = lunchStart.clone().add(30, "minutes")

            let luid = "l/" + lunchStart.format("X") + "/" + slugify(employee)
            let summary = " -- " + employee
            regEvents.push([employee, summary, luid, lunchStart.format(),
              lunchEnd.format(), 0])
          }
          else {
            Logger.log("Invalid Date - row=" + iRow +
              "; col=" + iCol + "; emp=" + employee)
          }
        }
      }
    }
  }
  return regEvents
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
        Option={"template": template})
    }
    const events = regEventsByEmp[employeeSlug]
    const range = sheet.getRange(sheet.getLastRow() + 1, 1, events.length, 6)
    range.setValues(events)
    sheet.sort(4)
  }
}

function fixDates(brokenDates, msgDate) {
  Logger.log("-- fixDates()")
  dates = brokenDates.map((d) => {
    d = moment(d, "ddd, MMM D").year(msgDate.year())
    if (d < msgDate) {
      d.add(1, "year")
    }
    return d
  })
  return dates
}
