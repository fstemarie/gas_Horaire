function test_876() {
  Logger.clear()
  preparePhase1()
  let testsJSON = cache.get('tests')
  let tests = JSON.parse(testsJSON)
  processSheets(tests)
}

function processSheets(msgInfos) {
  Logger.log('-- processSheets()')
  let filteredRegData, regData = []
  let procFolder = DriveApp.getFolderById(procFolderId)
  for (let msgInfo of msgInfos) {
    const { msgDate, fileId } = msgInfo
    const file = DriveApp.getFileById(fileId)
    const sheet = SpreadsheetApp.open(file).getSheets()[0]
    // ... ouvre la spreadsheet et ajoute les infos dans regData
    regData = regData.concat(processSheet(sheet, msgDate))
    file.moveTo(procFolder)
    // Va ajouter les donnees dans le registre
    Logger.log('Ajoute les donnees au registre pour Ste-Marie, François')
    filteredRegData = regData.filter((r) => {
      return r[0] == 'Ste-Marie, François'
    })
    addToRegistry(filteredRegData, 'Moi')

    filteredRegData = regData.filter((r) => {
      return r[0] != 'Ste-Marie, François'
    })
    addToRegistry(filteredRegData, 'Collegues')
  }
}

function processSheet(sheet, msgDate) {
  Logger.log('-- processSheet()')
  /*	Dans cette section, on transforme les rangees
    de date dans un format utilisable. */

  let regData = [], dates
  let schedule = sheet.getDataRange().getDisplayValues()
  // Pour chaque employe (ou dates)...
  for (let row = 0; row < schedule.length; row++) {
    let employee
    if (schedule[row][1].length == 0) {
      Logger.log('Skipped empty row')
      continue
    }
    else if (schedule[row][0].length == 0) {
      // ------------- La rangee se trouve a etre les dates de la semaine
      schedule[row].shift() // On enleve la cellule vide du debut
      // On garde en memoire les dates transformees en momentJS
      dates = fixDates(schedule[row], msgDate)
    }
    else {
      // if (scheduleData[row][0] != 'Ste-Marie, François') continue
      // ------------- La rangee se trouve a etre la cedule d'un employe
      employee = schedule[row].shift()
      Logger.log(employee)

      // Pour chaque journee de la semaine...
      for (let col = 0; col < schedule[row].length; col++) {
        let start, end, lunch, workStart, workEnd,
          lunchStart, lunchEnd, workDay = schedule[row][col]

        // Si la cellule est vide, on saute
        if (workDay.length == 0) continue
        // Si la cellule est mal formattee, on saute
        if (/[a-z]/i.test(workDay[0])) continue
        // Debarrasse du texte inutile
        workDay = workDay.replace(' PST', '')
        workDay = workDay.replace(/\s*NO LUNCH.*/, '')
        workDay = workDay.replace(' - ', '|')
        workDay = workDay.replace(/\s*LUNCH : /, '|')
        // De cette facon, on se retrouve avec 3 tokens
        ;[start, end, lunch] = workDay.split('|')
        //----------------------- Work ---------------------------------
        workStart = dates[col].format('YYYY-MM-DD ')
        switch (start.length) {
          case 4: case 5:
            workStart = moment(workStart + start, 'YYYY-MM-DD hh A')
            break

          case 7: case 8:
            workStart = moment(workStart + start, 'YYYY-MM-DD hh:mm A')
            break
        }
        if (!workStart.isValid()) {
          Logger.log('Invalid Date - row=' + row +
            '; col=' + col + '; emp=' + employee)
        }
        workEnd = dates[col].format('YYYY-MM-DD ')
        switch (end.length) {
          case 4: case 5:
            workEnd = moment(workEnd + end, 'YYYY-MM-DD hh A')
            break

          case 7: case 8:
            workEnd = moment(workEnd + end, 'YYYY-MM-DD hh:mm A')
            break
        }
        if (workEnd.isBefore(workStart)) workEnd.add(1, 'day')
        if (!workEnd.isValid()) {
          Logger.log('Invalid Date - row=' + row +
            '; col=' + col + '; emp=' + employee)
        }
        if (!(workStart.isValid() && workEnd.isValid())) {
          regData.push([employee, 'Travail', wuid, 'Invalide',
            'Invalide', 0])
        }
        else {
          let wuid = 'w/' + workStart.format('X') + '/geeksquad.ca'
          regData.push([employee, 'Travail', wuid, workStart.format(),
            workEnd.format(), 0])
        }
        //---------------------- Lunch ---------------------------------
        // if (lunch.length > 0 && employee == 'Ste-Marie, François') {
        if (typeof lunch != 'undefined') {
          lunchStart = dates[col].format('YYYY-MM-DD ')
          switch (lunch.length) {
            case 4: case 5:
              lunchStart = moment(lunchStart + lunch, 'YYYY-MM-DD hh A')
              break

            case 7: case 8:
              lunchStart = moment(lunchStart + lunch, 'YYYY-MM-DD hh:mm A')
              break
          }
          if (lunchStart.isBefore(workStart)) lunchStart.add(1, 'day')
          lunchEnd = lunchStart.clone().add(30, 'minutes')
          if (!lunchStart.isValid()) {
            Logger.log('Invalid Date - row=' + row +
              '; col=' + col + '; emp=' + employee)
          }
          else {
            let luid = 'l/' + lunchStart.format('X') + '/geeksquad.ca'
            regData.push([employee, 'Lunch', luid, lunchStart.format(),
              lunchEnd.format(), 0])
          }
        }
      }
    }
  }
  return regData
}

function addToRegistry(regData, sheetName) {
  Logger.log('-- addToRegistry()')
  if (regData.length >= 1) {
    let sheet = SpreadsheetApp.openById(registryId).getSheetByName(sheetName)
    let range = sheet.getRange(sheet.getLastRow() + 1, 1, regData.length, 6)
    range.setValues(regData)
  }
}

function fixDates(brokenDates, msgDate) {
  Logger.log('-- fixDates()')
  dates = brokenDates.map((d) => {
    d = moment(d, 'ddd, MMM D')
    d.year(msgDate.year)
    if (d < msgDate) {
      d.add(1, 'year')
    }
    return d
  }) // On garde en memoire les dates transformees en momentJS
  return dates
}
