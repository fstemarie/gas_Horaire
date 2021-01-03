function test_876() {
	Logger.clear();

	prepare();
	processSheets();
}

function processSheets() {
	var folder, files, file, sheet, regRange;
	var regData = Array(), newRegData;

	Logger.log('processSheets()')
	folder = DriveApp.getFolderById(folderId);
	files = folder.getFiles();
	while (files.hasNext()) { // Pour chacun des fichiers qui...
		file = files.next();
		if (file.isStarred()) { // ... n'ont pas ete traite...
			spreadsheet = SpreadsheetApp.open(file);
			sheet = spreadsheet.getSheetByName('Schedule');
			newRegData = processSheet(sheet); // ... ouvre la spreadsheet et ajoute les infos dans regData
			regData = regData.concat(newRegData);
			file.setStarred(false); // Marque le fichier comme etant traite.
		}
	}

	// Va ajouter les donnees dans le registre
	Logger.log('Ajoute donnees au registre pour Ste-Marie, François')
	newRegData = regData.filter((r) => { return r[0] == 'Ste-Marie, François' });
	Logger.log(newRegData);
	if (newRegData.length >= 1) {
		sheet = SpreadsheetApp.openById(registryId).getSheetByName('Moi');
		sheet.getRange(sheet.getLastRow() + 1, 1, newRegData.length, 5).setValues(newRegData);
	}

	newRegData = regData.filter((r) => { return r[0] != 'Ste-Marie, François' });
	Logger.log(newRegData);
	if (newRegData.length >= 1) {
		sheet = SpreadsheetApp.openById(registryId).getSheetByName('Collegues');
		sheet.getRange(sheet.getLastRow() + 1, 1, newRegData.length, 5).setValues(newRegData);
	}
}

function processSheet(sheet, filter) {
	var row, col, start, end, lunch, workStart, workEnd, lunchStart, lunchEnd;
	var dataRange, filter, days, scheduleData, employee;
	var regData = Array();

	Logger.log('processSheet()')
	Logger.log(sheet.getParent().getName());

	//--------------------- Dans cette section, on transforme les rangees de date dans un format utilisable.
	dataRange = sheet.getDataRange();
	if (!dataRange.getFilter() == null) dataRange.getFilter().remove();
	filter = dataRange.createFilter();
	filter.setColumnFilterCriteria(1, SpreadsheetApp.newFilterCriteria().whenCellEmpty().build());
	filter.setColumnFilterCriteria(2, SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build());
	dataRange.offset(0, 1).setNumberFormat('yyyy/mm/dd');
	dataRange.getFilter().remove();
	scheduleData = dataRange.getDisplayValues();

	for (row = 0; row < scheduleData.length; row++) { // Pour chaque employe (ou dates)...
		if (scheduleData[row][1] == '') continue;

		// ------------------ La rangee se trouve a etre les dates de la semaine
		if (scheduleData[row][0] == '') {
			scheduleData[row].shift(); // On enleve la cellule vide du debut,
			days = scheduleData[row].map((x) => { return moment(x, 'YYYY/MM/DD') }); // On garde en memoire les dates transformees en momentJS
			continue; // ... et on passe a la prochaine rangee
		}
		if (scheduleData[row][0] != 'Ste-Marie, François') continue;

		// ------------------ La rangee se trouve a etre la cedule d'un employe
		employee = scheduleData[row].shift();
		Logger.log(employee);

		for (var col = 0; col < scheduleData[row].length; col++) { // Pour chaque journee de la semaine...
			var thatDay = scheduleData[row][col];

			if (thatDay == '') continue; // Si la cellule est vide, on passe a la prochaine
			if (thatDay.toLowerCase() == 'off') continue; // Si l'employe est OFF cette journee la, on passe a la prochaine
			if (thatDay.substring(0, 8).toLowerCase() == 'vacation') continue; // Si l'employe est en vacance cette journee la, on passe a la prochaine
			thatDay = thatDay.replace(/ PST/i, '').replace(/\nLUNCH : /i, ',').replace(/ - /, ','); // Debarrasse du texte inutile
			[start, end, lunch] = thatDay.split(','); //De cette facon, on se retrouve avec 3 tokens
			if (typeof lunch == 'undefined') lunch = '';

			//------------------------------------ Work ------------------------------------
			workStart = days[col].format('YYYY/MM/DD');
			switch (start.length) {
				case 4:
				case 5:
					start = moment(start, 'hh A').format('h:mm A');
					workStart = moment(workStart + ' ' + start);
					break;

				case 7:
				case 8:
					start = moment(start, 'hh:mm A').format('h:mm A');
					workStart = moment(workStart + ' ' + start);
			}

			workEnd = days[col].format('YYYY/MM/DD');
			switch (end.length) {
				case 4:
				case 5:
					end = moment(end, 'hh A').format('h:mm A');
					workEnd = moment(workEnd + ' ' + end);
					break;

				case 7:
				case 8:
					end = moment(end, 'hh:mm A').format('h:mm A');
					workEnd = moment(workEnd + ' ' + end);
			}
			if (workEnd.isBefore(workStart)) workEnd.add(1, 'day');

			//------------------------------------ Lunch ------------------------------------
			if (lunch != '' && employee == 'Ste-Marie, François') {
				lunchStart = days[col].format('YYYY/MM/DD');
				switch (lunch.length) {
					case 4:
					case 5:
						lunch = moment(lunch, 'hh A').format('h:mm A');
						lunchStart = moment(lunchStart + ' ' + lunch);
						break;

					case 7:
					case 8:
						lunch = moment(lunch, 'hh:mm A').format('h:mm A');
						lunchStart = moment(lunchStart + ' ' + lunch);
				}
				if (lunchStart.isBefore(workStart)) lunchStart.add(1, 'day');
				lunchEnd = lunchStart.clone().add(30, 'minutes');

				regData.push([employee, '<< Travail', workStart.format(), lunchStart.format(), 0]);
				regData.push([employee, '-- Lunch', lunchStart.format(), lunchEnd.format(), 0]);
				regData.push([employee, '>> Travail', lunchEnd.format(), workEnd.format(), 0]);
			}
			else {
				regData.push([employee, 'Travail', workStart.format(), workEnd.format(), 0]);
			}
		}
	}
	Logger.log(regData);
	return regData;
}
