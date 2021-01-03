function test_764() {
	Logger.clear();

	prepare();
	fillCalendars();
}

function fillCalendars() {
	const MAXLIMIT = 40;
	let regRange, regData, filter, row, cal, event, limit;
	let employee, eventName, eventStart, eventEnd;

	Logger.log('registryId: ' + registryId);
	regRange = SpreadsheetApp.open(DriveApp.getFileById(registryId)).getDataRange();
	regRange.offset(1, 0);
	regData = regRange.getValues();

	limit = 0;
	for (row = 1; row < regData.length; row++) {
		if (regData[row][4] == 1) continue; // Si deja traite, on le saute.

		employee = regData[row][0];
		eventName = regData[row][1];
		eventStart = moment(regData[row][2]);
		eventEnd = moment(regData[row][3]);

		switch (employee) {
			case 'Ste-Marie, François':
				cal = getCalendar('Best Buy - Francois');

				/*
				cal.getEventsForDay(eventStart.toDate()).forEach(function (ev){
				  ev.deleteEvent();
				  Utilities.sleep(500);
				});
				*/

				event = cal.createEvent(eventName, eventStart.toDate(), eventEnd.toDate());
				Utilities.sleep(1000);
				switch (eventName) {
					case '<< Travail':
						event.addPopupReminder(15);
						event.addPopupReminder(5);
						event.setColor(CalendarApp.EventColor.ORANGE)
						break;

					case '>> Travail':
						event.addPopupReminder(15);
						event.addPopupReminder(5);
						event.setColor(CalendarApp.EventColor.ORANGE)
						break;

					case '-- Lunch':
						event.addPopupReminder(10);
						event.addPopupReminder(5);
						event.setColor(CalendarApp.EventColor.MAUVE)
						break;
				}
				regData[row][4] = 1;
				break;

			default:
				cal = getCalendar('Best Buy - Collegues');
				cal.setColor(CalendarApp.Color.INDIGO);

				event = cal.createEvent(employee, eventStart.toDate(), eventEnd.toDate());
				Utilities.sleep(500);
				event.setColor(CalendarApp.EventColor.MAUVE);
				regData[row][4] = 1;
				break;
		}
		limit++;
		if (limit == MAXLIMIT) break;
	}
	regRange.setValues(regData);
}