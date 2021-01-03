function prepareSheets_tt() {
	Logger.clear();
	Logger.log('prepareSheets_tt');

	prepare();
	prepareSheets();
	processSheets();
}

function fillCalendars_tt() {
	Logger.clear();
	Logger.log('fillCalendars_tt');

	prepare();
	fillCalendars();
}
