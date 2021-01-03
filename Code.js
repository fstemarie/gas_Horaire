const LBLUNPROCESSED = 'Horaire',
    LBLPROCESSED = 'Processed',
    FOLDERID = '1_oCEevGAQsho4RledOZHsH5OuPj7eDsz',
    REGISTRYID = '1P9iSRig6zK6OwePOBRaAhm_5vNu2PyT4S-gZn5hQjes',
    MOMENTURL = 'https://momentjs.com/downloads/moment.min.js'

var cache, lblUnprocessed, lblProcessed, registryId,
	rootFolderId, unprocFolderId, procFolderId
	

function install() {
	Logger.log('-- install()')
    ScriptApp.newTrigger('prepareSheets_tt').timeBased()
        .atHour(17).everyDays(1).create();
    ScriptApp.newTrigger('fillCalendars_tt').timeBased()
        .atHour(17).nearMinute(05).everyDays(1).create();
    ScriptApp.newTrigger('fillCalendars_tt').timeBased()
        .atHour(17).nearMinute(15).everyDays(1).create();
    ScriptApp.newTrigger('fillCalendars_tt').timeBased()
        .atHour(17).nearMinute(25).everyDays(1).create();
    ScriptApp.newTrigger('fillCalendars_tt').timeBased()
        .atHour(17).nearMinute(35).everyDays(1).create();
    ScriptApp.newTrigger('fillCalendars_tt').timeBased()
        .atHour(17).nearMinute(45).everyDays(1).create();
}

function unInstall() {
	Logger.log('-- unInstall()')
    ScriptApp.getProjectTriggers().forEach((trigger) => {
        ScriptApp.deleteTrigger(trigger);
    });
}

function prepare() {
	Logger.log('-- prepare()')
	cache = CacheService.getUserCache()
	lblUnprocessed = PropertiesService.getScriptProperties()
        .getProperty('lblUnprocessed');
    if (lblUnprocessed == null) {
        PropertiesService.getScriptProperties()
            .setProperty('lblUnprocessed', LBLUNPROCESSED);
        lblUnprocessed = LBLUNPROCESSED;
    }

    lblProcessed = PropertiesService.getScriptProperties()
        .getProperty('lblProcessed');
    if (lblUnprocessed == null) {
        PropertiesService.getScriptProperties()
            .setProperty('lblProcessed', LBLPROCESSED);
        lblProcessed = LBLPROCESSED;
    }

    let momentJS = cache.get('momentJS')
    if (momentJS == null) {
		momentJS = UrlFetchApp.fetch(momentJS).getContentText()
		cache.put('momentJS', momentJS, 7200)
	}
	eval(momentJS);
    moment.defaultFormat = "YYYY/MM/DD h:mm A";

    rootFolderId = PropertiesService.getScriptProperties().getProperty('folderId');
    if (rootFolderId == null) {
        rootFolderId = DriveApp.createFolder('Horaire').getId();
        PropertiesService.getScriptProperties()
            .setProperty('folderId', rootFolderId);
    }

    // Prepare the registry
    registryId = PropertiesService.getScriptProperties()
        .getProperty('registryId');
    if (registryId == null) {
        registryId = createRegistry();
        PropertiesService.getScriptProperties()
            .setProperty('registryId', registryId);
    }
    else {
        try {
            DriveApp.getFileById(registryId);
        }
        catch (err) {
            registryId = createRegistry();
            PropertiesService.getScriptProperties()
                .setProperty('registryId', registryId);
        }
    }
}

function createRegistry() {
    var registry, sheet;

    Logger.log('-- createRegistry()');
    registry = SpreadsheetApp.create('Horaire registre');
    sheet = registry.getSheets()[0];
    sheet.setName('Moi');
    sheet.appendRow(['Employé', 'Évenement', 'Date début',
        'Date fin', 'Traité']);
    sheet.getRange('A1:E1').setFontSize(14).setFontWeight('bold')
        .setHorizontalAlignment('center').setBorder(null, null, true,
            null, null, null);
    sheet.setFrozenRows(1);
    sheet = registry.insertSheet('Collegues');
    sheet.appendRow(['Employé', 'Évenement', 'Date début',
        'Date fin', 'Traité']);
    sheet.getRange('A1:E1').setFontSize(14).setFontWeight('bold')
        .setHorizontalAlignment('center').setBorder(null, null, true,
            null, null, null);
    sheet.setFrozenRows(1);
    return registry.getId();
}

function getCalendar(name) {
    var cals = CalendarApp.getCalendarsByName(name), cal;

	Logger.log('-- getCalendar()')
	if (cals.length == 0) {
        // Le calendrier n'existe pas, donc on le cree
        cal = CalendarApp.createCalendar(name);
        // On set la timezone a PST
        cal.setTimeZone('America/Vancouver');
        Utilities.sleep(500);
    }
    else {
        cal = cals[0];
    }
    return cal;
}
