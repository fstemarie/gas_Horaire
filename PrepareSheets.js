function test_234() {
	prepare();
	prepareSheets();
}

function prepareSheets() {
	let threads, message, attachment, sheet;
	let attachments = new Array();
	let messages = new Array();

	Logger.log('-- prepareSheets()');
	// Process Threads
	// Gets all threads that haven't been processed already
	Logger.log('Process Threads');
	threads = GmailApp.getUserLabelByName(lblUnprocessed).getThreads().filter((th) => {
		let labels = th.getLabels().map((lbl) => lbl.getName());
		for (let lbl of labels) {
			if (lbl == lblProcessed) return false;
		}
		return true;
	});
	for (let thread of threads) {
		for (let message of thread.getMessages()) {
			messages.push(message);
		}
	}

	// Process Messages
	Logger.log('Process Messages');
	for (let message of messages) {
		Logger.log('Message : ' + message.getSubject());
		attachments = attachments.concat(processMessage(message));
	}

	// Process Attachments
	Logger.log('Process Attachments');
	for (let attachment of attachments) {
		Logger.log('Attachment: ' + attachment.getName());
		processAttachment(attachment);
	}
}

function processMessage(message) {
	let attachments = new Array();

	Logger.log('-- processMessage()');
	for (let attachment of message.getAttachments()) {
		if (attachment.getName().slice(-5) === '.xlsx') {
			attachments.push(attachment);
		}
	}
	message.unstar().markRead();
	message.getThread().addLabel(GmailApp.getUserLabelByName(lblProcessed));
	return attachments;
}

function processAttachment(attachment) {
	let fileId, spreadsheet, resources;

	Logger.log('-- processAttachment()');
	resources = {
		title: attachment.getName().replace(/.xlsx?/, ""),
		parents: [{ id: folderId }]
	}
	fileId = Drive.Files.insert(resources, attachment.copyBlob(), { convert: true }).id; // Creates new Google Sheet from transforming original excel file
	Logger.log(fileId);
	DriveApp.getFileById(fileId).setStarred(true);
}
