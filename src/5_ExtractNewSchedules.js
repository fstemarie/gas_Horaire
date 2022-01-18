function extractNewSchedules() {
  Logger.log("-- prepareSpreadsheets()")
  let threads, messages = []
  // Process Threads
  // Gets all threads that haven't been processed already
  Logger.log("Process Threads")
  threads = GmailApp.getUserLabelByName(LABEL_UNPROCESSED).getThreads()
  threads = threads.filter((th) => {
  	let labels = th.getLabels().map((lbl) => lbl.getName())
  	for (const lbl of labels) {
  		if (lbl == LABEL_PROCESSED) return false
  	}
  	return true
  })
  for (const thread of threads) {
    messages = messages.concat(thread.getMessages())
  }
  // Process Messages
  Logger.log("Process Messages")
  for (const message of messages) {
    Logger.log("Message : " + message.getSubject())
    processMessage(message)
  }
}

function processMessage(message) {
  Logger.log("-- processMessage()")
  // Pour chaque fichier attache au message...
  for (let attachment of message.getAttachments()) {
    let msgDate, fileId, ss
    if (attachment.getName().slice(-5) === ".xlsx") {
      msgDate = moment(message.getDate())
      fileId = processAttachment(attachment)
      ss = SpreadsheetApp.openById(fileId)
      ss.addDeveloperMetadata("gas_Horaire.msgDate", msgDate.format())
    }
  }
  message.getThread().addLabel(GmailApp.getUserLabelByName(LABEL_PROCESSED))
}

function processAttachment(attachment) {
  Logger.log("-- processAttachment()")
  let resources = {
    title: attachment.getName().replace(/.xlsx?/, ""),
    parents: [{ id: newFolder.getId() }],
    mimeType: "application/vnd.google-apps.spreadsheet"
  }
  // Creates new Google Sheet from transforming original excel file
  let fileId = Drive.Files.insert(resources, attachment.copyBlob(),
    { convert: true }).id
  return fileId
}
