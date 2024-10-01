function extractNewSchedules() {
  // Logger.log("-- extractNewSchedules()")
  // Process Threads
  // Gets all threads that haven't been processed already
  Logger.log("Looking for new messages")
  let messages = [], threads = GmailApp.search(GMAIL_QUERY)
  for (const thread of threads) {
    messages = messages.concat(thread.getMessages())
  }
  // Process Messages
  if (messages.length == 0) Logger.log("No new messages")
  for (const message of messages) {
    processMessage(message)
  }
}

function processMessage(message) {
  // Logger.log("-- processMessage()")
  // Pour chaque fichier attache au message...
  const subject = message.getSubject()
  const strMsgDate = moment(message.getDate()).format()
  Logger.log(`Processing new message : ${subject} | ${strMsgDate}`)
  for (let attachment of message.getAttachments()) {
    if (attachment.getName().slice(-5) === ".xlsx") {
      const fileId = processAttachment(attachment)
      const ss = SpreadsheetApp.openById(fileId)
      ss.addDeveloperMetadata("gas_Horaire.msgDate", strMsgDate)
    }
  }
  message.getThread().addLabel(GmailApp.getUserLabelByName(LABEL_PROCESSED))
}

function processAttachment(attachment) {
  // Logger.log("-- processAttachment()")
  let resources = {
    title: attachment.getName().replace(/.xlsx?/, ""),
    parents: [{ id: newFolder.getId() }],
    mimeType: MimeType.GOOGLE_SHEETS
  }
  // Creates new Google Sheet from transforming original excel file
  Logger.log(`Extracting Attachment: ${attachment.getName()}`)
  let fileId = Drive.Files.insert(resources, attachment.copyBlob(),
    { convert: true }).id
  return fileId
}
