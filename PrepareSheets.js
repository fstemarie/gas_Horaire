function test_234() {
  Logger.clear()
  preparePhase1()
  let tests = prepareSheets()
  let testsJSON = JSON.stringify(tests)
  CacheService.getUserCache().put('tests', testsJSON, 7200)
}

function prepareSheets() {
  Logger.log('-- prepareSheets()')
  let threads, messages = [], results = []
  // Process Threads
  // Gets all threads that haven't been processed already
  Logger.log('Process Threads')
  threads = GmailApp.getUserLabelByName(LABEL_UNPROCESSED).getThreads()
  threads = threads.filter((th) => {
  	let labels = th.getLabels().map((lbl) => lbl.getName())
  	for (let lbl of labels) {
  		if (lbl == LABEL_PROCESSED) return false
  	}
  	return true
  })
  for (let thread of threads) {
    messages = messages.concat(thread.getMessages())
  }
  // Process Messages
  Logger.log('Process Messages')
  for (let message of messages) {
    Logger.log('Message : ' + message.getSubject())
    results = results.concat(processMessage(message))
  }
  return results
}

function processMessage(message) {
  Logger.log('-- processMessage()')
  let results = []
  for (let attachment of message.getAttachments()) {
    let result = {}
    if (attachment.getName().slice(-5) === '.xlsx') {
      result['msgDate'] = moment(message.getDate()).format()
      result['msgId'] = message.getId()
      result['threadId'] = message.getThread().getId()
      result['fileId'] = processAttachment(attachment)
      results.push(result)
    }
  }
  message.getThread().addLabel(GmailApp.getUserLabelByName(LABEL_PROCESSED))
  return results
}

function processAttachment(attachment) {
  Logger.log('-- processAttachment()')
  let resources = {
    title: attachment.getName().replace(/.xlsx?/, ""),
    parents: [{ id: unprocFolder.getId() }]
  }
  // Creates new Google Sheet from transforming original excel file
  let fileId = Drive.Files.insert(resources, attachment.copyBlob(),
    { convert: true }).id
  // Logger.log('File written and converted :' + fileId)
  // DriveApp.getFileById(fileId).setStarred(true)
  return fileId
}
