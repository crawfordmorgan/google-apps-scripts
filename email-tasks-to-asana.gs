function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createAddonMenu()
    .addItem('Info', 'info')
    .addItem('Send Tasks', 'sendTasks')
    .addToUi();
}

function info() {
  var doc = DocumentApp.getActiveDocument();
  var ui = DocumentApp.getUi();
  ui.alert("To use this tool, you need to add the Asana project's email address. Go to the Asana project you want to add tasks to, go to the menu (⌄), select 'Import' and then 'email,' and copy the email address. Paste it in this document after 'Asana email:'\n\nTo run the task compiler, go to Extensions>Send tasks to Asana>Send Tasks. You will probably need to authorize the script the first time you run it. If you see a security warning, proceed ahead.\n\nAny text following the phrase ‘TASK:’ will be sent as a task to Asana, until the next linebreak.\n\nThe task compiler will run until it hits a horizontal line, or the end of the document, whichever comes first. To avoid sending duplicate tasks to Asana, remember to add a new horizontal line between your previous meetings and your current one (go to Insert>Horizontal Line).\n\nTASK: this is an example of a task\n* TASK: so is this\n\nAlso … TASK: this is a task too\n\nTASK this is not a task\n\nTASK - neither is this")
}

function findHorizontalRule() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paras = body.getParagraphs();
    for (var i = 0; i < paras.length; i++) {
    var elem = paras[i].findElement(DocumentApp.ElementType.HORIZONTAL_RULE);
    if (elem != null) {
      return i;
    }
  }
  return -1;
    }

function sendTasks() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paras = body.getParagraphs();
  var text = "";
  var ui = DocumentApp.getUi();
  var horizontalLineIndex = findHorizontalRule();
if(horizontalLineIndex !== -1){
  for (var i = 0; i < horizontalLineIndex; i++) {
    text += paras[i].getText() + "\n";
}
}
    else {
      text = body.getText();
}


  var tasks = text.match(/TASK: (.*)/g); // match all instances of "TASK: " followed by any characters
  var emailMatch = text.match(/Asana email: (.*)/);
    if (emailMatch) {
      var email = emailMatch[1];
    } else {
      ui.alert("Email not found in text\nGo to the Asana project you want to add tasks to, go to the menu (⌄), select 'Import' and then 'email,' and copy the email address. Paste it in this document after 'Asana email:'");
      return;
    }
  
  var emailRegex = /^[a-zA-Z0-9+._-]+@[a-zA-Z0-9.-]+\.asana\.com$/;
if (emailRegex.test(emailMatch[1])) {
  var email = emailMatch[1];
} else {
  ui.alert("Invalid email address found in text.\nGo to the Asana project you want to add tasks to, go to the menu (⌄), select 'Import' and then 'email,' and copy the email address. Paste it in this document after 'Asana email:'");
  return;
}

  // remove the "TASK: " prefix from the tasks
  for (var i = 0; i < tasks.length; i++) {
    tasks[i] = tasks[i].substring(6);
  }

  // do not remove the matched text from the document
   for (var i = 0; i < tasks.length; i++) {
     GmailApp.sendEmail(email, tasks[i], "added from google docs");
   }

  var timestamp = new Date();
  
  if (horizontalLineIndex !== -1) {
    body.insertParagraph(horizontalLineIndex - 2, "Asana tasks sent: " + timestamp);
  } else {
    body.appendParagraph("Asana tasks sent: " + timestamp);
  }

  var taskList = "Task Review:\n\n";
  for (var i = 0; i < tasks.length; i++) {
    taskList += tasks[i] + "\n";
  }
  ui.alert(taskList);
}
