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
  
  var emailRegex = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;
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
/*
  // do not remove the matched text from the document
   for (var i = 0; i < tasks.length; i++) {
     GmailApp.sendEmail(email, tasks[i], "added from google docs");
   }
  */
  var taskList = "Task Review:\n\n";
  for (var i = 0; i < tasks.length; i++) {
    taskList += tasks[i] + "\n";
  }
  ui.alert(taskList);
}
