
function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Menu").addItem("Cash App", "getGmailEmails").addToUi();
}

function getGmailEmails(){

  var label = GmailApp.getUserLabelByName("Add Your Lebel Here");
  var threads = label.getThreads();
  
  for(var i = threads.length - 1; i >=0; i--){
    var messages = threads[i].getMessages();
    
    for (var j = 0; j <messages.length; j++){
      var message = messages[j];
      extractDetails(message);
    }
  }
  
}

function extractDetails(message){

  var emailData = {
    date : "",
    body : ""
  }

  emailData.date = message.getDate();
  emailData.body = message.getPlainBody();

  var regTransactionId = /(?<=Payment to|Payment from)[\s\$\d].*/;
  var transactionId = emailData.body.match(regTransactionId);
  if(!transactionId) return false;
  var  cashTag= transactionId.toString();

  var regStatus = /(received|completed)/i;
  var status = emailData.body.match(regStatus);
  if(!status) return false;
  var type = status.toString().split(",")[1];
 
  var regAmount = /^Amount[\s][\n].*/m;
  var amount = emailData.body.match(regAmount);
  if(!amount) return false;
  var dollar = amount.toString().substring(9);

  var regIdentifier = /(?<=Identifier)[\s][\n].*/m;
  var identifier = emailData.body?.match(regIdentifier);
  if(!identifier) return false;
  var hastag = identifier.toString();

  var regSender = /(?<=From)[\s][\n].*/m; 
  var sender = emailData.body.match(regSender);
  if(!sender) return false;
  var senderName = sender.toString();


  var regReciever = /(?<=To)[\s][\n].*/m; 
  var reciever = emailData.body.match(regReciever);
  if(!reciever) return false;
  var recieverName = reciever.toString();

  var regData = /(?<=Date:)[\s].*/;
  var date = emailData.body.match(regData);
  if(!date) return false;
  var date_time = date.toString();

  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  activeSheet.appendRow([cashTag, type, dollar, hastag ,senderName, recieverName, date_time]);
  removeDuplicates();
}


function removeDuplicates() {
  var sh=SpreadsheetApp.getActiveSheet();
  var dt=sh.getDataRange().getValues();
  var uA=[];
  var d=0;
  for(var i=0;i<dt.length;i++) {
    if(uA.indexOf(dt[i][3])==-1) {
      uA.push(dt[i][3]);
    }else{
      sh.deleteRow(i+1-d++);
    }
  }
}



















