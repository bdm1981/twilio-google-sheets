//Custom Menu
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Twilio SMS")
    .addItem("Send Messages", "getNumbers")
  .addItem("Fetch Stats", "fetchStats")
    .addToUi();
}

function getAPIparams(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let settings = ss.getSheets()[2];
  let accountSID = settings.getRange(2, 2).getValue();
  let key = settings.getRange(3, 2).getValue();
  let secret = settings.getRange(4, 2).getValue();
  let serviceSID = settings.getRange(5, 2).getValue();

  let auth = Utilities.base64Encode(`${key}:${secret}`);
  
  return {accountSID: accountSID, auth: auth, serviceSID: serviceSID}
}
                  

function SendMessage(params) {

  let apiParams = getAPIparams();

  // Make a POST request with form data.
  var formData = {
    To: params.To,
    Body: getMessage(),
    MessagingServiceSid: apiParams.serviceSID
  };

  var options = {
    method: "post",
    headers: {
      Authorization: `Basic ${apiParams.auth}`
    },
    payload: formData
  };
  UrlFetchApp.fetch(
    `https://api.twilio.com/2010-04-01/Accounts/${apiParams.accountSID}/Messages.json`,
    options
  );
}

function buildMessages(numbers) {
  numbers.forEach(num => {
    SendMessage({
      To: num
    });
  });
}

function getNumbers() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let contacts = ss.getSheets()[0];
  let len = contacts.getLastRow();

  let phoneNumbers = contacts.getRange(2, 2, len, 2).getValues();
  let sendTo = [];

  phoneNumbers.forEach(num => {
    if (num[0] != "") {
      sendTo.push(num[0]);
    }
  });

  buildMessages(sendTo);
}

function getMessage() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let messageForm = ss.getSheets()[1];

  let message = messageForm.getRange(2, 1).getValue();
  Logger.log(message);
  return message;
}


/**
These functions support the pulling of Stats
*/

function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) 
        month = '0' + month;
    if (day.length < 2) 
        day = '0' + day;

    return [year, month, day].join('/');
}

function fetchStats() {
  const apiParams = getAPIparams();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let statsSheet = ss.getSheets()[3];
  let start = formatDate(statsSheet.getRange(2,8).getValue());
  let stop = formatDate(statsSheet.getRange(3,8).getValue());
  
  Logger.log(start)
  
  var options = {
    method: "get",
    headers: {
      Authorization: `Basic ${apiParams.auth}`
    }
  };
  let response = UrlFetchApp.fetch(
    `https://api.twilio.com/2010-04-01/Accounts/${apiParams.accountSID}/Messages.json?DateSent%3E=${start}&DateSent%3C=${stop}&PageSize=1000`,
    options
  );

  let json = response.getContentText();

  let history = JSON.parse(json);
  let messages = [];

  history.messages.forEach(msg => {
    messages.push([msg.to, msg.status, msg.error, msg.date_sent]);
  });



  let stats = ss.getSheets()[3];
stats.getRange(2,1,1001, 4).clearContent();
  stats.getRange(2, 1, messages.length, 4).setValues(messages);
}
