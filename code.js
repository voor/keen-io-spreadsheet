function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{
    name: 'Send Data to Keen (Insert)',
    functionName: 'insert'
  }];
  sheet.addMenu('Keen', menuEntries);

  var helpMenuEntries = [{
    name: 'README',
    functionName: 'showInstructions'
  }, {
    name: 'Insert',
    functionName: 'insertHelp'
  }];
  sheet.addMenu('Keen Help', helpMenuEntries);
}


function showInstructions() {
  instructions = '<font face="Courier New"><h3>BEFORE YOU BEGIN:</h3>' +
    '<p>This spreadsheet is read-only, so you will need to make your own copy by going to ' +
    'File -> Make a Copy... (Note: you will need to make a copy of THIS spreadsheet, not create a new one).</p>' +
    '<p><i>To display these instructions again, go to: "Keen Help" -> "README"</i></p></font>';

  instructions_html = HtmlService.createHtmlOutput(instructions);
  instructions_html.setHeight(300);
  instructions_html.setWidth(800);

  SpreadsheetApp.getActiveSpreadsheet().show(instructions_html);
}

function insertHelp() {
  Browser.msgBox("Sends the current selection into a Keen IO project.  Make sure the first row consists of the headers, and do not select the row number column.");
}

function getProjectId() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var projectId = sheet.getRange(1, 2).getValue();
  if (projectId == "") {
    // Project id is not set, so we prompt user.
    projectId = Browser.inputBox("Enter project id:");
    sheet.getRange(1, 2).setValue(projectId);
  }
  return projectId + '';
}

function getEventCollection() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var eventCollection = sheet.getRange(2, 2).getValue();
  if (eventCollection == "") {
    // Event collection is not set, so we prompt user.
    eventCollection = Browser.inputBox("Enter event collection name:");
    sheet.getRange(2, 2).setValue(eventCollection);
  }
  return eventCollection;
}

function getWriteKey() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var writeKey = sheet.getRange(3, 2).getValue();
  if (writeKey == "") {
    // Event collection is not set, so we prompt user.
    writeKey = Browser.inputBox("Enter write key:");
    sheet.getRange(3, 2).setValue(writeKey);
  }
  return writeKey;
}

function clearOutput() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 4).setValue('');
}

function displayOutput(output) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 4).setValue(sheet.getRange(1, 4).getValue() + JSON.stringify(output) + '\n');
}


function toEvent(header, data) {
  var event = {};

  for (var i = 0; i < header.length; ++i) {
    event[header[i]] = data[i];
  }

  return event;
};

function insert() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var selection = sheet.getActiveSelection();
    var instances = selection.getValues();

    var projectId = getProjectId();
    var eventCollection = getEventCollection();
    var writeKey = getWriteKey();

    var url = 'https://api.keen.io/3.0/projects/' + projectId + '/events/' + eventCollection + '?api_key=' + writeKey;

    var options = {
      'method': 'POST',
      'contentType': 'application/json'
    };

    clearOutput();

    var events = [];
    var header = instances[0];

    for (var i = 1; i < instances.length; ++i) {
      var event = toEvent(header, instances[i]);
      events.push(event);
    }

    var user_reply = Browser.msgBox('Insert events:', JSON.stringify(events), Browser.Buttons.OK_CANCEL);
    if (user_reply == 'ok') {
      events.forEach(function(event) {
        options.payload = JSON.stringify(event);
        var response = UrlFetchApp.fetch(url, options);
        displayOutput(JSON.stringify(response));
      });

    }
  } catch (e) {
    Browser.msgBox('ERROR:' + e, Browser.Buttons.OK);
  }
}
