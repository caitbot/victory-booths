/**
 * Use this to do one-off operations.
 */
function __manualInit() {
  // Store the ID for the Calendar, which is needed to retrieve events by ID.
  var cal = CalendarApp.getCalendarsByName("Victory Booths @ Manny's")[0];
  var calID = cal.getId(); //.split("@")[0];
  ScriptProperties.setProperty('calId', calID);
  Logger.log('The calendar ID is "%s".', calID);
  Logger.log('The calendar ID is "%s".', ScriptProperties.getProperty('calId'));
}


// setup trigger
function __manualInit2() {
  var ss = SpreadsheetApp.openById('1SvZpIMxI0QvL0BSV-QDxU9Et5-Y9_BC7U_oLJtYjFMQ');
  ScriptApp.newTrigger('_onFormSubmit').forSpreadsheet(ss).onFormSubmit()
      .create();
}


// test cal events
function testCal() {

}


// Twilio test
function twilioTest() {
  sendSms('2016610071', "This is a test via GoogleApps");
}

function sendSms(to, body) {
  var messages_url = "https://api.twilio.com/2010-04-01/Accounts/AC8165e299f98be421d561984e13dfa51f/Messages.json";

  var payload = {
    "To": to,
    "Body" : body,
    "From" : "+12055742251‬"
  };

  var options = {
    "method" : "post",
    "payload" : payload
  };

  options.headers = {
    "Authorization" : "Basic " + Utilities.base64Encode("AC8165e299f98be421d561984e13dfa51f:94d62ecc50254bfd8b92049b568eef2d")
  };

  UrlFetchApp.fetch(messages_url, options);
}

function sendAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var numRows = sheet.getLastRow() - 1;
  var dataRange = sheet.getRange(startRow, 1, numRows, 2)
  var data = dataRange.getValues();

  for (i in data) {
    var row = data[i];
    try {
      response_data = sendSms(row[0], row[1]);
      status = "sent";
    } catch(err) {
      Logger.log(err);
      status = "error";
    }
    sheet.getRange(startRow + Number(i), 3).setValue(status);
  }
}



// test
function test1() {
//  Logger.log(['a', 'b', 'c'].indexOf('d'));
  var string = "sub some stuff into here: {NAME}, then say hi!";
  Logger.log(subStringIntoTag("{NAME}", "Noam", string));
}


// test social
function testSocial() {
  var email = '3153iv@gmail.com'; // get email address

  var subject = "Victory Booth social test";

//  // get doc as html
//  var id = '1Faj6agPM6BUc-URy8ZBlbN77R9A405cNVvSHpKPpB1Y';
//  var html = getDocAsHTML_(id);
//  var content = html;

  // get html from doc
  var id = '1fuK2YPX2AOniRonBQS4wFaK9ZPOwDf5cmTRd9vcGGog';
  var content = DocumentApp.openById(id).getBody().editAsText().getText();

  var body = "To view this email, please enable html in your email client.";

  Logger.log(content);

  sendEmail_(email, subject, content, content);
}


// test function
function testingTesting() {

  var email = '3153iv@gmail.com'; // get email address

  // send info email
  var subject = "Victory Booth @ Manny’s: information for your upcoming shifts!";

//  var doc = DocumentApp.openById('11EH4RUDuCZMh6E7UL2WuvxOmA6z9kUOsHTSwHJ5yOeM');
////  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/11EH4RUDuCZMh6E7UL2WuvxOmA6z9kUOsHTSwHJ5yOeM/edit');
//  var body = doc.getBody().editAsText();

  // get doc as html
  var id = '11EH4RUDuCZMh6E7UL2WuvxOmA6z9kUOsHTSwHJ5yOeM';
  var html = getDocAsHTML_(id);

  var content = html.split('{TIME_AND_DATE}')[0] +
    'Thursday, Sept 10, 2:00pm' +
    html.split('{TIME_AND_DATE}')[1];
  var body = "To view this email, please enable html in your email client.";

  sendEmail_(email, subject, body, content);
}


function test() {
  var a = 'line 1';
  var b = 'line 2';
  var array = [a,b];
  Logger.log(array.join(', '));
}




// ************************************************************************************************

/**
 * Insert Custom menu when the spreadsheet opens.
 */
function onOpen() {
  var menu = []
  menu.push({name: 'Send Calendar Invites', functionName: 'sendInvites'});
  menu.push({name: 'Send Followups', functionName: 'sendFollowups'});
  menu.push({name: 'Create Check-in', functionName: 'createCheckin'});

  SpreadsheetApp.getActive().addMenu('Custom', menu);
}




/**
 * Auto-trigger on form submit
 */
function _onFormSubmit(e) {
  sendInvites(); // TODO: improve efficiency by not running everything
}


/**
 * Add the user as a guest for every session he or she selected.
 * @param {object} user An object that contains the user's name and email.
 * @param {Array<String[]>} response An array of data for the user's session choices.
 */
function sendInvites() {
  var cal = CalendarApp.getCalendarById(ScriptProperties.getProperty('calId'));

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Form Responses 1');
  var range = sheet.getDataRange();
  var values = range.getValues();

  // define columns
  var EMAIL = 1;
  var FNAME = 2;
  var LNAME = 3;
  var PHONE = 4;
  var INVITED = 63;

  // TODO: change date range to use, test email suppression for empties
  var LAUNCHDAY = 11;
  var ENDDAY    = 61;

  // email content
  var subject = "Victory Booth @ Manny’s: information for your upcoming shifts!";
  var html = getDocAsHTML_('11EH4RUDuCZMh6E7UL2WuvxOmA6z9kUOsHTSwHJ5yOeM');

  if (!(values[0][INVITED] == "Invites Sent?")) {
    Logger.log("Columns may have shifted!!");
    return;
  }

  // iterate through spreadsheet: all submissions (rows)
  for (var i = 1; i < values.length; i++) { // TODO: run in reverse order, stop at first already processed
    if (!values[i][INVITED]) {

      // get details
      var fname = values[i][FNAME];
      var lname = values[i][LNAME];
      var phone = values[i][PHONE];
      var email = values[i][EMAIL];

      Logger.log('Row ' + i + ': ' + email);

      var timeslots = []; // initialize

      // iterate through dates (columns)
      for (var j = LAUNCHDAY; j <= ENDDAY; j++) {
        // TODO: check it's not in the past

        var entry = values[i][j];
        if (entry && !(entry == '-')) { // TODO: test
          Logger.log('row, entry: %s', i, entry);

          // assemble date & time info
          var time = startTimeFromSession_(entry);
          date = new Date(values[0][j] + ', 2020');
          var timezone = (date > new Date('10/31/2020')) ? ' -0800': ' -0700';

          startDateTime = new Date((date.getMonth()+1) + '/' + date.getDate() + ', 2020 ' + time + timezone);
          endDateTime = new Date(startDateTime);
          endDateTime.setHours(endDateTime.getHours()+1);

          // create a calendar event
          var event = createVolunteerSlot_(cal, fname, email, startDateTime, endDateTime, []); // 52*60, 90
          event.setDescription([fname, lname, phone].join(' '));
          timeslots.push(Utilities.formatDate(startDateTime, "US/Pacific", "EEEE, MMM d, h:mmaaa")); // see https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html
          Logger.log(timeslots);

          // TODO: schedule followup email
        }
      }

      if (timeslots.length > 0) {
        // send info email
//        var content = html.split('{TIME_AND_DATE}')[0] + '• ' +
//          timeslots.join('</span></=p><p class=3D"c0"><span class=3D"c3">• ') +
//            html.split('{TIME_AND_DATE}')[1];
        var content = subStringIntoTag('{TIME_AND_DATE}', '• ' + timeslots.join('</span></=p><p class=3D"c0"><span class=3D"c3">• '), html);
//        Logger.log(content);
        content = subStringIntoTag('{NAME}', fname, content);
//        Logger.log(content);
        var body = "To view this email, please enable html in your email client.";
        Logger.log(email + ': ' + timeslots.join(', '));
        sendEmail_(email, subject, body, content);

        // mark as processed
        values[i][INVITED] = true;
        range.setValues(values); // write values to mark processed
      }
    }
  }
  Logger.log('done');
}



// fill tag in string
function subStringIntoTag(tag, sub, string) {
  return string.split(tag).join(sub);
}



/*
 * Send "thank you" followup emails
 */
function sendFollowups() {

  var activeSheet = SpreadsheetApp.getActiveSheet();
//  var rangeList = activeSheet.getRangeList(['A1:B4', 'D1:E4']);
//  rangeList.activate();

  var selection = activeSheet.getSelection();
//  // Current Cell: D1
//  Logger.log('Current Cell: ' + selection.getCurrentCell().getA1Notation());
//  // Active Range: D1:E4
//  Logger.log('Active Range: ' + selection.getActiveRange().getA1Notation());
//  // Active Ranges: A1:B4, D1:E4
//  var ranges =  selection.getActiveRangeList().getRanges();
//  for (var i = 0; i < ranges.length; i++) {
//    Logger.log('Active Ranges: ' + ranges[i].getA1Notation());
//  }
//  Logger.log('Active Sheet: ' + selection.getActiveSheet().getName());

  // set up dialog boxes
  var ui = SpreadsheetApp.getUi();

  // check only one range selected
  if (selection.getActiveRangeList().getRanges().length > 1) {
    ui.alert('Unable to process multiple selections.');
    return;
  }

  // check selection starts at the top
  if (selection.getActiveRange().getRow() > 1) {
    ui.alert('Select a whole column, from the very top.');
    return;
  }

  // get sheet data
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Form Responses 1');
  var range = sheet.getDataRange();
  var values = range.getValues();

  // check selection is full height
  if (selection.getActiveRange().getHeight() < range.getHeight()) {
    ui.alert('Select a whole column.');
    return;
  }

  // check selection is single column
  if (selection.getActiveRange().getWidth() > 1) {
    ui.alert('Select a single column.');
    return;
  }
  // TODO: handle multiple column selection

  var thankyous = selection.getActiveRange().getValues();

  // confirm date to thank
  var senddate = new Date(thankyous[0] + ' 2020');
  var result = ui.alert(
    'Please confirm: ' + Utilities.formatDate(senddate, "US/Pacific", "EEEE, MMM d"),
      'Are you sure you want to send thank you emails for the following?\n   • '
      + Utilities.formatDate(senddate, "US/Pacific", "EEEE, MMM d")
      + '\n   • (column ' + selection.getActiveRange().getA1Notation().split(':')[0] + ')',
        ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    // define column map
    var EMAIL = 1;
    var FNAME = 2;

    // email content
    var subject = "You did something to help us win!";
    var html = getDocAsHTML_('1b1zCV_xhLJ45Vt6nTHpRqCHeRsbzFAJ1ms6nNmV4RVo');
    var content = html;
    var body = "To view this email, please enable html in your email client.";

    Logger.log(thankyous[0]);
    var emails = [];
    var names = [];
    for (var i = 1; i < values.length; i++) {
      var addemail = values[i][EMAIL];
      if (!(thankyous[i] == "")) {
        if (emails.indexOf(addemail) < 0) {
          emails.push(addemail);
          names.push(values[i][FNAME]);
        }
      }
    }

    // send emails
    for (var i = 0; i < emails.length; i++) {
      sendEmail_(emails[i], subject, body, subStringIntoTag('{NAME}', names[i], content));
      Logger.log(emails[i]);
    }
    ui.alert(emails.length + ' emails sent!');
  }
}

function createCheckin() {
  var activeSheet = SpreadsheetApp.getActiveSheet();

  var selection = activeSheet.getSelection();

  // set up dialog boxes
  var ui = SpreadsheetApp.getUi();

  // check only one range selected
  if (selection.getActiveRangeList().getRanges().length > 1) {
    ui.alert('Unable to process multiple selections.');
    return;
  }

  // check selection starts at the top
  if (selection.getActiveRange().getRow() > 1) {
    ui.alert('Select a whole column, from the very top.');
    return;
  }

  // get sheet data
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Form Responses 1');
  var range = sheet.getDataRange();
  var values = range.getValues();

  // check selection is full height
  if (selection.getActiveRange().getHeight() < range.getHeight()) {
    ui.alert('Select a whole column.');
    return;
  }

  // check selection is single column
  if (selection.getActiveRange().getWidth() > 1) {
    ui.alert('Select a single column.');
    return;
  }
  // TODO: handle multiple column selection

  var daysSessions = selection.getActiveRange().getValues();

  // define column map
  var EMAIL = 1;
  var FNAME = 2;
  var LNAME = 3;
  var PHONE = 4;

  var sessions = new Map();
  // create list of indexes for each session
  for (var i = 1; i < daysSessions.length; i++) {
    var sessionTimeDate = daysSessions[i]
    if (!sessions.has(sessionTimeDate)) {
      sessions.set(sessionTimeDate, []);
    }
    sessions.get(sessionTimeDate).append(i);
  }

  // assemble date & time info
  var date = new Date(values[0][j] + ', 2020');
  var fileName = Utilities.formatDate(date, "US/Pacific", "EEEE, MMM d, h:mmaaa"); // see https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html

  // create file
  var file = folder.createFile(fileName, "");

  var heading = Utilities.formatDate(date, "US/Pacific", "EEEE, MMM d, h:mmaaa"); // see https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html
  var body = file.getBody();

  // create checkin for each session
  sessions.forEach((personIndexList, sessionDateTime, map) => {
  body.appendParagraph(heading);
  body.appendParagraph(sessionDateTime);

  contents = []
  // build table contents
  personIndexList.forEach((index) => {
                          contents.append([values[index][FNAME] + values[index][LNAME], ' ', ' ', values[index][PHONE] + '\n' + values[index][EMAIL], ' ']);
});

var table = [['Name', 'L / P / T', 'Qty', 'Phone/Email', 'Food/Drink']].join(contents);
body.appendTable(table);
  });
}

function createCheckin() {
  var activeSheet = SpreadsheetApp.getActiveSheet();

  var selection = activeSheet.getSelection();

  // set up dialog boxes
  var ui = SpreadsheetApp.getUi();

  // check only one range selected
  if (selection.getActiveRangeList().getRanges().length > 1) {
    ui.alert('Unable to process multiple selections.');
    return;
  }

  // check selection starts at the top
  if (selection.getActiveRange().getRow() > 1) {
    ui.alert('Select a whole column, from the very top.');
    return;
  }

  // get sheet data
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Form Responses 1');
  var range = sheet.getDataRange();
  var values = range.getValues();

  // check selection is full height
  if (selection.getActiveRange().getHeight() < range.getHeight()) {
    ui.alert('Select a whole column.');
    return;
  }

  // check selection is single column
  if (selection.getActiveRange().getWidth() > 1) {
    ui.alert('Select a single column.');
    return;
  }

  var daysSessions = selection.getActiveRange().getDisplayValues();

  // define column map
  var EMAIL = 1;
  var FNAME = 2;
  var LNAME = 3;
  var PHONE = 4;

  var sessions = new Map();
  // create list of indexes for each session
  for (var i = 1; i < daysSessions.length; i++) {
    var sessionTimeDate = daysSessions[i][0]
    if (sessionTimeDate !== "") {
      if (!sessions.has(sessionTimeDate)) {
        sessions.set(sessionTimeDate, []);
      }
      sessions.get(sessionTimeDate).push(i);
    }
  }

  // assemble date & time info
  var date = new Date(daysSessions[0][0] + ' 2020');
  var fileName = Utilities.formatDate(date, "US/Pacific", "EEEE, MMM d"); // see https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html

  // create file
  var folder = DriveApp.getFolderById("10hnMB9pIpdKsAhGiiH1EBP154XQ2NgNQ");
  var file = folder.createFile(fileName, "");

  var heading = Utilities.formatDate(date, "US/Pacific", "EEEE, MMM d"); // see https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html
  var body = file.getBody();

  // create page for each session
  sessions.forEach((personIndexList, sessionDateTime, map) => {
    // session heading
    body.appendParagraph(heading.toString()).setBold(true);
    body.appendParagraph(sessionDateTime.toString()).setBold(false);

    contents = []
    // build table contents
    personIndexList.forEach((index) => {
      contents.push(["🗹 " + values[index][FNAME] + " " + values[index][LNAME], ' ', ' ', values[index][PHONE] + '\n' + values[index][EMAIL], ' ']);
    });

    var table = [['Name', 'L / P / T', 'Qty', 'Phone/Email', 'Food/Drink']].concat(contents);
    body.appendTable(table);

    body.appendPageBreak();
  });
}


// TODO: create event, send email, schedule reminder email, all upon form submission
//function _onFormSubmit(e) {
//  var user = {name: e.namedValues['First Name'][0] + ' '+ e.namedValues['Last Name'][0],
//              email: e.namedValues['Email address'][0]};
//
//  // Grab the session data again so that we can match it to the user's choices.
//  var response = [];
//  var values = SpreadsheetApp.getActive().getSheetByName('Conference Setup').getDataRange().getValues();
//  for (var i = 1; i < values.length; i++) {
//    var session = values[i];
//    var title = session[0];
//    var day = session[1].toLocaleDateString();
//    var time = session[2].toLocaleTimeString();
//    var timeslot = time + ' ' + day;
//
//    // For every selection in the response, find the matching timeslot and title
//    // in the spreadsheet and add the session data to the response array.
//    if (e.namedValues[timeslot] && e.namedValues[timeslot] == title) {
//      response.push(session);
//    }
//  }
//  sendInvites_(user, response);
//  sendDoc_(user, response);
//}



// TODO: reformat into checkbox grid form, auto update grid


// create calendar event & send invite
function createVolunteerSlot_(cal, name, email, startDateTime, endDateTime, reminders) {
  var event = cal.createEvent(name + ' -Victory Booth shift',
                              new Date(startDateTime),
                              new Date(endDateTime),
                              {location: "Victory Booths @ Manny's, 485 Valencia St, San Francisco, CA 94103, USA",
                               guests: email,
                               sendInvites: true}).setGuestsCanSeeGuests(false);
  for (var i = 0; i < 5; i++) {
    if (i < reminders.length) {
      event.addEmailReminder(reminders[i]);
    } else {
      break;
    }
  }

  return event;
}


// send details email
function sendEmail_(email, subject, body, content) {
  MailApp.sendEmail(
    email,           // recipient
    subject,         // subject
    body, {          // body
      htmlBody: content // advanced options
    }
  );
}


/**
 * get doc contents as HTML
 * see https://stackoverflow.com/questions/39779550/how-to-send-rich-text-emails-with-gmailapp
 */
function getDocAsHTML_(id) {
  var forDriveScope = DriveApp.getStorageUsed(); //needed to get Drive Scope requested
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+id+"&exportFormat=html";
  var param = {
    method      : "get",
    headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions:true,
  };
  var html = UrlFetchApp.fetch(url,param).getContentText();
  return html;
}


// Decode the time from a session string of form "Session 1: 12:00PM"
function startTimeFromSession_(seshString) {
  var sesh = seshString.split(': ')[1].split(':')[0]; // look at hour only, assumes all PM
  switch(sesh) {
    case '12':
      return '12:00:00';
    case '1':
      return '13:00:00';
    case '2':
      return '14:00:00';
    case '3':
      return '15:00:00';
    case '4':
      return '16:00:00';
    case '5':
      return '17:00:00';
    case '6':
      return '18:00:00';
  }
}
