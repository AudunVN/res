function onOpen() {
  var menu = [{name: 'Set up reservation system', functionName: 'setUpReservationSystem'}];
  if (ScriptProperties.getProperty('formId')) {
    var updateForm = {name: 'Update form', functionName: 'setFormEntries'};
    menu.push(updateForm);
  }
  SpreadsheetApp.getActive().addMenu('Reservation system', menu);
}

/**
 * A set-up function that uses the conference data in the spreadsheet to create
 * Google Calendar events, a Google Form, and a trigger that allows the script
 * to react to form responses.
 */
function setUpReservationSystem() {
  if (ScriptProperties.getProperty('calId')) {
    Browser.msgBox('Your reservation system is already set up. Look in Google Drive!');
  }
  var ss = SpreadsheetApp.getActive();
  ss.insertSheet('Reservations', 0);
  var sheet = ss.getSheetByName('Reservations');
  var range = sheet.getDataRange();
  var values = range.getValues();
  setUpCalendar(values, range);
  setUpForm(ss, values);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
      .create();
  ss.removeMenu('Reservation system');
}

/**
 * Creates a Google Calendar with events for each conference session in the
 * spreadsheet, then writes the event IDs to the spreadsheet for future use.
 *
 * @param {String[][]} values Cell values for the spreadsheet range.
 * @param {Range} range A spreadsheet range that contains conference data.
 */
function setUpCalendar(values, range) {
  var cal = CalendarApp.createCalendar('Gunnerus Reservation Calendar');
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = session[0];
    var start = joinDateAndTime_(session[1], session[2]);
    var end = joinDateAndTime_(session[1], session[3]);
    var options = {location: session[4], sendInvites: true};
    var event = cal.createEvent(title, start, end, options)
        .setGuestsCanSeeGuests(false);
    session[5] = event.getId();
  }
  range.setValues(values);

  // Store the ID for the Calendar, which is needed to retrieve events by ID.
  ScriptProperties.setProperty('calId', cal.getId());
}

/**
 * Creates a single Date object from separate date and time cells.
 *
 * @param {Date} date A Date object from which to extract the date.
 * @param {Date} time A Date object from which to extract the time.
 * @return {Date} A Date object representing the combined date and time.
 */
function joinDateAndTime_(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}

/**
 * Creates a Google Form that allows respondents to select which conference
 * sessions they would like to attend, grouped by date and start time.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {String[][]} values Cell values for the spreadsheet range.
 */

function setFormEntries() {
  var form = FormApp.openById(ScriptProperties.getProperty('formId'));
  form.setTitle('Reservation Form')
  .setDescription('Reserve a time slot on the Gunnerus.')
  .setConfirmationMessage('Your registration has been received. A receipt has been sent to your email address.')
  .setAllowResponseEdits(true)
  .setAcceptingResponses(true);
  Logger.log(form.getItems());
  var formLength = form.getItems().length;
  Logger.log("Form length: "+formLength);
  for (var i = 0; i < formLength; i++) {
    form.deleteItem(0);
  }
  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addDateItem().setTitle('Start time').setRequired(true);
  form.addDateItem().setTitle('End time').setRequired(true);
}

function setUpForm(ss, values) {
  // Group the sessions by date and time so that they can be passed to the form.
  var schedule = {};
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var day = session[1].toLocaleDateString();
    var time = session[2].toLocaleTimeString();
    if (!schedule[day]) {
      schedule[day] = {};
    }
    if (!schedule[day][time]) {
      schedule[day][time] = [];
    }
    schedule[day][time].push(session[0]);
  }

  // Create the form and add a multiple-choice question for each timeslot.
  var form = FormApp.create('Reservation Form');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  setFormEntries();
  /*for (var day in schedule) {
    var header = form.addSectionHeaderItem().setTitle('Sessions for ' + day);
    for (var time in schedule[day]) {
      var item = form.addMultipleChoiceItem().setTitle(time + ' ' + day)
          .setChoiceValues(schedule[day][time]);
    }
  }*/
  ScriptProperties.setProperty('formId', form.getId());
}

/**
 * A trigger-driven function that sends out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  var editLink = getLastURL(ScriptProperties.getProperty('formId'));
  var user = {name: e.namedValues['Name'][0], email: e.namedValues['Email'][0]};

  // Grab the session data again so that we can match it to the user's choices.
  var response = [];
  var values = SpreadsheetApp.getActive().getSheetByName('Reservations')
     .getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = session[0];
    var day = session[1].toLocaleDateString();
    var time = session[2].toLocaleTimeString();
    var timeslot = time + ' ' + day;

    // For every selection in the response, find the matching timeslot and title
    // in the spreadsheet and add the session data to the response array.
    if (e.namedValues[timeslot] && e.namedValues[timeslot] == title) {
      response.push(session);
    }
  }
  /*sendInvites_(user, response);*/
  sendDoc(user, response, editLink);
}

/**
 * Add the user as a guest for every session he or she selected.
 *
 * @param {Object} user An object that contains the user's name and email.
 * @param {String[][]} response An array of data for the user's session choices.
 *//*
function sendInvites_(user, response) {
  var id = ScriptProperties.getProperty('calId');
  var cal = CalendarApp.getCalendarById(id);
  for (var i = 0; i < response.length; i++) {
    cal.getEventSeriesById(response[i][5]).addGuest(user.email);
  }
}*/

/**
 * Create and share a personalized Google Doc that shows the user's itinerary.
 *
 * @param {Object} user An object that contains the user's name and email.
 * @param {String[][]} response An array of data for the user's session choices.
 */
function sendDoc(user, response, editLink) {
  var doc = DocumentApp.create('Reservation summary for ' + user.name)
      .addEditor(user.email);
  var body = doc.getBody();
  var table = [['Session', 'Date', 'Time', 'Location']];
  for (var i = 0; i < response.length; i++) {
    table.push([response[i][0], response[i][1].toLocaleDateString(),
        response[i][2].toLocaleTimeString(), response[i][4]]);
  }
  body.insertParagraph(0, doc.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  table = body.appendTable(table);
  table.getRow(0).editAsText().setBold(true);
  doc.saveAndClose();
  
  // Email a link to the Doc as well as a PDF copy.
  MailApp.sendEmail({
    to: user.email,
    subject: doc.getName(),
    body: 'Thanks for registering! Here\'s your itinerary: ' + doc.getUrl() + ', and your edit link is: test ' + editLink,
    attachments: doc.getAs(MimeType.PDF),
  });
}

function getLastURL(FormID) {
  var form = FormApp.openById(FormID); //ID of the form 
  var responses = form.getResponses();
  var lasturl = responses[responses.length-1].getEditResponseUrl();
  return lasturl;
}
