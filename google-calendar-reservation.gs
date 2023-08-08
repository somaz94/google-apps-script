// Add your Google Sheets URL here
var spreadsheetUrl = '';

function logError(error) {
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheet = spreadsheet.getSheetByName("Reply_01"); // Change this to your sheet name
  var lastRow = sheet.getLastRow() + 1; // Get the next row
  sheet.getRange(lastRow, 1).setValue(new Date());
  sheet.getRange(lastRow, 2).setValue(error);
}

// User Slack applications. Insert the token of the Slack Outgoing WebHook
var slackTokens = ['', '']; 

// ADD your Calendar ID
var calendarIds = {
  "Large Conference Room": 'c_xxxx@resource.calendar.google.com',
  "Small Conference Room": 'c_xxxx@resource.calendar.google.com',
};

// ADD your roomName
var roomNamesInKorean = {
  "Large Conference Room": "10-1 대회의실",
  "Small Conference Room": "10-2 소회의실",
};

// ADD your slackWebhookUrl
var slackWebhookUrl = '';

function doPost(e) {
  Logger.log('doPost started');
  var eventId;
  try {
    var data = e.parameter;
    Logger.log('Data parsed: ' + data.text);

    // Check Token Value
    if ( !slackTokens.includes(data.token) ) {
      Logger.log('Invalid token');
      sendSlackMessage('Error: Invalid token');
      return;
    }

    var text = data.text;
    var roomName, roomCalendarName;
    var roomPattern = /회의실 : (10-\d [가-힣]+)/;
    var match = text.match(roomPattern);
    if (match) {
      roomCalendarName = match[1];
      if (roomCalendarName === "10-1 대회의실") {
        roomName = "Large Conference Room";
      } else if (roomCalendarName === "10-2 소회의실") {
        roomName = "Small Conference Room";
      } 
    } else {
      Logger.log('No valid room name found in the text');
      sendSlackMessage('Error: No valid room name found in the text');
      return;
    }

    if (roomName) {
      var success = true;
      var message;
      if (text.includes("회의실 예약")) {
        eventId = createReservation(text, roomName);
        if (!eventId) {
          success = false;
        } else {
          message = '회의실 예약이 정상적으로 처리되었습니다.';
        }
      } else if (text.includes("예약취소")) {
        eventId = getEventIdFromText(text);
        Logger.log('Returned eventId for cancellation: ' + eventId);
        if (eventId) {
          eventId = cancelReservation(eventId, roomName);
          if (!eventId) {
            success = false;
          } else {
            message = '회의실 예약 취소가 정상적으로 처리되었습니다.';
          }
        } else {
          Logger.log('Event ID not provided for cancellation');
          sendSlackMessage('Error: Event ID not provided for cancellation', eventId, roomName);
          return;
        }
      } else if (text.includes("예약수정")) {
          Logger.log('Reservation modification text: ' + text); // Add this line for debugging
          eventId = getEventIdFromText(text);
          if (eventId) {
              eventId = modifyReservation(eventId, text, roomName);
              if (!eventId) {
                  success = false;
              } else {
                  message = '회의실 예약 수정이 정상적으로 처리되었습니다.';
              }
          } else {
              Logger.log('Event ID not provided for modification');
              sendSlackMessage('Error: Event ID not provided for modification', eventId, roomName);
              return;
          }
      }

      if (success) {
        sendSlackMessage(message, eventId, roomName);
      }
    }
  } catch (error) {
    Logger.log('Error in doPost: ' + error);
    logError(error.message);
    sendSlackMessage('Error: ' + error.message, eventId, roomName);
  }
  Logger.log('doPost ended');
}

function createReservation(text, roomName) {
  try {
    var calendarId = calendarIds[roomName];
    var calendar = CalendarApp.getCalendarById(calendarId);
    var startTime = getDateTimeFromText(text, "시작 시간 :");
    var endTime = getDateTimeFromText(text, "종료 시간 :");

    if (!startTime) {
      sendSlackMessage('Error: Invalid start time');
      throw new Error('Invalid start time');
    }

    if (!endTime) {
      endTime = new Date(startTime.getTime() + 60 * 60 * 1000); // Add 1 hour to the start time
    }

    var meetingName = getMeetingNameFromText(text);

    Logger.log('Start Time: ' + startTime);
    Logger.log('End Time: ' + endTime);

    try {
      // Check for conflicting events before creating the new reservation
      var conflicts = calendar.getEvents(startTime, endTime);
      if (conflicts.length > 0) {
        var message = '동일한 시간에 이미 예약이 있습니다. 캘린더를 확인해 주세요.';
        sendSlackMessage('Error: ' + message);
        throw new Error(message);
      }
    } catch (error) {
      Logger.log('Error in checking for conflicts: ' + error);
      return null;
    }

    var event = calendar.createEvent(meetingName, startTime, endTime);
    var eventId = event.getId();
    event.setDescription('Event ID: ' + eventId);
    // Wait for a while for Google Calendar to sync
    Utilities.sleep(1000);  // Increase the waiting time
    event = calendar.getEventById(event.getId());  // Reload the event
    Logger.log('Created event ID: ' + eventId + ' in ' + roomName);
    return eventId;

  } catch (error) {
    Logger.log('Error in createReservation: ' + error);
    sendSlackMessage('Error: ' + error.message);
    logError(error.message);
  }
}

function getMeetingNameFromText(text) {
  var namePattern = /회의 이름 : (.*)/;
  var match = text.match(namePattern);
  if (match) {
    return match[1];
  } else {
    Logger.log('No valid meeting name found in the text');
    return null;
  }
}

function getEventIdFromText(text) {
  var eventIdPattern = /회의 ID : (\S+)/;
  var match = text.match(eventIdPattern);
  if (match) {
    var decodedId = Utilities.newBlob(Utilities.base64Decode(match[1])).getDataAsString();
    return decodedId;
  } else {
    Logger.log('No valid event ID found in the text');
    return null;
  }
}

function getEventIdFromEvent(event) {
  var description = event.getDescription();
  var eventIdPattern = /Event ID: (\S+)/;
  var match = description.match(eventIdPattern);
  if (match) {
    return match[1];
  } else {
    Logger.log('No valid event ID found in the event description');
    return null;
  }
}

function getDateTimeFromText(text, timeLabel) {
  var datePattern = /회의 날짜 : (\d{4}-\d{2}-\d{2})/;
  var timePattern = new RegExp(timeLabel + " (\\d{2}:\\d{2})");
  var dateMatch = text.match(datePattern);
  var timeMatch = text.match(timePattern);
  if (dateMatch && timeMatch) {
    return new Date(dateMatch[1] + 'T' + timeMatch[1] + ':00');
  } else if (dateMatch && timeLabel === "종료 시간 :") {  // If end time is not provided
    var startTime = getDateTimeFromText(text, "시작 시간 :");
    if (startTime) {
      var endTime = new Date(startTime.getTime() + 60 * 60 * 1000); // Add 1 hour to the start time
      return endTime;
    }
  } else {
    Logger.log('No valid date and time found in the text');
    return null;
  }
}

function cancelReservation(eventId, roomName) {
  var calendarId = calendarIds[roomName];
  var calendar = CalendarApp.getCalendarById(calendarId);

  var now = new Date();
  var future = new Date();
  future.setFullYear(now.getFullYear() + 1);  // Look for the event in the next one year

  var events = calendar.getEvents(now, future);

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    if (getEventIdFromEvent(event) === eventId) {
      event.deleteEvent();
      Utilities.sleep(2000);
      Logger.log('Cancelled event with ID: ' + eventId);
      return eventId;
    }
  }
  Logger.log('No event found with ID: ' + eventId);
  // sendSlackMessage('Error: No event found with provided ID', eventId, roomName);
  return null;
}

function modifyReservation(eventId, text, roomName) {
  var calendarId = calendarIds[roomName];
  var calendar = CalendarApp.getCalendarById(calendarId);

  var now = new Date();
  var future = new Date();
  future.setFullYear(now.getFullYear() + 1);  // Look for the event in the next one year

  var events = calendar.getEvents(now, future);

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    if (getEventIdFromEvent(event) === eventId) {
      var newStartTime = getDateTimeFromText(text, "시작 시간 :");
      var newEndTime = getDateTimeFromText(text, "종료 시간 :");
      if (!newStartTime) {
        Logger.log('Invalid new start time');
        return null;
      }
      if (!newEndTime) {
        newEndTime = new Date(newStartTime.getTime() + 60 * 60 * 1000);
      }
      event.setTime(newStartTime, newEndTime);
      Utilities.sleep(2000);
      Logger.log('Modified event with ID: ' + eventId);
      return eventId;
    }
  }
  Logger.log('No event found with ID: ' + eventId);
  // sendSlackMessage('Error: No event found with provided ID', eventId, roomName);
  return null;
}

function sendSlackMessage(message, eventId, roomName) {
  var payload;
  if (message.includes('동일한 시간에 이미 예약이 있습니다.')) {
    payload = { text: message };
  } else {
    var encodedId = eventId ? Utilities.base64Encode(Utilities.newBlob(eventId).getBytes()) : 'None';
    payload = {
      attachments: [{
        pretext: message,
        color: "#D00000",
        fields: [{
          title: "회의 ID",
          value: encodedId,
          short: false
        }, {
          title: roomNamesInKorean[roomName] + " 캘린더",
          value: "https://calendar.google.com/calendar/embed?src=" + calendarIds[roomName] + "&ctz=Asia%2FSeoul",
          short: false
        }]
      }]
    };
  }

  var options = {
    method: 'post',
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(slackWebhookUrl, options);
}

