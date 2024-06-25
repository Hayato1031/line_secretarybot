var CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("CHANNEL_ACCESS_TOKEN")
var USER_ID = PropertiesService.getScriptProperties().getProperty("USER_ID")
var MAIL = PropertiesService.getScriptProperties().getProperty("MAIL")

var googleCalendar = {
  MAIL : "予定",
  "ja.japanese#holiday@group.v.calendar.google.com": "祝日",
};

var weekday = ["日", "月", "火", "水", "木", "金", "土"];
function pushWeekly() {

  var calendars = CalendarApp.getAllCalendars();
  var dt = new Date()
  var message = "1週間の予定です。\n\n";

  for ( var i = 0;  i < 7;  i++ ) {

    dt.setDate(dt.getDate() + 1);
    message += Utilities.formatDate(dt, 'JST', '★ MM/dd(' + weekday[dt.getDay()] + ')') + "\n";

    var dayText = "";
    for(g in calendars) {
      var calendar = calendars[g];

      var calendarName = googleCalendar[calendar.getId()]
      if ( calendarName == undefined ) {
        continue;
      }

      var events = calendar.getEventsForDay(dt);
      if( events.length == 0 ) {
        continue;
      }

      dayText += "< " + calendarName + " >\n";
      for(g in events) {
        dayText += DayText(events[g]);
      }
      dayText += "\n"
    }

    if ( dayText == "") {
        dayText += "予定なし\n\n";
    }
    message += dayText;
  }

  sendToLine(message);
}

function sendToLine(message){
    //deleteTrigger();
  var postData = {
    "to": USER_ID,
    "messages": [{
      "type": "text",
      "text": message,
    }]
  };

  var url = "https://api.line.me/v2/bot/message/push";
  var headers = {
    "Content-Type": "application/json",
    'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(postData)
  };
  var response = UrlFetchApp.fetch(url, options);
}

function DayText(event) {
  return TimeText(event.getStartTime()) + ' - ' + TimeText(event.getEndTime()) + " " + event.getTitle() + '\n';
}

function TimeText(str){
  return Utilities.formatDate(str, 'JST', 'HH:mm');
}