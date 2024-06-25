var CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("CHANNEL_ACCESS_TOKEN");
var SHEET_ID = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
var dateExp = /(\d+)\/(\d+)\s(\d+):(\d+)/;
var ss = SpreadsheetApp.openById(SHEET_ID);
var wishlist = ss.getSheetByName("WishList");

//doPost関数（Lineからメッセージを受け取る）
function doPost(e) {
    GetMessage(e);
}

//受け取ったメッセージの処理
function GetMessage(e) {
  var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
  if (typeof replyToken === 'undefined') {
    return;
  }
  var messageText = JSON.parse(e.postData.contents).events[0].message.text;
  var cache = CacheService.getScriptCache();
  var type = cache.get("type");

  if (type === null) {
    //初期メッセージ
      if (messageText === "SelectA") {
        //予定メニュー
        cache.put("type", "plan_menu");
        reply(replyToken, "予約を選択しました。\n予定追加→SelectA\n今日の予定→SelectB\n今週の予定→SelectC\nを選択してください");
      } else if (messageText.match("SelectB")) {
        //欲しいものメニュー
        cache.put("type", "wish_menu");
        reply(replyToken, "欲しいものリストを選択しました。\n欲しいものリスト→SelectA\n欲しいもの追加→SelectBを選択してください");
      }  else if (messageText.match("SelectC")) {
        ;
      } else {
      //処理方法の返答
        replyPlans(replyToken, "「SelectA」で予定に関する項目が使用できます", "「SelectB」は欲しいものリストです。", "「SelectF」でキャンセル扱いです");
      }    
    
  } else {
    //キャンセル処理
    if (messageText === "SelectF") {
      cache.remove("type");
      reply(replyToken, "実行中の動作をキャンセルをしました");
      return;
    } 

    switch(type) {
      case "plan_menu":
        //予定の追加
        if (messageText === "SelectA") {
          cache.put("type", "plan_add_1");
        //開始日時の質問
          replyPlans(replyToken, "予定日をいずれかの形式で教えてください。", "12/1\n3:00", "4/1 13:00");
        //今日、７日間の予定の取得
        } else if (messageText.match("SelectB")) {
          reply(replyToken, getEvents());
        }  else if (messageText.match("SelectC")) {
          reply(replyToken, notifyWeekly());
        } else {
        //処理方法の返答
          replyPlans(replyToken, "「予定の追加」で予定追加します", "「今日の予定」で今日の予定をお知らせします。", "「今週の予定」で7日間の予定をお知らせします。");
        }

      case "plan_add_1":
        // 開始日時の追加
        if ( messageText.match(dateExp)) {
          var [matched, start_month, start_day, start_Hour, start_Min] = messageText.match(dateExp);
          cache.put("type", "plan_add_2");
          cache.put("start_month", start_month);
          cache.put("start_day", start_day);
          cache.put("start_hour", start_Hour);
          cache.put("start_min", start_Min);
          //終了日時の質問
          var year = new Date().getFullYear();
          //var year = 2020;
          var startDate = new Date(year, cache.get("start_month") - 1, cache.get("start_day"), cache.get("start_hour"), cache.get("start_min"));
          reply(replyToken,"開始日時は\n" + EventFormat(startDate) + "\nですね。\n\n次に予定の終了日時をお知らせください。");
          break;
        }else{
          reply(replyToken, "予定追加処理中です。\n「キャンセル」\nで追加作業をキャンセルします。");
          break;
        }

      case "plan_add_2":
        // 終了日時の追加
        if ( messageText.match(dateExp)) {
          var [matched, end_month, end_day, end_Hour, end_Min] = messageText.match(dateExp);
          cache.put("type", "plan_add_3");
          cache.put("end_month", end_month);
          cache.put("end_day", end_day);
          cache.put("end_hour", end_Hour);
          cache.put("end_min", end_Min);
          //予定名の質問
          var year = new Date().getFullYear();
          //var year = 2020;
          var endDate = new Date(year, cache.get("end_month") - 1, cache.get("end_day"), cache.get("end_hour"), cache.get("end_min"));
          reply(replyToken,"終了日時は\n" + EventFormat(endDate) + "\nですね。\n\n最後に予定名を教えてください。");
          break;
        }else{
          reply(replyToken, "予定追加処理中です。\n「キャンセル」\nで追加作業をキャンセルします。");
          break;
        }

      case "plan_add_3":
        // 最終確認
        cache.put("type", "plan_add_4");
        cache.put("title", messageText);
        var [title, startDate, endDate] = createData(cache);
        //予定追加の確認
        replyPlans(replyToken, "予定名：" + title, "開始日時：\n" + EventFormat(startDate)+ "\n終了日時：\n" + EventFormat(endDate), "予定を追加しますか？\n 「はい」か「いいえ」でお知らせください。");
        break;

      case "plan_add_4":
        if (messageText === "はい") {
          cache.remove("type");
          var [title, startDate, endDate] = createData(cache);
          CalendarApp.getDefaultCalendar().createEvent(title, startDate, endDate);
          reply(replyToken, "Googleカレンダーに予定を追加しました");
        } else if (messageText === "いいえ") {
          cache.remove("type");
          reply(replyToken, "予定の追加をキャンセルしました。");
        } else{
          reply(replyToken, "「はい」か「いいで」でお答えください。");
          break;
        }
        break;

      case "wish_menu":
        if (messageText === "SelectA") {
          //欲しいものリスト確認
          wishlist_message(replyToken);
          reply(replyToken,"欲しいものリストのメニューに戻ります。")
          
        } else if (messageText.match("SelectB")) {
          //欲しいものリスト追加
          cache.put("type", "wish_add_1");
          reply(replyToken, "カテゴリーを選択してください。\n「電子機器類」→SelectA\n「飲食物」→SelectB\n「服・装飾品等」→SelectC\n「チケットなど」→SelectD\n「その他」→SelectE");
          break;
        }  else if (messageText.match("SelectC")) {
          ;
        } else {
        //処理方法の返答
          replyPlans(replyToken, "「SelectA」でほしいものリストを確認できます", "「SelectB」は欲しいものリストの追加です。", "「SelectF」でキャンセル扱いです");
        }

      case "wish_add_1":
        if (messageText === "SelectA"){
          cache.put("category", "電子機器類");
          cache.put("type", "wish_add_2");
          reply(replyToken, "次に欲しいものの名前を入力してください。");
          break;
        } else if (messageText === "SelectB"){
          cache.put("category", "飲食物");
          cache.put("type", "wish_add_2");
          reply(replyToken, "次に欲しいものの名前を入力してください。");
          break;
        } else if (messageText === "SelectC"){
          cache.put("category", "服・装飾品等");
          cache.put("type", "wish_add_2");
          reply(replyToken, "次に欲しいものの名前を入力してください。");
          break;
        } else if (messageText === "SelectD"){
          cache.put("category", "チケットなど");
          cache.put("type", "wish_add_2");
          reply(replyToken, "次に欲しいものの名前を入力してください。");
          break;
        } else if (messageText === "SelectE"){
          cache.put("category", "その他");
          cache.put("type", "wish_add_2");
          reply(replyToken, "次に欲しいものの名前を入力してください。");
          break;
        }

      case "wish_add_2":
        cache.put("name", messageText);
        cache.put("type", "wish_add_3");
        reply(replyToken, "次に欲しいもののリンクを入力してください。\nない場合は「SelectA」を入力してください");
        break;

      case "wish_add_3":
        if(messageText === "SelectA"){
          cache.put("link", "なし");
          cache.put("type", "wish_add_4");
          reply(replyToken, "次に欲しいものの金額を入力してください。\n不明な場合は「SelectA」を入力してください");
          break;
        }else{
          cache.put("link", messageText);
          cache.put("type", "wish_add_4");
          reply(replyToken, "次に欲しいものの金額を入力してください。\n不明な場合は「SelectA」を入力してください");
          break;
        }

      case "wish_add_4":
        if (/^[0-9]+$/.test(messageText)) {
          let id = wishlist.getLastRow();
          cache.put("ID", id);
          cache.put("cost", messageText);
          cache.put("type", "wish_add_5");
          let name = cache.get("name");
          let category = cache.get("category");
          let link = cache.get("link");
          let cost = cache.get("cost");
          reply(replyToken, "名称:" + name + "\nカテゴリー:" + category + "\nリンク:" + link + "\n金額:" + cost + "\n\n欲しいものを追加しますか？\n「はい」か「いいえ」で答えてください。");
          break;
        } else if (messageText === "SelectA") {
          let id = wishlist.getLastRow();
          cache.put("ID", id);
          cache.put("cost", "不明");
          cache.put("type", "wish_add_5");
          let name = cache.get("name");
          let category = cache.get("category");
          let link = cache.get("link");
          let cost = cache.get("cost");
          reply(replyToken, "名称:" + name + "\nカテゴリー:" + category + "\nリンク:" + link + "\n金額:" + cost + "\n\n欲しいものを追加しますか？\n「はい」か「いいえ」で答えてください。");
          break;
        } else {
          reply(replyToken, "半角数字もしくは「SelectA」を選択してください。\nキャンセルの場合は「SelectF」を選択してください。");
          break;
        }

      case "wish_add_5":
        if (messageText === "はい") {
          let id = cache.get("ID");
          let name = cache.get("name");
          let category = cache.get("category");
          let link = cache.get("link");
          let cost = cache.get("cost");
          
          wishlist.appendRow([id, name, category, link, cost]);
          reply(replyToken, "「" + name + "」をリストに追加しました。");
          cache.remove("type");
        } else if (messageText === "いいえ") {
          reply(replyToken, "追加を中止しました。");
          cache.remove("type");
        } else {
          reply(replyToken, "「はい」か「いいえ」で答えてください。");
        }
        break;
    }
  }
}

function createData(cache) {
  var year = new Date().getFullYear();
  //var year = 2020;
  var title = cache.get("title");
  var startDate = new Date(year, cache.get("start_month") - 1, cache.get("start_day"), cache.get("start_hour"), cache.get("start_min"));
  var endDate = new Date(year, cache.get("end_month") - 1, cache.get("end_day"), cache.get("end_hour"), cache.get("end_min"));
  return [title, startDate, endDate];
}

function EventFormat(Date) {
  var y = Date.getFullYear();
  var m = Date.getMonth() + 1;
  var d = Date.getDate();
  var w = Date.getDay();
  var H = Date.getHours();
  var M = Date.getMinutes();
  var weekname = ['日', '月', '火', '水', '木', '金', '土'];
  m = ('0' + m).slice(-2);
  d = ('0' + d).slice(-2);
  return y + '年' + m + '月' + d + '日 (' + weekname[w] + ')\n' + H + '時' + M + '分';
}

function replyPlans(replyToken, message, message2, message3) {
  var url = "https://api.line.me/v2/bot/message/reply";
  UrlFetchApp.fetch(url, {
    "headers": {
      "Content-Type": "application/json; charset=UTF-8",
      "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN,
    },
    "method": "post",
    "payload": JSON.stringify({
      "replyToken": replyToken,
      "messages": [{
        "type": "text",
        "text": message,
      },{
        "type": "text",
        "text": message2,
      },{
        "type": "text",
        "text": message3,
      }],
    }),
  });
  return ContentService.createTextOutput(JSON.stringify({"content": "post ok"})).setMimeType(ContentService.MimeType.JSON);
}

function reply(replyToken, message) {
  var url = "https://api.line.me/v2/bot/message/reply";
  UrlFetchApp.fetch(url, {
    "headers": {
      "Content-Type": "application/json; charset=UTF-8",
      "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN,
    },
    "method": "post",
    "payload": JSON.stringify({
      "replyToken": replyToken,
      "messages": [{
        "type": "text",
        "text": message,
      }],
    }),
  });
  return ContentService.createTextOutput(JSON.stringify({"content": "post ok"})).setMimeType(ContentService.MimeType.JSON);
}

//今日の予定
function getEvents() {
  var events = CalendarApp.getDefaultCalendar().getEventsForDay(new Date());
  var body = "今日の予定は";

  if (events.length === 0) {
    body += "ありません。";
    return body;
  }

  body += "\n";
  events.forEach(function(event) {
    var title = event.getTitle();
    var start = HmFormat(event.getStartTime());
    var end = HmFormat(event.getEndTime());
    body += "★" + title + ": " + start + " ~ " + end + "\n";
  });
  body += "です。";
  return body;
}

//７日間の予定
function notifyWeekly() {
  var  body = "7日間の予定は\n";
  var weekday = ["日", "月", "火", "水", "木", "金", "土"];
  for ( var i = 0;  i < 7;  i++ ) {

    var dt = new Date();
    dt.setDate(dt.getDate()+i);
    var events = CalendarApp.getDefaultCalendar().getEventsForDay(dt);
    body += Utilities.formatDate(dt, "JST", '★ MM/dd(' + weekday[dt.getDay()] + ')') + "\n";
    if (events.length === 0) {
      body += "ありません。\n";
    }

    events.forEach(function(event) {
      var title = event.getTitle();
      var start = HmFormat(event.getStartTime());
      var end = HmFormat(event.getEndTime());
      body += title + ": " + start + " ~ " + end + "\n";
    });
  }
    return body;
}

function HmFormat(date){
  return Utilities.formatDate(date, "JST", "HH:mm");
}

function wishList_add(data){
  let writeRow = wishlist.getLastRow + 1;
  wishlist.getRange(writeRow, 1).setValue(data.id);
  wishlist.getRange(writeRow, 2).setValue(data.category);
  wishlist.getRange(writeRow, 3).setValue(data.name);
  wishlist.getRange(writeRow, 4).setValue(data.link);
  wishlist.getRange(writeRow, 5).setValue(data.cost);
}

function wishlist_message(replyToken){
  let data_list = []
  let last = wishlist.getLastRow();
  let replyMessage = "";

  for(let i = 2; i <= last; i++){
    if(wishlist.getRange(i,6).getValue() === ""){
      let data = {}
      data.id = wishlist.getRange(i,1).getValue()
      data.name = wishlist.getRange(i,2).getValue()
      data_list.push(data)
    }
  }

  for(let i in data_list){
    let reply = "ID: " + data_list[i].id + "  " + data_list[i].name + "\n"
    replyMessage += reply;
  }

  reply(replyToken,replyMessage);
}
