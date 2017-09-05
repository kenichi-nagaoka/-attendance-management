var today;
var time;
var body;
var DATE_COL = 1;
var ATTENDANCE_TIME_COL = DATE_COL + 1;
var LEAVE_TIME_COL = ATTENDANCE_TIME_COL + 1;
var START_ROW = 2;

function doPost(e) {

  // メッセージを入力した日時をリクエストパラメータから算出
  calcDateTime(e);

  writeSpreadsheet(e);
  
  // Slack API(chat.postMessage)で指定したチャンネルに通知メッセージ
  postMessage(createSlackApp(), e);
  
  MailApp.sendEmail(getProperty("TO"), "【勤怠連絡】" + today, body);
}

function createSlackApp() {
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');
  return SlackApp.create(token);
}

function calcDateTime(e) {
  var date = new Date(e.parameter.timestamp*1000);
  time = ('0' + date.getHours()).slice(-2) + ":" + ('0' + date.getMinutes()).slice(-2);
  today = formatDate(date);
}

function writeSpreadsheet(e) {
  var spreadsheet = SpreadsheetApp.openByUrl(getProperty("SPREADSHEETS_URL"));
  //var sheet = spreadsheet.getSheetByName(e.parameter.user_name);
  var sheet = spreadsheet.getSheetByName("KenichiNagaoka");
  
  // A列の2行目から最下行までの日付を取得
  var lastRow = sheet.getLastRow();
  var dateValues = sheet.getRange(START_ROW, DATE_COL, lastRow).getValues();

  // 既に日付が入力されていればそこを最下行として時間を書き込む
  for (var row in dateValues) {
    if (formatDate(new Date(dateValues[row])) == today) {
      writeTime(e, sheet, Number(row) + 2);
      return;
    }
  }
  
  // 日付が入力されていなければ最下行+1に日付と時間を書き込む
  var targetRow = lastRow + 1;
  sheet.getRange(targetRow, DATE_COL).setValue(today);
  writeTime(e, sheet, targetRow);
}

function writeTime(e, sheet, targetRow) { 
  if (isAttendance(e)) {
    sheet.getRange(targetRow, ATTENDANCE_TIME_COL).setValue(time);
  } else {
    sheet.getRange(targetRow, LEAVE_TIME_COL).setValue(time);
  }
}

function postMessage(slackApp, e) {
  var botName = getProperty("BOT_NAME");
  var icon_url = getProperty("ICON_URL");
  return slackApp.postMessage("#勤怠管理", createMessage(e), {
    username: botName,
    icon_url: icon_url
  });
}

function createMessage(e) {
  if (isAttendance(e)) {
    body = e.parameter.user_name + "は" + time + "に出社しました。";
    return today + "の出社時間は" + time + "で記録しました。";
  }
  body = e.parameter.user_name + "は" + time + "に退社しました。";
  return today + "の退社時間は" + time + "で記録しました。";
}

function isAttendance(e) {
  return e.parameter.trigger_word == getProperty("TRIGGER_WORD");;
}

function formatDate(date) {
  return Utilities.formatDate(date, 'JST', 'yyyy年M月d日')
}

function getProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}
