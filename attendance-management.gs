var today;
var time;
var DATE_COL = 1;
var ATTENDANCE_TIME_COL = DATE_COL + 1;
var LEAVE_TIME_COL = ATTENDANCE_TIME_COL + 1;
var START_ROW = 2;

function doPost(e) {

  // ���b�Z�[�W����͂������������N�G�X�g�p�����[�^����Z�o
  calcDateTime(e);

  // SpreadSheet�ɏ�������
  writeSpreadsheet(e);
  
  // Slack API(chat.postMessage)�Ŏw�肵���`�����l���ɒʒm���b�Z�[�W
  postMessage(createSlackApp(), e);
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
  var spreadsheet = SpreadsheetApp.openByUrl('SPREADSHEETS_URL');
  //var sheet = spreadsheet.getSheetByName(e.parameter.user_name);
  var sheet = spreadsheet.getSheetByName("KenichiNagaoka");
  
  // A���2�s�ڂ���ŉ��s�܂ł̓��t���擾
  var lastRow = sheet.getLastRow();
  var dateValues = sheet.getRange(START_ROW, DATE_COL, lastRow).getValues();

  // ���ɓ��t�����͂���Ă���΂������ŉ��s�Ƃ��Ď��Ԃ���������
  for (var row in dateValues) {
    if (formatDate(new Date(dateValues[row])) == today) {
      writeTime(e, sheet, Number(row) + 2);
      return;
    }
  }
  
  // ���t�����͂���Ă��Ȃ���΍ŉ��s+1�ɓ��t�Ǝ��Ԃ���������
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
  var botName = PropertiesService.getScriptProperties().getProperty('BOT_NAME');
  var icon_url = PropertiesService.getScriptProperties().getProperty('ICON_URL');
  return slackApp.postMessage("#�ΑӊǗ�", createMessage(e), {
    username: botName,
    icon_url: icon_url
  });
}

function createMessage(e) {
  if (isAttendance(e)) {
    return today + "�̏o�Ў��Ԃ�" + time + "�ŋL�^���܂���:sunny:";
  }
  return today + "�̑ގЎ��Ԃ�" + time + "�ŋL�^���܂���:night_with_stars:";
}

function isAttendance(e) {
  return e.parameter.trigger_word == PropertiesService.getScriptProperties().getProperty('TRIGGER_WORD');;
}

function formatDate(date) {
  return Utilities.formatDate(date, 'JST', 'yyyy�NM��d��')
}