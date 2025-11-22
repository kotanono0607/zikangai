//=================================================================================================================================================
// TimeReport.gs
function handleTimeReport(e, ss) {
  var user = e.parameter.ログインID;
  Logger.log("時間外報告 action triggered for user: " + user);
  
  var sheet年月 = ss.getSheetByName("年月");
  var sheetテーブル = ss.getSheetByName("テーブル");

  var 年月リスト = sheet年月
                    .getRange(2, 1, sheet年月.getLastRow()-1, 1)
                    .getValues();
  Logger.log("年月リスト: " + JSON.stringify(年月リスト));

  var テーブルデータ = [];
  if(sheetテーブル.getLastRow() > 1) {
    テーブルデータ = sheetテーブル
                      .getRange(2, 1, sheetテーブル.getLastRow()-1, 4)
                      .getValues();
  }
  Logger.log("テーブルデータ: " + JSON.stringify(テーブルデータ));

  var options = [];
  年月リスト.forEach(function(年月行) {
    var 年月Str = String(年月行[0]);
    var 時間外Val = "", 振替時間Val = "";
    for(var k = 0; k < テーブルデータ.length; k++) {
      var テーブル年月 = "";
      if(テーブルデータ[k][0] instanceof Date) {
        テーブル年月 = Utilities.formatDate(テーブルデータ[k][0], Session.getScriptTimeZone(), "yyyy年M月");
      } else {
        テーブル年月 = String(テーブルデータ[k][0]);
      }
      var テーブルユーザー = String(テーブルデータ[k][1]);
      if(テーブル年月 == 年月Str && テーブルユーザー == String(user)) {
        時間外Val = テーブルデータ[k][2];
        振替時間Val = テーブルデータ[k][3];
        break;
      }
    }
    options.push({年月: 年月Str, B: 時間外Val, C: 振替時間Val});
  });
  Logger.log("options: " + JSON.stringify(options));

  var tmpl = HtmlService.createTemplateFromFile('select');
  tmpl.ログインID = user;
  tmpl.options = options;
  return tmpl.evaluate();
}
function handleSendText(e, ss) {
  var 入力ID = e.parameter.入力ID;
  var 選択年月 = e.parameter.年月;
  var 時間外値 = e.parameter.時間外;
  var 振替時間値 = e.parameter.振替時間;
  
  Logger.log("sendText action, 入力ID: " + 入力ID + ", 選択年月: " + 選択年月 + ", 時間外: " + 時間外値 + ", 振替時間: " + 振替時間値);

  var sheetテーブル = ss.getSheetByName("テーブル");
  var テーブルデータ = [];
  if(sheetテーブル.getLastRow() > 1) {
    テーブルデータ = sheetテーブル
                      .getRange(2, 1, sheetテーブル.getLastRow()-1, 4)
                      .getValues();
  }
  Logger.log("既存テーブルデータ: " + JSON.stringify(テーブルデータ));

  var 上書き済み = false;
  for(var i = 0; i < テーブルデータ.length; i++){
    var テーブル年月 = "";
    if(テーブルデータ[i][0] instanceof Date) {
      テーブル年月 = Utilities.formatDate(テーブルデータ[i][0], Session.getScriptTimeZone(), "yyyy年M月");
    } else {
      テーブル年月 = String(テーブルデータ[i][0]);
    }
    var テーブルユーザー = String(テーブルデータ[i][1]);
    if(テーブル年月 == String(選択年月) && テーブルユーザー == String(入力ID)) {
       sheetテーブル.getRange(i+2, 3, 1, 2).setValues([[時間外値, 振替時間値]]);
       Logger.log("データ上書き：行 " + (i+2));
       上書き済み = true;
       break;
    }
  }
  if(!上書き済み) {
    sheetテーブル.appendRow([選択年月, 入力ID, 時間外値, 振替時間値]);
    Logger.log("新規行追加");
  }

  // **時間外報告後にメニュー画面に戻るように修正**
  var tmpl = HtmlService.createTemplateFromFile('メニュー');
  tmpl.ログインID = 入力ID;
  return tmpl.evaluate();
}
