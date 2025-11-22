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
                      .getRange(2, 1, sheetテーブル.getLastRow()-1, 5)
                      .getValues();
  }
  Logger.log("テーブルデータ: " + JSON.stringify(テーブルデータ));

  var options = [];
  年月リスト.forEach(function(年月行) {
    var 年月Str = String(年月行[0]);
    var 時間外Val = "", 振替時間Val = "", 備考Val = "";
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
        備考Val = テーブルデータ[k][4] || "";
        break;
      }
    }
    options.push({年月: 年月Str, B: 時間外Val, C: 振替時間Val, D: 備考Val});
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
  var 時間外値 = parseFloat(e.parameter.時間外) || 0;
  var 振替時間値 = parseFloat(e.parameter.振替時間) || 0;
  var 備考値 = e.parameter.備考 || "";

  // 入力バリデーション
  if (時間外値 < 0 || 振替時間値 < 0) {
    return HtmlService.createHtmlOutput("エラー: 負の値は入力できません");
  }
  if (時間外値 > 500 || 振替時間値 > 500) {
    return HtmlService.createHtmlOutput("エラー: 入力値が大きすぎます（最大500時間）");
  }

  Logger.log("sendText action, 入力ID: " + 入力ID + ", 選択年月: " + 選択年月 + ", 時間外: " + 時間外値 + ", 振替時間: " + 振替時間値 + ", 備考: " + 備考値);

  var sheetテーブル = ss.getSheetByName("テーブル");
  var テーブルデータ = [];
  if(sheetテーブル.getLastRow() > 1) {
    テーブルデータ = sheetテーブル
                      .getRange(2, 1, sheetテーブル.getLastRow()-1, 5)
                      .getValues();
  }
  Logger.log("既存テーブルデータ: " + JSON.stringify(テーブルデータ));

  var 上書き済み = false;
  var メッセージ = "";
  for(var i = 0; i < テーブルデータ.length; i++){
    var テーブル年月 = "";
    if(テーブルデータ[i][0] instanceof Date) {
      テーブル年月 = Utilities.formatDate(テーブルデータ[i][0], Session.getScriptTimeZone(), "yyyy年M月");
    } else {
      テーブル年月 = String(テーブルデータ[i][0]);
    }
    var テーブルユーザー = String(テーブルデータ[i][1]);
    if(テーブル年月 == String(選択年月) && テーブルユーザー == String(入力ID)) {
       sheetテーブル.getRange(i+2, 3, 1, 3).setValues([[時間外値, 振替時間値, 備考値]]);
       Logger.log("データ上書き：行 " + (i+2));
       上書き済み = true;
       メッセージ = "✅ " + 選択年月 + " のデータを更新しました";
       break;
    }
  }
  if(!上書き済み) {
    sheetテーブル.appendRow([選択年月, 入力ID, 時間外値, 振替時間値, 備考値]);
    Logger.log("新規行追加");
    メッセージ = "✅ " + 選択年月 + " のデータを保存しました";
  }

  // **成功メッセージ付きでメニュー画面に戻る**
  var tmpl = HtmlService.createTemplateFromFile('メニュー');
  tmpl.ログインID = 入力ID;
  tmpl.successMessage = メッセージ;
  return tmpl.evaluate();
}

/**
 * 時間外データ削除処理
 */
function handleDeleteRecord(e, ss) {
  var 入力ID = e.parameter.ログインID;
  var 選択年月 = e.parameter.年月;

  Logger.log("deleteRecord action, 入力ID: " + 入力ID + ", 選択年月: " + 選択年月);

  var sheetテーブル = ss.getSheetByName("テーブル");
  var テーブルデータ = [];
  if(sheetテーブル.getLastRow() > 1) {
    テーブルデータ = sheetテーブル
                      .getRange(2, 1, sheetテーブル.getLastRow()-1, 5)
                      .getValues();
  }

  var 削除済み = false;
  for(var i = 0; i < テーブルデータ.length; i++){
    var テーブル年月 = "";
    if(テーブルデータ[i][0] instanceof Date) {
      テーブル年月 = Utilities.formatDate(テーブルデータ[i][0], Session.getScriptTimeZone(), "yyyy年M月");
    } else {
      テーブル年月 = String(テーブルデータ[i][0]);
    }
    var テーブルユーザー = String(テーブルデータ[i][1]);
    if(テーブル年月 == String(選択年月) && テーブルユーザー == String(入力ID)) {
       sheetテーブル.deleteRow(i+2);
       Logger.log("データ削除：行 " + (i+2));
       削除済み = true;
       break;
    }
  }

  // **削除結果メッセージ付きでメニュー画面に戻る**
  var tmpl = HtmlService.createTemplateFromFile('メニュー');
  tmpl.ログインID = 入力ID;
  if (削除済み) {
    tmpl.successMessage = "🗑️ " + 選択年月 + " のデータを削除しました";
  } else {
    tmpl.successMessage = "⚠️ 削除対象のデータが見つかりませんでした";
  }
  return tmpl.evaluate();
}
