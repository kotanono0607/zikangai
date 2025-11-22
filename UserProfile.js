//=================================================================================================================================================
// UserProfile.js
// 所属変更機能（全ユーザー利用可）

/**
 * 所属変更画面表示
 */
function handleProfileEdit(e, ss) {
  var ログインID = e.parameter.ログインID;

  var sheet = ss.getSheetByName("ユーザー名");
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

  // 自分のユーザー情報を検索
  var ユーザー情報 = null;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(ログインID)) {
      ユーザー情報 = {
        ログインID: String(data[i][0]),
        氏名: String(data[i][1]),
        所属１: String(data[i][2]),
        所属２: String(data[i][3]),
        状態: String(data[i][4]),
        権限: String(data[i][5])
      };
      break;
    }
  }

  if (!ユーザー情報) {
    return HtmlService.createHtmlOutput("❌ ユーザー情報が見つかりません");
  }

  Logger.log("所属変更画面表示: " + ログインID);

  var tmpl = HtmlService.createTemplateFromFile('profile');
  tmpl.ログインID = ログインID;
  tmpl.ユーザー = ユーザー情報;
  return tmpl.evaluate();
}

/**
 * 所属情報更新
 */
function handleProfileUpdate(e, ss) {
  var ログインID = e.parameter.ログインID;
  var 所属１ = e.parameter.shozoku1;
  var 所属２ = e.parameter.shozoku2;

  // バリデーション
  if (!所属１) {
    return HtmlService.createHtmlOutput("❌ 所属１は必須です");
  }

  var sheet = ss.getSheetByName("ユーザー名");
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

  // 自分の情報を更新
  var updated = false;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(ログインID)) {
      var rowIndex = i + 2;
      sheet.getRange(rowIndex, 3).setValue(所属１);
      sheet.getRange(rowIndex, 4).setValue(所属２);
      updated = true;
      Logger.log("所属情報更新: " + ログインID + " → " + 所属１ + " / " + 所属２);
      break;
    }
  }

  if (!updated) {
    return HtmlService.createHtmlOutput("❌ ユーザー情報が見つかりません");
  }

  // 成功メッセージと共にメニューに戻る
  var tmpl = HtmlService.createTemplateFromFile('メニュー');
  tmpl.ログインID = ログインID;
  tmpl.氏名 = data[i][1]; // 更新したユーザーの氏名
  tmpl.権限 = data[i][5]; // 権限
  tmpl.successMessage = "✅ 所属情報を更新しました";
  return tmpl.evaluate();
}
