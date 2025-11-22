/**
 * 所属マスタ管理機能
 *
 * 階層構造：所属1 → 所属2
 * シート構造：
 *   A列: 所属1
 *   B列: 所属2
 */

/**
 * 所属マスタシートを取得または作成
 */
function get所属マスタSheet(ss) {
  var sheetName = "所属マスタ";
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 2).setValues([["所属１", "所属２"]]);
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold");
  }

  return sheet;
}

/**
 * 既存ユーザーから所属データを自動抽出
 * Setup.js から呼び出される
 */
function 所属マスタ自動抽出(ss) {
  var userSheet = ss.getSheetByName("ユーザー名");
  var masterSheet = get所属マスタSheet(ss);

  if (!userSheet || userSheet.getLastRow() < 2) {
    Logger.log("ユーザーデータが存在しません");
    return;
  }

  // 既存のユーザーデータから所属ペアを抽出
  var userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 6).getValues();
  var pairs = {}; // {所属1: Set(所属2)}

  for (var i = 0; i < userData.length; i++) {
    var 所属1 = String(userData[i][2]).trim();
    var 所属2 = String(userData[i][3]).trim();

    if (所属1 && 所属2) {
      if (!pairs[所属1]) {
        pairs[所属1] = {};
      }
      pairs[所属1][所属2] = true;
    }
  }

  // 所属マスタシートに書き込み
  masterSheet.getRange(2, 1, masterSheet.getLastRow(), 2).clearContent();

  var rows = [];
  for (var dept1 in pairs) {
    for (var dept2 in pairs[dept1]) {
      rows.push([dept1, dept2]);
    }
  }

  if (rows.length > 0) {
    masterSheet.getRange(2, 1, rows.length, 2).setValues(rows);
    Logger.log("所属マスタ自動抽出完了: " + rows.length + "件");
  }
}

/**
 * 所属マスタデータを階層構造で取得
 * @return {Object} {所属1: [所属2の配列]}
 */
function get所属マスタ(ss) {
  var sheet = get所属マスタSheet(ss);

  if (sheet.getLastRow() < 2) {
    return {};
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var hierarchy = {};

  for (var i = 0; i < data.length; i++) {
    var 所属1 = String(data[i][0]).trim();
    var 所属2 = String(data[i][1]).trim();

    if (所属1 && 所属2) {
      if (!hierarchy[所属1]) {
        hierarchy[所属1] = [];
      }
      if (hierarchy[所属1].indexOf(所属2) === -1) {
        hierarchy[所属1].push(所属2);
      }
    }
  }

  return hierarchy;
}

/**
 * 所属1のリストを取得
 */
function get所属1リスト(ss) {
  var hierarchy = get所属マスタ(ss);
  var list = [];

  for (var dept1 in hierarchy) {
    list.push(dept1);
  }

  return list.sort();
}

/**
 * 指定された所属1に対応する所属2のリストを取得
 */
function get所属2リスト(ss, 所属1) {
  var hierarchy = get所属マスタ(ss);
  return hierarchy[所属1] || [];
}

/**
 * 所属マスタ管理画面の表示
 */
function handleDepartmentMaster(e, ss) {
  var ログインID = e.parameter.ログインID;

  // 管理者チェック
  if (!isAdmin(ログインID, ss)) {
    var errorTmpl = HtmlService.createTemplate('<p>この機能は管理者のみ利用可能です</p>');
    return errorTmpl.evaluate();
  }

  var tmpl = HtmlService.createTemplateFromFile('department-master');
  tmpl.ログインID = ログインID;
  tmpl.所属マスタ = get所属マスタ(ss);

  return tmpl.evaluate();
}

/**
 * 所属ペアの追加
 */
function handleAddDepartment(e, ss) {
  var ログインID = e.parameter.ログインID;

  if (!isAdmin(ログインID, ss)) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "権限がありません"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var 所属1 = e.parameter.所属1;
  var 所属2 = e.parameter.所属2;

  if (!所属1 || !所属2) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "所属1と所属2を入力してください"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = get所属マスタSheet(ss);

  // 既存チェック
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === 所属1 && String(data[i][1]) === 所属2) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: "既に登録されています"
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }

  sheet.appendRow([所属1, 所属2]);

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: "追加しました"
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 所属ペアの削除
 */
function handleDeleteDepartment(e, ss) {
  var ログインID = e.parameter.ログインID;

  if (!isAdmin(ログインID, ss)) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "権限がありません"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var 所属1 = e.parameter.所属1;
  var 所属2 = e.parameter.所属2;

  var sheet = get所属マスタSheet(ss);

  if (sheet.getLastRow() < 2) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "データが存在しません"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === 所属1 && String(data[i][1]) === 所属2) {
      sheet.deleteRow(i + 2);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: "削除しました"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    message: "該当データが見つかりません"
  })).setMimeType(ContentService.MimeType.JSON);
}
