//Aggregation.gs

function handleShozokuBetsuSuikei(e, ss) {
  var sheetユーザー名 = ss.getSheetByName("ユーザー名");
  var sheetテーブル = ss.getSheetByName("テーブル");

  // ユーザー名シートからデータ取得
  var userRows = sheetユーザー名
                   .getRange(2, 1, sheetユーザー名.getLastRow() - 1, 4)
                   .getValues();
  Logger.log("ユーザー名行数: " + userRows.length);

  var userDeptMap = {};
  // ユーザーごとに所属情報をマッピング
  userRows.forEach(function(row) {
    var id = String(row[0]);
    userDeptMap[id] = {課: row[2] || "未設定", 係: row[3] || "未設定"};
  });

  // テーブルシートからデータ取得
  var tableRows = [];
  if (sheetテーブル.getLastRow() > 1) {
    tableRows = sheetテーブル
                  .getRange(2, 1, sheetテーブル.getLastRow() - 1, 4)
                  .getValues();
  }
  Logger.log("テーブル行数: " + tableRows.length);

  var hierarchical = {};
  tableRows.forEach(function(row) {
    var dateVal = row[0];
    var login = String(row[1]);
    var overtime = Number(row[2]) || 0;
    var transfer = Number(row[3]) || 0;

    var ym = "";
    if (dateVal instanceof Date) {
      ym = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy年M月");
    } else {
      ym = String(dateVal);
    }

    var userDept = userDeptMap[login] || {課:"未設定", 係:"未設定"};
    var course = userDept.課;
    var section = userDept.係;
    if (!hierarchical[course]) hierarchical[course] = {};
    if (!hierarchical[course][section]) hierarchical[course][section] = {};
    if (!hierarchical[course][section][ym]) {
      hierarchical[course][section][ym] = {時間外: 0, 振替時間: 0};
    }
    hierarchical[course][section][ym].時間外 += overtime;
    hierarchical[course][section][ym].振替時間 += transfer;
  });

  Logger.log("階層集計結果: " + JSON.stringify(hierarchical));

  var tmpl = HtmlService.createTemplateFromFile('所属別集計結果');
  tmpl.hierarchical = hierarchical;
  return tmpl.evaluate();
}

function handleNengetsuBetsuSuikei(e, ss) {
  var sheetテーブル = ss.getSheetByName("テーブル");

  // テーブルシートからデータ取得（年月, ログインID, 時間外, 振替時間）
  var tableRows = [];
  if(sheetテーブル.getLastRow() > 1) {
    tableRows = sheetテーブル
                  .getRange(2, 1, sheetテーブル.getLastRow() - 1, 4)
                  .getValues();
  }

  var aggregationByYm = {};

  tableRows.forEach(function(row) {
    var dateVal = row[0];
    var overtime = Number(row[2]) || 0;
    var transfer = Number(row[3]) || 0;

    // 日付を「yyyy年M月」形式に変換
    var ym = "";
    if(dateVal instanceof Date) {
      ym = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy年M月");
    } else {
      ym = String(dateVal);
    }

    if(!aggregationByYm[ym]) {
      aggregationByYm[ym] = {時間外: 0, 振替時間: 0};
    }
    aggregationByYm[ym].時間外 += overtime;
    aggregationByYm[ym].振替時間 += transfer;
  });

  Logger.log("年月別集計: " + JSON.stringify(aggregationByYm));

  var tmpl = HtmlService.createTemplateFromFile('年月別集計結果');
  tmpl.aggregationByYm = aggregationByYm;
  return tmpl.evaluate();
}
