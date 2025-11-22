/**
 * Routing.gs 全文修正 Ver1.1
 * Google Apps Script: Web アプリのルーティング定義
 */

/**
 * GET リクエストを処理
 * @param {Object} e イベントパラメータ
 * @returns {HtmlOutput} HTML 出力
 */
function doGet(e) {
  var page = e.parameter.page;
  Logger.log("doGet called, page: " + page);

  if (page === 'input') {
    var inputTmpl = HtmlService.createTemplateFromFile('input');
    return inputTmpl.evaluate();

  } else if (page === 'menu') {
    // メニュー画面へ遷移：ログインID をテンプレートに設定
    var menuTmpl = HtmlService.createTemplateFromFile('メニュー');
    menuTmpl.ログインID = e.parameter.ログインID;
    return menuTmpl.evaluate();

  } else {
    // デフォルト：ログイン画面
    var loginTmpl = HtmlService.createTemplateFromFile('login');
    return loginTmpl.evaluate();
  }
}

/**
 * POST リクエストを処理
 * @param {Object} e イベントパラメータ
 * @returns {HtmlOutput} HTML 出力
 */
function doPost(e) {
  Logger.log("doPost action: " + e.parameter.action);
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
  var action = e.parameter.action;

  switch (action) {
    // ログイン処理
    case "login":
      return ログイン処理(e, ss);

    // 時間外報告（年月選択画面へ）
    case "時間外報告":
      return handleTimeReport(e, ss);

    // 時間外集計メニュー
    case "時間外集計":
      var 集計Tmpl = HtmlService.createTemplateFromFile('集計メニュー');
      集計Tmpl.ログインID = e.parameter.ログインID;
      return 集計Tmpl.evaluate();

    // 所属別集計
    case "所属別集計":
      return handleShozokuBetsuSuikei(e, ss);

    // 年月別集計
    case "年月別集計":
      return handleNengetsuBetsuSuikei(e, ss);

    // select アクション（年月選択→入力画面）
    case "select":
      Logger.log("select action with 年月: " + e.parameter.年月);
      var selectTmpl = HtmlService.createTemplateFromFile('input');
      selectTmpl.ログインID   = e.parameter.ログインID;
      selectTmpl.選択年月     = e.parameter.年月;
      selectTmpl.前回時間外   = e.parameter.時間外;
      selectTmpl.前回振替時間 = e.parameter.振替時間;
      return selectTmpl.evaluate();

    // sendText アクション（入力値保存→メニュー画面へ）
    case "sendText":
      return handleSendText(e, ss);

    // deleteRecord アクション（データ削除）
    case "deleteRecord":
      return handleDeleteRecord(e, ss);

    // それ以外はログイン画面へ
    default:
      return HtmlService.createTemplateFromFile('login').evaluate();
  }
}

/**
 * ログイン処理
 * @param {Object} e イベントパラメータ
 * @param {Spreadsheet} ss スプレッドシート
 * @returns {HtmlOutput} メニュー画面 or エラーメッセージ
 */
function ログイン処理(e, ss) {
  var sheetユーザー名 = ss.getSheetByName("ユーザー名");
  var dataユーザー = sheetユーザー名
    .getRange(2, 1, sheetユーザー名.getLastRow() - 1, 1)
    .getValues();
  Logger.log("ユーザー名データ: " + JSON.stringify(dataユーザー));

  // フォーム値とシート値を文字列比較
  var loginValid = dataユーザー.some(function(row) {
    return String(row[0]) === String(e.parameter.user);
  });
  Logger.log("loginValid: " + loginValid);

  if (!loginValid) {
    return HtmlService.createHtmlOutput("Invalid login");
  }

  // ログイン成功後はメニュー画面
  var menuTmpl = HtmlService.createTemplateFromFile('メニュー');
  menuTmpl.ログインID = String(e.parameter.user);
  return menuTmpl.evaluate();
}
