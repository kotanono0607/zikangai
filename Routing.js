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
  var action = e.parameter.action;
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");

  Logger.log("doGet called, page: " + page + ", action: " + action);

  // QRログイン処理（GETリクエスト）
  if (action === 'qrLogin') {
    return handleQrLogin(e, ss);
  }

  if (page === 'input') {
    var inputTmpl = HtmlService.createTemplateFromFile('input');
    return inputTmpl.evaluate();

  } else if (page === 'menu') {
    // メニュー画面へ遷移：ユーザー情報を取得してテンプレートに設定
    var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
    var sheet = ss.getSheetByName("ユーザー名");
    var ログインID = e.parameter.ログインID;

    var menuTmpl = HtmlService.createTemplateFromFile('メニュー');
    menuTmpl.ログインID = ログインID;

    // ユーザー情報を取得
    if (sheet.getLastRow() >= 2) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]) === String(ログインID)) {
          menuTmpl.氏名 = String(data[i][1]);
          menuTmpl.権限 = String(data[i][5]);
          break;
        }
      }
    }

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

      // ユーザー情報を取得（氏名を表示するため）
      var ログインID = e.parameter.ログインID;
      var 氏名 = ログインID; // デフォルトはログインID
      var userSheet = ss.getSheetByName("ユーザー名");
      if (userSheet && userSheet.getLastRow() >= 2) {
        var userDataList = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 2).getValues();
        for (var i = 0; i < userDataList.length; i++) {
          if (String(userDataList[i][0]) === String(ログインID)) {
            氏名 = String(userDataList[i][1]);
            break;
          }
        }
      }

      var selectTmpl = HtmlService.createTemplateFromFile('input');
      selectTmpl.ログインID   = ログインID;
      selectTmpl.氏名         = 氏名;
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

    // ユーザー管理（管理者専用）
    case "userList":
      return handleUserList(e, ss);

    case "addUserForm":
      return handleAddUserForm(e, ss);

    case "editUser":
      return handleEditUser(e, ss);

    case "saveUser":
      return handleSaveUser(e, ss);

    case "retireUser":
      return handleRetireUser(e, ss);

    // 所属変更（全ユーザー）
    case "profileEdit":
      return handleProfileEdit(e, ss);

    case "profileUpdate":
      return handleProfileUpdate(e, ss);

    // 所属マスタ管理（管理者専用）
    case "departmentMaster":
      return handleDepartmentMaster(e, ss);

    case "addDepartment":
      return handleAddDepartment(e, ss);

    case "deleteDepartment":
      return handleDeleteDepartment(e, ss);

    // CSV エクスポート/インポート（管理者専用）
    case "csvExport":
      return handleCsvExport(e, ss);

    case "csvImport":
      return handleCsvImport(e, ss);

    // QRコード生成（管理者専用）
    case "generateQrPdf":
      return handleGenerateQrPdf(e, ss);

    case "generateUserQr":
      return handleGenerateUserQr(e, ss);

    // 分析・レポート
    case "analytics":
      return handleAnalytics(e, ss);

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

  // 全列取得（ID, 氏名, 所属１, 所属２, 状態, 権限）
  var dataユーザー = sheetユーザー名
    .getRange(2, 1, sheetユーザー名.getLastRow() - 1, 6)
    .getValues();

  var 入力ID = String(e.parameter.user);

  Logger.log("ログイン試行: " + 入力ID);

  // ユーザー情報を検索
  var ユーザー情報 = null;
  for (var i = 0; i < dataユーザー.length; i++) {
    if (String(dataユーザー[i][0]) === 入力ID) {
      ユーザー情報 = {
        ログインID: String(dataユーザー[i][0]),
        氏名: String(dataユーザー[i][1]),
        所属１: String(dataユーザー[i][2]),
        所属２: String(dataユーザー[i][3]),
        状態: String(dataユーザー[i][4]),
        権限: String(dataユーザー[i][5])
      };
      break;
    }
  }

  // ユーザーが見つからない
  if (!ユーザー情報) {
    Logger.log("ユーザーが見つかりません: " + 入力ID);
    return HtmlService.createHtmlOutput("ログインIDが違います");
  }

  // 退職者チェック
  if (ユーザー情報.状態 === "退職") {
    Logger.log("退職者のログイン試行: " + 入力ID);
    return HtmlService.createHtmlOutput("このアカウントは利用できません（退職済み）");
  }

  // ログイン成功
  Logger.log("ログイン成功: " + 入力ID + " (" + ユーザー情報.権限 + ")");

  var menuTmpl = HtmlService.createTemplateFromFile('メニュー');
  menuTmpl.ログインID = ユーザー情報.ログインID;
  menuTmpl.氏名 = ユーザー情報.氏名;
  menuTmpl.権限 = ユーザー情報.権限;
  return menuTmpl.evaluate();
}
