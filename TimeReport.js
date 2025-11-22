//=================================================================================================================================================
// TimeReport.gs

/**
 * 指定された年月が編集可能かチェック（サーバーサイドバリデーション）
 */
function is年月編集可能(年月文字列) {
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1; // 1-12
  var currentFiscalYear = (currentMonth >= 4) ? currentYear : currentYear - 1;
  var is猶予期間 = (currentMonth >= 4 && currentMonth <= 5);

  // 開始年度（固定：2024年度）
  var START_FISCAL_YEAR = 2024;

  // 年月文字列から年と月を抽出（例: "2025年11月" → 2025, 11）
  var match = 年月文字列.match(/(\d{4})年(\d{1,2})月/);
  if (!match) return false;

  var targetYear = parseInt(match[1]);
  var targetMonth = parseInt(match[2]);

  // 対象年月の年度を判定
  var targetFiscalYear = (targetMonth >= 4) ? targetYear : targetYear - 1;

  // 2024年度より前は編集不可
  if (targetFiscalYear < START_FISCAL_YEAR) {
    return false;
  }

  // 年度オフセットを計算
  var fiscalYearOffset = targetFiscalYear - currentFiscalYear;

  // 編集可否を判定
  if (fiscalYearOffset === 0) {
    // 当年度：常に編集可能
    return true;
  } else if (fiscalYearOffset === -1 && is猶予期間) {
    // 前年度：4月〜5月のみ編集可能
    return true;
  }

  // それ以外：編集不可
  return false;
}

/**
 * 年月リストを自動生成（年度ベース：4月〜3月）
 * 2024年度から当年度まで（最大10年度分）を生成
 * 各年月に編集可能フラグと年度情報を付与
 */
function 年月リスト自動生成() {
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1; // 1-12

  // 現在の年度を判定（4月以降なら当年、3月以前なら前年）
  var currentFiscalYear = (currentMonth >= 4) ? currentYear : currentYear - 1;

  // 5月末までは前年度も編集可能
  var is猶予期間 = (currentMonth >= 4 && currentMonth <= 5);

  // 開始年度（固定：2024年度）
  var START_FISCAL_YEAR = 2024;

  var 年度別データ = {};

  // 2024年度から当年度までを生成（最大10年度）
  var startYear = Math.max(START_FISCAL_YEAR, currentFiscalYear - 9); // 10年以上前は表示しない
  for (var fiscalYear = startYear; fiscalYear <= currentFiscalYear; fiscalYear++) {
    var 年度名 = fiscalYear + "年度";

    // 編集可否を判定
    var editable = false;
    if (fiscalYear === currentFiscalYear) {
      // 当年度：常に編集可能
      editable = true;
    } else if (fiscalYear === currentFiscalYear - 1 && is猶予期間) {
      // 前年度：4月〜5月のみ編集可能
      editable = true;
    }
    // それ以外：閲覧のみ

    年度別データ[年度名] = {
      fiscalYear: fiscalYear,
      editable: editable,
      months: []
    };

    // 4月〜12月（当年）
    for (var month = 4; month <= 12; month++) {
      年度別データ[年度名].months.push({
        年月: fiscalYear + "年" + month + "月",
        editable: editable
      });
    }

    // 1月〜3月（翌年）
    for (var month = 1; month <= 3; month++) {
      年度別データ[年度名].months.push({
        年月: (fiscalYear + 1) + "年" + month + "月",
        editable: editable
      });
    }
  }

  Logger.log("自動生成された年度別データ: " + JSON.stringify(年度別データ));
  return 年度別データ;
}

function handleTimeReport(e, ss) {
  var user = e.parameter.ログインID;
  Logger.log("時間外報告 action triggered for user: " + user);

  // ユーザー情報を取得（氏名を表示するため）
  var sheetユーザー = ss.getSheetByName("ユーザー名");
  var 氏名 = user; // デフォルトはログインID
  if (sheetユーザー && sheetユーザー.getLastRow() >= 2) {
    var userData = sheetユーザー.getRange(2, 1, sheetユーザー.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < userData.length; i++) {
      if (String(userData[i][0]) === String(user)) {
        氏名 = String(userData[i][1]);
        break;
      }
    }
  }

  var sheetテーブル = ss.getSheetByName("テーブル");

  // 年月リストを年度別に自動生成
  var 年度別データ = 年月リスト自動生成();

  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1;
  var currentFiscalYear = (currentMonth >= 4) ? currentYear : currentYear - 1;

  var テーブルデータ = [];
  if(sheetテーブル.getLastRow() > 1) {
    テーブルデータ = sheetテーブル
                      .getRange(2, 1, sheetテーブル.getLastRow()-1, 5)
                      .getValues();
  }
  Logger.log("テーブルデータ: " + JSON.stringify(テーブルデータ));

  // 年度別にデータをマージ
  var 年度配列 = [];
  for (var 年度名 in 年度別データ) {
    var 年度Info = 年度別データ[年度名];
    var monthsWithData = [];

    年度Info.months.forEach(function(monthObj) {
      var 年月Str = monthObj.年月;
      var 時間外Val = "", 振替時間Val = "", 備考Val = "";

      // テーブルから既存データを検索
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

      monthsWithData.push({
        年月: 年月Str,
        時間外: 時間外Val,
        振替時間: 振替時間Val,
        備考: 備考Val,
        editable: monthObj.editable
      });
    });

    年度配列.push({
      年度名: 年度名,
      fiscalYear: 年度Info.fiscalYear,
      editable: 年度Info.editable,
      isCurrentYear: (年度Info.fiscalYear === currentFiscalYear),
      months: monthsWithData
    });
  }

  // 年度を降順にソート（新しい年度が上）
  年度配列.sort(function(a, b) { return b.fiscalYear - a.fiscalYear; });

  Logger.log("年度配列: " + JSON.stringify(年度配列));

  var tmpl = HtmlService.createTemplateFromFile('select');
  tmpl.ログインID = user;
  tmpl.氏名 = 氏名;
  tmpl.年度データ = 年度配列;
  return tmpl.evaluate();
}
function handleSendText(e, ss) {
  var 入力ID = e.parameter.入力ID;
  var 選択年月 = e.parameter.年月;
  var 時間外値 = parseFloat(e.parameter.時間外) || 0;
  var 振替時間値 = parseFloat(e.parameter.振替時間) || 0;
  var 備考値 = e.parameter.備考 || "";

  // 編集権限チェック（サーバーサイドバリデーション）
  if (!is年月編集可能(選択年月)) {
    Logger.log("編集権限エラー: " + 選択年月 + " は編集不可期間です");
    return HtmlService.createHtmlOutput("❌ エラー: " + 選択年月 + " は編集できません（確定済み期間）");
  }

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
  // ユーザー情報を取得
  var sheetユーザー = ss.getSheetByName("ユーザー名");
  var 氏名 = 入力ID;
  var 権限 = "";
  if (sheetユーザー && sheetユーザー.getLastRow() >= 2) {
    var userData = sheetユーザー.getRange(2, 1, sheetユーザー.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < userData.length; i++) {
      if (String(userData[i][0]) === String(入力ID)) {
        氏名 = String(userData[i][1]);
        権限 = String(userData[i][5]);
        break;
      }
    }
  }

  var tmpl = HtmlService.createTemplateFromFile('メニュー');
  tmpl.ログインID = 入力ID;
  tmpl.氏名 = 氏名;
  tmpl.権限 = 権限;
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

  // 編集権限チェック（サーバーサイドバリデーション）
  if (!is年月編集可能(選択年月)) {
    Logger.log("削除権限エラー: " + 選択年月 + " は編集不可期間です");
    var tmpl = HtmlService.createTemplateFromFile('メニュー');
    tmpl.ログインID = 入力ID;
    tmpl.successMessage = "❌ エラー: " + 選択年月 + " は削除できません（確定済み期間）";
    return tmpl.evaluate();
  }

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
  // ユーザー情報を取得
  var sheetユーザー = ss.getSheetByName("ユーザー名");
  var 氏名 = 入力ID;
  var 権限 = "";
  if (sheetユーザー && sheetユーザー.getLastRow() >= 2) {
    var userData = sheetユーザー.getRange(2, 1, sheetユーザー.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < userData.length; i++) {
      if (String(userData[i][0]) === String(入力ID)) {
        氏名 = String(userData[i][1]);
        権限 = String(userData[i][5]);
        break;
      }
    }
  }

  var tmpl = HtmlService.createTemplateFromFile('メニュー');
  tmpl.ログインID = 入力ID;
  tmpl.氏名 = 氏名;
  tmpl.権限 = 権限;
  if (削除済み) {
    tmpl.successMessage = "🗑️ " + 選択年月 + " のデータを削除しました";
  } else {
    tmpl.successMessage = "⚠️ 削除対象のデータが見つかりませんでした";
  }
  return tmpl.evaluate();
}
