/**
 * Analytics.js - 時間外分析機能
 * 個人別・所属別の時間外集計とレポート機能
 */

/**
 * 分析画面を表示
 */
function handleAnalytics(e, ss) {
  var ログインID = e.parameter.ログインID;

  // ユーザー情報を取得
  var userSheet = ss.getSheetByName("ユーザー名");
  var userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 6).getValues();

  var ユーザー情報 = null;
  for (var i = 0; i < userData.length; i++) {
    if (String(userData[i][0]) === String(ログインID)) {
      ユーザー情報 = {
        ログインID: String(userData[i][0]),
        氏名: String(userData[i][1]),
        所属１: String(userData[i][2]),
        所属２: String(userData[i][3]),
        状態: String(userData[i][4]),
        権限: String(userData[i][5])
      };
      break;
    }
  }

  if (!ユーザー情報) {
    return HtmlService.createHtmlOutput("ユーザー情報が取得できません");
  }

  // 分析画面テンプレート
  var tmpl = HtmlService.createTemplateFromFile('analytics');
  tmpl.ログインID = ユーザー情報.ログインID;
  tmpl.氏名 = ユーザー情報.氏名;
  tmpl.権限 = ユーザー情報.権限;
  tmpl.所属１ = ユーザー情報.所属１;
  tmpl.所属２ = ユーザー情報.所属２;

  return tmpl.evaluate();
}

/**
 * 個人の月別時間外データを取得（年度ベース）
 */
function getMyMonthlyData(ログインID, 年度) {
  Logger.log("getMyMonthlyData called: ログインID=" + ログインID + ", 年度=" + 年度);

  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
  var sheet = ss.getSheetByName("テーブル");

  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log("データなし: シートが空またはヘッダーのみ");
    return [];
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  Logger.log("取得行数: " + data.length);

  var monthlyData = [];
  var allUserData = []; // デバッグ用：このユーザーの全データ

  // 年度の開始年と終了年を計算
  var startYear = parseInt(年度);
  var endYear = startYear + 1;

  Logger.log("対象年度: " + startYear + "年度 (期間: " + startYear + "年4月〜" + endYear + "年3月)");

  for (var i = 0; i < data.length; i++) {
    var 年月 = String(data[i][0]);
    var userID = String(data[i][1]).trim();
    var 時間外 = parseFloat(data[i][2]) || 0;
    var 振替時間 = parseFloat(data[i][3]) || 0;

    // ログインIDの照合（デバッグ用に最初の数件だけログ出力）
    if (i < 10) {
      Logger.log("行" + (i+2) + ": 年月=" + 年月 + ", userID=[" + userID + "], ログインID=[" + ログインID + "], 一致=" + (userID === String(ログインID).trim()));
    }

    // ユーザーIDが一致しない場合はスキップ
    if (userID !== String(ログインID).trim()) continue;

    // 年月から年と月を抽出
    var match = 年月.match(/(\d{4})年(\d{1,2})月/);
    if (!match) {
      Logger.log("年月フォーマットエラー: " + 年月);
      continue;
    }

    var year = parseInt(match[1]);
    var month = parseInt(match[2]);

    // このユーザーの全データを記録（デバッグ用）
    allUserData.push(年月);

    // 年度に属するかチェック（4月〜翌年3月）
    var inFiscalYear = false;
    if (month >= 4 && year === startYear) {
      inFiscalYear = true;
    } else if (month <= 3 && year === endYear) {
      inFiscalYear = true;
    }

    Logger.log("年月=" + 年月 + ", year=" + year + ", month=" + month + ", 年度判定=" + inFiscalYear);

    if (inFiscalYear) {
      Logger.log("データ追加: " + 年月 + " (時間外=" + 時間外 + ", 振替=" + 振替時間 + ")");
      monthlyData.push({
        年月: 年月,
        時間外: 時間外,
        振替時間: 振替時間,
        勤務時間: 時間外 - 振替時間,
        year: year,
        month: month
      });
    } else {
      Logger.log("データ除外: " + 年月 + " (年度=" + startYear + "に該当せず)");
    }
  }

  Logger.log("このユーザーの全データ: " + allUserData.join(", "));
  Logger.log("最終データ件数: " + monthlyData.length);

  // 月順にソート（4月から順番に）
  monthlyData.sort(function(a, b) {
    var aOrder = a.month >= 4 ? a.month - 4 : a.month + 8;
    var bOrder = b.month >= 4 ? b.month - 4 : b.month + 8;
    return aOrder - bOrder;
  });

  return monthlyData;
}

/**
 * 所属1別の月次集計（管理者用）
 */
function getDepartmentMonthlyData(年月) {
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
  var timeSheet = ss.getSheetByName("テーブル");
  var userSheet = ss.getSheetByName("ユーザー名");

  if (timeSheet.getLastRow() < 2) {
    return [];
  }

  // ユーザーマスタを取得（ID → 所属１のマッピング）
  var userMap = {};
  var userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < userData.length; i++) {
    var id = String(userData[i][0]);
    userMap[id] = {
      氏名: String(userData[i][1]),
      所属１: String(userData[i][2]),
      所属２: String(userData[i][3]),
      状態: String(userData[i][4])
    };
  }

  // 時間外データを取得
  var timeData = timeSheet.getRange(2, 1, timeSheet.getLastRow() - 1, 5).getValues();
  var deptData = {};

  for (var i = 0; i < timeData.length; i++) {
    var 年月str = String(timeData[i][0]);
    var userID = String(timeData[i][1]);
    var 時間外 = parseFloat(timeData[i][2]) || 0;
    var 振替時間 = parseFloat(timeData[i][3]) || 0;

    // 指定年月のデータのみ
    if (年月str !== 年月) continue;

    // ユーザー情報取得
    var user = userMap[userID];
    if (!user) continue;

    var 所属１ = user.所属１ || "未所属";

    // 所属1別に集計
    if (!deptData[所属１]) {
      deptData[所属１] = {
        所属１: 所属１,
        時間外合計: 0,
        振替時間合計: 0,
        勤務時間合計: 0,
        人数: 0,
        詳細: []
      };
    }

    deptData[所属１].時間外合計 += 時間外;
    deptData[所属１].振替時間合計 += 振替時間;
    deptData[所属１].勤務時間合計 += (時間外 - 振替時間);
    deptData[所属１].人数++;

    deptData[所属１].詳細.push({
      ログインID: userID,
      氏名: user.氏名,
      所属２: user.所属２,
      時間外: 時間外,
      振替時間: 振替時間,
      勤務時間: 時間外 - 振替時間
    });
  }

  // 配列に変換してソート
  var result = [];
  for (var dept in deptData) {
    // 詳細を勤務時間でソート
    deptData[dept].詳細.sort(function(a, b) {
      return b.勤務時間 - a.勤務時間;
    });
    result.push(deptData[dept]);
  }

  // 勤務時間合計でソート
  result.sort(function(a, b) {
    return b.勤務時間合計 - a.勤務時間合計;
  });

  return result;
}

/**
 * 個人別ランキング取得（管理者用）
 */
function getPersonalRanking(年月, limit) {
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
  var timeSheet = ss.getSheetByName("テーブル");
  var userSheet = ss.getSheetByName("ユーザー名");

  if (timeSheet.getLastRow() < 2) {
    return [];
  }

  limit = limit || 10;

  // ユーザーマスタを取得
  var userMap = {};
  var userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < userData.length; i++) {
    var id = String(userData[i][0]);
    userMap[id] = {
      氏名: String(userData[i][1]),
      所属１: String(userData[i][2]),
      所属２: String(userData[i][3])
    };
  }

  // 時間外データを取得
  var timeData = timeSheet.getRange(2, 1, timeSheet.getLastRow() - 1, 5).getValues();
  var ranking = [];

  for (var i = 0; i < timeData.length; i++) {
    var 年月str = String(timeData[i][0]);
    var userID = String(timeData[i][1]);
    var 時間外 = parseFloat(timeData[i][2]) || 0;
    var 振替時間 = parseFloat(timeData[i][3]) || 0;

    if (年月str !== 年月) continue;

    var user = userMap[userID];
    if (!user) continue;

    ranking.push({
      ログインID: userID,
      氏名: user.氏名,
      所属１: user.所属１,
      所属２: user.所属２,
      時間外: 時間外,
      振替時間: 振替時間,
      勤務時間: 時間外 - 振替時間
    });
  }

  // 勤務時間でソート
  ranking.sort(function(a, b) {
    return b.勤務時間 - a.勤務時間;
  });

  // 上位N件を返す
  return ranking.slice(0, limit);
}

/**
 * 利用可能な年度リストを取得
 */
function getAvailableFiscalYears() {
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1;
  var currentFiscalYear = (currentMonth >= 4) ? currentYear : currentYear - 1;

  var years = [];
  for (var y = 2024; y <= currentFiscalYear; y++) {
    years.push(y);
  }

  return years;
}

/**
 * 利用可能な年月リストを取得（最近12ヶ月）
 */
function getAvailableMonths() {
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
  var sheet = ss.getSheetByName("テーブル");

  if (sheet.getLastRow() < 2) {
    return [];
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var monthSet = {};

  for (var i = 0; i < data.length; i++) {
    var 年月 = String(data[i][0]);
    monthSet[年月] = true;
  }

  var months = Object.keys(monthSet);

  // ソート（新しい順）
  months.sort(function(a, b) {
    var matchA = a.match(/(\d{4})年(\d{1,2})月/);
    var matchB = b.match(/(\d{4})年(\d{1,2})月/);
    if (!matchA || !matchB) return 0;

    var dateA = parseInt(matchA[1]) * 100 + parseInt(matchA[2]);
    var dateB = parseInt(matchB[1]) * 100 + parseInt(matchB[2]);

    return dateB - dateA;
  });

  return months.slice(0, 12); // 最近12ヶ月
}
