//=================================================================================================================================================
// Setup.js
// 初期セットアップ用スクリプト

/**
 * 初期セットアップ: 管理者ユーザーを作成
 *
 * 実行方法:
 * 1. GASエディタで このファイルを開く
 * 2. 関数リストから「初期セットアップ_管理者作成」を選択
 * 3. 「実行」ボタンをクリック
 * 4. 実行ログを確認
 */
function 初期セットアップ_管理者作成() {
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
  var sheet = ss.getSheetByName("ユーザー名");

  // 既に管理者ユーザーが存在するかチェック
  if (sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === "admin") {
        Logger.log("❌ 管理者ユーザー（admin）は既に存在します");
        return;
      }
    }
  }

  // 管理者ユーザーを追加
  sheet.appendRow([
    "admin",        // ログインID
    "管理者",       // 氏名
    "総務課",       // 所属１
    "庶務",         // 所属２
    "在職",         // 状態
    "管理者"        // 権限
  ]);

  Logger.log("✅ 管理者ユーザーを作成しました");
  Logger.log("ログインID: admin");
  Logger.log("");
  Logger.log("⚠️ 注意: パスワード認証は実装されていません（将来的にQRログインなどに変更予定）");
}

/**
 * ヘッダー行が正しいかチェック
 */
function シート構造チェック() {
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
  var sheet = ss.getSheetByName("ユーザー名");

  if (!sheet) {
    Logger.log("❌ 「ユーザー名」シートが見つかりません");
    return;
  }

  var headers = sheet.getRange(1, 1, 1, 6).getValues()[0];
  var expectedHeaders = ["ログインID", "氏名", "所属１", "所属２", "状態", "権限"];

  Logger.log("現在のヘッダー:");
  Logger.log(JSON.stringify(headers));
  Logger.log("");
  Logger.log("期待されるヘッダー:");
  Logger.log(JSON.stringify(expectedHeaders));
  Logger.log("");

  var isCorrect = true;
  for (var i = 0; i < expectedHeaders.length; i++) {
    if (String(headers[i]) !== expectedHeaders[i]) {
      isCorrect = false;
      Logger.log("❌ 列" + (i + 1) + "が不正: 「" + headers[i] + "」→「" + expectedHeaders[i] + "」に修正してください");
    }
  }

  if (isCorrect) {
    Logger.log("✅ シート構造は正しいです");
  }
}

/**
 * テストユーザー5人を作成（開発・テスト用）
 */
function テストユーザー作成() {
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
  var sheet = ss.getSheetByName("ユーザー名");

  var testUsers = [
    ["user001", "山田太郎", "総務課", "庶務係", "在職", "一般"],
    ["user002", "佐藤花子", "営業課", "第一係", "在職", "一般"],
    ["user003", "鈴木一郎", "経理課", "会計係", "在職", "一般"],
    ["user004", "田中美咲", "人事課", "労務係", "在職", "一般"],
    ["user005", "伊藤健太", "IT課", "システム係", "退職", "一般"]
  ];

  for (var i = 0; i < testUsers.length; i++) {
    sheet.appendRow(testUsers[i]);
  }

  Logger.log("✅ テストユーザー " + testUsers.length + "人を作成しました");
  Logger.log("");
  Logger.log("作成されたユーザー:");
  for (var i = 0; i < testUsers.length; i++) {
    Logger.log("ID: " + testUsers[i][0] + " / 状態: " + testUsers[i][4]);
  }
}

/**
 * 完全セットアップ: テストユーザー + 所属マスタ自動抽出
 *
 * 実行手順:
 * 1. テストユーザー5人を作成
 * 2. 既存ユーザーから所属マスタを自動抽出
 */
function 完全セットアップ() {
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");

  Logger.log("=== 完全セットアップ開始 ===");
  Logger.log("");

  // 1. テストユーザー作成
  Logger.log("▶ テストユーザー作成中...");
  テストユーザー作成();
  Logger.log("");

  // 2. 所属マスタ自動抽出
  Logger.log("▶ 所属マスタ自動抽出中...");
  所属マスタ自動抽出(ss);
  Logger.log("");

  Logger.log("=== 完全セットアップ完了 ===");
}
