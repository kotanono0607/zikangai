/**
 * CSV エクスポート/インポート機能
 *
 * エクスポート: ヘッダー付きCSV出力
 * インポート: IDで照合し、存在すれば更新、なければ新規追加（削除なし）
 */

/**
 * CSV エクスポート
 * ヘッダー付きでユーザーデータを出力
 */
function handleCsvExport(e, ss) {
  var ログインID = e.parameter.ログインID;

  // 管理者チェック
  if (!isAdmin(ログインID, ss)) {
    return ContentService.createTextOutput("アクセス権限がありません")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  var sheet = ss.getSheetByName("ユーザー名");

  if (sheet.getLastRow() < 2) {
    return ContentService.createTextOutput("エクスポートするデータがありません")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // ヘッダー行を取得
  var headers = sheet.getRange(1, 1, 1, 6).getValues()[0];

  // データ行を取得
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  // CSV形式に変換
  var csvContent = [];

  // ヘッダー行を追加
  csvContent.push(headers.join(','));

  // データ行を追加
  for (var i = 0; i < data.length; i++) {
    var row = [];
    for (var j = 0; j < data[i].length; j++) {
      var cell = String(data[i][j]);
      // カンマや改行を含む場合はダブルクォートで囲む
      if (cell.indexOf(',') !== -1 || cell.indexOf('\n') !== -1 || cell.indexOf('"') !== -1) {
        cell = '"' + cell.replace(/"/g, '""') + '"';
      }
      row.push(cell);
    }
    csvContent.push(row.join(','));
  }

  var csv = csvContent.join('\n');

  // BOM付きUTF-8で出力（Excel対応）
  var bom = '\uFEFF';
  var output = ContentService.createTextOutput(bom + csv);
  output.setMimeType(ContentService.MimeType.CSV);

  // ファイル名に日付を含める
  var today = new Date();
  var dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var fileName = 'ユーザーマスタ_' + dateStr + '.csv';

  output.downloadAsFile(fileName);

  Logger.log("CSV エクスポート完了: " + data.length + "件");

  return output;
}

/**
 * CSV インポート処理
 * IDで照合し、存在すれば更新、なければ新規追加
 */
function handleCsvImport(e, ss) {
  var ログインID = e.parameter.ログインID;

  // 管理者チェック
  if (!isAdmin(ログインID, ss)) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "アクセス権限がありません"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var csvData = e.parameter.csvData;

  if (!csvData) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "CSVデータがありません"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // CSV解析
  var lines = csvData.split(/\r?\n/);

  if (lines.length < 2) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "CSVデータが不正です（最低2行必要）"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // ヘッダー行をスキップ
  var header = lines[0].split(',');

  // ヘッダー検証
  var expectedHeaders = ["ログインID", "氏名", "所属１", "所属２", "状態", "権限"];
  var headerValid = true;

  for (var i = 0; i < expectedHeaders.length; i++) {
    if (header[i] !== expectedHeaders[i]) {
      headerValid = false;
      break;
    }
  }

  if (!headerValid) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "CSVヘッダーが不正です。正しい形式: " + expectedHeaders.join(',')
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = ss.getSheetByName("ユーザー名");

  // 既存ユーザーのマップを作成（ログインID → 行番号）
  var existingUsers = {};
  if (sheet.getLastRow() >= 2) {
    var existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < existingData.length; i++) {
      var id = String(existingData[i][0]);
      existingUsers[id] = i + 2; // 行番号（1-indexed、ヘッダー考慮）
    }
  }

  var addedCount = 0;
  var updatedCount = 0;
  var errorCount = 0;
  var errors = [];

  // データ行を処理
  for (var i = 1; i < lines.length; i++) {
    var line = lines[i].trim();

    if (!line) {
      continue; // 空行スキップ
    }

    var cells = parseCsvLine(line);

    if (cells.length !== 6) {
      errorCount++;
      errors.push("行" + (i + 1) + ": 列数が不正（6列必要、" + cells.length + "列）");
      continue;
    }

    var ログインID = String(cells[0]).trim();
    var 氏名 = String(cells[1]).trim();
    var 所属１ = String(cells[2]).trim();
    var 所属２ = String(cells[3]).trim();
    var 状態 = String(cells[4]).trim();
    var 権限 = String(cells[5]).trim();

    // バリデーション
    if (!ログインID || !氏名) {
      errorCount++;
      errors.push("行" + (i + 1) + ": ログインIDまたは氏名が空です");
      continue;
    }

    if (状態 !== "在職" && 状態 !== "退職") {
      errorCount++;
      errors.push("行" + (i + 1) + ": 状態は「在職」または「退職」である必要があります");
      continue;
    }

    if (権限 !== "一般" && 権限 !== "管理者") {
      errorCount++;
      errors.push("行" + (i + 1) + ": 権限は「一般」または「管理者」である必要があります");
      continue;
    }

    // 既存ユーザーか確認
    if (existingUsers[ログインID]) {
      // 更新
      var rowIndex = existingUsers[ログインID];
      sheet.getRange(rowIndex, 1, 1, 6).setValues([[ログインID, 氏名, 所属１, 所属２, 状態, 権限]]);
      updatedCount++;
    } else {
      // 新規追加
      sheet.appendRow([ログインID, 氏名, 所属１, 所属２, 状態, 権限]);
      addedCount++;
    }
  }

  var message = "インポート完了\n";
  message += "新規追加: " + addedCount + "件\n";
  message += "更新: " + updatedCount + "件";

  if (errorCount > 0) {
    message += "\nエラー: " + errorCount + "件\n\n" + errors.join("\n");
  }

  Logger.log("CSV インポート完了: 追加=" + addedCount + " 更新=" + updatedCount + " エラー=" + errorCount);

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: message,
    added: addedCount,
    updated: updatedCount,
    errors: errorCount
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * CSV行をパースする（カンマやダブルクォートを考慮）
 */
function parseCsvLine(line) {
  var cells = [];
  var cell = '';
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    var char = line[i];
    var nextChar = (i + 1 < line.length) ? line[i + 1] : '';

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        // エスケープされたダブルクォート
        cell += '"';
        i++; // 次の文字をスキップ
      } else {
        // クォート開始/終了
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // セル区切り
      cells.push(cell);
      cell = '';
    } else {
      cell += char;
    }
  }

  cells.push(cell); // 最後のセルを追加

  return cells;
}
