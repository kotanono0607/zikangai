/**
 * QRログイン管理機能
 *
 * 外部依存: QuickChart.io (https://quickchart.io/)
 * - QRコード画像生成に使用
 * - 完全無料、制限なし
 * - Google Charts API の代替サービス
 */

/**
 * QRコードでのログイン処理
 * URLパラメータ: ?action=qrLogin&id=user001
 */
function handleQrLogin(e, ss) {
  var userId = e.parameter.id;

  if (!userId) {
    return HtmlService.createHtmlOutput("無効なQRコードです（IDが見つかりません）");
  }

  Logger.log("QRログイン試行: " + userId);

  // ユーザー情報を取得
  var sheet = ss.getSheetByName("ユーザー名");

  if (!sheet || sheet.getLastRow() < 2) {
    return HtmlService.createHtmlOutput("ユーザーデータが見つかりません");
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  var ユーザー情報 = null;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      // 退職者チェック
      if (String(data[i][4]) === "退職") {
        Logger.log("QRログイン失敗: 退職済みユーザー - " + userId);
        return HtmlService.createHtmlOutput("このアカウントは利用できません（退職済み）");
      }

      ユーザー情報 = {
        ログインID: String(data[i][0]),
        氏名: String(data[i][1]),
        権限: String(data[i][5])
      };
      break;
    }
  }

  if (!ユーザー情報) {
    Logger.log("QRログイン失敗: ユーザーが見つかりません - " + userId);
    return HtmlService.createHtmlOutput("無効なQRコードです（ユーザーが見つかりません）");
  }

  Logger.log("QRログイン成功: " + userId + " (" + ユーザー情報.氏名 + ")");

  // メニュー画面へ自動遷移
  var tmpl = HtmlService.createTemplateFromFile('メニュー');
  tmpl.ログインID = ユーザー情報.ログインID;
  tmpl.氏名 = ユーザー情報.氏名;
  tmpl.権限 = ユーザー情報.権限;
  return tmpl.evaluate();
}

/**
 * QRログイン用URLを生成
 * @param {string} userId ログインID
 * @return {string} QRログインURL
 */
function generateQrLoginUrl(userId) {
  var webAppUrl = ScriptApp.getService().getUrl();
  return webAppUrl + '?action=qrLogin&id=' + encodeURIComponent(userId);
}

/**
 * QuickChart.io でQRコード画像URLを生成
 * @param {string} userId ログインID
 * @param {number} size QRコードのサイズ（ピクセル）デフォルト: 300
 * @return {string} QRコード画像URL
 */
function getQrCodeImageUrl(userId, size) {
  size = size || 300;
  var loginUrl = generateQrLoginUrl(userId);

  // QuickChart.io API を使用（無料）
  // https://quickchart.io/documentation/
  var qrImageUrl = 'https://quickchart.io/qr?text=' +
                   encodeURIComponent(loginUrl) +
                   '&size=' + size +
                   '&margin=2'; // マージン（余白）

  return qrImageUrl;
}

/**
 * 全ユーザーのQRカードPDFを生成
 * 管理者のみ実行可能
 */
function handleGenerateQrPdf(e, ss) {
  var ログインID = e.parameter.ログインID;

  // 管理者チェック
  if (!isAdmin(ログインID, ss)) {
    return HtmlService.createHtmlOutput("❌ この機能は管理者のみ利用可能です");
  }

  var sheet = ss.getSheetByName("ユーザー名");

  if (!sheet || sheet.getLastRow() < 2) {
    return HtmlService.createHtmlOutput("ユーザーデータが見つかりません");
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  // 在職ユーザーのみ抽出
  var users = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][4]) === "在職") {
      users.push({
        ログインID: String(data[i][0]),
        氏名: String(data[i][1]),
        所属１: String(data[i][2]),
        所属２: String(data[i][3])
      });
    }
  }

  if (users.length === 0) {
    return HtmlService.createHtmlOutput("在職ユーザーが見つかりません");
  }

  Logger.log("QRカード一括生成: " + users.length + "人");

  // QRカードHTMLテンプレートを生成
  var tmpl = HtmlService.createTemplateFromFile('qr-card');
  tmpl.users = users;
  tmpl.isBulk = true;

  return tmpl.evaluate()
    .setTitle('QRログインカード - 一括生成 (' + users.length + '人)')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * 個別ユーザーのQRカードを生成
 * 管理者のみ実行可能
 */
function handleGenerateUserQr(e, ss) {
  var ログインID = e.parameter.ログインID;
  var targetUserId = e.parameter.targetUserId;

  // 管理者チェック
  if (!isAdmin(ログインID, ss)) {
    return HtmlService.createHtmlOutput("❌ この機能は管理者のみ利用可能です");
  }

  if (!targetUserId) {
    return HtmlService.createHtmlOutput("対象ユーザーIDが指定されていません");
  }

  var sheet = ss.getSheetByName("ユーザー名");
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  var targetUser = null;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(targetUserId)) {
      targetUser = {
        ログインID: String(data[i][0]),
        氏名: String(data[i][1]),
        所属１: String(data[i][2]),
        所属２: String(data[i][3])
      };
      break;
    }
  }

  if (!targetUser) {
    return HtmlService.createHtmlOutput("ユーザーが見つかりません");
  }

  Logger.log("QRカード個別生成: " + targetUserId + " (" + targetUser.氏名 + ")");

  // QRカードHTMLテンプレートを生成
  var tmpl = HtmlService.createTemplateFromFile('qr-card');
  tmpl.users = [targetUser]; // 配列として渡す
  tmpl.isBulk = false;

  return tmpl.evaluate()
    .setTitle('QRログインカード - ' + targetUser.氏名)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
