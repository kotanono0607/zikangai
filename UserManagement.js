//=================================================================================================================================================
// UserManagement.js
// ユーザー管理機能（管理者専用）

/**
 * 管理者権限チェック
 */
function isAdmin(ログインID, ss) {
  var sheet = ss.getSheetByName("ユーザー名");
  if (sheet.getLastRow() < 2) return false;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(ログインID) && String(data[i][5]) === "管理者") {
      return true;
    }
  }
  return false;
}

/**
 * ユーザー一覧表示
 */
function handleUserList(e, ss) {
  var ログインID = e.parameter.ログインID;

  // 管理者権限チェック
  if (!isAdmin(ログインID, ss)) {
    return HtmlService.createHtmlOutput("❌ アクセス権限がありません");
  }

  var sheet = ss.getSheetByName("ユーザー名");
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    // ユーザーが1人もいない場合
    var tmpl = HtmlService.createTemplateFromFile('admin');
    tmpl.ログインID = ログインID;
    tmpl.users = [];
    tmpl.totalCount = 0;
    return tmpl.evaluate();
  }

  // 全ユーザーデータ取得
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // 検索条件取得
  var 検索キーワード = e.parameter.search || "";
  var 状態フィルタ = e.parameter.statusFilter || "all";
  var 権限フィルタ = e.parameter.roleFilter || "all";

  // フィルタリング
  var filteredUsers = [];
  for (var i = 0; i < data.length; i++) {
    var user = {
      ログインID: String(data[i][0]),
      氏名: String(data[i][1]),
      所属１: String(data[i][2]),
      所属２: String(data[i][3]),
      状態: String(data[i][4]),
      権限: String(data[i][5])
    };

    // 検索キーワードでフィルタ
    if (検索キーワード) {
      var searchText = (user.ログインID + user.氏名 + user.所属１ + user.所属２).toLowerCase();
      if (searchText.indexOf(検索キーワード.toLowerCase()) === -1) {
        continue;
      }
    }

    // 状態フィルタ
    if (状態フィルタ !== "all" && user.状態 !== 状態フィルタ) {
      continue;
    }

    // 権限フィルタ
    if (権限フィルタ !== "all" && user.権限 !== 権限フィルタ) {
      continue;
    }

    filteredUsers.push(user);
  }

  // ページネーション
  var page = parseInt(e.parameter.page) || 1;
  var pageSize = 20;
  var totalCount = filteredUsers.length;
  var totalPages = Math.ceil(totalCount / pageSize);
  var startIndex = (page - 1) * pageSize;
  var endIndex = Math.min(startIndex + pageSize, totalCount);

  var pagedUsers = filteredUsers.slice(startIndex, endIndex);

  Logger.log("ユーザー一覧表示: " + pagedUsers.length + "件 (total: " + totalCount + ")");

  var tmpl = HtmlService.createTemplateFromFile('admin');
  tmpl.ログインID = ログインID;
  tmpl.users = pagedUsers;
  tmpl.currentPage = page;
  tmpl.totalPages = totalPages;
  tmpl.totalCount = totalCount;
  tmpl.検索キーワード = 検索キーワード;
  tmpl.状態フィルタ = 状態フィルタ;
  tmpl.権限フィルタ = 権限フィルタ;
  return tmpl.evaluate();
}

/**
 * ユーザー編集画面表示
 */
function handleEditUser(e, ss) {
  var ログインID = e.parameter.ログインID;
  var 編集対象ID = e.parameter.editUserId;

  // 管理者権限チェック
  if (!isAdmin(ログインID, ss)) {
    return HtmlService.createHtmlOutput("❌ アクセス権限がありません");
  }

  var sheet = ss.getSheetByName("ユーザー名");
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  // ユーザー情報を検索
  var ユーザー情報 = null;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === 編集対象ID) {
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
    return HtmlService.createHtmlOutput("❌ ユーザーが見つかりません");
  }

  var tmpl = HtmlService.createTemplateFromFile('edit-user');
  tmpl.ログインID = ログインID;
  tmpl.所属マスタ = get所属マスタ(ss);
  tmpl.ユーザー = ユーザー情報;
  tmpl.isNew = false;
  return tmpl.evaluate();
}

/**
 * 新規ユーザー登録画面表示
 */
function handleAddUserForm(e, ss) {
  var ログインID = e.parameter.ログインID;

  // 管理者権限チェック
  if (!isAdmin(ログインID, ss)) {
    return HtmlService.createHtmlOutput("❌ アクセス権限がありません");
  }

  // 次のログインID を生成
  var sheet = ss.getSheetByName("ユーザー名");
  var lastRow = sheet.getLastRow();
  var 次のID = lastRow; // 自動採番

  var tmpl = HtmlService.createTemplateFromFile('edit-user');
  tmpl.ログインID = ログインID;
  tmpl.所属マスタ = get所属マスタ(ss);
  tmpl.ユーザー = {
    ログインID: String(次のID),
    氏名: "",
    所属１: "",
    所属２: "",
    状態: "在職",
    権限: "一般"
  };
  tmpl.isNew = true;
  return tmpl.evaluate();
}

/**
 * ユーザー情報保存
 */
function handleSaveUser(e, ss) {
  var ログインID = e.parameter.ログインID;

  // 管理者権限チェック
  if (!isAdmin(ログインID, ss)) {
    return HtmlService.createHtmlOutput("❌ アクセス権限がありません");
  }

  var 対象ID = e.parameter.userId;
  var 氏名 = e.parameter.userName;
  var 所属１ = e.parameter.shozoku1;
  var 所属２ = e.parameter.shozoku2;
  var 状態 = e.parameter.status;
  var 権限 = e.parameter.role;
  var isNew = e.parameter.isNew === "true";

  // バリデーション
  if (!対象ID || !氏名) {
    return HtmlService.createHtmlOutput("❌ 必須項目を入力してください");
  }

  var sheet = ss.getSheetByName("ユーザー名");
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  if (isNew) {
    // 新規登録
    Logger.log("新規ユーザー登録: " + 対象ID);
    sheet.appendRow([対象ID, 氏名, 所属１, 所属２, 状態, 権限]);
  } else {
    // 既存ユーザー更新
    Logger.log("ユーザー情報更新: " + 対象ID);
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === 対象ID) {
        var rowIndex = i + 2;
        sheet.getRange(rowIndex, 2).setValue(氏名);
        sheet.getRange(rowIndex, 3).setValue(所属１);
        sheet.getRange(rowIndex, 4).setValue(所属２);
        sheet.getRange(rowIndex, 5).setValue(状態);
        sheet.getRange(rowIndex, 6).setValue(権限);
        break;
      }
    }
  }

  // 成功メッセージと共にユーザー一覧に戻る
  return handleUserList(e, ss);
}

/**
 * 退職処理（状態を「退職」に変更）
 */
function handleRetireUser(e, ss) {
  var ログインID = e.parameter.ログインID;
  var 対象ID = e.parameter.retireUserId;

  // 管理者権限チェック
  if (!isAdmin(ログインID, ss)) {
    return HtmlService.createHtmlOutput("❌ アクセス権限がありません");
  }

  // 自分自身を退職処理できないようにする
  if (対象ID === ログインID) {
    return HtmlService.createHtmlOutput("❌ 自分自身を退職処理できません");
  }

  var sheet = ss.getSheetByName("ユーザー名");
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === 対象ID) {
      var rowIndex = i + 2;
      sheet.getRange(rowIndex, 5).setValue("退職");
      Logger.log("退職処理完了: " + 対象ID);
      break;
    }
  }

  // ユーザー一覧に戻る
  return handleUserList(e, ss);
}
