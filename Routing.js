/**
 * Routing.gs 蜈ｨ譁・ｿｮ豁｣ Ver1.1
 * Google Apps Script: Web 繧｢繝励Μ縺ｮ繝ｫ繝ｼ繝・ぅ繝ｳ繧ｰ螳夂ｾｩ
 */

/**
 * GET 繝ｪ繧ｯ繧ｨ繧ｹ繝医ｒ蜃ｦ逅・ * @param {Object} e 繧､繝吶Φ繝医ヱ繝ｩ繝｡繝ｼ繧ｿ
 * @returns {HtmlOutput} HTML 蜃ｺ蜉・ */
function doGet(e) {
  var page = e.parameter.page;
  Logger.log("doGet called, page: " + page);

  if (page === 'input') {
    var inputTmpl = HtmlService.createTemplateFromFile('input');
    return inputTmpl.evaluate();

  } else if (page === 'menu') {
    // 繝｡繝九Η繝ｼ逕ｻ髱｢縺ｸ驕ｷ遘ｻ・壹Ο繧ｰ繧､繝ｳID 繧偵ユ繝ｳ繝励Ξ繝ｼ繝医↓險ｭ螳・    var menuTmpl = HtmlService.createTemplateFromFile('繝｡繝九Η繝ｼ');
    menuTmpl.繝ｭ繧ｰ繧､繝ｳID = e.parameter.繝ｭ繧ｰ繧､繝ｳID;
    return menuTmpl.evaluate();

  } else {
    // 繝・ヵ繧ｩ繝ｫ繝茨ｼ壹Ο繧ｰ繧､繝ｳ逕ｻ髱｢
    var loginTmpl = HtmlService.createTemplateFromFile('login');
    return loginTmpl.evaluate();
  }
}

/**
 * POST 繝ｪ繧ｯ繧ｨ繧ｹ繝医ｒ蜃ｦ逅・ * @param {Object} e 繧､繝吶Φ繝医ヱ繝ｩ繝｡繝ｼ繧ｿ
 * @returns {HtmlOutput} HTML 蜃ｺ蜉・ */
function doPost(e) {
  Logger.log("doPost action: " + e.parameter.action);
  var ss = SpreadsheetApp.openById("1eabKd-YqMH48rX5BdhFd_MU6KWFdWHAWt2t5-Y96reA");
  var action = e.parameter.action;

  switch (action) {
    // 繝ｭ繧ｰ繧､繝ｳ蜃ｦ逅・    case "login":
      return 繝ｭ繧ｰ繧､繝ｳ蜃ｦ逅・e, ss);

    // 譎る俣螟門ｱ蜻奇ｼ亥ｹｴ譛磯∈謚樒判髱｢縺ｸ・・    case "譎る俣螟門ｱ蜻・:
      return handleTimeReport(e, ss);

    // 譎る俣螟夜寔險医Γ繝九Η繝ｼ
    case "譎る俣螟夜寔險・:
      var 髮・ｨ・mpl = HtmlService.createTemplateFromFile('髮・ｨ医Γ繝九Η繝ｼ');
      髮・ｨ・mpl.繝ｭ繧ｰ繧､繝ｳID = e.parameter.繝ｭ繧ｰ繧､繝ｳID;
      return 髮・ｨ・mpl.evaluate();

    // 謇螻槫挨髮・ｨ・    case "謇螻槫挨髮・ｨ・:
      return handleShozokuBetsuSuikei(e, ss);

    // 蟷ｴ譛亥挨髮・ｨ・    case "蟷ｴ譛亥挨髮・ｨ・:
      return handleNengetsuBetsuSuikei(e, ss);

    // select 繧｢繧ｯ繧ｷ繝ｧ繝ｳ・亥ｹｴ譛磯∈謚樞・蜈･蜉帷判髱｢・・    case "select":
      Logger.log("select action with 蟷ｴ譛・ " + e.parameter.蟷ｴ譛・;
      var selectTmpl = HtmlService.createTemplateFromFile('input');
      selectTmpl.繝ｭ繧ｰ繧､繝ｳID   = e.parameter.繝ｭ繧ｰ繧､繝ｳID;
      selectTmpl.驕ｸ謚槫ｹｴ譛・    = e.parameter.蟷ｴ譛・
      selectTmpl.蜑榊屓譎る俣螟・  = e.parameter.譎る俣螟・
      selectTmpl.蜑榊屓謖ｯ譖ｿ譎る俣 = e.parameter.謖ｯ譖ｿ譎る俣;
      return selectTmpl.evaluate();

    // sendText 繧｢繧ｯ繧ｷ繝ｧ繝ｳ・亥・蜉帛､菫晏ｭ倪・繝｡繝九Η繝ｼ逕ｻ髱｢縺ｸ・・    case "sendText":
      return handleSendText(e, ss);

    // 縺昴ｌ莉･螟悶・繝ｭ繧ｰ繧､繝ｳ逕ｻ髱｢縺ｸ
    default:
      return HtmlService.createTemplateFromFile('login').evaluate();
  }
}

/**
 * 繝ｭ繧ｰ繧､繝ｳ蜃ｦ逅・ * @param {Object} e 繧､繝吶Φ繝医ヱ繝ｩ繝｡繝ｼ繧ｿ
 * @param {Spreadsheet} ss 繧ｹ繝励Ξ繝・ラ繧ｷ繝ｼ繝・ * @returns {HtmlOutput} 繝｡繝九Η繝ｼ逕ｻ髱｢ or 繧ｨ繝ｩ繝ｼ繝｡繝・そ繝ｼ繧ｸ
 */
function 繝ｭ繧ｰ繧､繝ｳ蜃ｦ逅・e, ss) {
  var sheet繝ｦ繝ｼ繧ｶ繝ｼ蜷・= ss.getSheetByName("繝ｦ繝ｼ繧ｶ繝ｼ蜷・);
  var data繝ｦ繝ｼ繧ｶ繝ｼ = sheet繝ｦ繝ｼ繧ｶ繝ｼ蜷・    .getRange(2, 1, sheet繝ｦ繝ｼ繧ｶ繝ｼ蜷・getLastRow() - 1, 1)
    .getValues();
  Logger.log("繝ｦ繝ｼ繧ｶ繝ｼ蜷阪ョ繝ｼ繧ｿ: " + JSON.stringify(data繝ｦ繝ｼ繧ｶ繝ｼ));

  // 繝輔か繝ｼ繝蛟､縺ｨ繧ｷ繝ｼ繝亥､繧呈枚蟄怜・豈碑ｼ・  var loginValid = data繝ｦ繝ｼ繧ｶ繝ｼ.some(function(row) {
    return String(row[0]) === String(e.parameter.user);
  });
  Logger.log("loginValid: " + loginValid);

  if (!loginValid) {
    return HtmlService.createHtmlOutput("Invalid login");
  }

  // 繝ｭ繧ｰ繧､繝ｳ謌仙粥蠕後・繝｡繝九Η繝ｼ逕ｻ髱｢
  var menuTmpl = HtmlService.createTemplateFromFile('繝｡繝九Η繝ｼ');
  menuTmpl.繝ｭ繧ｰ繧､繝ｳID = String(e.parameter.user);
  return menuTmpl.evaluate();
}
