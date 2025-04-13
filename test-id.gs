// ===== 定数設定 =====
const LOG_SPREADSHEET_ID = "1BzLZG2B1T_r5bPd2NQmj_dr4nGO5LMtm_Bu8G45jhrM";
//const LINE_CHANNEL_ACCESS_TOKEN = "YOUR_LINE_CHANNEL_ACCESS_TOKEN";

/**
 * LINEのWebhookイベントを受信し、個別チャットの場合はメッセージ、表示名、ユーザーIDをスプレッドシートに保存
 */
function doPost(e) {
  var json = JSON.parse(e.postData.contents);

  json.events.forEach(function(event) {
    // 個別チャットの場合の処理
    if (event.source && event.source.type === "user" && event.source.userId) {
      var userId = event.source.userId;
      var messageText = event.message && event.message.text ? event.message.text : "";
      var displayName = getUserProfile(userId);

      saveUserDataToSheet(userId, displayName, messageText);
      Logger.log("個別チャットのユーザーID: " + userId + " (名前: " + displayName + ") のメッセージ: " + messageText);
    }
  });

  return ContentService.createTextOutput("OK");
}

/**
 * LINE APIを使用して、ユーザープロフィール（表示名）を取得
 * @param {string} userId - ユーザーのLINE ID
 * @return {string} ユーザーの表示名（取得できなかった場合は "Unknown"）
 */
function getUserProfile(userId) {
  var url = "https://api.line.me/v2/bot/profile/" + userId;
  var options = {
    "method": "get",
    "headers": {
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var jsonResponse = JSON.parse(response.getContentText());
    return jsonResponse.displayName || "Unknown";
  } catch (e) {
    Logger.log("プロフィール取得エラー: " + e.toString());
    return "Unknown";
  }
}

/**
 * ユーザーのデータ（ユーザーID、表示名、メッセージ）をスプレッドシートに保存（重複登録を防ぐ）
 * @param {string} userId - 取得したユーザーのLINE ID
 * @param {string} displayName - ユーザーの表示名
 * @param {string} messageText - 受信したメッセージ
 */
function saveUserDataToSheet(userId, displayName, messageText) {
  var ss = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
  var sheet = ss.getSheetByName("LINE_User_Messages");

  if (!sheet) {
    // シートが存在しなければ新規作成
    sheet = ss.insertSheet("LINE_User_Messages");
    sheet.appendRow(["ユーザーID", "表示名", "メッセージ", "日時"]);
  }

  // 必要に応じて重複チェックなどの処理を追加できます
  // 今回は単に受信した内容を記録する例です
  var now = new Date();
  sheet.appendRow([userId, displayName, messageText, now]);
  Logger.log("ユーザーデータを保存しました: " + userId);
}
