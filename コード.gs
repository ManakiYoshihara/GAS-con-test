/**
 * ===== 定数設定 =====
 */
// LINE公式アカウントのチャネルアクセストークン
const LINE_CHANNEL_ACCESS_TOKEN = "dZpUYeX01inGjmuvPX0Rs3bUnvZl4ukhopc1UVnOrrBn2aWgpdhhPqkC5tknfu1yPp6suWSPxeYxGx2OQ6GPXDeUxp/lk1fedvn7U/TohJYm9Tqgztbozs5or3hExTSRvkhmztxhS4Z+u1/I0eGmIgdB04t89/1O/w1cDnyilFU=";

// スプレッドシートID（シフト管理シート）
const SPREADSHEET_ID = "1FB4Ly73WM9KyTceyUU94HBnWK342Gaag3Rrf0t6tLkU";

// ユーザー名とユーザーIDの対応表
const USER_MAPPING = {
  "倉迫": "Uda9a94c9ebee0913879094fab46c3d53",
  "久野": "U4216eea3ccecabffb5407f6288144222",
  "金森": "Uab1afcc281a07f8f9eaf9c86f62c673a"
};

// 担当者が不明の場合のフォールバックユーザーID（適切なユーザーIDに置き換えてください）
const FALLBACK_USER_ID = "U56fb0ab66318998e7c36e9c9c7426a66";


/**
 * ===== LINE公式アカウントAPIを利用して個別チャットでメッセージ送信 =====
 */
function sendLineMessage(target, message) {
  var url = "https://api.line.me/v2/bot/message/push";
  var payload = {
    "to": target,
    "messages": [
      {
        "type": "text",
        "text": message
      }
    ]
  };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    "payload": JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}


/**
 * ===== スプレッドシートから担当者名を取得する関数 =====
 */
function getResponsibleName() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheets()[0];
  
  if (!sheet) {
    Logger.log("エラー: 指定したシートが見つかりません");
    return "担当";
  }

  var data = sheet.getDataRange().getValues();
  var todayStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "M/d"); 
  Logger.log("今日の日付 (比較用): " + todayStr);

  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0];
    var responsible = data[i][2];

    var typeStr = Object.prototype.toString.call(rowDate);
    Logger.log("【デバッグ】Row " + (i+1) + " - A列の型判定: " + typeStr + " / 生データ: " + rowDate);
    
    if (typeStr === "[object Date]") {
      var rowDateStr = Utilities.formatDate(rowDate, "Asia/Tokyo", "M/d");
      Logger.log("【デバッグ】Row " + (i+1) + " - Date型としてフォーマットされた日付: " + rowDateStr + " / 担当者: " + responsible);
      if (rowDateStr === todayStr) {
        Logger.log("【デバッグ】Row " + (i+1) + " - 一致！担当者: " + responsible);
        return responsible || "担当";
      }
    } else if (typeof rowDate === "string") {
      var trimmedRowDate = rowDate.trim();
      Logger.log("【デバッグ】Row " + (i+1) + " - 文字列として取得された日付: " + trimmedRowDate + " / 担当者: " + responsible);
      if (trimmedRowDate === todayStr) {
        Logger.log("【デバッグ】Row " + (i+1) + " - 文字列一致！担当者: " + responsible);
        return responsible || "担当";
      }
    } else {
      Logger.log("【デバッグ】Row " + (i+1) + " - 想定外の型: " + typeStr + " / 値: " + rowDate);
    }
  }

  Logger.log("該当する担当者が見つかりませんでした");
  return "担当";
}


/**
 * ===== 【朝7時用】担当者へ個別チャットでメッセージ送信する関数 =====
 */
function sendMorningIndividualMessage() {
  var responsibleName = getResponsibleName();
  var userId = USER_MAPPING[responsibleName];
  
  // 朝のメッセージ文
  var morningMessage = "おはようございます！\n" +
                       "本日もよろしくお願いいたします！\n\n" +
                       "午前中の段階で\n" +
                       "・Slackの営業予約の確認\n" +
                       "・本日営業予定の方にリマインド\n\n" +
                       "こちらの2点をお願いいたします！ \n" +
                       "完了したら「完了！」と営業LINEグループに投稿してください!";

  // フォールバック用メッセージ文
  var fallbackMessage = "担当者不明のため送信されています。\n" + morningMessage;

  if (!userId) {
    Logger.log("担当者がUSER_MAPPINGに存在しないため、フォールバックユーザーに送信します");
    userId = FALLBACK_USER_ID;
    sendLineMessage(userId, fallbackMessage);
    Logger.log("フォールバックメッセージ送信成功");
  } else {
    try {
      sendLineMessage(userId, morningMessage);
      Logger.log("朝の個別メッセージ送信成功: " + responsibleName);
    } catch (e) {
      Logger.log("朝の個別メッセージ送信エラー: " + e.toString());
    }
  }
}


/**
 * ===== 【昼12時用】USER_MAPPINGの全員へ個別チャットでメッセージ送信する関数 =====
 */
function sendNoonIndividualMessage() {
  var noonMessage = "お疲れ様です！\n" +
                    "本日も\n\n" +
                    "・Slackにて自分宛のメンションがないかの確認と対応\n" +
                    "・公式LINEにて自分の担当の状況の確認\n" +
                    "こちらの2点をお願いいたします！ \n" +
                    "完了したら「完了！」と営業LINEグループに投稿してください!";
                    
  for (var name in USER_MAPPING) {
    var userId = USER_MAPPING[name];
    try {
      sendLineMessage(userId, noonMessage);
      Logger.log("昼の個別メッセージ送信成功: " + name);
    } catch (e) {
      Logger.log("昼の個別メッセージ送信エラー (" + name + "): " + e.toString());
    }
  }
}


/**
 * ===== トリガーの設定（初回実行時に実行してください） =====
 * 朝7時と昼12時に自動実行するトリガーを作成します。
 */
function setupTriggers() {
  // 既存のトリガーを全て削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // 朝7時に送信するトリガーを設定
  ScriptApp.newTrigger("sendMorningIndividualMessage")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  
  // 昼12時に送信するトリガーを設定
  ScriptApp.newTrigger("sendNoonIndividualMessage")
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
}
