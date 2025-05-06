/**
 * SPREADシートの「処理用（YYYY/MM）」シートのQ列がTRUEの場合、
 * O列の値に一致する「メッセージ文」から「問題」「解答」「授業動画」を取得し、
 * 月間報告シートのQ,R,S列に出力します。
 * ただし「問題」はQ列の値にかかわらず一致すればセットします。
 *
 * @param {Spreadsheet} dataSpreadsheet - データ格納用スプレッドシート
 * @param {string} monthlySheetName - 月間報告シート名（例: "2025年03月度"）
 */
function appendMessageDataToDataSheet(dataSpreadsheet, monthlySheetName) {
  // メッセージ文シートを取得
  var ss3 = SpreadsheetApp.openById(SPREADSHEET3_ID);
  var messageSheet = ss3.getSheetByName("メッセージ文");
  if (!messageSheet) {
    Logger.log("メッセージ文シートが見つかりません。");
    return;
  }
  // ヘッダー行取得
  var headers = messageSheet.getRange(1, 1, 1, messageSheet.getLastColumn()).getValues()[0];
  var contentIdx = headers.indexOf("内容");
  var problemIdx = headers.indexOf("問題");
  var answerIdx  = headers.indexOf("解答");
  var videoIdx   = headers.indexOf("授業動画");
  if (contentIdx < 0 || problemIdx < 0 || answerIdx < 0 || videoIdx < 0) {
    Logger.log("メッセージ文シートの必要なヘッダーが見つかりません。");
    return;
  }

  // 月間報告シートから年月を抽出し、対応する処理用シート名を作成
  var m = monthlySheetName.match(/^(\d{4})年(\d{2})月度$/);
  if (!m) {
    Logger.log("月間報告シート名の形式が不正です: " + monthlySheetName);
    return;
  }
  var processSheetName = '処理用（' + m[1] + '/' + m[2] + '）';
  var processSheet = dataSpreadsheet.getSheetByName(processSheetName);
  if (!processSheet) {
    Logger.log("処理用シートが見つかりません: " + processSheetName);
    return;
  }

  // 処理用シートのO列(15)とQ列(17)を取得（1行目から最終行まで）
  var lastProcRow = processSheet.getLastRow();
  if (lastProcRow < 1) return;
  var procData = processSheet.getRange(1, 15, lastProcRow, 3).getValues(); // [O,P,Q]

  // メッセージ文データを一度だけ取得
  var lastMsgRow = messageSheet.getLastRow();
  var msgData = messageSheet.getRange(2, 1, lastMsgRow - 1, messageSheet.getLastColumn()).getValues();
  // msgData をキー(content)→行配列のマップに変換
  var msgMap = {};
  msgData.forEach(function(row) {
    var key = row[contentIdx];
    if (key !== "") {
      msgMap[key] = row;
    }
  });

  // 月間報告シート取得
  var monthlySheet = dataSpreadsheet.getSheetByName(monthlySheetName);
  if (!monthlySheet) {
    Logger.log("月間報告シートが見つかりません: " + monthlySheetName);
    return;
  }

  // QRS列用配列作成（9行目以降）
  var pqrValues = [];
  for (var i = 0; i < lastProcRow; i++) {
    var content = procData[i][0];
    var flag    = procData[i][2] === true; // Q列がTRUEなら解答/動画も取得
    var row     = msgMap[content];
    var prob    = row ? row[problemIdx] : "";                   // 問題は常に取得
    var ans     = (row && flag) ? row[answerIdx] : "";         // 解答はフラグ依存
    var vid     = (row && flag) ? row[videoIdx] : "";          // 動画はフラグ依存
    pqrValues.push([prob, ans, vid]);
  }

  // 月間報告シートのQ,R,S列（9行目以降）へ一括出力
  monthlySheet.getRange(9, 17, pqrValues.length, 3).setValues(pqrValues);
}
