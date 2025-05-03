/**
 * SPREADSHEET3の「メッセージ文」シートからデータを取得し、
 * データ格納用スプレッドシートの月間報告シートのO列（9行目以降）の
 * 値に一致する「問題」「解答」「授業動画」をP,Q,R列に追加します。
 *
 * @param {Spreadsheet} dataSpreadsheet - データ格納用スプレッドシート
 * @param {string} monthlySheetName - 処理対象の月間報告シート名
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
    // データ行取得
    var lastRowMsg = messageSheet.getLastRow();
    var msgData = messageSheet.getRange(2, 1, lastRowMsg - 1, messageSheet.getLastColumn()).getValues();
  
    // データ格納用シートの対象月シート取得
    var dataSheet = dataSpreadsheet.getSheetByName(monthlySheetName);
    if (!dataSheet) {
      Logger.log("データ格納用シートに " + monthlySheetName + " が見つかりません。");
      return;
    }
    // O列9行目以降の取得
    var lastRowData = dataSheet.getLastRow();
    var oValues = dataSheet.getRange(9, 15, lastRowData - 8, 1).getValues(); // O列は15列目
  
    // PQR列用に値を準備
    var pqrValues = [];
    oValues.forEach(function(row) {
      var match = msgData.find(function(msgRow) {
        return msgRow[contentIdx] === row[0];
      });
      if (match) {
        pqrValues.push([match[problemIdx], match[answerIdx], match[videoIdx]]);
      } else {
        pqrValues.push(["", "", ""]);
      }
    });
    // P,Q,R列にセット
    dataSheet.getRange(9, 16, pqrValues.length, 3).setValues(pqrValues);
  }
  