/**
 * 月報生成ボタンによるトリガーを設定する関数（※作成時は既存のトリガーを削除し、processMonthlyReportTrigger を設定）
 */
function createEditTrigger() {
    // 既存のトリガーを削除（handler 関数名が processMonthlyReportTrigger で、MAIN_SHEET_ID に紐付くもの）
    const allTriggers = ScriptApp.getProjectTriggers();
    allTriggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processMonthlyReportTrigger' &&
          trigger.getTriggerSourceId() === MAIN_SHEET_ID) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 新しい編集トリガーを作成（processMonthlyReportTrigger を呼び出す）
    ScriptApp.newTrigger('processMonthlyReportTrigger')
      .forSpreadsheet(MAIN_SHEET_ID)
      .onEdit()
      .create();
    
    Logger.log('編集トリガーが設定されました。');
  }

/**
 * 書類生成ボタンによるトリガーを設定する関数
 */
function createEditTrigger() {
    // 既存のトリガーを削除（重複を避けるため）
    const allTriggers = ScriptApp.getProjectTriggers();
    
    allTriggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'handleEdit' && trigger.getTriggerSourceId() === MAIN_SHEET_ID) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 新しい編集トリガーを作成
    ScriptApp.newTrigger('handleEdit')
      .forSpreadsheet(MAIN_SHEET_ID)
      .onEdit()
      .create();
    
    Logger.log('編集トリガーが設定されました。');
  }
  
/**
 * SPREADSHEET2（個別指導シート日報フォーム送信）用のトリガーを設定する関数
 */
function createSpreadsheet2Trigger() {
  ScriptApp.newTrigger('processMonthlyReportTrigger')
    .forSpreadsheet(SPREADSHEET2_ID)
    .onFormSubmit()
    .create();
}


/**
 * SPREADSHEET3（集団授業シート参加チェックボックス編集）用のトリガーを設定する関数
 */
function createSpreadsheet3Trigger() {
  ScriptApp.newTrigger('processMonthlyReportTrigger')
    .forSpreadsheet(SPREADSHEET3_ID)
    .onEdit()
    .create();
}