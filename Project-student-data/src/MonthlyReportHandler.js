/**
 * メイン・フォーム送信・SPREADSHEET3の編集イベントのいずれからも呼び出すトリガーハンドラ
 * ※編集イベントの場合は、MAIN_SHEETではG列（7列目ではなく8列目チェックになっていたので注意）、SPREADSHEET3ではN列（14列目）の変更のみ処理
 */
function processMonthlyReportTrigger(e) {
  // 編集イベントの場合、セルの値が "TRUE" でなければ処理を中断
  if (e.range && e.value !== "TRUE") return;
  
  var sheet, row, headers, studentName, teacherName = "";
  
  // 編集イベントの場合（チェックボックス編集など）
  if (e.range) {
    sheet = e.range.getSheet();
    row = e.range.getRow();
    var editedCol = e.range.getColumn();
    var ssId = e.source.getId();
    
    // MAIN_SHEETの場合は8列目以外はスキップ（※G列が8列目ならこのままでOK）
    if (ssId === MAIN_SHEET_ID && editedCol !== 8) {
      return;
    }
    // SPREADSHEET3の場合は16(P)列目以外はスキップ
    if (ssId === SPREADSHEET3_ID && editedCol !== 16) {
      return;
    }
    
    // シートごとにヘッダー行や対象のヘッダーを切り分ける
    if (ssId === MAIN_SHEET_ID) {
      // MAIN_SHEET：ヘッダーは1行目
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var idx = headers.indexOf("生徒様のお名前");
      if (idx === -1) {
        Logger.log("MAIN_SHEETで生徒名のヘッダー「生徒様のお名前」が見つかりません。");
        return;
      }
      studentName = sheet.getRange(row, idx + 1).getValue();
    } else if (ssId === SPREADSHEET3_ID) {
      // SPREADSHEET3：ヘッダーは2行目
      headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      var idx = headers.indexOf("LINEの名前");
      if (idx === -1) {
        Logger.log("SPREADSHEET3で生徒名のヘッダー「LINEの名前」が見つかりません。");
        return;
      }
      studentName = sheet.getRange(row, idx + 1).getValue();
    }
    
    // 担当教師（あれば）の取得
    var teacherIdx = headers.indexOf("担当教師");
    if (teacherIdx !== -1) {
      teacherName = sheet.getRange(row, teacherIdx + 1).getValue();
    }
    
  // フォーム送信（onFormSubmit）イベントの場合
  } else if (e.values) {
    // ※スタンドアロンの場合、getActiveSheet()は期待通り動作しないことがあるため、
    //    e.source からシートを明示的に取得するか、シート名が既知なら getSheetByName() を利用してください。
    // ここでは例として、フォーム送信先シート名が "フォームの回答 1" であると仮定しています。
    sheet = e.source.getSheetByName("フォームの回答 1");
    if (!sheet) {
      Logger.log("フォーム送信先シートが見つかりません。");
      return;
    }
    // フォーム回答のC列（3列目）を studentName として取得
    studentName = e.values[2]; // e.values は 0-indexed
    teacherName = ""; // 担当教師名は空文字列
  }
  
  if (!studentName) {
    Logger.log("生徒名が空のため処理を中断します。");
    return;
  }
  
  // 一度だけ開く
  var ss2 = SpreadsheetApp.openById(SPREADSHEET2_ID);
  var ss3 = SpreadsheetApp.openById(SPREADSHEET3_ID);

  // メイン処理呼び出しに ss2, ss3 を渡す
  processMonthlyReport(studentName, teacherName, ss2, ss3);
  
  // 編集イベントの場合、チェックボックスを OFF に戻す（MAIN_SHEET のときのみ）
  if (e.range && e.source.getId() === MAIN_SHEET_ID) {
    e.range.setValue(false);
  }
}



/**
 * メイン処理：対象生徒のファイル群作成＆データ転記（最適化版）
 * @param {string} studentName
 * @param {string} teacherName
 * @param {Spreadsheet} ss2
 * @param {Spreadsheet} ss3
 */
function processMonthlyReport(studentName, teacherName, ss2, ss3) {
  // 対象フォルダの取得
  var targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  // フォルダ名に studentName を含むフォルダを検索（部分一致対応）
  var folderIterator = targetFolder.getFolders();
  var studentFolder = targetFolder;
  while (folderIterator.hasNext()) {
    var f = folderIterator.next();
    if (f.getName().indexOf(studentName) !== -1) {
      studentFolder = f;
      break;
    }
  }
  
  // ファイル名の作成
  var dataFileName   = studentName + "さん_月間報告（データ格納用）";
  var sharedFileName = studentName + "さん_月間報告（共有用）";
  
  // 上書き処理：同名ファイルがあれば削除してからコピー
  var dataFile   = getOrOverwriteFile(SPREADSHEET4_ID, dataFileName, studentFolder);
  var sharedFile = getOrOverwriteFile(SPREADSHEET5_ID, sharedFileName, studentFolder);
  
  var dataSpreadsheet = SpreadsheetApp.open(dataFile);

  // --- 個別指導記録用：ヘッダー追加 & 一括転記 ---
  var dstIndiv = dataSpreadsheet.getSheetByName("個別指導記録用");
  ensureHeader(dstIndiv, ss2.getSheets()[0], 2);
  batchCopyRows(ss2.getSheets()[0], dstIndiv, "生徒名（漢字）", studentName);

  // --- 集団授業記録用：ヘッダー追加 & 一括転記 ---
  var dstGroup = dataSpreadsheet.getSheetByName("集団授業記録用");
  ensureHeader(dstGroup, ss3.getSheets()[0], 2);
  batchCopyRows(ss3.getSheets()[0], dstGroup, "LINEの名前", studentName);

  // --- 集団授業シートのソート（教科→日程） ---
  sortSheet(dstGroup, ["教科", "日程"]);
  

  
  // ★ 月間報告生成：当月分＋25日以降は翌月分も実行 ★
  var today = new Date();
  var dates = [today];
  if (today.getDate() >= 25) {
    // 翌月1日を追加
    dates.push(new Date(today.getFullYear(), today.getMonth() + 1, 1));
  }
  dates.forEach(function(d) {
    // createMonthlyReport を日付指定で呼び出し
    var sheetName = createMonthlyReport(studentName, teacherName, dataSpreadsheet, d);
    appendMessageDataToDataSheet(dataSpreadsheet, sheetName);
    updateSharedMonthlyReport(sheetName, dataSpreadsheet, sharedFile);
  });

  // アナウンス文の更新
  updateAnnouncementDoc(studentName, sharedFile.getUrl());
}

/**
 * ensureHeader:
 *  dstSheet が空なら、srcSheet の headerRowIndex 行目を見出しとして設定
 */
function ensureHeader(dstSheet, srcSheet, headerRowIndex) {
  if (dstSheet.getLastRow() === 0) {
    var headers = srcSheet.getRange(headerRowIndex, 1, 1, srcSheet.getLastColumn()).getValues()[0];
    dstSheet.appendRow(headers);
  }
}

/**
 * batchCopyRows:
 *   srcSheet の全データから keyHeader 列が keyValue と一致する行を
 *   まとめて dstSheet に転記します。
 *   （ヘッダー行を 1 行目・2 行目のいずれかから自動検出）
 */
function batchCopyRows(srcSheet, dstSheet, keyHeader, keyValue) {
  var data = srcSheet.getDataRange().getValues();
  // まず 1 行目にヘッダーがあるか試し、なければ 2 行目をヘッダーとみなす
  var headerRowIndex = data[0].indexOf(keyHeader) >= 0 ? 0
                       : (data[1].indexOf(keyHeader) >= 0 ? 1 : -1);
  if (headerRowIndex < 0) return;  // ヘッダーが見つからなければ終了

  var idx = data[headerRowIndex].indexOf(keyHeader);
  var rows = [];
  for (var i = headerRowIndex + 1; i < data.length; i++) {
    if (data[i][idx] === keyValue) rows.push(data[i]);
  }
  if (rows.length) {
    dstSheet
      .getRange(dstSheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
  }
}



/**
 * sortSheet:
 *  sheet のデータ行を headersToSort の順序で昇順ソートします。
 *  ヘッダーは1行目、データは2行目以降を対象
 */
function sortSheet(sheet, headersToSort) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) return;
  // 1行目をヘッダーとして取得
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var specs = [];
  headersToSort.forEach(function(h) {
    var ci = headerRow.indexOf(h);
    if (ci >= 0) specs.push({ column: ci + 1, ascending: true });
  });
  if (specs.length) {
    // データ行は2行目以降
    sheet.getRange(2, 1, lastRow - 1, lastCol).sort(specs);
  }
}




/**
 * テンプレートファイルのコピーを作成する際、以下の処理を行います。
 * ・共有用テンプレート（SPREADSHEET5_ID）の場合：
 *    - 同名ファイルが複数存在する場合、最も更新日時が新しいものを返します。
 *    - 存在しなければ新規コピーを作成します。
 * ・それ以外（データ格納用）は、同名ファイルがあれば削除（ゴミ箱へ移動）してから新規コピーを作成します。
 *
 * @param {string} templateId - テンプレートのファイルID
 * @param {string} fileName - 作成するファイル名
 * @param {Folder} folder - コピー先のフォルダオブジェクト
 * @returns {File} - 作成または取得したファイル
 */
function getOrOverwriteFile(templateId, fileName, folder) {
  var files = folder.getFilesByName(fileName);
  if (templateId === SPREADSHEET5_ID) {
    // 共有用の場合：同名ファイルがあれば最新のものを返す（上書き更新対象）
    var latestFile = null;
    while (files.hasNext()) {
      var f = files.next();
      if (!latestFile || f.getLastUpdated() > latestFile.getLastUpdated()) {
        latestFile = f;
      }
    }
    if (latestFile) {
      return latestFile;
    } else {
      return DriveApp.getFileById(templateId).makeCopy(fileName, folder);
    }
  } else {
    // データ格納用の場合：同名ファイルがあれば削除してから新規コピー
    if (files.hasNext()) {
      var existingFile = files.next();
      existingFile.setTrashed(true);
    }
    return DriveApp.getFileById(templateId).makeCopy(fileName, folder);
  }
}



/**
 * 月間報告シートを作成する関数
 * ※一日前の日付から対象の年月を決定し、対象シートがなければテンプレートからコピー
 * @param {string} studentName
 * @param {string} teacherName
 * @param {Spreadsheet} dataSpreadsheet - データ格納用スプレッドシート
 * @returns {string} - 作成または取得したシート名（例："2025年03月度"）
 */
function createMonthlyReport(studentName, teacherName, dataSpreadsheet, targetDate) {
  // targetDate が未指定なら従来どおり一日前の日付を使用
  if (!targetDate) {
    targetDate = new Date();
    targetDate.setDate(targetDate.getDate() - 1);
  }
  var year = targetDate.getFullYear();
  var month = targetDate.getMonth() + 1;
  var monthStr = (month < 10 ? "0" : "") + month;
  var targetSheetName = year + "年" + monthStr + "月度";
  
  // 1. 「処理用（YYYY/MM）」シートの作成（既に存在する場合はスキップ）
  var processSheetName = "処理用（" + year + "/" + monthStr + "）";
  var processSheet = dataSpreadsheet.getSheetByName(processSheetName);
  if (!processSheet) {
    var processTemplate = dataSpreadsheet.getSheetByName("処理用テンプレート");
    if (processTemplate) {
      processSheet = processTemplate.copyTo(dataSpreadsheet);
      processSheet.setName(processSheetName);
      // A1～I1のDATE(2025, 3, 1)部分を実行月初に置換
      updateDynamicDateInRange(processSheet, "A1", "I1");
    } else {
      Logger.log("『処理用テンプレート』シートが見つかりません。");
    }
  }
  
  // 2. 「YYYY年MM月度」シートの作成（既に存在する場合はスキップ）
  var monthlySheet = dataSpreadsheet.getSheetByName(targetSheetName);
  if (!monthlySheet) {
    var templateSheet = SpreadsheetApp.openById(SPREADSHEET4_ID).getSheetByName("月間報告テンプレート");
    monthlySheet = templateSheet.copyTo(dataSpreadsheet);
    monthlySheet.setName(targetSheetName);
    monthlySheet.getRange("A1").setValue(year);
    monthlySheet.getRange("C1").setValue(monthStr);
    monthlySheet.getRange("B3").setValue(studentName);
    // A9にARRAYFORMULAを追加
    var arrFormula = '=ARRAYFORMULA(\'' + processSheetName + '\'!A1:P)';
    monthlySheet.getRange("A9").setFormula(arrFormula);
  }
  
  // 3. 「YYYY年MM月度」シートを一番左に移動
  dataSpreadsheet.setActiveSheet(monthlySheet);
  dataSpreadsheet.moveActiveSheet(1);
  
  return targetSheetName;
}

/**
 * 指定シートのA1～O1範囲に設定されている数式の中で
 * DATE(2025, 3, 1)と記述された部分を、実行時の月初（DATE(YYYY, MM, 1)）に置換する
 * @param {Sheet} sheet - 対象のシート
 * @param {string} startCell - 開始セル（例："A1"）
 * @param {string} endCell - 終了セル（例："O1"）
 */
function updateDynamicDateInRange(sheet, startCell, endCell) {
  var range = sheet.getRange(startCell + ":" + endCell);
  var formulas = range.getFormulas()[0];
  var today = new Date();
  var year = today.getFullYear();
  var month = today.getMonth() + 1;
  var newStartDateStr = "DATE(" + year + ", " + month + ", 1)";
  
  // A～Oは15列。H列は8番目 (0-indexedでは7番目) なので、そのセルはスキップする
  // A1とI1のみ処理（i===0 がA列、i===8 がI列）
  for (var i = 0; i < formulas.length; i++) {
    if (i !== 0 && i !== 8) continue;
    if (formulas[i]) {
      formulas[i] = formulas[i].replace(/DATE\(\s*2025\s*,\s*3\s*,\s*1\s*\)/g, newStartDateStr);
    }
  }
  range.setFormulas([formulas]);
}




/**
 * 共有用スプレッドシートへの転記および共有設定を行う関数
 * 数式がある場合でも、見た目の表示値のみを転記し、書式（セル幅・色・文字装飾等）も保持します。
 * 同名シートがすでに存在しても中身のみを上書きします。
 *
 * @param {string} monthlySheetName - 対象月のシート名
 * @param {Spreadsheet} dataSpreadsheet - データ格納用スプレッドシート
 * @param {File} sharedFile - 共有用ファイル（テンプレートからコピーしたもの）
 */
function updateSharedMonthlyReport(monthlySheetName, dataSpreadsheet, sharedFile) {
  const sharedSpreadsheet = SpreadsheetApp.open(sharedFile);
  const monthlySheet = dataSpreadsheet.getSheetByName(monthlySheetName);
  if (!monthlySheet) {
    Logger.log("データ格納用シートに対象の月間報告シートがありません。");
    return;
  }

  // 対象の共有用シートを取得（存在しなければ新規作成）
  let copiedSheet = sharedSpreadsheet.getSheetByName(monthlySheetName);
  if (!copiedSheet) {
    copiedSheet = monthlySheet.copyTo(sharedSpreadsheet);
    copiedSheet.setName(monthlySheetName);
    // 新規作成したシートをアクティブにし、左端に移動する処理
    sharedSpreadsheet.setActiveSheet(copiedSheet);
    sharedSpreadsheet.moveActiveSheet(1);
  }

  // コピー元の範囲とプロパティの取得
  const sourceRange = monthlySheet.getDataRange();
  const numRows = sourceRange.getNumRows();
  const numCols = sourceRange.getNumColumns();
  const values = sourceRange.getValues();
  const numberFormats = sourceRange.getNumberFormats();
  const backgrounds = sourceRange.getBackgrounds();
  const fontColors = sourceRange.getFontColors();
  const fontWeights = sourceRange.getFontWeights();
  const fontStyles = sourceRange.getFontStyles();
  const fontSizes = sourceRange.getFontSizes();
  const horizontalAlignments = sourceRange.getHorizontalAlignments();
  const verticalAlignments = sourceRange.getVerticalAlignments();
  const wrapStrategies = sourceRange.getWrapStrategies();

  // 貼り付け先の範囲に対し、まずセル内容のみをクリア（書式はそのまま）
  const targetRange = copiedSheet.getRange(1, 1, numRows, numCols);
  targetRange.clear({ contentsOnly: true });
  
  // 値の転記
  targetRange.setValues(values);

  // コピー元の書式設定（数値・日付のフォーマットも含む）を転記
  targetRange.setNumberFormats(numberFormats);
  targetRange.setBackgrounds(backgrounds);
  targetRange.setFontColors(fontColors);
  targetRange.setFontWeights(fontWeights);
  targetRange.setFontStyles(fontStyles);
  targetRange.setFontSizes(fontSizes);
  targetRange.setHorizontalAlignments(horizontalAlignments);
  targetRange.setVerticalAlignments(verticalAlignments);
  targetRange.setWrapStrategies(wrapStrategies);

  // セル幅のコピー
  for (let col = 1; col <= numCols; col++) {
    copiedSheet.setColumnWidth(col, monthlySheet.getColumnWidth(col));
  }
  // 行の高さのコピー
  for (let row = 1; row <= numRows; row++) {
    copiedSheet.setRowHeight(row, monthlySheet.getRowHeight(row));
  }
  // 範囲外の余分な行・列をクリア
  const existingMaxRows = copiedSheet.getMaxRows();
  const existingMaxCols = copiedSheet.getMaxColumns();
  if (existingMaxRows > numRows) {
    copiedSheet.getRange(numRows + 1, 1, existingMaxRows - numRows, existingMaxCols).clear();
  }
  if (existingMaxCols > numCols) {
    copiedSheet.getRange(1, numCols + 1, numRows, existingMaxCols - numCols).clear();
  }
  
  // シート更新の反映
  SpreadsheetApp.flush();

  // B列9行目以降のみ、手動で「表示形式→数字→自動」にチェックしたのと同等の処理を適用
  if (numRows >= 9) {
    // B列は2列目。9行目以降の行数は (numRows - 8)
    const bRange = copiedSheet.getRange(9, 2, numRows - 8, 1);
    bRange.setNumberFormat("General");

    // K列は11列目。9行目以降の行数は (numRows - 8)
    const kRange = copiedSheet.getRange(9, 11, numRows - 8, 1);
    kRange.setNumberFormat("HH:mm");
  }

  // ▼▼▼ ここから条件付き書式の設定 ▼▼▼
  //L列からM列に変更したが変数名はそのまま
  const maxRow = copiedSheet.getMaxRows();
  const jRange = copiedSheet.getRange("J9:J" + maxRow);
  const lRange = copiedSheet.getRange("M9:M" + maxRow);

  // 追加したいルールをまとめて生成
  const newRules = [
    // J列：今日
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=J9=TODAY()")
      .setBackground("#d9ead3")
      .setRanges([jRange])
      .build(),
    // J列：未来
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=J9>TODAY()")
      .setBackground("#FFF2CC")
      .setRanges([jRange])
      .build(),
    // L列：英語
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("英語")
      .setBackground("#F4CCCC")
      .setRanges([lRange])
      .build(),
    // L列：数学
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("数学")
      .setBackground("#CFE2F3")
      .setRanges([lRange])
      .build(),
    // L列：理科
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("理科")
      .setBackground("#FFF2CC")
      .setRanges([lRange])
      .build()
  ];

  // 既存のルールを取得し、重複しないものだけ追加
  let rules = copiedSheet.getConditionalFormatRules();
  newRules.forEach(newRule => {
    const exists = rules.some(oldRule => {
      const oldRanges = oldRule.getRanges().map(r => r.getA1Notation()).join();
      const newRanges = newRule.getRanges().map(r => r.getA1Notation()).join();
      const oldBg = oldRule.getBooleanCondition()?.getBackground();
      const newBg = newRule.getBooleanCondition()?.getBackground();
      return oldRanges === newRanges && oldBg === newBg;
    });
    if (!exists) rules.push(newRule);
  });

  // まとめて適用
  copiedSheet.setConditionalFormatRules(rules);
  // ▲▲▲ 条件付き書式の設定ここまで ▲▲▲

  // リンク共有設定
  sharedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // テンプレートからコピーされた際に残る「月間報告」シートがあれば削除
  const templateSheet = sharedSpreadsheet.getSheetByName("月間報告");
  if (templateSheet) {
    sharedSpreadsheet.deleteSheet(templateSheet);
  }
}



/**
 * 対象フォルダ内で、studentName が含まれるフォルダ内にある
 * "studentNameさん_先生共有用アナウンス文" ドキュメントの本文中の "[monthlysheet]" を
 * sharedSpreadsheetLink に置換します。
 * 存在しない場合はスキップします。
 *
 * @param {string} studentName - 生徒名
 * @param {string} sharedSpreadsheetLink - 共有用スプレッドシートのURL
 */
function updateAnnouncementDoc(studentName, sharedSpreadsheetLink) {
  // TARGET_FOLDER_ID で指定されたフォルダを取得
  var targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  var folderIterator = targetFolder.getFolders();
  var studentFolder = null;
  
  // studentName が含まれるフォルダを探す
  while (folderIterator.hasNext()) {
    var folder = folderIterator.next();
    if (folder.getName().indexOf(studentName) !== -1) {
      studentFolder = folder;
      break;
    }
  }
  
  if (!studentFolder) {
    Logger.log("該当フォルダが見つかりません: " + studentName);
    return;
  }
  
  // ドキュメント名を作成
  var docName = studentName + "さん_先生共有用アナウンス文";
  var fileIterator = studentFolder.getFilesByName(docName);
  if (!fileIterator.hasNext()) {
    Logger.log("該当ドキュメントが見つかりません: " + docName);
    return;
  }
  
  var docFile = fileIterator.next();
  var doc = DocumentApp.openById(docFile.getId());
  var body = doc.getBody();
  
  // "[monthlysheet]" を共有用スプレッドシートのリンクに置換
  body.replaceText("\\[monthlysheet\\]", sharedSpreadsheetLink);
  
  doc.saveAndClose();
  Logger.log("アナウンス文の [monthlysheet] 部分を置換しました。");
}
