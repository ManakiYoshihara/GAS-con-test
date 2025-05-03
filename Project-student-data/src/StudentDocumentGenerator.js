/**
 * メインスプレッドシートでの編集を検知するトリガー関数
 * @param {Object} e - 編集イベントオブジェクト
 */
function handleEdit(e) {
  const sheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const editedColumn = e.range.getColumn();
  
  // メインスプレッドシートか確認
  if (e.source.getId() !== MAIN_SHEET_ID) return;
  
  // 編集されたシート名がメインシートか確認
  if (sheet.getName() !== MAIN_SHEET_NAME) return;
  
  // G列（7列目）に編集があったか確認
  if (editedColumn === 7) {
    const isChecked = e.value === 'TRUE';
    if (isChecked) {
      processRow(editedRow);
    }
  }
}

/**
 * 指定された行のデータを処理する関数
 * @param {number} row - 処理する行番号
 */
function processRow(row) {
  const mainSpreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
  const mainSheet = mainSpreadsheet.getSheetByName(MAIN_SHEET_NAME);
  
  // ヘッダー行の取得（1行目と仮定）
  const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  
  // 処理対象の行データ取得
  const rowData = mainSheet.getRange(row, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  
  // ヘッダーとデータをマッピング
  let data = {};
  headers.forEach((header, index) => {
    data[header] = rowData[index];
  });
  
  // 変数へのマッピング
  const student_name = data['生徒様のお名前'] || '名前不明';
  const student_id = data['ID'] || 'ID不明';
  const teacher_name = data['担当教師'] || '担当教師不明';
  const teacher_mail = data['Gmail'] || 'メールアドレス不明';
  const first_month_num = data['初月指導回数'] || '初月指導回数不明';
  const one_class_time = data['1コマ指導時間'] || '1コマ指導時間不明';
  const each_month_num = data['毎月指導回数'] || '毎月指導回数不明';
  
  // 生徒専用フォルダを作成（Folderオブジェクトを返す）
  const studentFolder = createStudentFolder(student_name, student_id);
  const studentFolderUrl = studentFolder.getUrl();
  
  // 生徒専用フォルダ内に授業動画フォルダを作成する（存在しない場合のみ）
  const lessonFolderUrl = createLessonVideoFolder(studentFolder, student_name, teacher_mail);
  
  // 生徒カルテ・指導報告フォームリンク（既存の場合は削除して新規生成）
  const newSpreadsheetUrl = copyTemplateSpreadsheet(data, student_name, studentFolder);
  Logger.log(`スプレッドシートが作成されました: ${newSpreadsheetUrl}`);

  // 学習プランシート（既に存在する場合は新規生成しない）
  const learningPlanUrl = copyLearningPlanTemplate(student_name, studentFolder);
  Logger.log(`学習プランシートが作成されました: ${learningPlanUrl}`);
  
  // 挨拶文（既に存在する場合は新規生成しない）
  const newDocUrl = copyAndReplaceDoc(student_name, teacher_name, studentFolder);
  
  // フォームURLを生成
  const formUrl = generatePrefilledFormUrl(teacher_name, student_name);
  const repformUrl = generatePrefilledRepFormUrl(teacher_name, student_name);
  
  if (formUrl) { // formUrlが有効な場合のみ記録
    recordFormLink(newSpreadsheetUrl, formUrl);
  } else {
    Logger.log('フォームURLが生成されませんでした。フォームリンクへの記録をスキップします。');
  }

  // 共有処理： teacher_mail が「メールアドレス不明」ではない場合のみ実施
  if (teacher_mail !== 'メールアドレス不明') {
    shareFileWithUser(newSpreadsheetUrl, teacher_mail);
    shareFileWithUser(learningPlanUrl, teacher_mail);
    // ※ 授業動画フォルダは作成時に共有設定済み
  } else {
    Logger.log('teacher_mail が「メールアドレス不明」のため、共有はスキップします。');
  }

  // 先生共有用アナウンス文（既存の場合は削除して新規生成、[lessonfolderUrl]を置換）
  const announcementDocUrl = generateAnnouncementDocFromTemplate(
    student_name,
    newSpreadsheetUrl,
    formUrl,
    repformUrl,
    learningPlanUrl,
    studentFolder,
    lessonFolderUrl,
    first_month_num,
    one_class_time,
    each_month_num
  );
  Logger.log(`アナウンス文書が作成されました: ${announcementDocUrl}`);
  
  // 処理完了後にチェックボックスをオフにする
  mainSheet.getRange(row, 7).setValue(false);
}

/**
 * 生徒専用フォルダを作成し、そのFolderオブジェクトを返す関数
 * @param {string} student_name - 生徒様のお名前
 * @param {string} student_id - 生徒様のID
 * @returns {Folder} - 生徒専用フォルダのFolderオブジェクト
 */
function createStudentFolder(student_name, student_id) {
  const destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  const folderName = `${student_id}_${student_name}さん`;
  
  // 既にフォルダが存在するか確認
  const folders = destinationFolder.getFoldersByName(folderName);
  let studentFolder;
  
  if (folders.hasNext()) {
    studentFolder = folders.next();
  } else {
    // 新しいフォルダを作成
    studentFolder = destinationFolder.createFolder(folderName);
    Logger.log(`新しい生徒専用フォルダが作成されました: ${studentFolder.getUrl()}`);
  }
  
  return studentFolder;
}

/**
 * 生徒専用フォルダ内に「オンライン学習塾のLaf_[student_name]様授業動画」フォルダを作成し、共有設定を行う関数
 * @param {Folder} studentFolder - 生徒専用フォルダのFolderオブジェクト
 * @param {string} student_name - 生徒様のお名前
 * @param {string} teacher_mail - 担当教師のメールアドレス
 * @returns {string} - 作成された授業動画フォルダのURL
 */
function createLessonVideoFolder(studentFolder, student_name, teacher_mail) {
  const folderName = `オンライン学習塾のLaf_${student_name}様授業動画`;
  const folders = studentFolder.getFoldersByName(folderName);
  let lessonFolder;
  
  if (folders.hasNext()) {
    lessonFolder = folders.next();
  } else {
    lessonFolder = studentFolder.createFolder(folderName);
    Logger.log(`新しい授業動画フォルダが作成されました: ${lessonFolder.getUrl()}`);
    
    // teacher_mail に編集権限を付与（有効な場合）
    if (teacher_mail && teacher_mail !== 'メールアドレス不明') {
      lessonFolder.addEditor(teacher_mail);
    }
    // リンクを知っている人全員に閲覧権限を付与
    lessonFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  
  return lessonFolder.getUrl();
}

/**
 * テンプレートスプレッドシートをコピーし、データを「記録用」シートに挿入する関数  
 * 生徒カルテ・指導報告フォームリンクは、同名ファイルが存在する場合は削除して新規生成する
 * @param {Object} data - 行データのオブジェクト
 * @param {string} student_name - 生徒様のお名前
 * @param {Folder} folder - 生徒専用フォルダのFolderオブジェクト
 * @returns {string} - 新しく作成されたスプレッドシートのURL
 */
function copyTemplateSpreadsheet(data, student_name, folder) {
  const newFileName = `${student_name}さん_生徒カルテ・指導報告フォームリンク`;
  // 同名ファイルの存在チェックと削除
  const existingFiles = folder.getFilesByName(newFileName);
  while (existingFiles.hasNext()) {
    const file = existingFiles.next();
    file.setTrashed(true);
    Logger.log(`既存の生徒カルテファイルを削除しました: ${file.getUrl()}`);
  }
  
  const templateFile = DriveApp.getFileById(TEMPLATE_SHEET_ID);
  const newFile = templateFile.makeCopy(newFileName, folder);
  
  // 新しいスプレッドシートを開く
  const newSpreadsheet = SpreadsheetApp.open(newFile);
  
  // 「記録用」シートを取得または作成
  let recordSheet = newSpreadsheet.getSheetByName(RECORD_SHEET_NAME);
  if (!recordSheet) {
    recordSheet = newSpreadsheet.insertSheet(RECORD_SHEET_NAME);
  }
  
  const mainSpreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
  const mainSheet = mainSpreadsheet.getSheetByName(MAIN_SHEET_NAME);
  const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  
  const rowData = [];
  headers.forEach(header => {
    rowData.push(data[header]);
  });
  
  // ヘッダーが存在しない場合、ヘッダーを追加
  const lastRow = recordSheet.getLastRow();
  if (lastRow === 0) {
    recordSheet.appendRow(headers);
  }
  recordSheet.appendRow(rowData);
  
  return newSpreadsheet.getUrl();
}

/**
 * 学習プランテンプレートをコピーする関数  
 * 学習プランシートは、同名ファイルが存在する場合は新規生成しない
 * @param {string} student_name - 生徒様のお名前
 * @param {Folder} folder - 生徒専用フォルダのFolderオブジェクト
 * @returns {string} - 学習プランスプレッドシートのURL
 */
function copyLearningPlanTemplate(student_name, folder) {
  const newFileName = `${student_name}さん_学習プランシート`;
  const existingFiles = folder.getFilesByName(newFileName);
  if (existingFiles.hasNext()) {
    const existingFile = existingFiles.next();
    Logger.log(`既存の学習プランシートが使用されます: ${existingFile.getUrl()}`);
    return SpreadsheetApp.open(existingFile).getUrl();
  }
  
  try {
    const templateFile = DriveApp.getFileById(LEARNING_PLAN_TEMPLATE_ID);
    const newFile = templateFile.makeCopy(newFileName, folder);
    const newSpreadsheet = SpreadsheetApp.open(newFile);
    const newUrl = newSpreadsheet.getUrl();
    Logger.log(`学習プランシートが作成されました: ${newUrl}`);
    return newUrl;
  } catch (error) {
    Logger.log(`copyLearningPlanTemplateでエラーが発生しました: ${error}`);
    return null;
  }
}

/**
 * ドキュメントテンプレートをコピーし、プレースホルダーを置換する関数  
 * 挨拶文は、同名ファイルが存在する場合は新規生成しない
 * @param {string} student_name - 生徒様のお名前
 * @param {string} teacher_name - 担当教師の名前
 * @param {Folder} folder - 生徒専用フォルダのFolderオブジェクト
 * @returns {string|null} - 挨拶文ドキュメントのURL（失敗時はnull）
 */
function copyAndReplaceDoc(student_name, teacher_name, folder) {
  const newDocName = `${student_name}さん_挨拶文`;
  const existingDocs = folder.getFilesByName(newDocName);
  if (existingDocs.hasNext()) {
    const existingDoc = existingDocs.next();
    Logger.log(`既存の挨拶文ドキュメントが使用されます: ${existingDoc.getUrl()}`);
    return DocumentApp.openById(existingDoc.getId()).getUrl();
  }
  
  try {
    const docTemplateFile = DriveApp.getFileById(DOC_TEMPLATE_ID);
    const newDocFile = docTemplateFile.makeCopy(newDocName, folder);
    const newDoc = DocumentApp.openById(newDocFile.getId());
    const body = newDoc.getBody();
    
    body.replaceText('\\[student_name\\]', student_name);
    body.replaceText('\\[teacher_name\\]', teacher_name);
    
    newDoc.saveAndClose();
    
    const newDocUrl = newDoc.getUrl();
    Logger.log(`ドキュメントが作成されました: ${newDocUrl}`);
    
    return newDocUrl;
  } catch (error) {
    Logger.log(`copyAndReplaceDocでエラーが発生しました: ${error}`);
    return null;
  }
}

/**
 * URLからGoogleドライブフォルダのIDを抽出する関数
 * @param {string} folderUrl - フォルダのURL
 * @returns {string} - フォルダID
 */
function extractFolderIdFromUrl(folderUrl) {
  const match = folderUrl.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * 教師名と生徒名を基に事前入力フォームURLを生成する関数
 * @param {string} teacher_name - 担当教師の名前
 * @param {string} student_name - 生徒様のお名前
 * @returns {string} - 生成されたフォームURL
 */
function generatePrefilledFormUrl(teacher_name, student_name) {
  const encodedTeacherName = encodeURIComponent(teacher_name);
  const encodedStudentName = encodeURIComponent(student_name);
  return `${FORM_BASE_URL}?${TEACHER_ENTRY_ID}=${encodedTeacherName}&${STUDENT_ENTRY_ID}=${encodedStudentName}`;
}

/**
 * 教師名と生徒名を基に事前入力面談報告フォームURLを生成する関数
 * @param {string} teacher_name - 担当教師の名前
 * @param {string} student_name - 生徒様のお名前
 * @returns {string} - 生成されたフォームURL
 */
function generatePrefilledRepFormUrl(teacher_name, student_name) {
  const encodedTeacherName = encodeURIComponent(teacher_name);
  const encodedStudentName = encodeURIComponent(student_name);
  return `${REP_FORM_BASE_URL}?${TEACHER_ENTRY_ID}=${encodedTeacherName}&${STUDENT_ENTRY_ID}=${encodedStudentName}`;
}

/**
 * コピー先スプレッドシートの「フォームリンク」シートにフォームURLを記録する関数
 * @param {string} spreadsheetUrl - コピーされたスプレッドシートのURL
 * @param {string} formUrl - 生成されたフォームのURL
 */
function recordFormLink(spreadsheetUrl, formUrl) {
  if (!formUrl) {
    Logger.log('フォームURLが存在しません。recordFormLinkをスキップします。');
    return;
  }
  
  const copiedSpreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  let linkSheet = copiedSpreadsheet.getSheetByName(LINK_SHEET_NAME);
  if (!linkSheet) {
    linkSheet = copiedSpreadsheet.insertSheet(LINK_SHEET_NAME);
    linkSheet.appendRow(['フォームURL']);
  }
  
  try {
    linkSheet.appendRow([formUrl]);
    Logger.log(`フォームリンクが記録されました: ${formUrl}`);
  } catch (error) {
    Logger.log(`recordFormLinkでエラーが発生しました: ${error}`);
  }
}

/**
 * 指定されたURLからファイルIDを抽出し、対象ファイルを指定のユーザーに編集権限で共有する関数
 * @param {string} fileUrl - 共有するファイルのURL
 * @param {string} userEmail - 共有先のメールアドレス
 */
function shareFileWithUser(fileUrl, userEmail) {
  if (!fileUrl || !userEmail) return;
  const idMatch = fileUrl.match(/[-\w]{25,}/);
  if (idMatch) {
    const fileId = idMatch[0];
    const file = DriveApp.getFileById(fileId);
    file.addEditor(userEmail);
    Logger.log(`${file.getName()} を ${userEmail} に共有しました。`);
  } else {
    Logger.log(`ファイルIDがURLから取得できませんでした: ${fileUrl}`);
  }
}

/**
 * アナウンス文用のテンプレートドキュメントをコピーし、プレースホルダーを置換する関数  
 * 先生共有用アナウンス文は、同名ファイルが存在する場合は削除して新規生成する  
 * また、[lessonfolderUrl] プレースホルダーを授業動画フォルダのリンクに置換する
 * @param {string} student_name - 生徒様のお名前
 * @param {string} spreadsheetUrl - 生徒カルテ・指導報告フォームリンクのURL
 * @param {string} formUrl - 指導報告フォームのURL
 * @param {string} repformUrl - 面談報告フォームのURL
 * @param {string} learningPlanUrl - 学習プランシートのURL
 * @param {Folder} studentFolder - 生徒専用フォルダのFolderオブジェクト
 * @param {string} lessonFolderUrl - 授業動画フォルダのURL
 * @param {string} first_month_num - 初月指導回数
 * @returns {string} - 生成されたアナウンス文ドキュメントのURL
 */
function generateAnnouncementDocFromTemplate(student_name, spreadsheetUrl, formUrl, repformUrl, learningPlanUrl, studentFolder, lessonFolderUrl, first_month_num, one_class_time, each_month_num) {
  const newDocName = `${student_name}さん_先生共有用アナウンス文`;
  
  // 生徒専用フォルダ内に同名ファイルが存在する場合は削除
  const existingDocs = studentFolder.getFilesByName(newDocName);
  while (existingDocs.hasNext()) {
    const file = existingDocs.next();
    file.setTrashed(true);
    Logger.log(`既存の先生共有用アナウンス文を削除しました: ${file.getUrl()}`);
  }
  
  const templateFile = DriveApp.getFileById(ANNOUNCEMENT_TEMPLATE_ID);
  // 一旦ルートにコピー作成（後で生徒専用フォルダに移動）
  const newDocFile = templateFile.makeCopy(newDocName);
  
  const doc = DocumentApp.openById(newDocFile.getId());
  const body = doc.getBody();
  
  // プレースホルダーの置換
  body.replaceText('\\[student_name\\]', student_name);
  body.replaceText('\\[spreadsheetUrl\\]', spreadsheetUrl);
  body.replaceText('\\[formUrl\\]', formUrl);
  body.replaceText('\\[repformUrl\\]', repformUrl);
  body.replaceText('\\[learningPlanUrl\\]', learningPlanUrl);
  body.replaceText('\\[first_month_num\\]', first_month_num);
  body.replaceText('\\[one_class_time\\]', one_class_time);
  body.replaceText('\\[each_month_num\\]', each_month_num);
  body.replaceText('\\[lessonfolderUrl\\]', lessonFolderUrl);
  
  doc.saveAndClose();
  
  // 作成したドキュメントを生徒専用フォルダへ移動
  studentFolder.addFile(newDocFile);
  DriveApp.getRootFolder().removeFile(newDocFile);
  
  Logger.log('アナウンス文ドキュメントが作成されました: ' + doc.getUrl());
  return doc.getUrl();
}
