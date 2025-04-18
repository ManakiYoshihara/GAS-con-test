const scriptProps = PropertiesService.getScriptProperties().getProperties();

// ▼ スプレッドシートおよびテンプレート関連のID
const MAIN_SHEET_ID = scriptProps.MAIN_SHEET_ID;
const TEMPLATE_SHEET_ID = scriptProps.TEMPLATE_SHEET_ID;
const LEARNING_PLAN_TEMPLATE_ID = scriptProps.LEARNING_PLAN_TEMPLATE_ID;
const DOC_TEMPLATE_ID = scriptProps.DOC_TEMPLATE_ID;
const ANNOUNCEMENT_TEMPLATE_ID = scriptProps.ANNOUNCEMENT_TEMPLATE_ID;
const DESTINATION_FOLDER_ID = scriptProps.DESTINATION_FOLDER_ID;

// ▼ 月報処理関連のID
const SPREADSHEET2_ID = scriptProps.SPREADSHEET2_ID; // 個別指導
const SPREADSHEET3_ID = scriptProps.SPREADSHEET3_ID; // 集団授業
const SPREADSHEET4_ID = scriptProps.SPREADSHEET4_ID; // 月報テンプレート
const SPREADSHEET5_ID = scriptProps.SPREADSHEET5_ID; // 共有用テンプレート
const TARGET_FOLDER_ID = scriptProps.TARGET_FOLDER_ID; // 出力フォルダ

// ▼ シート名
const MAIN_SHEET_NAME = scriptProps.MAIN_SHEET_NAME;
const RECORD_SHEET_NAME = scriptProps.RECORD_SHEET_NAME;
const LINK_SHEET_NAME = scriptProps.LINK_SHEET_NAME;

// ▼ Googleフォーム関連
const FORM_BASE_URL = scriptProps.FORM_BASE_URL;
const REP_FORM_BASE_URL = scriptProps.REP_FORM_BASE_URL;
const TEACHER_ENTRY_ID = scriptProps.TEACHER_ENTRY_ID;
const STUDENT_ENTRY_ID = scriptProps.STUDENT_ENTRY_ID;
