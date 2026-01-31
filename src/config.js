/**
 * 列ヘッダー名の定義
 * ヘッダー名を変更する場合は、ここの値を修正するだけで全関数に適用されます。
 */
const COLUMN_DEF = {
  TEAM_NAME: 'チーム名',
  TO: '送付先メアド(TO)',
  CC: '送付先メアド(CC)',
  PERIOD: '抽出期間(Nか月)',
  START_TIME: '開始時刻',
  END_TIME: '終了時刻',
  HOLIDAY: '土日祝日',
  ORDER_NUMBER: 'ブイキューブ発注番号'
};


// Configuration - Update these values with your actual spreadsheet IDs
const CONFIG = {
  // Devin Sheet (Status tracking and Master list)
  DEVIN_SPREADSHEET_ID: '1AMnu87tF2nSZlHTzdIaOseLlppS3NTGeP6faEvJGFoU', // "IoT_サブスク管理台帳_Devin"
  HISTORY_SHEET_NAME: 'チーム追加履歴', // Changed from Devin管理シート
  DUPLICATE_REPORT_SHEET_NAME: '重複チェックレポート',

  // Source spreadsheet ID (the one to read from)
  SOURCE_SPREADSHEET_ID: '1AMnu87tF2nSZlHTzdIaOseLlppS3NTGeP6faEvJGFoU',
  SOURCE_SHEET_NAME: '人感センサー設定シート',

  // ★★★ 新規登録チェック用（調達側シート）の設定 ★★★
  PROCUREMENT_SPREADSHEET_ID: '1JsrXXLbKfZTtcCCKX5HS62BOeZ4N39WwhrvUvBCjlZk', // 調達側シートのID
  PROCUREMENT_SHEET_NAME: '人感センサー設定シート', // ※仮設定（必要に応じて変更してください）
  PROCUREMENT_DATA_RANGE: 'B21:L',                 // ※変更: 最終行まで自動で読み込む
  PROCUREMENT_HEADER_ROW_NUM: 18,                  // ★追加: ヘッダー行の番号

  // Target spreadsheet ID (the one to write to)
  TARGET_SPREADSHEET_ID: '1AMnu87tF2nSZlHTzdIaOseLlppS3NTGeP6faEvJGFoU',
  TARGET_SHEET_NAME: '出力',
  CLEANED_DATA_SHEET_NAME: '4.To・CCなし行削除', // The final, cleaned data list


  // Column mapping configuration spreadsheet
  MAPPING_SPREADSHEET_ID: '1AMnu87tF2nSZlHTzdIaOseLlppS3NTGeP6faEvJGFoU',
  MAPPING_SHEET_NAME: 'マッピング設定',
  MAPPING_DATA_RANGE: 'A2:C100',    // Range containing mapping rules (Source Column, Target Column, Description)

  // Source data range settings
  SOURCE_TITLE_RANGE: 'B18:L18',     // Header/title row range
  SOURCE_DATA_RANGE: 'B21:L1150',   // Data rows range (excluding headers)

  // Target data range settings
  TARGET_TITLE_RANGE: 'A1:K1',     // Where to write headers in target
  TARGET_DATA_START_CELL: 'A2',    // Where to start writing data (after headers)

  // ★★★「メール送付機能が情報を取得するファイル」のIDをここに設定してください
  // NEW: Configuration for the Mail Settings Spreadsheet
  MAIL_SETTINGS_SPREADSHEET_ID: '1AMnu87tF2nSZlHTzdIaOseLlppS3NTGeP6faEvJGFoU',
  MAIL_SETTINGS_SHEET_NAME: '5.設定上書き',

  // NEW: Configuration for the results sheet
  RESULTS_SHEET_NAME: '6.レポート設定結果', // ★★★ 差分を出力するシート名

  // Notification settings
  NOTIFICATION_EMAIL: 'suzuki@example.com', // ★★★ 通知先メールアドレスを設定してください

  // Options
  INCLUDE_HEADERS: true,            // Whether to transfer headers
  CLEAR_TARGET_BEFORE_WRITE: true  // Whether to clear target sheet before writing
};