// 売上管理スプレッドシート構築スクリプト

/**
 * 定数定義
 */
const CONSTANTS = {
  // データ行数と合計行
  DATA_ROWS: 65,
  TOTAL_ROW: 67,
  MAX_CLIENT_ROWS: 1000,
  YEARLY_SHEET_MONTHS: 12,
  
  // 行の高さ
  HEADER_ROW_HEIGHT: 35,
  DATA_ROW_HEIGHT: 28,
  TOTAL_ROW_HEIGHT: 35,
  
  // ステータスオプション
  STATUS_OPTIONS: ['確認中', '入金前', '入金済み', '納品済み', '固定費', '振り分け済み', '作成済み', '振込前', '振込済み'],
  
  // ステータスごとの色設定
  STATUS_COLORS: {
    '確認中': '#FFF9C4',      // 薄い黄色
    '入金前': '#FFE0B2',      // 薄いオレンジ
    '入金済み': '#C8E6C9',    // 薄い緑
    '納品済み': '#BBDEFB',    // 薄い青
    '固定費': '#E0E0E0',       // 薄いグレー
    '振り分け済み': '#E1BEE7', // 薄い紫
    '作成済み': '#B2EBF2',     // 薄い青緑
    '振込前': '#F8BBD0',       // 薄いピンク
    '振込済み': '#A5D6A7'      // 薄い緑
  },
  
  // クライアント列の背景色（値が入力されている場合）
  CLIENT_COLOR: '#E3F2FD',     // 薄い青
  
  // デフォルト値
  DEFAULT_START_MONTH: 11,
  DEFAULT_EXCHANGE_RATE: 150,
  EXCHANGE_RATE_TIMEOUT: 10000
};

/**
 * ログシートを作成
 */
function createLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'ログ';
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    // 既存のログシートがある場合は、フォーマットのみ更新
    formatLogSheet(sheet);
    return sheet;
  }
  
  sheet = ss.insertSheet(sheetName);
  // シートを先頭に移動
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
  
  // ヘッダー行を設定
  const headers = ['日時', 'レベル', '関数名', 'メッセージ'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // フィルタを適用
  const dataRange = sheet.getRange(1, 1, 10000, headers.length);
  dataRange.createFilter();
  
  // フォーマットを適用
  formatLogSheet(sheet);
  
  // 行の高さを調整
  sheet.setRowHeight(1, CONSTANTS.HEADER_ROW_HEIGHT);
  
  return sheet;
}

/**
 * ログシートのフォーマット設定
 */
function formatLogSheet(sheet) {
  // ヘッダー行のフォーマット（統一色：グレー）
  const headerColor = '#9E9E9E'; // グレー
  
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setBackground(headerColor);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setBorder(true, true, true, true, false, false);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 180); // 日時
  sheet.setColumnWidth(2, 80);   // レベル
  sheet.setColumnWidth(3, 200);  // 関数名
  sheet.setColumnWidth(4, 500);  // メッセージ
}

/**
 * ログを記録（ログシートに書き込む）
 * @param {string} level ログレベル（INFO, WARNING, ERROR）
 * @param {string} functionName 関数名
 * @param {string} message メッセージ
 */
function writeLog(level, functionName, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('ログ');
    
    // ログシートが存在しない場合は作成
    if (!logSheet) {
      logSheet = createLogSheet();
    }
    
    // ログを追加（最新のログを上に追加）
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const logData = [[timestamp, level, functionName || '', message || '']];
    
    // 2行目に挿入（既存のログは下に移動）
    logSheet.insertRowBefore(2);
    logSheet.getRange(2, 1, 1, 4).setValues(logData);
    
    // ログレベルに応じて色を設定
    const rowRange = logSheet.getRange(2, 1, 1, 4);
    if (level === 'ERROR') {
      rowRange.setBackground('#ffebee'); // 薄い赤
    } else if (level === 'WARNING') {
      rowRange.setBackground('#fff3e0'); // 薄いオレンジ
    } else {
      rowRange.setBackground('#ffffff'); // 白
    }
    
    // Logger.logにも記録（デバッグ用）
    Logger.log(`[${level}] ${functionName}: ${message}`);
    
    // ログが1000行を超えた場合、古いログを削除
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1000) {
      logSheet.deleteRows(1001, lastRow - 1000);
    }
  } catch (error) {
    // ログ記録に失敗した場合はLogger.logのみ使用
    Logger.log(`ログ記録エラー: ${error.toString()}`);
    Logger.log(`[${level}] ${functionName}: ${message}`);
  }
}

/**
 * クライアントシートを作成
 */
function createClientSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'クライアント';
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
    // シートを先頭に移動
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  }
  
  // ヘッダー行を設定
  const headers = ['名前', '種別', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // サンプルデータを追加（必要に応じて削除可能）
  sheet.getRange(2, 1, 1, 3).setValues([['サンプルクライアント', 'クライアント', '']]);
  
  // フィルタを適用（テーブル形式）- データ範囲全体を指定
  const dataRange = sheet.getRange(1, 1, CONSTANTS.MAX_CLIENT_ROWS, 3);
  dataRange.createFilter();
  
  // フォーマットを適用
  formatClientSheet(sheet);
  
  // 行の高さを調整
  sheet.setRowHeight(1, CONSTANTS.HEADER_ROW_HEIGHT); // ヘッダー行
}

/**
 * クライアントシートのフォーマット設定
 */
function formatClientSheet(sheet) {
  // ヘッダー行のフォーマット（統一色：紫）- パフォーマンス最適化
  const headerColor = '#9C27B0'; // 紫
  
  const headerRange = sheet.getRange(1, 1, 1, 3);
  // 複数のプロパティを一度に設定（API呼び出しを最小化）
  headerRange.setBackground(headerColor);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setBorder(true, true, true, true, false, false);
  
  // 列幅の調整 - パフォーマンス最適化（必要な列のみ設定）
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 300);
}

/**
 * 年度別シートを作成
 * @param {string} sheetName シート名（例: "2026年度"）
 * @param {number} startMonth 開始月（1-12）
 */
function createYearlySheet(sheetName, startMonth = CONSTANTS.DEFAULT_START_MONTH) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
    // シートを先頭に移動
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  }
  
  // ヘッダー行を設定
  const headers = ['対象月', '経費', '売上', '利益', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 開始月から12ヶ月分のデータ行を作成
  const months = [];
  for (let i = 0; i < CONSTANTS.YEARLY_SHEET_MONTHS; i++) {
    const monthNum = ((startMonth - 1 + i) % 12) + 1;
    months.push(monthNum + '月');
  }
  
  const dataRows = [];
  months.forEach((month, index) => {
    const row = index + 2;
    dataRows.push([
      month,
      `=IFERROR(INDIRECT("'${month}'!L7"),0)`, // 経費合計の参照（当月全体）
      `=IFERROR(INDIRECT("'${month}'!L8"),0)`, // 売上合計の参照（当月全体）
      `=IFERROR(INDIRECT("'${month}'!L9"),0)`, // 利益合計の参照（当月全体）
      ''
    ]);
  });
  
  sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  
  // 合計行を追加 - バッチ処理で最適化
  const totalRow = dataRows.length + 2;
  sheet.getRange(totalRow, 1).setValue('合計');
  const totalFormulas = [
    [`=SUM(B2:B${dataRows.length + 1})`],
    [`=SUM(C2:C${dataRows.length + 1})`],
    [`=SUM(D2:D${dataRows.length + 1})`]
  ];
  sheet.getRange(totalRow, 2, 1, 3).setFormulas([totalFormulas.map(f => f[0])]);
  
  // フォーマットを適用
  formatYearlySheet(sheet);
  
  // 行の高さを調整（広く余裕を持たせる）- パフォーマンス最適化
  // SheetのsetRowHeightメソッドを使用（RangeではなくSheetから呼び出す）
  sheet.setRowHeight(1, CONSTANTS.HEADER_ROW_HEIGHT); // ヘッダー行
  // データ行の高さを一括設定（範囲の最初と最後の行を指定）
  for (let i = 2; i <= dataRows.length + 1; i++) {
    sheet.setRowHeight(i, CONSTANTS.DATA_ROW_HEIGHT);
  }
  sheet.setRowHeight(dataRows.length + 2, CONSTANTS.TOTAL_ROW_HEIGHT); // 合計行
}

/**
 * 月別シートを作成
 * @param {string} sheetName シート名（例: "11月"）
 * @param {string} yearlySheetName 年度別シート名（例: "2026年度"）
 */
function createMonthlySheet(sheetName, yearlySheetName = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
    // 年度別シートの後に配置
    if (yearlySheetName) {
      const yearlySheet = ss.getSheetByName(yearlySheetName);
      if (yearlySheet) {
        const yearlyIndex = yearlySheet.getIndex();
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(yearlyIndex + 1);
      }
    } else {
      // 年度別シートが見つからない場合は、クライアントシートの後に配置
      const clientSheet = ss.getSheetByName('クライアント');
      if (clientSheet) {
        const clientIndex = clientSheet.getIndex();
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(clientIndex + 1);
      }
    }
  }
  
  // ヘッダー行を設定
  const headers = ['項番', '受注日', '納期', 'クライアント', '業種', '経費', '経費（ドル）', '売上', '利益', 'ステータス', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 65行のデータ行を作成
  const dataRows = [];
  for (let i = 1; i <= CONSTANTS.DATA_ROWS; i++) {
    dataRows.push([i, '', '', '', '', '', '', '', '', '', '']);
  }

  sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);

  // 利益列の計算式を設定（売上 - 経費 - 経費（ドル換算））
  // パフォーマンス最適化：配列を生成して一括設定
  const profitFormulas = [];
  for (let i = 2; i <= CONSTANTS.DATA_ROWS + 1; i++) {
    profitFormulas.push([`=IF(H${i}="","",H${i}-F${i}-IF(G${i}="",0,G${i}*getExchangeRate()))`]);
  }
  sheet.getRange(2, 9, CONSTANTS.DATA_ROWS, 1).setFormulas(profitFormulas);
  sheet.getRange(2, 9, CONSTANTS.DATA_ROWS, 1).setNumberFormat('¥#,##0');
  
  // 合計行を追加（67行目）- バッチ処理で最適化
  const totalRow = CONSTANTS.TOTAL_ROW;
  sheet.getRange(totalRow, 1).setValue('合計');
  const totalFormulas = [
    [`=SUM(F2:F${CONSTANTS.DATA_ROWS + 1})`],  // 経費（円）
    [`=SUM(G2:G${CONSTANTS.DATA_ROWS + 1})*getExchangeRate()`],  // 経費（ドル）を円換算して合計
    [`=SUM(H2:H${CONSTANTS.DATA_ROWS + 1})`],  // 売上
    [`=SUM(I2:I${CONSTANTS.DATA_ROWS + 1})`]   // 利益
  ];
  sheet.getRange(totalRow, 6, 1, 4).setFormulas([totalFormulas.map(f => f[0])]);
  sheet.getRange(totalRow, 9).setNumberFormat('¥#,##0');
  
  // 月末請求シートの利益を参照するエリア（K列）- バッチ処理で最適化
  const endOfMonthSheetName = sheetName + '（月末請求分）';
  const kColumnValues = [
    ['月末請求分'],
    ['経費合計'],
    ['売上合計'],
    ['利益合計']
  ];
  sheet.getRange(1, 11, 4, 1).setValues(kColumnValues);
  const kColumnFormulas = [
    [`=IFERROR(INDIRECT("'${endOfMonthSheetName}'!E${CONSTANTS.TOTAL_ROW}"),0)`],
    [`=IFERROR(INDIRECT("'${endOfMonthSheetName}'!F${CONSTANTS.TOTAL_ROW}"),0)`],
    [`=IFERROR(INDIRECT("'${endOfMonthSheetName}'!G${CONSTANTS.TOTAL_ROW}"),0)`]
  ];
  sheet.getRange(2, 12, 3, 1).setFormulas(kColumnFormulas);

  // 当月全体の合計（L列）- バッチ処理で最適化
  const monthlyTotalValues = [
    ['当月全体'],
    ['経費合計'],
    ['売上合計'],
    ['利益合計']
  ];
  sheet.getRange(6, 11, 4, 1).setValues(monthlyTotalValues);
  // 当月全体の合計数式（L列、7-9行目）: 月別シートの合計 + 月末請求分の合計
  const monthlyTotalFormulas = [
    [`=F${CONSTANTS.TOTAL_ROW}+L2`], // 経費合計 = 月別シートの経費合計 + 月末請求分の経費合計（L列の値）
    [`=H${CONSTANTS.TOTAL_ROW}+L3`], // 売上合計 = 月別シートの売上合計 + 月末請求分の売上合計（L列の値）
    [`=I${CONSTANTS.TOTAL_ROW}+L4`]  // 利益合計 = 月別シートの利益合計 + 月末請求分の利益合計（L列の値）
  ];
  sheet.getRange(7, 12, 3, 1).setFormulas(monthlyTotalFormulas);
  
  // 受注日列のデータ検証（カレンダーから選択可能）
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, CONSTANTS.DATA_ROWS, 1).setDataValidation(dateRule);

  // 納期列のデータ検証（カレンダーから選択可能）
  sheet.getRange(2, 3, CONSTANTS.DATA_ROWS, 1).setDataValidation(dateRule);

  // ステータスのデータ検証を設定
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONSTANTS.STATUS_OPTIONS)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 10, CONSTANTS.DATA_ROWS, 1).setDataValidation(statusRule);

  // クライアント列のデータ検証を設定（クライアントシートから参照）
  const clientSheet = ss.getSheetByName('クライアント');
  if (clientSheet) {
    const clientRange = clientSheet.getRange(`A2:A${CONSTANTS.MAX_CLIENT_ROWS}`);
    const clientValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(clientRange)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 4, CONSTANTS.DATA_ROWS, 1).setDataValidation(clientValidation);
  }

  // 条件付き書式を設定（一度に設定してパフォーマンス最適化）
  setConditionalFormatting(sheet, 10, 4); // ステータス列（10列目）、クライアント列（4列目）

  // 納期列の条件付き書式を設定（当日: 赤、1日前: 黄色）
  setDeliveryDateConditionalFormatting(sheet, 3); // 納期列（3列目）
  
  // フォーマットを適用
  formatMonthlySheet(sheet);
  
  // 行の高さを調整（広く余裕を持たせる）- パフォーマンス最適化
  // SheetのsetRowHeightメソッドを使用（RangeではなくSheetから呼び出す）
  sheet.setRowHeight(1, CONSTANTS.HEADER_ROW_HEIGHT); // ヘッダー行
  // データ行の高さを一括設定
  for (let i = 2; i <= CONSTANTS.DATA_ROWS + 1; i++) {
    sheet.setRowHeight(i, CONSTANTS.DATA_ROW_HEIGHT);
  }
  sheet.setRowHeight(CONSTANTS.TOTAL_ROW, CONSTANTS.TOTAL_ROW_HEIGHT); // 合計行
}

/**
 * クライアントごとに異なる色を設定する条件付き書式
 * @param {Sheet} sheet シートオブジェクト
 * @param {number} clientColumn クライアント列の番号
 */
function setClientConditionalFormattingByClient(sheet, clientColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName('クライアント');
  
  if (!clientSheet) {
    return;
  }
  
  const rules = [];
  const columnLetter = String.fromCharCode(64 + clientColumn);
  
  // クライアントシートからクライアント名を取得
  const clientData = clientSheet.getRange(2, 1, CONSTANTS.MAX_CLIENT_ROWS, 1).getValues();
  const clients = clientData.filter(row => row[0] && row[0].toString().trim() !== '');
  
  // クライアントごとに異なる色を割り当てる（色のパレット）
  const clientColors = [
    '#E3F2FD', '#F3E5F5', '#E8F5E9', '#FFF3E0', '#FCE4EC',
    '#E1F5FE', '#F1F8E9', '#FFF8E1', '#EFEBE9', '#E0F2F1',
    '#E8EAF6', '#F9FBE7', '#FFFDE7', '#FBE9E7', '#E8F0FE',
    '#F5F5F5', '#FFF9C4', '#FFE0B2', '#C8E6C9', '#BBDEFB'
  ];
  
  // 各クライアントごとに条件付き書式ルールを作成
  clients.forEach((clientRow, index) => {
    const clientName = clientRow[0].toString().trim();
    const color = clientColors[index % clientColors.length];
    
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=${columnLetter}2="${clientName}"`)
      .setBackground(color)
      .setRanges([sheet.getRange(2, clientColumn, CONSTANTS.DATA_ROWS, 1)])
      .build();
    rules.push(rule);
  });
  
  // すべてのルールを一度に設定（パフォーマンス最適化）
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
  }
}

/**
 * ステータス列とクライアント列の条件付き書式を一度に設定（パフォーマンス最適化）
 * @param {Sheet} sheet シートオブジェクト
 * @param {number} statusColumn ステータス列の番号
 * @param {number} clientColumn クライアント列の番号
 */
function setConditionalFormatting(sheet, statusColumn, clientColumn) {
  const rules = [];
  const statusColumnLetter = String.fromCharCode(64 + statusColumn);
  const clientColumnLetter = String.fromCharCode(64 + clientColumn);
  
  // ステータス列の条件付き書式ルールを作成
  CONSTANTS.STATUS_OPTIONS.forEach(status => {
    const color = CONSTANTS.STATUS_COLORS[status];
    if (color) {
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${statusColumnLetter}2="${status}"`)
        .setBackground(color)
        .setRanges([sheet.getRange(2, statusColumn, CONSTANTS.DATA_ROWS, 1)])
        .build();
      rules.push(rule);
    }
  });
  
  // クライアント列の条件付き書式ルールを作成（クライアントごとに異なる色）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName('クライアント');
  
  if (clientSheet) {
    const clientData = clientSheet.getRange(2, 1, CONSTANTS.MAX_CLIENT_ROWS, 1).getValues();
    const clients = clientData.filter(row => row[0] && row[0].toString().trim() !== '');
    
    // クライアントごとに異なる色を割り当てる（色のパレット）
    const clientColors = [
      '#E3F2FD', '#F3E5F5', '#E8F5E9', '#FFF3E0', '#FCE4EC',
      '#E1F5FE', '#F1F8E9', '#FFF8E1', '#EFEBE9', '#E0F2F1',
      '#E8EAF6', '#F9FBE7', '#FFFDE7', '#FBE9E7', '#E8F0FE',
      '#F5F5F5', '#FFF9C4', '#FFE0B2', '#C8E6C9', '#BBDEFB'
    ];
    
    // 各クライアントごとに条件付き書式ルールを作成
    clients.forEach((clientRow, index) => {
      const clientName = clientRow[0].toString().trim();
      const color = clientColors[index % clientColors.length];
      
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${clientColumnLetter}2="${clientName}"`)
        .setBackground(color)
        .setRanges([sheet.getRange(2, clientColumn, CONSTANTS.DATA_ROWS, 1)])
        .build();
      rules.push(rule);
    });
  }
  
  // すべてのルールを一度に設定（パフォーマンス最適化）
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
  }
}

/**
 * ステータス列に条件付き書式を設定
 * @param {Sheet} sheet シートオブジェクト
 * @param {number} statusColumn ステータス列の番号（月別シート: 8列目、月末請求シート: 7列目）
 */
function setStatusConditionalFormatting(sheet, statusColumn) {
  const rules = [];
  const columnLetter = String.fromCharCode(64 + statusColumn); // A=65, B=66, ...
  
  // 各ステータスごとに条件付き書式ルールを作成
  CONSTANTS.STATUS_OPTIONS.forEach(status => {
    const color = CONSTANTS.STATUS_COLORS[status];
    if (color) {
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${columnLetter}2="${status}"`)
        .setBackground(color)
        .setRanges([sheet.getRange(2, statusColumn, CONSTANTS.DATA_ROWS, 1)])
        .build();
      rules.push(rule);
    }
  });
  
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
  }
}

/**
 * クライアント列に条件付き書式を設定（値が入力されている場合に色を付ける）
 * @param {Sheet} sheet シートオブジェクト
 * @param {number} clientColumn クライアント列の番号（月別シート: 2列目、月末請求シート: 3列目）
 */
function setClientConditionalFormatting(sheet, clientColumn) {
  const columnLetter = String.fromCharCode(64 + clientColumn); // A=65, B=66, ...
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=${columnLetter}2<>""`)
    .setBackground(CONSTANTS.CLIENT_COLOR)
    .setRanges([sheet.getRange(2, clientColumn, CONSTANTS.DATA_ROWS, 1)])
    .build();
  
  sheet.setConditionalFormatRules([rule]);
}

/**
 * 納期列に条件付き書式を設定（当日: 赤、1日前: 黄色）
 * @param {Sheet} sheet シートオブジェクト
 * @param {number} deliveryColumn 納期列の番号
 */
function setDeliveryDateConditionalFormatting(sheet, deliveryColumn) {
  const rules = [];
  const columnLetter = String.fromCharCode(64 + deliveryColumn); // A=65, B=66, ...

  // 納期が当日の場合: 赤色背景
  const todayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=${columnLetter}2=TODAY()`)
    .setBackground('#FF0000')
    .setFontColor('#FFFFFF')
    .setRanges([sheet.getRange(2, deliveryColumn, CONSTANTS.DATA_ROWS, 1)])
    .build();
  rules.push(todayRule);

  // 納期が1日前の場合: 黄色背景
  const tomorrowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=${columnLetter}2=TODAY()+1`)
    .setBackground('#FFFF00')
    .setRanges([sheet.getRange(2, deliveryColumn, CONSTANTS.DATA_ROWS, 1)])
    .build();
  rules.push(tomorrowRule);

  // 既存のルールを取得して追加
  const existingRules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules(existingRules.concat(rules));
}

/**
 * 既存の月末請求シートに納期列を追加（受注日の後に挿入）
 * @param {string} sheetName シート名
 */
function addDeliveryColumnToEndOfMonthSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }

    // 既に納期列が存在するか確認（3列目のヘッダーが「納期」かどうか）
    const thirdColumnHeader = sheet.getRange(1, 3).getValue();
    if (thirdColumnHeader === '納期') {
      writeLog('INFO', 'addDeliveryColumnToEndOfMonthSheet', `シート「${sheetName}」には既に納期列が存在します`);
      return;
    }

    // ステップ1: 既存のデータ検証と条件付き書式をすべてクリア
    sheet.clearDataValidations();
    sheet.clearConditionalFormatRules();

    // ステップ2: C列（クライアントの位置）の前に新しい列を挿入
    sheet.insertColumnBefore(3);

    // ステップ3: ヘッダーを設定
    sheet.getRange(1, 3).setValue('納期');

    // ステップ4: 利益列の数式を更新（G列、旧数式はE-Dだったが、F-Eに更新）
    const profitFormulas = [];
    for (let i = 2; i <= CONSTANTS.DATA_ROWS + 1; i++) {
      profitFormulas.push([`=IF(F${i}="","",F${i}-E${i})`]);
    }
    sheet.getRange(2, 7, CONSTANTS.DATA_ROWS, 1).setFormulas(profitFormulas);

    // ステップ5: データ検証を新しい列番号で設定
    const dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();

    // 受注日列（B列）
    sheet.getRange(2, 2, CONSTANTS.DATA_ROWS, 1).setDataValidation(dateRule);

    // 納期列（C列）
    sheet.getRange(2, 3, CONSTANTS.DATA_ROWS, 1).setDataValidation(dateRule);

    // クライアント列（D列）
    const clientSheet = ss.getSheetByName('クライアント');
    if (clientSheet) {
      try {
        const clientRange = clientSheet.getRange(`A2:A${CONSTANTS.MAX_CLIENT_ROWS}`);
        const clientValidation = SpreadsheetApp.newDataValidation()
          .requireValueInRange(clientRange)
          .setAllowInvalid(false)
          .build();
        sheet.getRange(2, 4, CONSTANTS.DATA_ROWS, 1).setDataValidation(clientValidation);
      } catch (error) {
        writeLog('WARNING', 'addDeliveryColumnToEndOfMonthSheet', `クライアント列のデータ検証設定に失敗: ${error.toString()}`);
      }
    }

    // ステータス列（H列）
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONSTANTS.STATUS_OPTIONS)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 8, CONSTANTS.DATA_ROWS, 1).setDataValidation(statusRule);

    // ステップ6: 条件付き書式を新しい列番号で設定
    setConditionalFormatting(sheet, 8, 4); // ステータス列（8列目）、クライアント列（4列目）
    setDeliveryDateConditionalFormatting(sheet, 3); // 納期列（3列目）

    // ステップ7: フォーマットを適用
    formatEndOfMonthSheet(sheet);

    // ステップ8: 合計行の数式を更新
    const totalRow = CONSTANTS.TOTAL_ROW;
    const totalFormulas = [
      [`=SUM(E2:E${CONSTANTS.DATA_ROWS + 1})`],
      [`=SUM(F2:F${CONSTANTS.DATA_ROWS + 1})`],
      [`=SUM(G2:G${CONSTANTS.DATA_ROWS + 1})`]
    ];
    sheet.getRange(totalRow, 5, 1, 3).setFormulas([totalFormulas.map(f => f[0])]);

    // ステップ9: サマリーエリアの数式を更新
    const summaryFormulas = [
      [`=G${CONSTANTS.TOTAL_ROW}`],
      [`=E${CONSTANTS.TOTAL_ROW}`],
      [`=G${CONSTANTS.TOTAL_ROW}`]
    ];
    sheet.getRange(2, 11, 3, 1).setFormulas(summaryFormulas);

    // ステップ10: クライアント別売上一覧の数式を更新
    const clientListRow = 8;
    sheet.getRange(clientListRow, 10).setFormula(`=UNIQUE(FILTER(D2:D${CONSTANTS.DATA_ROWS + 1},D2:D${CONSTANTS.DATA_ROWS + 1}<>""))`);

    const clientSumFormulas = [];
    for (let i = 0; i < 20; i++) {
      const row = clientListRow + i;
      clientSumFormulas.push([`=IF(J${row}="","",SUMIF(D2:D${CONSTANTS.DATA_ROWS + 1},J${row},F2:F${CONSTANTS.DATA_ROWS + 1}))`]);
    }
    sheet.getRange(clientListRow, 11, 20, 1).setFormulas(clientSumFormulas);

    writeLog('INFO', 'addDeliveryColumnToEndOfMonthSheet', `シート「${sheetName}」に納期列を追加しました`);
  } catch (error) {
    writeLog('ERROR', 'addDeliveryColumnToEndOfMonthSheet', `エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 既存の月別シートに受注日と納期列を追加（項番の後に挿入）
 * @param {string} sheetName シート名
 */
function addOrderAndDeliveryColumnsToMonthlySheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }

    // 既に受注日列が存在するか確認（2列目のヘッダーが「受注日」かどうか）
    const secondColumnHeader = sheet.getRange(1, 2).getValue();
    if (secondColumnHeader === '受注日') {
      writeLog('INFO', 'addOrderAndDeliveryColumnsToMonthlySheet', `シート「${sheetName}」には既に受注日と納期列が存在します`);
      return;
    }

    // ステップ1: 既存のデータ検証と条件付き書式をすべてクリア
    sheet.clearDataValidations();
    sheet.clearConditionalFormatRules();

    // ステップ2: B列（クライアントの位置）の前に2列を挿入
    sheet.insertColumnsBefore(2, 2);

    // ステップ3: ヘッダーを設定
    sheet.getRange(1, 2).setValue('受注日');
    sheet.getRange(1, 3).setValue('納期');

    // ステップ4: 利益列の数式を更新（I列、旧数式はF-D-IF(E...)だったが、H-F-IF(G...)に更新）
    const profitFormulas = [];
    for (let i = 2; i <= CONSTANTS.DATA_ROWS + 1; i++) {
      profitFormulas.push([`=IF(H${i}="","",H${i}-F${i}-IF(G${i}="",0,G${i}*getExchangeRate()))`]);
    }
    sheet.getRange(2, 9, CONSTANTS.DATA_ROWS, 1).setFormulas(profitFormulas);

    // ステップ5: データ検証を新しい列番号で設定
    const dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();

    // 受注日列（B列）
    sheet.getRange(2, 2, CONSTANTS.DATA_ROWS, 1).setDataValidation(dateRule);

    // 納期列（C列）
    sheet.getRange(2, 3, CONSTANTS.DATA_ROWS, 1).setDataValidation(dateRule);

    // クライアント列（D列）
    const clientSheet = ss.getSheetByName('クライアント');
    if (clientSheet) {
      try {
        const clientRange = clientSheet.getRange(`A2:A${CONSTANTS.MAX_CLIENT_ROWS}`);
        const clientValidation = SpreadsheetApp.newDataValidation()
          .requireValueInRange(clientRange)
          .setAllowInvalid(false)
          .build();
        sheet.getRange(2, 4, CONSTANTS.DATA_ROWS, 1).setDataValidation(clientValidation);
      } catch (error) {
        writeLog('WARNING', 'addOrderAndDeliveryColumnsToMonthlySheet', `クライアント列のデータ検証設定に失敗: ${error.toString()}`);
      }
    }

    // ステータス列（J列）
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONSTANTS.STATUS_OPTIONS)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 10, CONSTANTS.DATA_ROWS, 1).setDataValidation(statusRule);

    // ステップ6: 条件付き書式を新しい列番号で設定
    setConditionalFormatting(sheet, 10, 4); // ステータス列（10列目）、クライアント列（4列目）
    setDeliveryDateConditionalFormatting(sheet, 3); // 納期列（3列目）

    // ステップ7: フォーマットを適用
    formatMonthlySheet(sheet);

    // ステップ8: 合計行の数式を更新
    const totalRow = CONSTANTS.TOTAL_ROW;
    const totalFormulas = [
      [`=SUM(F2:F${CONSTANTS.DATA_ROWS + 1})`],  // 経費（円）
      [`=SUM(G2:G${CONSTANTS.DATA_ROWS + 1})*getExchangeRate()`],  // 経費（ドル）を円換算して合計
      [`=SUM(H2:H${CONSTANTS.DATA_ROWS + 1})`],  // 売上
      [`=SUM(I2:I${CONSTANTS.DATA_ROWS + 1})`]   // 利益
    ];
    sheet.getRange(totalRow, 6, 1, 4).setFormulas([totalFormulas.map(f => f[0])]);

    // ステップ9: 月末請求シートを参照するエリアの数式を更新
    const endOfMonthSheetName = sheetName + '（月末請求分）';
    const kColumnFormulas = [
      [`=IFERROR(INDIRECT("'${endOfMonthSheetName}'!E${CONSTANTS.TOTAL_ROW}"),0)`],
      [`=IFERROR(INDIRECT("'${endOfMonthSheetName}'!F${CONSTANTS.TOTAL_ROW}"),0)`],
      [`=IFERROR(INDIRECT("'${endOfMonthSheetName}'!G${CONSTANTS.TOTAL_ROW}"),0)`]
    ];
    sheet.getRange(2, 12, 3, 1).setFormulas(kColumnFormulas);

    // ステップ10: 当月全体の合計数式を更新
    const monthlyTotalFormulas = [
      [`=F${CONSTANTS.TOTAL_ROW}+L2`], // 経費合計
      [`=H${CONSTANTS.TOTAL_ROW}+L3`], // 売上合計
      [`=I${CONSTANTS.TOTAL_ROW}+L4`]  // 利益合計
    ];
    sheet.getRange(7, 12, 3, 1).setFormulas(monthlyTotalFormulas);

    writeLog('INFO', 'addOrderAndDeliveryColumnsToMonthlySheet', `シート「${sheetName}」に受注日と納期列を追加しました`);
  } catch (error) {
    writeLog('ERROR', 'addOrderAndDeliveryColumnsToMonthlySheet', `エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * すべての既存シートに新しい列を追加する
 */
function addColumnsToAllExistingSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let updatedSheets = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();

      // システムシート（ログ、クライアント、年度別シート）はスキップ
      if (sheetName === 'ログ' || sheetName === 'クライアント' || sheetName.includes('年度')) {
        return;
      }

      try {
        if (sheetName.endsWith('（月末請求分）')) {
          addDeliveryColumnToEndOfMonthSheet(sheetName);
          updatedSheets.push(sheetName);
        } else if (sheetName.includes('月')) {
          addOrderAndDeliveryColumnsToMonthlySheet(sheetName);
          updatedSheets.push(sheetName);
        }
      } catch (error) {
        writeLog('WARNING', 'addColumnsToAllExistingSheets', `シート「${sheetName}」の更新に失敗しました: ${error.toString()}`);
      }
    });

    SpreadsheetApp.getUi().alert(`${updatedSheets.length}個のシートを更新しました:\n${updatedSheets.join('\n')}`);
    writeLog('INFO', 'addColumnsToAllExistingSheets', `${updatedSheets.length}個のシートを更新しました`);
  } catch (error) {
    writeLog('ERROR', 'addColumnsToAllExistingSheets', `エラー: ${error.toString()}`);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${error.toString()}`);
  }
}

/**
 * シート名から月名を抽出（末尾の「（月末請求分）」を削除）
 * @param {string} sheetName シート名
 * @returns {string} 月名（例: "11月"）
 */
function extractMonthName(sheetName) {
  if (!sheetName.endsWith('（月末請求分）')) {
    throw new Error(`Invalid sheet name format: ${sheetName}`);
  }
  return sheetName.replace(/（月末請求分）$/, '');
}

/**
 * 月末請求シートを作成（完全再構築版）
 * @param {string} sheetName シート名（例: "11月（月末請求分）"）
 */
function createEndOfMonthSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    // ステップ1: シートの作成またはクリア
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet(sheetName);
      // 対応する月別シートの後に配置（失敗しても続行）
      try {
        const monthName = extractMonthName(sheetName);
        const monthSheet = ss.getSheetByName(monthName);
        if (monthSheet) {
          const monthIndex = monthSheet.getIndex();
          ss.setActiveSheet(sheet);
          ss.moveActiveSheet(monthIndex + 1);
        }
      } catch (error) {
        writeLog('WARNING', 'createEndOfMonthSheet', `シートの移動に失敗しました: ${error.toString()}`);
      }
    }
    
    // ステップ2: ヘッダー行を設定
    const headers = ['項番', '受注日', '納期', 'クライアント', '経費', '売上', '利益', 'ステータス', '備考'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ステップ3: データ行を作成（バッチ処理）
    const dataRows = [];
    for (let i = 1; i <= CONSTANTS.DATA_ROWS; i++) {
      dataRows.push([i, '', '', '', '', '', '', '', '']);
    }
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);

    // ステップ4: 利益列の計算式を設定（バッチ処理）
    const profitFormulas = [];
    for (let i = 2; i <= CONSTANTS.DATA_ROWS + 1; i++) {
      profitFormulas.push([`=IF(F${i}="","",F${i}-E${i})`]);
    }
    sheet.getRange(2, 7, CONSTANTS.DATA_ROWS, 1).setFormulas(profitFormulas);
    
    // ステップ5: 合計行を設定
    const totalRow = CONSTANTS.TOTAL_ROW;
    sheet.getRange(totalRow, 1).setValue('合計');
    const totalFormulas = [
      [`=SUM(E2:E${CONSTANTS.DATA_ROWS + 1})`],
      [`=SUM(F2:F${CONSTANTS.DATA_ROWS + 1})`],
      [`=SUM(G2:G${CONSTANTS.DATA_ROWS + 1})`]
    ];
    sheet.getRange(totalRow, 5, 1, 3).setFormulas([totalFormulas.map(f => f[0])]);
    
    // ステップ6: 別エリアに合計を表示（J列）
    const summaryValues = [
      ['当月の合計'],
      ['売上'],
      ['経費'],
      ['利益']
    ];
    sheet.getRange(1, 10, 4, 1).setValues(summaryValues);
    const summaryFormulas = [
      [`=G${CONSTANTS.TOTAL_ROW}`],
      [`=E${CONSTANTS.TOTAL_ROW}`],
      [`=G${CONSTANTS.TOTAL_ROW}`]
    ];
    sheet.getRange(2, 11, 3, 1).setFormulas(summaryFormulas);
    
    // ステップ7: クライアント別売上一覧を作成
    const clientListHeaders = [
      ['クライアント別売上一覧', ''],  // 2列にする
      ['クライアント', '売上']
    ];
    sheet.getRange(6, 10, 2, 2).setValues(clientListHeaders);

    const clientListRow = 8;
    sheet.getRange(clientListRow, 10).setFormula(`=UNIQUE(FILTER(D2:D${CONSTANTS.DATA_ROWS + 1},D2:D${CONSTANTS.DATA_ROWS + 1}<>""))`);

    const clientSumFormulas = [];
    for (let i = 0; i < 20; i++) {
      const row = clientListRow + i;
      clientSumFormulas.push([`=IF(J${row}="","",SUMIF(D2:D${CONSTANTS.DATA_ROWS + 1},J${row},F2:F${CONSTANTS.DATA_ROWS + 1}))`]);
    }
    sheet.getRange(clientListRow, 11, 20, 1).setFormulas(clientSumFormulas);
    
    // ステップ8: データ検証を設定
    // 受注日列のデータ検証（カレンダーから選択可能）
    const dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 2, CONSTANTS.DATA_ROWS, 1).setDataValidation(dateRule);

    // 納期列のデータ検証（カレンダーから選択可能）
    sheet.getRange(2, 3, CONSTANTS.DATA_ROWS, 1).setDataValidation(dateRule);

    // ステータス列のデータ検証
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONSTANTS.STATUS_OPTIONS)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 8, CONSTANTS.DATA_ROWS, 1).setDataValidation(statusRule);

    // クライアント列のデータ検証
    const clientSheet = ss.getSheetByName('クライアント');
    if (clientSheet) {
      try {
        const clientRange = clientSheet.getRange(`A2:A${CONSTANTS.MAX_CLIENT_ROWS}`);
        const clientValidation = SpreadsheetApp.newDataValidation()
          .requireValueInRange(clientRange)
          .setAllowInvalid(false)
          .build();
        sheet.getRange(2, 4, CONSTANTS.DATA_ROWS, 1).setDataValidation(clientValidation);
      } catch (error) {
        writeLog('WARNING', 'createEndOfMonthSheet', `クライアント列のデータ検証設定に失敗しました: ${error.toString()}`);
      }
    }

    // 条件付き書式を設定（一度に設定してパフォーマンス最適化）
    setConditionalFormatting(sheet, 8, 4); // ステータス列（8列目）、クライアント列（4列目）

    // 納期列の条件付き書式を設定（当日: 赤、1日前: 黄色）
    setDeliveryDateConditionalFormatting(sheet, 3); // 納期列（3列目）
    
    // ステップ9: フォーマットを適用（最後に実行）
    formatEndOfMonthSheet(sheet);
    
    // ステップ10: 行の高さを調整（フォーマット後に実行、確実に完了させる）
    sheet.setRowHeight(1, CONSTANTS.HEADER_ROW_HEIGHT);
    // データ行の高さを設定（パフォーマンス最適化：ループを最適化）
    for (let i = 2; i <= CONSTANTS.DATA_ROWS + 1; i++) {
      sheet.setRowHeight(i, CONSTANTS.DATA_ROW_HEIGHT);
    }
    sheet.setRowHeight(CONSTANTS.TOTAL_ROW, CONSTANTS.TOTAL_ROW_HEIGHT);
    
    writeLog('INFO', 'createEndOfMonthSheet', `シート作成完了: ${sheetName}`);
  } catch (error) {
    writeLog('ERROR', 'createEndOfMonthSheet', `シート作成エラー (${sheetName}): ${error.toString()}`);
    throw error;
  }
}

/**
 * 年度別シートのフォーマット設定（モダンなUI）
 */
function formatYearlySheet(sheet) {
  // ヘッダー行のフォーマット（年度別シート統一色：青）- パフォーマンス最適化
  const headerColor = '#4285F4'; // 青
  
  const headerRange = sheet.getRange(1, 1, 1, 5);
  // 複数のプロパティを一度に設定（API呼び出しを最小化）
  headerRange.setBackground(headerColor);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setBorder(true, true, true, true, false, false);
  
  // データ行のフォーマット - バッチ処理
  const dataRange = sheet.getRange(2, 1, CONSTANTS.YEARLY_SHEET_MONTHS, 5);
  dataRange.setBorder(true, true, true, true, false, false);
  
  // 交互の行に背景色を設定（バッチ処理で最適化）
  const evenRowRange = sheet.getRange(2, 1, 6, 5);
  const oddRowRange = sheet.getRange(3, 1, 6, 5);
  evenRowRange.setBackground('#f7f9fa');
  oddRowRange.setBackground('#ffffff');
  
  // 数値列のフォーマット（バッチ処理）- 年度別シートは数式参照のため除外
  // 年度別シートの数値列は数式で参照されているため、数値形式の設定をスキップ
  
  // 合計行のフォーマット（バッチ処理）
  const totalRow = CONSTANTS.YEARLY_SHEET_MONTHS + 2;
  const totalRange = sheet.getRange(totalRow, 1, 1, 5);
  totalRange.setBackground('#e8f5e9');
  totalRange.setFontWeight('bold');
  totalRange.setFontSize(11);
  totalRange.setBorder(true, true, true, true, false, false);
  sheet.getRange(totalRow, 2, 1, 3).setNumberFormat('¥#,##0');
  
  // 列幅の調整 - パフォーマンス最適化（必要な列のみ設定）
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 200);
}

/**
 * 月別シートのフォーマット設定（モダンなUI）
 */
function formatMonthlySheet(sheet) {
  // ヘッダー行のフォーマット（月別シート統一色：緑）- パフォーマンス最適化
  const headerColor = '#34A853'; // 緑

  const headerRange = sheet.getRange(1, 1, 1, 11);
  // 複数のプロパティを一度に設定（API呼び出しを最小化）
  headerRange.setBackground(headerColor);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setBorder(true, true, true, true, false, false);

  // データ行のフォーマット - バッチ処理
  const dataRange = sheet.getRange(2, 1, CONSTANTS.DATA_ROWS, 11);
  dataRange.setBorder(true, true, true, true, false, false);

  // 交互の行に背景色を設定（バッチ処理で最適化）
  const evenRowRange = sheet.getRange(2, 1, 33, 11);
  const oddRowRange = sheet.getRange(3, 1, 32, 11);
  evenRowRange.setBackground('#f7f9fa');
  oddRowRange.setBackground('#ffffff');

  // 日付列のフォーマット
  sheet.getRange(2, 2, CONSTANTS.DATA_ROWS, 1).setNumberFormat('mm/dd'); // 受注日（月日のみ表示）
  sheet.getRange(2, 3, CONSTANTS.DATA_ROWS, 1).setNumberFormat('mm/dd'); // 納期（月日のみ表示）

  // 数値列のフォーマット（バッチ処理で最適化）
  const numberFormatRange = sheet.getRange(2, 6, CONSTANTS.DATA_ROWS, 4);
  const numberFormats = [];
  for (let i = 0; i < CONSTANTS.DATA_ROWS; i++) {
    numberFormats.push(['¥#,##0', '$#,##0.00', '¥#,##0', '¥#,##0']); // 経費、経費（ドル）、売上、利益
  }
  numberFormatRange.setNumberFormats(numberFormats);

  // 合計行の数値フォーマット
  sheet.getRange(CONSTANTS.TOTAL_ROW, 6, 1, 4).setNumberFormat('¥#,##0');
  
  // 月末請求分のエリアのフォーマット（モダンなスタイル）- パフォーマンス最適化
  // 範囲をまとめて処理（API呼び出しを最小化）
  const summaryArea = sheet.getRange(1, 11, 9, 2);
  summaryArea.setFontSize(10);
  
  // ヘッダー部分のフォーマット（バッチ処理）
  const headerK11 = sheet.getRange(1, 11, 1, 2);
  headerK11.setFontWeight('bold');
  headerK11.setBackground('#fff3e0');
  
  // ラベル部分のフォーマット
  sheet.getRange(2, 11, 3, 1).setFontWeight('bold');
  sheet.getRange(7, 11, 3, 1).setFontWeight('bold');
  sheet.getRange(7, 11).setBackground('#fff3e0');
  
  // 数値フォーマット（まとめて設定）
  sheet.getRange(2, 12, 3, 1).setNumberFormat('¥#,##0');
  sheet.getRange(7, 12, 3, 1).setNumberFormat('¥#,##0');
  
  // 列幅の調整 - パフォーマンス最適化（必要な列のみ設定）
  // 0以外の列幅のみを設定（API呼び出しを最小化）
  const columnWidths = [40, 70, 70, 150, 100, 100, 120, 100, 100, 120, 800, 0, 120, 120];
  for (let i = 0; i < columnWidths.length; i++) {
    if (columnWidths[i] > 0) {
      sheet.setColumnWidth(i + 1, columnWidths[i]);
    }
  }
}

/**
 * 月末請求シートのフォーマット設定（モダンなUI）
 */
function formatEndOfMonthSheet(sheet) {
  // ヘッダー行のフォーマット（月末請求シート統一色：オレンジ）- パフォーマンス最適化
  const headerColor = '#FF9800'; // オレンジ

  const headerRange = sheet.getRange(1, 1, 1, 9);
  // 複数のプロパティを一度に設定（API呼び出しを最小化）
  headerRange.setBackground(headerColor);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setBorder(true, true, true, true, false, false);

  // データ行のフォーマット - バッチ処理
  const dataRange = sheet.getRange(2, 1, CONSTANTS.DATA_ROWS, 9);
  dataRange.setBorder(true, true, true, true, false, false);

  // 交互の行に背景色を設定（バッチ処理で最適化）
  const evenRowRange = sheet.getRange(2, 1, 33, 9);
  const oddRowRange = sheet.getRange(3, 1, 32, 9);
  evenRowRange.setBackground('#f7f9fa');
  oddRowRange.setBackground('#ffffff');

  // 数値列のフォーマット（バッチ処理）- 利益列は数式のため除外
  sheet.getRange(2, 2, CONSTANTS.DATA_ROWS, 1).setNumberFormat('mm/dd'); // 受注日（月日のみ表示）
  sheet.getRange(2, 3, CONSTANTS.DATA_ROWS, 1).setNumberFormat('mm/dd'); // 納期（月日のみ表示）
  sheet.getRange(2, 5, CONSTANTS.DATA_ROWS, 1).setNumberFormat('¥#,##0'); // 経費
  sheet.getRange(2, 6, CONSTANTS.DATA_ROWS, 1).setNumberFormat('¥#,##0'); // 売上
  // 利益列（G列）は数式のため数値形式の設定をスキップ

  // 合計行のフォーマット（バッチ処理）
  const totalRange = sheet.getRange(CONSTANTS.TOTAL_ROW, 1, 1, 9);
  totalRange.setBackground('#e8f5e9');
  totalRange.setFontWeight('bold');
  totalRange.setFontSize(11);
  totalRange.setBorder(true, true, true, true, false, false);
  sheet.getRange(CONSTANTS.TOTAL_ROW, 5, 1, 3).setNumberFormat('¥#,##0');
  
  // 別エリアのフォーマット（モダンなスタイル）- パフォーマンス最適化
  // 範囲をまとめて処理（API呼び出しを最小化）
  const summaryArea = sheet.getRange(1, 10, 9, 2);
  summaryArea.setFontSize(10);
  
  // ヘッダー部分のフォーマット（バッチ処理）
  const headerJ1 = sheet.getRange(1, 10, 1, 2);
  headerJ1.setFontWeight('bold');
  headerJ1.setBackground('#fff3e0');
  
  const headerJ6 = sheet.getRange(6, 10, 1, 2);
  headerJ6.setFontWeight('bold');
  headerJ6.setBackground('#fff3e0');
  
  const headerJ7 = sheet.getRange(7, 10, 1, 2);
  headerJ7.setFontWeight('bold');
  headerJ7.setBackground('#e8f5e9');
  
  // ラベル部分のフォーマット
  sheet.getRange(2, 10, 3, 1).setFontWeight('bold');
  
  // 数値フォーマット（まとめて設定）
  sheet.getRange(2, 11, 3, 1).setNumberFormat('¥#,##0');
  sheet.getRange(8, 11, 100, 1).setNumberFormat('¥#,##0');
  
  // 列幅の調整 - パフォーマンス最適化（必要な列のみ設定）
  const columnWidths = [40, 70, 70, 150, 100, 100, 100, 120, 800, 0, 150, 120];
  for (let i = 0; i < columnWidths.length; i++) {
    if (columnWidths[i] > 0) {
      sheet.setColumnWidth(i + 1, columnWidths[i]);
    }
  }
}

/**
 * 為替レート取得関数（USD/JPY）
 * exchangerate-api.comの無料APIを使用
 * @returns {number} USD/JPYの為替レート（エラー時はデフォルト値150を返す）
 */
function getExchangeRate() {
  try {
    // exchangerate-api.comの無料APIを使用
    const url = 'https://api.exchangerate-api.com/v4/latest/USD';
    const options = {
      'muteHttpExceptions': true,
      'timeout': CONSTANTS.EXCHANGE_RATE_TIMEOUT
    };
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      writeLog('WARNING', 'getExchangeRate', `為替レートAPIエラー: HTTP ${response.getResponseCode()}`);
      return CONSTANTS.DEFAULT_EXCHANGE_RATE;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (data && data.rates && data.rates.JPY) {
      return data.rates.JPY;
    } else {
      writeLog('WARNING', 'getExchangeRate', '為替レートデータが不正です');
      // フォールバックとして固定レートを返す
      return CONSTANTS.DEFAULT_EXCHANGE_RATE;
    }
  } catch (error) {
    writeLog('ERROR', 'getExchangeRate', `為替レート取得エラー: ${error.toString()}`);
    // エラー時は固定レートを返す
    return CONSTANTS.DEFAULT_EXCHANGE_RATE;
  }
}

/**
 * 新しい月のシートを作成（メニューから実行）
 */
function createNewMonthSheet() {
  const ui = SpreadsheetApp.getUi();
  
  // 月を選択するプロンプト
  const response = ui.prompt(
    '新しい月のシートを作成',
    '作成する月を入力してください（例: 12）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  let monthInput = response.getResponseText().trim();
  
  if (!monthInput) {
    ui.alert('月が入力されていません。');
    return;
  }
  
  // 数字のみかチェック
  if (!/^\d+$/.test(monthInput)) {
    ui.alert('数字のみを入力してください（例: 12）。');
    return;
  }
  
  const monthNum = parseInt(monthInput, 10);
  
  // 1-12の範囲かチェック
  if (monthNum < 1 || monthNum > 12) {
    ui.alert('1から12の間の数字を入力してください。');
    return;
  }
  
  // 「月」を付加
  const monthName = monthNum + '月';
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既に存在するかチェック
  if (ss.getSheetByName(monthName)) {
    ui.alert(`「${monthName}」シートは既に存在します。`);
    return;
  }
  
  // 年度別シートを検索（最初に見つかった年度別シートを使用）
  const allSheets = ss.getSheets();
  const yearlySheet = allSheets.find(sheet => sheet.getName().endsWith('年度'));
  const yearlySheetName = yearlySheet ? yearlySheet.getName() : null;
  
  try {
    // 月別シートを作成
    createMonthlySheet(monthName, yearlySheetName);
    
    // 月末請求シートを作成
    const endOfMonthSheetName = monthName + '（月末請求分）';
    if (ss.getSheetByName(endOfMonthSheetName)) {
      ui.alert(`「${endOfMonthSheetName}」シートは既に存在します。`);
      return;
    }
    createEndOfMonthSheet(endOfMonthSheetName);
    
    // 年度別シートを更新（該当する月の行があれば参照を更新）
    if (yearlySheetName) {
      updateYearlySheetReference(monthName, yearlySheetName);
    }
    
    ui.alert(`「${monthName}」と「${endOfMonthSheetName}」シートを作成しました。`);
  } catch (error) {
    writeLog('ERROR', 'createNewMonthSheet', error.toString());
    ui.alert('エラーが発生しました: ' + error.toString());
  }
}

/**
 * 年度別シートの参照を更新
 * @param {string} monthName 月名（例: "11月"）
 * @param {string} yearlySheetName 年度別シート名（例: "2026年度"）
 */
function updateYearlySheetReference(monthName, yearlySheetName = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 年度別シート名が指定されていない場合、すべての年度別シートを検索
  let yearlySheets = [];
  if (yearlySheetName) {
    const sheet = ss.getSheetByName(yearlySheetName);
    if (sheet) {
      yearlySheets = [sheet];
    }
  } else {
    // すべての年度別シートを検索
    const allSheets = ss.getSheets();
    yearlySheets = allSheets.filter(sheet => sheet.getName().endsWith('年度'));
  }
  
  if (yearlySheets.length === 0) {
    return;
  }
  
  // 各年度別シートに対して参照を更新
  yearlySheets.forEach(yearlySheet => {
    try {
      const monthRange = yearlySheet.getRange(2, 1, CONSTANTS.YEARLY_SHEET_MONTHS, 1);
      const monthValues = monthRange.getValues();
      
      for (let i = 0; i < monthValues.length; i++) {
        if (monthValues[i][0] === monthName) {
          const row = i + 2;
          // 経費、売上、利益の参照を更新（L列に変更）
          yearlySheet.getRange(row, 2).setFormula(`=IFERROR(INDIRECT("'${monthName}'!L7"),0)`);
          yearlySheet.getRange(row, 3).setFormula(`=IFERROR(INDIRECT("'${monthName}'!L8"),0)`);
          yearlySheet.getRange(row, 4).setFormula(`=IFERROR(INDIRECT("'${monthName}'!L9"),0)`);
          break;
        }
      }
    } catch (error) {
      writeLog('WARNING', 'updateYearlySheetReference', `エラー (${yearlySheet.getName()}): ${error.toString()}`);
    }
  });
}

/**
 * 新しい年度のシートを作成（メニューから実行）
 */
function createNewYearSheet() {
  const ui = SpreadsheetApp.getUi();
  
  // 年度を選択するプロンプト
  const response = ui.prompt(
    '新しい年度のシートを作成',
    '作成する年度を入力してください（例: 2027）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  let yearInput = response.getResponseText().trim();
  
  if (!yearInput) {
    ui.alert('年度が入力されていません。');
    return;
  }
  
  // 数字のみかチェック
  if (!/^\d+$/.test(yearInput)) {
    ui.alert('数字のみを入力してください（例: 2027）。');
    return;
  }
  
  // 「年度」を付加
  const yearName = yearInput + '年度';
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既に存在するかチェック
  if (ss.getSheetByName(yearName)) {
    ui.alert(`「${yearName}」シートは既に存在します。`);
    return;
  }
  
  // 年度別シートを作成（開始月はデフォルトで11月）
  try {
    createYearlySheet(yearName, CONSTANTS.DEFAULT_START_MONTH);
    ui.alert(`「${yearName}」シートを作成しました。`);
  } catch (error) {
    writeLog('ERROR', 'createNewYearSheet', error.toString());
    ui.alert('エラーが発生しました: ' + error.toString());
  }
}

/**
 * 既存の月別シートを更新（データ検証とフォーマットを反映）
 */
function updateExistingMonthlySheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return false;
  }
  
  // クライアント列のデータ検証を設定（クライアントシートから参照）
  const clientSheet = ss.getSheetByName('クライアント');
  if (clientSheet) {
    try {
      const clientRange = clientSheet.getRange(`A2:A${CONSTANTS.MAX_CLIENT_ROWS}`);
      const clientValidation = SpreadsheetApp.newDataValidation()
        .requireValueInRange(clientRange)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, 2, CONSTANTS.DATA_ROWS, 1).setDataValidation(clientValidation);
    } catch (error) {
      writeLog('WARNING', 'updateExistingMonthlySheet', `クライアント検証エラー: ${error.toString()}`);
    }
  }
  
  // ステータス列のデータ検証を設定
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONSTANTS.STATUS_OPTIONS)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 8, CONSTANTS.DATA_ROWS, 1).setDataValidation(statusRule);
  
  // 合計行の計算式を更新（E列の経費（ドル）を円換算）
  const totalRow = CONSTANTS.TOTAL_ROW;
  const totalFormulas = [
    [`=SUM(D2:D${CONSTANTS.DATA_ROWS + 1})`],  // 経費（円）
    [`=SUM(E2:E${CONSTANTS.DATA_ROWS + 1})*getExchangeRate()`],  // 経費（ドル）を円換算して合計
    [`=SUM(F2:F${CONSTANTS.DATA_ROWS + 1})`],  // 売上
    [`=SUM(G2:G${CONSTANTS.DATA_ROWS + 1})`]   // 利益
  ];
  sheet.getRange(totalRow, 4, 1, 4).setFormulas([totalFormulas.map(f => f[0])]);
  sheet.getRange(2, 7, CONSTANTS.DATA_ROWS, 1).setNumberFormat('¥#,##0');
  sheet.getRange(totalRow, 7).setNumberFormat('¥#,##0');
  
  // 月末請求シートの利益を参照するエリア（K列、L列）を更新
  const endOfMonthSheetName = sheetName + '（月末請求分）';
  
  // 古いM列、N列のデータをクリア（移動前の列）
  sheet.getRange(1, 13, 10, 2).clearContent();
  
  const kColumnValues = [
    ['月末請求分'],
    ['経費合計'],
    ['売上合計'],
    ['利益合計']
  ];
  sheet.getRange(1, 11, 4, 1).setValues(kColumnValues);
  const kColumnFormulas = [
    [`=IFERROR(INDIRECT("'${endOfMonthSheetName}'!D${CONSTANTS.TOTAL_ROW}"),0)`],
    [`=IFERROR(INDIRECT("'${endOfMonthSheetName}'!E${CONSTANTS.TOTAL_ROW}"),0)`],
    [`=IFERROR(INDIRECT("'${endOfMonthSheetName}'!F${CONSTANTS.TOTAL_ROW}"),0)`]
  ];
  sheet.getRange(2, 12, 3, 1).setFormulas(kColumnFormulas);
  
  // 当月全体の合計（L列）を更新
  const monthlyTotalValues = [
    ['当月全体'],
    ['経費合計'],
    ['売上合計'],
    ['利益合計']
  ];
  sheet.getRange(6, 11, 4, 1).setValues(monthlyTotalValues);
  // 当月全体の合計数式（L列、7-9行目）を更新
  const monthlyTotalFormulas = [
    [`=D${CONSTANTS.TOTAL_ROW}+L2`], // 経費合計 = 月別シートの経費合計 + 月末請求分の経費合計（L列の値）
    [`=F${CONSTANTS.TOTAL_ROW}+L3`], // 売上合計 = 月別シートの売上合計 + 月末請求分の売上合計（L列の値）
    [`=G${CONSTANTS.TOTAL_ROW}+L4`]  // 利益合計 = 月別シートの利益合計 + 月末請求分の利益合計（L列の値）
  ];
  sheet.getRange(7, 12, 3, 1).setFormulas(monthlyTotalFormulas);
  
  // 条件付き書式を設定（一度に設定してパフォーマンス最適化）
  setConditionalFormatting(sheet, 8, 2); // ステータス列（8列目）、クライアント列（2列目）
  
  // フォーマットを適用（既存のデータは保持）
  formatMonthlySheet(sheet);
  
  return true;
}

/**
 * 既存の月末請求シートを更新（データ検証とフォーマットを反映）
 */
function updateExistingEndOfMonthSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return false;
  }
  
  // 受注日列のデータ検証（カレンダーから選択可能）
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, CONSTANTS.DATA_ROWS, 1).setDataValidation(dateRule);
  
  // クライアント列のデータ検証を設定（クライアントシートから参照）
  const clientSheet = ss.getSheetByName('クライアント');
  if (clientSheet) {
    try {
      const clientRange = clientSheet.getRange(`A2:A${CONSTANTS.MAX_CLIENT_ROWS}`);
      const clientValidation = SpreadsheetApp.newDataValidation()
        .requireValueInRange(clientRange)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, 3, CONSTANTS.DATA_ROWS, 1).setDataValidation(clientValidation); // C列（3列目）がクライアント列
    } catch (error) {
      writeLog('WARNING', 'updateExistingEndOfMonthSheet', `クライアント検証エラー: ${error.toString()}`);
    }
  }
  
  // ステータス列のデータ検証を設定
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONSTANTS.STATUS_OPTIONS)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 7, CONSTANTS.DATA_ROWS, 1).setDataValidation(statusRule);
  
  // 条件付き書式を設定（一度に設定してパフォーマンス最適化）
  setConditionalFormatting(sheet, 7, 3); // ステータス列（7列目）、クライアント列（3列目）
  
  // フォーマットを適用（既存のデータは保持）
  formatEndOfMonthSheet(sheet);
  
  return true;
}

/**
 * 既存の年度別シートを更新（フォーマットを反映）
 */
function updateExistingYearlySheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return false;
  }
  
  // 月別シートの参照を更新（N列からL列に変更）
  const monthRange = sheet.getRange(2, 1, CONSTANTS.YEARLY_SHEET_MONTHS, 1);
  const monthValues = monthRange.getValues();
  
  for (let i = 0; i < monthValues.length; i++) {
    const monthName = monthValues[i][0];
    if (monthName && typeof monthName === 'string' && monthName.endsWith('月')) {
      const row = i + 2;
      // 参照をL列に更新
      sheet.getRange(row, 2).setFormula(`=IFERROR(INDIRECT("'${monthName}'!L7"),0)`);
      sheet.getRange(row, 3).setFormula(`=IFERROR(INDIRECT("'${monthName}'!L8"),0)`);
      sheet.getRange(row, 4).setFormula(`=IFERROR(INDIRECT("'${monthName}'!L9"),0)`);
    }
  }
  
  // フォーマットを適用（既存のデータは保持）
  formatYearlySheet(sheet);
  
  return true;
}

/**
 * すべての既存シートを更新（メニューから実行）
 */
function updateAllExistingSheets() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '既存シートの更新',
    'すべての既存シートに対して、最新のフォーマットとデータ検証を適用します。\n既存のデータは保持されます。\n\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let updatedCount = 0;
  let errorCount = 0;
  
  // クライアントシートが存在するか確認
  const clientSheet = ss.getSheetByName('クライアント');
  if (!clientSheet) {
    ui.alert('クライアントシートが見つかりません。先に初期シートを作成してください。');
    return;
  }
  
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    
    // クライアントシート自体はスキップ
    if (sheetName === 'クライアント') {
      continue;
    }
    
    // 年度別シート（「年度」で終わる）
    if (sheetName.endsWith('年度')) {
      if (updateExistingYearlySheet(sheetName)) {
        updatedCount++;
      } else {
        errorCount++;
      }
    }
    // 月末請求シート（「（月末請求分）」で終わる）
    else if (sheetName.includes('（月末請求分）')) {
      if (updateExistingEndOfMonthSheet(sheetName)) {
        updatedCount++;
      } else {
        errorCount++;
      }
    }
    // 月別シート（月の名前、例：「11月」「12月」など）
    else if (/^\d+月$/.test(sheetName)) {
      if (updateExistingMonthlySheet(sheetName)) {
        updatedCount++;
      } else {
        errorCount++;
      }
    }
  }
  
  ui.alert(`更新が完了しました。\n更新: ${updatedCount}シート\nエラー: ${errorCount}シート`);
}

/**
 * クライアントシートを作成（メニューから実行）
 */
function createClientSheetFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'クライアント';
  
  // 既に存在するかチェック
  const existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) {
    const response = ui.alert(
      'クライアントシートの作成',
      `「${sheetName}」シートは既に存在します。\n既存のシートをクリアして再作成しますか？`,
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (response === ui.Button.YES) {
      existingSheet.clear();
      createClientSheet();
      ui.alert(`「${sheetName}」シートを作成しました。`);
    } else if (response === ui.Button.NO) {
      // 既存のシートをそのまま使用し、フォーマットのみ更新
      formatClientSheet(existingSheet);
      // フィルタを再適用
      const dataRange = existingSheet.getRange(1, 1, CONSTANTS.MAX_CLIENT_ROWS, 3);
      if (!existingSheet.getFilter()) {
        dataRange.createFilter();
      }
      ui.alert(`「${sheetName}」シートのフォーマットを更新しました。`);
    }
  } else {
    createClientSheet();
    ui.alert(`「${sheetName}」シートを作成しました。`);
  }
}

/**
 * 初期設定用HTMLサイドバー（最適化版）
 * @returns {string} HTMLコンテンツ
 */
function getInitializeSetupHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body{font-family:Arial,sans-serif;padding:20px;margin:0;background:#fff}
    h2{margin:0 0 20px;color:#202124;font-size:20px;font-weight:500}
    .fg{margin-bottom:20px}
    label{display:block;margin-bottom:8px;color:#5f6368;font-size:14px;font-weight:500}
    input{width:100%;padding:10px;border:1px solid #dadce0;border-radius:4px;font-size:14px;box-sizing:border-box}
    input:focus{outline:none;border-color:#4285F4;box-shadow:0 0 0 2px rgba(66,133,244,.2)}
    .ht{font-size:12px;color:#80868b;margin-top:4px}
    .bg{display:flex;justify-content:flex-end;gap:10px;margin-top:24px}
    button{padding:10px 24px;border:none;border-radius:4px;font-size:14px;font-weight:500;cursor:pointer}
    .bp{background:#4285F4;color:#fff}
    .bp:hover{background:#3367d6}
    .bp:disabled{background:#dadce0;cursor:not-allowed}
    .bs{background:#f1f3f4;color:#202124}
    .bs:hover{background:#e8eaed}
    .e{color:#ea4335;font-size:12px;margin-top:4px;display:none}
    .loading{display:none;text-align:center;padding:20px;color:#5f6368}
    .success{display:none;padding:12px;background:#e8f5e9;color:#2e7d32;border-radius:4px;margin-top:16px}
  </style>
</head>
<body>
  <h2>初期設定</h2>
  <form id="sf">
    <div class="fg">
      <label for="y">年度</label>
      <input type="text" id="y" placeholder="例: 2026" required>
      <div class="ht">年度を入力してください</div>
      <div class="e" id="ye"></div>
    </div>
    <div class="fg">
      <label for="m">開始月</label>
      <input type="text" id="m" placeholder="例: 11" required>
      <div class="ht">年度の開始月を入力してください（1-12）</div>
      <div class="e" id="me"></div>
    </div>
    <div class="bg">
      <button type="button" class="bs" onclick="cancel()">キャンセル</button>
      <button type="submit" class="bp" id="submitBtn">実行</button>
    </div>
    <div class="loading" id="loading">処理中...</div>
    <div class="success" id="success"></div>
  </form>
  <script>
    var f=document.getElementById('sf'),y=document.getElementById('y'),m=document.getElementById('m'),ye=document.getElementById('ye'),me=document.getElementById('me'),submitBtn=document.getElementById('submitBtn'),loading=document.getElementById('loading'),success=document.getElementById('success');
    
    function cancel(){
      google.script.host.close();
    }
    
    f.addEventListener('submit',function(e){
      e.preventDefault();
      var yv=y.value.trim(),mv=m.value.trim(),err=false;
      
      ye.style.display='none';
      me.style.display='none';
      
      if(!yv){
        ye.textContent='年度が入力されていません。';
        ye.style.display='block';
        err=true;
      }else if(!/^\\d+$/.test(yv)){
        ye.textContent='年度は数字のみを入力してください。';
        ye.style.display='block';
        err=true;
      }
      
      if(!mv){
        me.textContent='開始月が入力されていません。';
        me.style.display='block';
        err=true;
      }else if(!/^\\d+$/.test(mv)){
        me.textContent='開始月は数字のみを入力してください。';
        me.style.display='block';
        err=true;
      }else{
        var mn=parseInt(mv,10);
        if(mn<1||mn>12){
          me.textContent='開始月は1から12の間の数字を入力してください。';
          me.style.display='block';
          err=true;
        }
      }
      
      if(!err){
        submitBtn.disabled=true;
        loading.style.display='block';
        success.style.display='none';
        
        google.script.run
          .withSuccessHandler(function(msg){
            loading.style.display='none';
            success.textContent=msg||'初期設定が完了しました。';
            success.style.display='block';
            setTimeout(function(){
              google.script.host.close();
            },2000);
          })
          .withFailureHandler(function(err){
            loading.style.display='none';
            submitBtn.disabled=false;
            alert('エラーが発生しました: '+err.toString());
          })
          .processInitializeSetup(yv,mv);
      }
    });
  </script>
</body>
</html>
  `;
}

/**
 * 初期設定の処理（HTMLサイドバーから呼び出し）- 完全再構築版
 * @param {string} yearInput 年度
 * @param {string} monthInput 開始月
 * @returns {string} 成功メッセージ
 */
function processInitializeSetup(yearInput, monthInput) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let createdSheets = [];
  
  try {
    const yearName = yearInput + '年度';
    const startMonth = parseInt(monthInput, 10);
    const startMonthName = startMonth + '月';
    const endOfMonthSheetName = startMonthName + '（月末請求分）';
    
    // ステップ1: 既存シートの確認と削除
    const existingSheets = [];
    if (ss.getSheetByName('ログ')) existingSheets.push('ログ');
    if (ss.getSheetByName('クライアント')) existingSheets.push('クライアント');
    if (ss.getSheetByName(yearName)) existingSheets.push(yearName);
    if (ss.getSheetByName(startMonthName)) existingSheets.push(startMonthName);
    if (ss.getSheetByName(endOfMonthSheetName)) existingSheets.push(endOfMonthSheetName);
    
    if (existingSheets.length > 0) {
      const response = ui.alert(
        '初期設定',
        `以下のシートが既に存在します：\n${existingSheets.join('\n')}\n\n既存のシートをクリアして再作成しますか？`,
        ui.ButtonSet.YES_NO_CANCEL
      );
      
      if (response === ui.Button.CANCEL) {
        throw new Error('ユーザーがキャンセルしました。');
      }
      
      if (response === ui.Button.YES) {
        existingSheets.forEach(sheetName => {
          try {
            const sheet = ss.getSheetByName(sheetName);
            if (sheet) {
              ss.deleteSheet(sheet);
            }
          } catch (error) {
            writeLog('WARNING', 'processInitializeSetup', `シート削除エラー (${sheetName}): ${error.toString()}`);
          }
        });
      }
    }
    
    // ステップ2: ログシートを作成（最初に作成してエラーを記録できるようにする）
    try {
      createLogSheet();
      createdSheets.push('ログ');
      writeLog('INFO', 'processInitializeSetup', 'ログシートを作成しました');
    } catch (error) {
      Logger.log(`ログシートの作成に失敗しました: ${error.toString()}`);
      // ログシートの作成失敗は続行可能
    }
    
    // ステップ3: クライアントシートを作成
    try {
      createClientSheet();
      createdSheets.push('クライアント');
      writeLog('INFO', 'processInitializeSetup', 'クライアントシートを作成しました');
    } catch (error) {
      writeLog('ERROR', 'processInitializeSetup', `クライアントシートの作成に失敗しました: ${error.toString()}`);
      throw new Error(`クライアントシートの作成に失敗しました: ${error.toString()}`);
    }
    
    // ステップ4: 年度別シートを作成
    try {
      createYearlySheet(yearName, startMonth);
      createdSheets.push(yearName);
      writeLog('INFO', 'processInitializeSetup', `年度別シート (${yearName}) を作成しました`);
    } catch (error) {
      writeLog('ERROR', 'processInitializeSetup', `年度別シートの作成に失敗しました: ${error.toString()}`);
      throw new Error(`年度別シートの作成に失敗しました: ${error.toString()}`);
    }
    
    // ステップ5: 月別シートを作成
    try {
      createMonthlySheet(startMonthName, yearName);
      createdSheets.push(startMonthName);
      writeLog('INFO', 'processInitializeSetup', `月別シート (${startMonthName}) を作成しました`);
    } catch (error) {
      writeLog('ERROR', 'processInitializeSetup', `月別シートの作成に失敗しました: ${error.toString()}`);
      throw new Error(`月別シートの作成に失敗しました: ${error.toString()}`);
    }
    
    // ステップ6: 月末請求シートを作成（確実に完了させる）
    try {
      writeLog('INFO', 'processInitializeSetup', `月末請求シート (${endOfMonthSheetName}) の作成を開始します`);
      createEndOfMonthSheet(endOfMonthSheetName);
      createdSheets.push(endOfMonthSheetName);
      writeLog('INFO', 'processInitializeSetup', `月末請求シート (${endOfMonthSheetName}) を作成しました`);
      
      // 作成されたシートが正しく存在するか確認
      const createdSheet = ss.getSheetByName(endOfMonthSheetName);
      if (!createdSheet) {
        throw new Error('月末請求シートが作成されませんでした');
      }
      
      // シートの基本構造が正しいか確認
      const headerRange = createdSheet.getRange(1, 1, 1, 8);
      const headerValues = headerRange.getValues()[0];
      if (headerValues[0] !== '項番') {
        throw new Error('月末請求シートのヘッダーが正しく設定されていません');
      }
      
    } catch (error) {
      writeLog('ERROR', 'processInitializeSetup', `月末請求シートの作成に失敗しました: ${error.toString()}`);
      throw new Error(`月末請求シートの作成に失敗しました: ${error.toString()}`);
    }
    
    // ステップ7: メニューを追加
    try {
      createMenu();
      writeLog('INFO', 'processInitializeSetup', 'メニューを追加しました');
    } catch (error) {
      writeLog('WARNING', 'processInitializeSetup', `メニューの追加に失敗しました: ${error.toString()}`);
      // メニューの追加失敗は続行可能
    }
    
    writeLog('INFO', 'processInitializeSetup', '初期設定が完了しました');
    return `初期設定が完了しました。\n\n作成されたシート：\n${createdSheets.map(s => `- ${s}`).join('\n')}`;
  } catch (error) {
    writeLog('ERROR', 'processInitializeSetup', `初期設定エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 初期設定（メニューから実行）
 */
function initializeSetup() {
  const html = HtmlService.createHtmlOutput(getInitializeSetupHtml())
    .setWidth(350)
    .setHeight(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * カスタムメニューを作成
 */
function createMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('売上管理')
    .addItem('初期設定', 'initializeSetup')
    .addSeparator()
    .addItem('新しい年度のシートを作成', 'createNewYearSheet')
    .addSeparator()
    .addItem('新しい月のシートを作成', 'createNewMonthSheet')
    .addSeparator()
    .addItem('クライアントシートを作成', 'createClientSheetFromMenu')
    .addSeparator()
    .addItem('既存シートを更新', 'updateAllExistingSheets')
    .addToUi();
}

/**
 * スプレッドシートを開いたときにメニューを追加
 */
function onOpen() {
  createMenu();
}

/**
 * ============================================
 * Web API エンドポイント（案件登録アプリ用）
 * ============================================
 */

/**
 * GETリクエストの処理（APIエンドポイント）
 * @param {Object} e イベントオブジェクト
 * @returns {Object} JSONレスポンス
 */
function doGet(e) {
  try {
    // 認証チェック（コメントアウト - Webアプリとして公開する場合、認証はWebアプリの設定で行う）
    // const user = Session.getActiveUser();
    // if (!user) {
    //   return ContentService.createTextOutput(JSON.stringify({
    //     success: false,
    //     error: '認証が必要です'
    //   })).setMimeType(ContentService.MimeType.JSON);
    // }
    
    const path = e.parameter.path || '';
    const action = e.parameter.action || '';
    
    // パスに応じて処理を分岐
    if (path === 'clients' || action === 'clients') {
      return handleGetClients();
    } else if (path === 'statuses' || action === 'statuses') {
      return handleGetStatuses();
    } else if (path === 'months' || action === 'months') {
      return handleGetMonths();
    } else if (action === 'registerCase' || e.parameter.sheetType) {
      // 案件登録処理（GETリクエストで処理）
      return handleRegisterCase(e.parameter);
    } else {
      return createCorsResponse({
        success: false,
        error: '無効なエンドポイントです'
      });
    }
  } catch (error) {
    writeLog('ERROR', 'doGet', error.toString());
    return createCorsResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * POSTリクエストの処理（案件登録）
 * @param {Object} e イベントオブジェクト
 * @returns {Object} JSONレスポンス
 */
function doPost(e) {
  try {
    // 認証チェック（コメントアウト - Webアプリとして公開する場合、認証はWebアプリの設定で行う）
    // const user = Session.getActiveUser();
    // if (!user) {
    //   return ContentService.createTextOutput(JSON.stringify({
    //     success: false,
    //     error: '認証が必要です'
    //   })).setMimeType(ContentService.MimeType.JSON);
    // }
    
    // リクエストボディをパース
    let requestData;
    if (e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      requestData = e.parameter;
    } else {
      throw new Error('リクエストデータがありません');
    }
    
    // 案件登録処理
    if (requestData.action === 'registerCase' || requestData.sheetType) {
      return handleRegisterCase(requestData);
    } else {
      return createCorsResponse({
        success: false,
        error: '無効なエンドポイントです'
      });
    }
  } catch (error) {
    writeLog('ERROR', 'doPost', error.toString());
    return createCorsResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * CORS対応のレスポンスを作成
 * @param {Object} data レスポンスデータ
 * @returns {TextOutput} CORSヘッダー付きレスポンス
 */
function createCorsResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  // Google Apps ScriptのWebアプリとして公開する場合、CORSヘッダーは自動的に設定される
  // ただし、開発環境では手動で設定する必要がある場合がある
  return output;
}

/**
 * ============================================
 * マスターデータ取得関数
 * ============================================
 */

/**
 * クライアント一覧を取得
 * @returns {TextOutput} JSONレスポンス
 */
function handleGetClients() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientSheet = ss.getSheetByName('クライアント');
    
    if (!clientSheet) {
      return createCorsResponse({
        success: false,
        error: 'クライアントシートが見つかりません'
      });
    }
    
    // クライアント名を取得（A列、2行目以降）
    const clientRange = clientSheet.getRange(2, 1, CONSTANTS.MAX_CLIENT_ROWS, 1);
    const clientValues = clientRange.getValues();
    
    // 空でないクライアント名をフィルタリング
    const clients = clientValues
      .map(row => row[0])
      .filter(client => client && client.toString().trim() !== '')
      .map(client => client.toString().trim());
    
    writeLog('INFO', 'handleGetClients', `${clients.length}件のクライアントを取得しました`);
    
    return createCorsResponse({
      success: true,
      clients: clients
    });
  } catch (error) {
    writeLog('ERROR', 'handleGetClients', error.toString());
    return createCorsResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * ステータス一覧を取得
 * @returns {TextOutput} JSONレスポンス
 */
function handleGetStatuses() {
  try {
    const statuses = CONSTANTS.STATUS_OPTIONS;
    
    writeLog('INFO', 'handleGetStatuses', `${statuses.length}件のステータスを取得しました`);
    
    return createCorsResponse({
      success: true,
      statuses: statuses
    });
  } catch (error) {
    writeLog('ERROR', 'handleGetStatuses', error.toString());
    return createCorsResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * 月一覧を取得（既存の月別シートの名前を取得）
 * @returns {TextOutput} JSONレスポンス
 */
function handleGetMonths() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    
    // 月別シートを抽出（「○月」という形式のシート名）
    const monthSheets = allSheets
      .map(sheet => sheet.getName())
      .filter(name => /^\d+月$/.test(name))
      .sort((a, b) => {
        // 月の順番でソート（11月、12月、1月...の順）
        const monthA = parseInt(a.replace('月', ''));
        const monthB = parseInt(b.replace('月', ''));
        return monthA - monthB;
      });
    
    writeLog('INFO', 'handleGetMonths', `${monthSheets.length}件の月別シートを取得しました`);
    
    return createCorsResponse({
      success: true,
      months: monthSheets
    });
  } catch (error) {
    writeLog('ERROR', 'handleGetMonths', error.toString());
    return createCorsResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * ============================================
 * 案件登録処理関数
 * ============================================
 */

/**
 * 案件登録処理
 * @param {Object} requestData リクエストデータ
 * @returns {TextOutput} JSONレスポンス
 */
function handleRegisterCase(requestData) {
  try {
    // 必須項目の検証
    const requiredFields = ['sheetType', 'month', 'clientName', 'revenue', 'status', 'notes'];
    for (const field of requiredFields) {
      if (!requestData[field] || requestData[field].toString().trim() === '') {
        return createCorsResponse({
          success: false,
          error: `${field}は必須項目です`
        });
      }
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetType = requestData.sheetType;
    const month = requestData.month;
    
    // 登録日を処理（未入力の場合は当日）
    let registrationDate = requestData.registrationDate;
    if (!registrationDate) {
      registrationDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    if (sheetType === '月別シート') {
      return registerToMonthlySheet(ss, month, requestData, registrationDate);
    } else if (sheetType === '月末請求シート') {
      return registerToEndOfMonthSheet(ss, month, requestData, registrationDate);
    } else {
      return createCorsResponse({
        success: false,
        error: '無効なシートタイプです'
      });
    }
  } catch (error) {
    writeLog('ERROR', 'handleRegisterCase', error.toString());
    return createCorsResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * 月別シートに案件を登録
 * @param {Spreadsheet} ss スプレッドシートオブジェクト
 * @param {string} month 月名（例: "11月"）
 * @param {Object} data 案件データ
 * @param {string} registrationDate 登録日
 * @returns {TextOutput} JSONレスポンス
 */
function registerToMonthlySheet(ss, month, data, registrationDate) {
  try {
    const sheet = ss.getSheetByName(month);
    if (!sheet) {
      return createCorsResponse({
        success: false,
        error: `「${month}」シートが見つかりません`
      });
    }
    
    // 空いている行を探す（2行目から67行目まで）
    let targetRow = null;
    for (let row = 2; row <= CONSTANTS.TOTAL_ROW; row++) {
      const clientValue = sheet.getRange(row, 4).getValue(); // D列（クライアント）
      if (!clientValue || clientValue.toString().trim() === '') {
        targetRow = row;
        break;
      }
    }

    if (!targetRow) {
      return createCorsResponse({
        success: false,
        error: '登録可能な空き行がありません'
      });
    }

    // 項番を設定（行番号 - 1）
    const itemNumber = targetRow - 1;

    // データを準備
    const clientName = data.clientName.toString().trim();
    const industry = data.industry ? data.industry.toString().trim() : '';
    const expense = data.expense ? parseFloat(data.expense) : '';
    const expenseUsd = data.expenseUsd ? parseFloat(data.expenseUsd) : '';
    const revenue = parseFloat(data.revenue);
    const status = data.status.toString().trim();

    // 受注日（orderDateが指定されている場合はそれを使用、なければregistrationDateを使用）
    const orderDateStr = data.orderDate || registrationDate;
    const orderDateObj = new Date(orderDateStr);

    // 納期
    const deliveryDateStr = data.deliveryDate || '';
    const deliveryDateObj = deliveryDateStr ? new Date(deliveryDateStr) : '';

    // 備考に登録日を追加（登録日: YYYY-MM-DD の形式で先頭に追加）
    let notes = data.notes.toString().trim();
    if (notes) {
      notes = `登録日: ${registrationDate}\n${notes}`;
    } else {
      notes = `登録日: ${registrationDate}`;
    }

    // データを書き込み
    sheet.getRange(targetRow, 1).setValue(itemNumber); // 項番
    sheet.getRange(targetRow, 2).setValue(orderDateObj); // 受注日
    sheet.getRange(targetRow, 3).setValue(deliveryDateObj); // 納期
    sheet.getRange(targetRow, 4).setValue(clientName); // クライアント
    sheet.getRange(targetRow, 5).setValue(industry); // 業種
    sheet.getRange(targetRow, 6).setValue(expense || ''); // 経費
    sheet.getRange(targetRow, 7).setValue(expenseUsd || ''); // 経費（ドル）
    sheet.getRange(targetRow, 8).setValue(revenue); // 売上
    // 利益（I列）は計算式で自動計算されるため設定不要
    sheet.getRange(targetRow, 10).setValue(status); // ステータス
    sheet.getRange(targetRow, 11).setValue(notes); // 備考
    
    writeLog('INFO', 'registerToMonthlySheet', `月別シート「${month}」に案件を登録しました（行${targetRow}）`);
    
    return createCorsResponse({
      success: true,
      message: '案件を登録しました',
      row: targetRow
    });
  } catch (error) {
    writeLog('ERROR', 'registerToMonthlySheet', error.toString());
    return createCorsResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * 月末請求シートに案件を登録
 * @param {Spreadsheet} ss スプレッドシートオブジェクト
 * @param {string} month 月名（例: "11月"）
 * @param {Object} data 案件データ
 * @param {string} registrationDate 登録日
 * @returns {TextOutput} JSONレスポンス
 */
function registerToEndOfMonthSheet(ss, month, data, registrationDate) {
  try {
    const endOfMonthSheetName = month + '（月末請求分）';
    const sheet = ss.getSheetByName(endOfMonthSheetName);
    
    if (!sheet) {
      return createCorsResponse({
        success: false,
        error: `「${endOfMonthSheetName}」シートが見つかりません`
      });
    }
    
    // 空いている行を探す（2行目から67行目まで）
    let targetRow = null;
    for (let row = 2; row <= CONSTANTS.TOTAL_ROW; row++) {
      const clientValue = sheet.getRange(row, 4).getValue(); // D列（クライアント）
      if (!clientValue || clientValue.toString().trim() === '') {
        targetRow = row;
        break;
      }
    }

    if (!targetRow) {
      return createCorsResponse({
        success: false,
        error: '登録可能な空き行がありません'
      });
    }

    // 項番を設定（行番号 - 1）
    const itemNumber = targetRow - 1;

    // データを準備
    const clientName = data.clientName.toString().trim();
    const expense = data.expense ? parseFloat(data.expense) : '';
    const revenue = parseFloat(data.revenue);
    const status = data.status.toString().trim();

    // 受注日（登録日を使用。orderDateが指定されている場合はそれを使用）
    let orderDate = data.orderDate || registrationDate;
    // 日付形式を変換（YYYY-MM-DD形式からDateオブジェクトへ）
    const dateParts = orderDate.split('-');
    const orderDateObj = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));

    // 納期
    const deliveryDateStr = data.deliveryDate || '';
    let deliveryDateObj = '';
    if (deliveryDateStr) {
      const deliveryParts = deliveryDateStr.split('-');
      deliveryDateObj = new Date(parseInt(deliveryParts[0]), parseInt(deliveryParts[1]) - 1, parseInt(deliveryParts[2]));
    }

    // 備考に登録日を追加（登録日が受注日と異なる場合のみ）
    let notes = data.notes.toString().trim();
    if (data.orderDate && data.orderDate !== registrationDate) {
      if (notes) {
        notes = `登録日: ${registrationDate}\n${notes}`;
      } else {
        notes = `登録日: ${registrationDate}`;
      }
    }

    // データを書き込み
    sheet.getRange(targetRow, 1).setValue(itemNumber); // 項番
    sheet.getRange(targetRow, 2).setValue(orderDateObj); // 受注日
    sheet.getRange(targetRow, 3).setValue(deliveryDateObj); // 納期
    sheet.getRange(targetRow, 4).setValue(clientName); // クライアント
    sheet.getRange(targetRow, 5).setValue(expense || ''); // 経費
    sheet.getRange(targetRow, 6).setValue(revenue); // 売上
    // 利益（G列）は計算式で自動計算されるため設定不要
    sheet.getRange(targetRow, 8).setValue(status); // ステータス
    sheet.getRange(targetRow, 9).setValue(notes); // 備考
    
    writeLog('INFO', 'registerToEndOfMonthSheet', `月末請求シート「${endOfMonthSheetName}」に案件を登録しました（行${targetRow}）`);
    
    return createCorsResponse({
      success: true,
      message: '案件を登録しました',
      row: targetRow
    });
  } catch (error) {
    writeLog('ERROR', 'registerToEndOfMonthSheet', error.toString());
    return createCorsResponse({
      success: false,
      error: error.toString()
    });
  }
}
