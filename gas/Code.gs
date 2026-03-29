/**
 * KADOMORI 日報ダッシュボード用 GAS (Google Apps Script)
 *
 * スプレッドシートID: 1sBQM_QMUXxm9ILdhZixgAi_bH5-qV-DpMvc-UxtHQDk
 * フォーム回答が自動でこのスプレッドシートに記録される。
 * このGASはダッシュボードからのAPI呼び出しに対してJSONデータを返す。
 *
 * デプロイ手順:
 * 1. Google スプレッドシートを開く
 * 2. 拡張機能 > Apps Script
 * 3. このコードを貼り付け
 * 4. デプロイ > 新しいデプロイ > ウェブアプリ
 *    - 実行: 自分
 *    - アクセス: 全員
 * 5. 生成されたURLをダッシュボードの設定画面に入力
 */

const SPREADSHEET_ID = '1sBQM_QMUXxm9ILdhZixgAi_bH5-qV-DpMvc-UxtHQDk';

// スプレッドシートの列ヘッダー（フォーム回答の順番）
const COLUMNS = [
  'タイムスタンプ',
  '出勤日',
  '出勤店舗',
  '入力者',
  '新規売上 (円)',
  '既存売上 (円)',
  '技術売上 (円)',
  '店販売上 (円)',
  '指名料金 (円)',
  '指名数',
  '既存数',
  '新規次回予約数',
  '既存次回予約数',
  'HPB 新規数',
  'HPB 契約数',
  'meta 新規数',
  'meta 契約数',
  'TikTok 新規数',
  'TikTok 契約数',
  'インバウンド 新規数',
  'インバウンド 契約数',
  'HP 新規数',
  'HP 契約数',
  '紹介 新規数',
  '紹介 契約数'
];

// 店舗名 → 店舗キーのマッピング
const STORE_NAME_TO_KEY = {
  '代官山KADOMORI': 'daikanyama_kadomori',
  '代官山SLEEPY': 'daikanyama_sleepy',
  '恵比寿SLEEPY': 'ebisu_sleepy',
  '大阪KADOMORI': 'osaka_kadomori',
  '大阪SLEEPY': 'osaka_sleepy',
  '福島SLEEPY': 'fukushima_sleepy'
};

const STORE_KEY_TO_NAME = {};
for (const [name, key] of Object.entries(STORE_NAME_TO_KEY)) {
  STORE_KEY_TO_NAME[key] = name;
}

/**
 * GETリクエストのハンドラ
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'load_data';

  try {
    let result;
    switch (action) {
      case 'load_data':
        result = loadData(e);
        break;
      case 'load_settings':
        result = loadSettings();
        break;
      case 'load_passwords':
        result = loadPasswords();
        break;
      case 'load_goals':
        result = loadGoals();
        break;
      default:
        result = loadData(e);
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * POSTリクエストのハンドラ
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';
    let result;

    switch (action) {
      case 'save_settings':
        result = saveSettings(body);
        break;
      case 'save_passwords':
        result = savePasswords(body);
        break;
      case 'save_goals':
        result = saveGoals(body);
        break;
      case 'save_data':
        result = saveData(body);
        break;
      default:
        result = { status: 'error', message: '不明なアクション: ' + action };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * スプレッドシートから日報データを読み込み
 */
function loadData(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0]; // 最初のシート（フォーム回答）
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return []; // ヘッダーのみ

  const headers = data[0];
  const records = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue; // 出勤日がない行をスキップ

    // 出勤日を YYYY/M/D 形式に変換
    let dateStr = '';
    const rawDate = row[1];
    if (rawDate instanceof Date) {
      dateStr = `${rawDate.getFullYear()}/${rawDate.getMonth() + 1}/${rawDate.getDate()}`;
    } else {
      // 文字列の場合はそのまま使うか変換
      const d = new Date(rawDate);
      if (!isNaN(d.getTime())) {
        dateStr = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
      } else {
        dateStr = String(rawDate);
      }
    }

    // 店舗名をキーに変換
    const storeName = String(row[2] || '').trim();
    const storeKey = STORE_NAME_TO_KEY[storeName] || storeName;

    const record = {
      id: i,
      date: dateStr,
      store: storeKey,
      storeName: storeName,
      staff: String(row[3] || '').trim(),
      sales: {
        newSales: toNum(row[4]),
        existingSales: toNum(row[5]),
        treatment: toNum(row[6]),
        retail: toNum(row[7]),
        nomination: toNum(row[8])
      },
      customers: {
        nominationCount: toInt(row[9]),
        existingCount: toInt(row[10]),
        newNextRes: toInt(row[11]),
        existingNextRes: toInt(row[12])
      },
      channels: {
        hpb: { newCount: toInt(row[13]), contractCount: toInt(row[14]) },
        meta: { newCount: toInt(row[15]), contractCount: toInt(row[16]) },
        tiktok: { newCount: toInt(row[17]), contractCount: toInt(row[18]) },
        inbound: { newCount: toInt(row[19]), contractCount: toInt(row[20]) },
        hp: { newCount: toInt(row[21]), contractCount: toInt(row[22]) },
        referral: { newCount: toInt(row[23]), contractCount: toInt(row[24]) }
      }
    };

    records.push(record);
  }

  return records;
}

/**
 * データの保存（編集内容をスプレッドシートに反映）
 */
function saveData(body) {
  const modifiedRows = body.data || [];
  if (modifiedRows.length === 0) return { status: 'success', message: '変更なし' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];

  modifiedRows.forEach(record => {
    const rowIndex = record.id + 1; // ヘッダー分 +1
    const sales = record.sales || {};
    const customers = record.customers || {};
    const channels = record.channels || {};

    // 各列を更新（E列〜Y列 = 5列目〜25列目）
    sheet.getRange(rowIndex, 5).setValue(sales.newSales || 0);
    sheet.getRange(rowIndex, 6).setValue(sales.existingSales || 0);
    sheet.getRange(rowIndex, 7).setValue(sales.treatment || 0);
    sheet.getRange(rowIndex, 8).setValue(sales.retail || 0);
    sheet.getRange(rowIndex, 9).setValue(sales.nomination || 0);
    sheet.getRange(rowIndex, 10).setValue(customers.nominationCount || 0);
    sheet.getRange(rowIndex, 11).setValue(customers.existingCount || 0);
    sheet.getRange(rowIndex, 12).setValue(customers.newNextRes || 0);
    sheet.getRange(rowIndex, 13).setValue(customers.existingNextRes || 0);

    const chOrder = ['hpb', 'meta', 'tiktok', 'inbound', 'hp', 'referral'];
    let col = 14;
    chOrder.forEach(ch => {
      const chData = channels[ch] || {};
      sheet.getRange(rowIndex, col).setValue(chData.newCount || 0);
      sheet.getRange(rowIndex, col + 1).setValue(chData.contractCount || 0);
      col += 2;
    });
  });

  return { status: 'success', message: `${modifiedRows.length} 件のデータを更新しました` };
}

// ========== 設定の保存・読み込み ==========

const SETTINGS_SHEET_NAME = '設定';

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function loadSettings() {
  try {
    const sheet = getOrCreateSheet(SETTINGS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const settings = {};
    data.forEach(row => {
      if (row[0]) {
        try { settings[row[0]] = JSON.parse(row[1]); }
        catch (e) { settings[row[0]] = row[1]; }
      }
    });
    return { status: 'success', settings: settings };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function saveSettings(body) {
  try {
    const sheet = getOrCreateSheet(SETTINGS_SHEET_NAME);
    const settings = body.settings || {};
    const rows = Object.entries(settings).map(([key, val]) => [key, JSON.stringify(val)]);
    sheet.clearContents();
    if (rows.length > 0) {
      sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    }
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function loadPasswords() {
  try {
    const sheet = getOrCreateSheet('パスワード');
    const data = sheet.getDataRange().getValues();
    const passwords = {};
    data.forEach(row => {
      if (row[0] && row[1]) {
        if (!passwords[row[0]]) passwords[row[0]] = {};
        passwords[row[0]][row[1]] = row[2] || '';
      }
    });
    return { status: 'success', passwords: passwords };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function savePasswords(body) {
  try {
    const sheet = getOrCreateSheet('パスワード');
    const passwords = body.passwords || {};
    const rows = [];
    for (const [store, staffMap] of Object.entries(passwords)) {
      for (const [staff, pw] of Object.entries(staffMap)) {
        rows.push([store, staff, pw]);
      }
    }
    sheet.clearContents();
    if (rows.length > 0) {
      sheet.getRange(1, 1, rows.length, 3).setValues(rows);
    }
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function loadGoals() {
  try {
    const sheet = getOrCreateSheet('目標');
    const data = sheet.getDataRange().getValues();
    const goals = {};
    data.forEach(row => {
      if (row[0]) {
        try { goals[row[0]] = JSON.parse(row[1]); }
        catch (e) { goals[row[0]] = row[1]; }
      }
    });
    return { status: 'success', goals: goals };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function saveGoals(body) {
  try {
    const sheet = getOrCreateSheet('目標');
    const goals = body.goals || {};
    const rows = Object.entries(goals).map(([key, val]) => [key, JSON.stringify(val)]);
    sheet.clearContents();
    if (rows.length > 0) {
      sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    }
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// ========== ユーティリティ ==========

function toNum(val) {
  if (val === '' || val === null || val === undefined) return 0;
  const n = parseFloat(val);
  return isNaN(n) ? 0 : n;
}

function toInt(val) {
  if (val === '' || val === null || val === undefined) return 0;
  const n = parseInt(val);
  return isNaN(n) ? 0 : n;
}
