/**
 * KADOMORI 日報ダッシュボード用 GAS (Google Apps Script)
 *
 * スプレッドシートID: 1sBQM_QMUXxm9ILdhZixgAi_bH5-qV-DpMvc-UxtHQDk
 *
 * シート構成:
 *   [フォーム回答] - 日報データ（Googleフォームから自動記録）
 *   [設定]         - staffRoster, geminiApiKey 等のKV設定
 *   [パスワード]   - 店舗, スタッフ, パスワード
 *   [目標]         - 月, 店舗, スタッフ, 平日日数, 休日日数, ...
 *   [基本給]       - 店舗, スタッフ, 基本給
 *   [スタッフ役割] - 店舗, スタッフ, 役割
 *   [キャンセル]   - id, 日付, 店舗, スタッフ, 媒体, 種別, 理由, メモ
 *   [在庫]         - id, 商品名, カテゴリ, 価格, 原価, 在庫数, 最低在庫, 単位
 *   [在庫履歴]     - id, 商品ID, 日付, 種別, 数量, メモ
 *   [物販売上]     - id, 日付, 店舗, スタッフ, 商品ID, 商品名, 数量, 価格, 原価
 *   [広告費]       - id, 日付, 店舗, 媒体, 金額, メモ
 *   [サブスクプラン] - id, プラン名, 店舗, 月額, ステータス
 *   [サブスクデータ] - 年月, 店舗, 月初人数, 新規, 解約, 月額
 *   [回数券]       - id, 顧客名, 店舗, 券名, 購入日, 総回数, 使用回数, 価格, 有効期限, ステータス
 */

const SPREADSHEET_ID = '1sBQM_QMUXxm9ILdhZixgAi_bH5-qV-DpMvc-UxtHQDk';

// 店舗名 → 店舗キーのマッピング
const STORE_NAME_TO_KEY = {
  '代官山KADOMORI': 'daikanyama_kadomori',
  '代官山SLEEPY': 'daikanyama_sleepy',
  '恵比寿SLEEPY': 'ebisu_sleepy',
  '大阪KADOMORI': 'osaka_kadomori',
  '大阪SLEEPY': 'osaka_sleepy',
  '福島SLEEPY': 'fukushima_sleepy'
};

// ========== リクエストハンドラ ==========

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'load_data';
  try {
    const handler = GET_HANDLERS[action] || loadData;
    const result = handler(e);
    return jsonResponse(result);
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';
    const handler = POST_HANDLERS[action];
    if (!handler) return jsonResponse({ status: 'error', message: '不明なアクション: ' + action });
    return jsonResponse(handler(body));
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.message });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

const GET_HANDLERS = {
  load_data: loadData,
  load_settings: loadSettings,
  load_passwords: loadPasswords,
  load_goals: loadGoals,
  load_salaries: loadSalaries,
  load_roles: loadRoles,
  load_cancellations: loadCancellations,
  load_inventory: loadInventory,
  load_inventory_history: loadInventoryHistory,
  load_retail_sales: loadRetailSales,
  load_ad_data: loadAdData,
  load_sub_plans: loadSubPlans,
  load_sub_data: loadSubData,
  load_tickets: loadTickets,
  load_all: loadAll
};

const POST_HANDLERS = {
  save_data: saveData,
  save_settings: saveSettings,
  save_passwords: savePasswords,
  save_goals: saveGoals,
  save_salaries: saveSalaries,
  save_roles: saveRoles,
  save_cancellations: saveCancellations,
  save_inventory: saveInventory,
  save_inventory_history: saveInventoryHistory,
  save_retail_sales: saveRetailSales,
  save_ad_data: saveAdData,
  save_sub_plans: saveSubPlans,
  save_sub_data: saveSubData,
  save_tickets: saveTickets
};

// ========== ユーティリティ ==========

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sheet;
}

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

function toStr(val) {
  return val == null ? '' : String(val).trim();
}

/** シートからヘッダー行を除いた全データを配列で返す */
function readSheetRows(sheetName, headers) {
  const sheet = getOrCreateSheet(sheetName, headers);
  const data = sheet.getDataRange().getValues();
  return data.length > 1 ? data.slice(1) : [];
}

/** シートを全消去してヘッダー+データを書き込む */
function writeSheetRows(sheetName, headers, rows) {
  const sheet = getOrCreateSheet(sheetName, headers);
  sheet.clearContents();
  if (headers) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

// ========== 一括読み込み ==========

function loadAll(e) {
  return {
    status: 'success',
    data: loadData(e),
    goals: loadGoals().goals || {},
    salaries: loadSalaries().salaries || {},
    roles: loadRoles().roles || {},
    cancellations: loadCancellations().cancellations || [],
    inventory: loadInventory().inventory || [],
    inventoryHistory: loadInventoryHistory().history || [],
    retailSales: loadRetailSales().retailSales || [],
    adData: loadAdData().adData || [],
    subPlans: loadSubPlans().subPlans || [],
    subData: loadSubData().subData || {},
    tickets: loadTickets().tickets || [],
    settings: loadSettings().settings || {},
    passwords: loadPasswords().passwords || {}
  };
}

// ========== 日報データ ==========

function loadData(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;

    let dateStr = '';
    const rawDate = row[1];
    if (rawDate instanceof Date) {
      dateStr = `${rawDate.getFullYear()}/${rawDate.getMonth() + 1}/${rawDate.getDate()}`;
    } else {
      const d = new Date(rawDate);
      dateStr = !isNaN(d.getTime()) ? `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}` : String(rawDate);
    }

    const storeName = toStr(row[2]);
    records.push({
      id: i,
      date: dateStr,
      store: STORE_NAME_TO_KEY[storeName] || storeName,
      storeName: storeName,
      staff: toStr(row[3]),
      sales: { newSales: toNum(row[4]), existingSales: toNum(row[5]), treatment: toNum(row[6]), retail: toNum(row[7]), nomination: toNum(row[8]) },
      customers: { nominationCount: toInt(row[9]), existingCount: toInt(row[10]), newNextRes: toInt(row[11]), existingNextRes: toInt(row[12]) },
      channels: {
        hpb: { newCount: toInt(row[13]), contractCount: toInt(row[14]) },
        meta: { newCount: toInt(row[15]), contractCount: toInt(row[16]) },
        tiktok: { newCount: toInt(row[17]), contractCount: toInt(row[18]) },
        inbound: { newCount: toInt(row[19]), contractCount: toInt(row[20]) },
        hp: { newCount: toInt(row[21]), contractCount: toInt(row[22]) },
        referral: { newCount: toInt(row[23]), contractCount: toInt(row[24]) }
      }
    });
  }
  return records;
}

function saveData(body) {
  const rows = body.data || [];
  if (rows.length === 0) return { status: 'success', message: '変更なし' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  const chOrder = ['hpb', 'meta', 'tiktok', 'inbound', 'hp', 'referral'];

  rows.forEach(r => {
    const ri = r.id + 1;
    const s = r.sales || {}, c = r.customers || {}, ch = r.channels || {};
    sheet.getRange(ri, 5, 1, 5).setValues([[s.newSales||0, s.existingSales||0, s.treatment||0, s.retail||0, s.nomination||0]]);
    sheet.getRange(ri, 10, 1, 4).setValues([[c.nominationCount||0, c.existingCount||0, c.newNextRes||0, c.existingNextRes||0]]);
    const chVals = [];
    chOrder.forEach(k => { const d = ch[k]||{}; chVals.push(d.newCount||0, d.contractCount||0); });
    sheet.getRange(ri, 14, 1, 12).setValues([chVals]);
  });
  return { status: 'success', message: `${rows.length} 件更新` };
}

// ========== 設定 (KV形式) ==========

function loadSettings() {
  try {
    const rows = readSheetRows('設定', ['キー', '値']);
    const settings = {};
    rows.forEach(r => {
      if (r[0]) { try { settings[r[0]] = JSON.parse(r[1]); } catch(e) { settings[r[0]] = r[1]; } }
    });
    return { status: 'success', settings };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveSettings(body) {
  try {
    const settings = body.settings || {};
    const rows = Object.entries(settings).map(([k,v]) => [k, JSON.stringify(v)]);
    writeSheetRows('設定', ['キー', '値'], rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== パスワード ==========

const PW_HEADERS = ['店舗', 'スタッフ', 'パスワード'];

function loadPasswords() {
  try {
    const rows = readSheetRows('パスワード', PW_HEADERS);
    const passwords = {};
    rows.forEach(r => { if (r[0]&&r[1]) { if(!passwords[r[0]]) passwords[r[0]]={}; passwords[r[0]][r[1]]=r[2]||''; }});
    return { status: 'success', passwords };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function savePasswords(body) {
  try {
    const pw = body.passwords || {};
    const rows = [];
    for (const [store, sm] of Object.entries(pw)) for (const [staff, p] of Object.entries(sm)) rows.push([store, staff, p]);
    writeSheetRows('パスワード', PW_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== 目標 ==========

const GOAL_HEADERS = ['月', '店舗', 'スタッフ', '平日日数', '休日日数', '平日目標', '休日目標', '物販目標', '新規目標', '既存目標', '客単価目標', '新規予約率目標', '予約率目標', '★5目標'];

function loadGoals() {
  try {
    const rows = readSheetRows('目標', GOAL_HEADERS);
    const goals = {};
    rows.forEach(r => {
      const month = toStr(r[0]), store = toStr(r[1]), staff = toStr(r[2]);
      if (!month || !store || !staff) return;
      if (!goals[month]) goals[month] = {};
      if (!goals[month][store]) goals[month][store] = {};
      goals[month][store][staff] = {
        weekdays: toInt(r[3]), weekends: toInt(r[4]),
        weekdayTarget: toInt(r[5]), weekendTarget: toInt(r[6]),
        retail: toInt(r[7]), newCustomers: toInt(r[8]), existingCustomers: toInt(r[9]),
        unitPrice: toInt(r[10]), newReservationRate: toInt(r[11]),
        reservationRate: toInt(r[12]), reviews5Star: toInt(r[13])
      };
    });
    return { status: 'success', goals };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveGoals(body) {
  try {
    const goals = body.goals || {};
    const rows = [];
    for (const [month, stores] of Object.entries(goals)) {
      for (const [store, staffMap] of Object.entries(stores)) {
        for (const [staff, g] of Object.entries(staffMap)) {
          rows.push([month, store, staff, g.weekdays||0, g.weekends||0, g.weekdayTarget||0, g.weekendTarget||0, g.retail||0, g.newCustomers||0, g.existingCustomers||0, g.unitPrice||0, g.newReservationRate||0, g.reservationRate||0, g.reviews5Star||0]);
        }
      }
    }
    writeSheetRows('目標', GOAL_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== 基本給 ==========

const SAL_HEADERS = ['店舗', 'スタッフ', '基本給'];

function loadSalaries() {
  try {
    const rows = readSheetRows('基本給', SAL_HEADERS);
    const salaries = {};
    rows.forEach(r => { if (r[0]&&r[1]) { if(!salaries[r[0]]) salaries[r[0]]={}; salaries[r[0]][r[1]]=toInt(r[2]); }});
    return { status: 'success', salaries };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveSalaries(body) {
  try {
    const sal = body.salaries || {};
    const rows = [];
    for (const [store, sm] of Object.entries(sal)) for (const [staff, v] of Object.entries(sm)) rows.push([store, staff, v]);
    writeSheetRows('基本給', SAL_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== スタッフ役割 ==========

const ROLE_HEADERS = ['店舗', 'スタッフ', '役割'];

function loadRoles() {
  try {
    const rows = readSheetRows('スタッフ役割', ROLE_HEADERS);
    const roles = {};
    rows.forEach(r => { if (r[0]&&r[1]) { if(!roles[r[0]]) roles[r[0]]={}; roles[r[0]][r[1]]=toStr(r[2])||'member'; }});
    return { status: 'success', roles };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveRoles(body) {
  try {
    const roles = body.roles || {};
    const rows = [];
    for (const [store, sm] of Object.entries(roles)) for (const [staff, r] of Object.entries(sm)) rows.push([store, staff, r]);
    writeSheetRows('スタッフ役割', ROLE_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== キャンセル ==========

const CANCEL_HEADERS = ['ID', '日付', '店舗', 'スタッフ', '媒体', '種別', '理由', 'メモ'];

function loadCancellations() {
  try {
    const rows = readSheetRows('キャンセル', CANCEL_HEADERS);
    const cancellations = rows.map(r => ({
      id: toInt(r[0]), date: toStr(r[1]), store: toStr(r[2]), staff: toStr(r[3]),
      channel: toStr(r[4]), type: toStr(r[5]), reason: toStr(r[6]), memo: toStr(r[7])
    })).filter(c => c.date);
    return { status: 'success', cancellations };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveCancellations(body) {
  try {
    const list = body.cancellations || [];
    const rows = list.map(c => [c.id||0, c.date||'', c.store||'', c.staff||'', c.channel||'', c.type||'', c.reason||'', c.memo||'']);
    writeSheetRows('キャンセル', CANCEL_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== 在庫 ==========

const INV_HEADERS = ['ID', '商品名', 'カテゴリ', '価格', '原価', '在庫数', '最低在庫', '単位'];

function loadInventory() {
  try {
    const rows = readSheetRows('在庫', INV_HEADERS);
    const inventory = rows.map(r => ({
      id: toStr(r[0]), name: toStr(r[1]), category: toStr(r[2]),
      price: toInt(r[3]), cost: toInt(r[4]), stock: toInt(r[5]), minStock: toInt(r[6]), unit: toStr(r[7])
    })).filter(i => i.id || i.name);
    return { status: 'success', inventory };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveInventory(body) {
  try {
    const list = body.inventory || [];
    const rows = list.map(i => [i.id||'', i.name||'', i.category||'', i.price||0, i.cost||0, i.stock||0, i.minStock||0, i.unit||'']);
    writeSheetRows('在庫', INV_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== 在庫履歴 ==========

const INV_HIST_HEADERS = ['ID', '商品ID', '日付', '種別', '数量', 'メモ'];

function loadInventoryHistory() {
  try {
    const rows = readSheetRows('在庫履歴', INV_HIST_HEADERS);
    const history = rows.map(r => ({
      id: toStr(r[0]), productId: toStr(r[1]), date: toStr(r[2]),
      type: toStr(r[3]), quantity: toInt(r[4]), note: toStr(r[5])
    })).filter(h => h.productId);
    return { status: 'success', history };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveInventoryHistory(body) {
  try {
    const list = body.history || [];
    const rows = list.map(h => [h.id||'', h.productId||'', h.date||'', h.type||'', h.quantity||0, h.note||'']);
    writeSheetRows('在庫履歴', INV_HIST_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== 物販売上 ==========

const RETAIL_HEADERS = ['ID', '日付', '店舗', 'スタッフ', '商品ID', '商品名', '数量', '価格', '原価'];

function loadRetailSales() {
  try {
    const rows = readSheetRows('物販売上', RETAIL_HEADERS);
    const retailSales = rows.map(r => ({
      id: toStr(r[0]), date: toStr(r[1]), store: toStr(r[2]), staff: toStr(r[3]),
      productId: toStr(r[4]), productName: toStr(r[5]),
      quantity: toInt(r[6]), price: toInt(r[7]), cost: toInt(r[8])
    })).filter(r => r.date);
    return { status: 'success', retailSales };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveRetailSales(body) {
  try {
    const list = body.retailSales || [];
    const rows = list.map(r => [r.id||'', r.date||'', r.store||'', r.staff||'', r.productId||'', r.productName||'', r.quantity||0, r.price||0, r.cost||0]);
    writeSheetRows('物販売上', RETAIL_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== 広告費 ==========

const AD_HEADERS = ['ID', '日付', '店舗', '媒体', '金額', 'メモ'];

function loadAdData() {
  try {
    const rows = readSheetRows('広告費', AD_HEADERS);
    const adData = rows.map(r => ({
      id: toStr(r[0]), date: toStr(r[1]), store: toStr(r[2]),
      channel: toStr(r[3]), amount: toInt(r[4]), memo: toStr(r[5])
    })).filter(a => a.date);
    return { status: 'success', adData };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveAdData(body) {
  try {
    const list = body.adData || [];
    const rows = list.map(a => [a.id||'', a.date||'', a.store||'', a.channel||'', a.amount||0, a.memo||'']);
    writeSheetRows('広告費', AD_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== サブスクプラン ==========

const SUB_PLAN_HEADERS = ['ID', 'プラン名', '店舗', '月額', 'ステータス'];

function loadSubPlans() {
  try {
    const rows = readSheetRows('サブスクプラン', SUB_PLAN_HEADERS);
    const subPlans = rows.map(r => ({
      id: toStr(r[0]), planName: toStr(r[1]), store: toStr(r[2]),
      monthlyFee: toInt(r[3]), status: toStr(r[4]) || 'active'
    })).filter(s => s.planName);
    return { status: 'success', subPlans };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveSubPlans(body) {
  try {
    const list = body.subPlans || [];
    const rows = list.map(s => [s.id||'', s.planName||'', s.store||'', s.monthlyFee||0, s.status||'active']);
    writeSheetRows('サブスクプラン', SUB_PLAN_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== サブスクデータ ==========

const SUB_DATA_HEADERS = ['年月', '店舗', '月初人数', '新規', '解約', '月額'];

function loadSubData() {
  try {
    const rows = readSheetRows('サブスクデータ', SUB_DATA_HEADERS);
    const subData = {};
    rows.forEach(r => {
      const ym = toStr(r[0]), store = toStr(r[1]);
      if (!ym || !store) return;
      if (!subData[ym]) subData[ym] = {};
      subData[ym][store] = {
        startCount: toInt(r[2]), newMembers: toInt(r[3]),
        cancelled: toInt(r[4]), monthlyFee: toInt(r[5])
      };
    });
    return { status: 'success', subData };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveSubData(body) {
  try {
    const subData = body.subData || {};
    const rows = [];
    for (const [ym, stores] of Object.entries(subData)) {
      for (const [store, d] of Object.entries(stores)) {
        rows.push([ym, store, d.startCount||0, d.newMembers||0, d.cancelled||0, d.monthlyFee||0]);
      }
    }
    writeSheetRows('サブスクデータ', SUB_DATA_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ========== 回数券 ==========

const TICKET_HEADERS = ['ID', '顧客名', '店舗', '券名', '購入日', '総回数', '使用回数', '価格', '有効期限', 'ステータス'];

function loadTickets() {
  try {
    const rows = readSheetRows('回数券', TICKET_HEADERS);
    const tickets = rows.map(r => ({
      id: toStr(r[0]), customerName: toStr(r[1]), store: toStr(r[2]),
      ticketName: toStr(r[3]), purchaseDate: toStr(r[4]),
      totalCount: toInt(r[5]), usedCount: toInt(r[6]),
      price: toInt(r[7]), expiryDate: toStr(r[8]), status: toStr(r[9]) || 'active'
    })).filter(t => t.id || t.customerName);
    return { status: 'success', tickets };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveTickets(body) {
  try {
    const list = body.tickets || [];
    const rows = list.map(t => [t.id||'', t.customerName||'', t.store||'', t.ticketName||'', t.purchaseDate||'', t.totalCount||0, t.usedCount||0, t.price||0, t.expiryDate||'', t.status||'active']);
    writeSheetRows('回数券', TICKET_HEADERS, rows);
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}
