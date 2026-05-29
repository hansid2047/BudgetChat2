// ═══════════════════════════════════════════════════════════════════
//  聊天記帳系統 — Google Apps Script 後端
//
//  設定步驟：
//  1. 開啟 Google 試算表 > 擴充功能 > Apps Script
//  2. 貼上此程式碼，儲存
//  3. 執行 initSheets() 初始化工作表
//  4. 部署 > 新增部署 > 網路應用程式
//     - 以身分執行：我（你的帳號）
//     - 誰可以存取：所有人
//  5. 複製部署 URL 貼到 budget-chat.html 的 GAS_URL
// ═══════════════════════════════════════════════════════════════════

const SHEET_RECORDS = "記帳紀錄";
const SHEET_SUMMARY = "月份摘要";

// ─── 初始化 ──────────────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_RECORDS)) {
    const s = ss.insertSheet(SHEET_RECORDS);
    s.appendRow(["日期","時間","類別","子類別","說明","金額","收/支","卡路里(kcal)","記錄方式","原始輸入"]);
    s.getRange(1,1,1,10).setFontWeight("bold").setBackground("#1C1C2E").setFontColor("#A78BFA");
    s.setFrozenRows(1);
    [100,80,80,100,200,80,60,100,80,200].forEach((w,i)=>s.setColumnWidth(i+1,w));
  }

  if (!ss.getSheetByName(SHEET_SUMMARY)) {
    const s = ss.insertSheet(SHEET_SUMMARY);
    s.appendRow(["月份","總收入","總支出","淨額","餐飲支出","總卡路里"]);
    s.getRange(1,1,1,6).setFontWeight("bold").setBackground("#1C1C2E").setFontColor("#A78BFA");
    s.setFrozenRows(1);
    [100,100,100,100,100,100].forEach((w,i)=>s.setColumnWidth(i+1,w));
  }
}

// ─── JSONP GET 路由 ───────────────────────────────────────────────
function doGet(e) {
  const p  = e.parameter || {};
  const cb = p.callback || "";
  let result;
  try {
    const action = p.action || "";
    if      (action === "getRecords") result = getRecords(p.month);
    else if (action === "addRecord")  result = addRecord(p);
    else if (action === "getSummary") result = getSummary();
    else result = { ok:false, error:"unknown action" };
  } catch(err) {
    result = { ok:false, error:err.message };
  }

  const json = JSON.stringify(result);
  if (cb) return ContentService.createTextOutput(`${cb}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ─── 取得記錄 ─────────────────────────────────────────────────────
function getRecords(month) {
  const s    = getSheet(SHEET_RECORDS);
  const rows = s.getDataRange().getValues();
  if (rows.length <= 1) return { ok:true, data:[] };

  let data = rows.slice(1).map(r => ({
    date:      fmtD(r[0]),
    time:      r[1],
    category:  r[2],
    subCat:    r[3],
    desc:      r[4],
    amount:    parseFloat(r[5]) || 0,
    type:      r[6],       // 收入 / 支出
    calories:  parseFloat(r[7]) || 0,
    inputMode: r[8],
    raw:       r[9],
  }));

  if (month) data = data.filter(r => r.date && r.date.startsWith(month));
  return { ok:true, data };
}

// ─── 新增記錄 ─────────────────────────────────────────────────────
function addRecord(p) {
  const s   = getSheet(SHEET_RECORDS);
  const now = new Date();
  const dateStr = Utilities.formatDate(now, "Asia/Taipei", "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(now, "Asia/Taipei", "HH:mm");

  s.appendRow([
    dateStr,
    timeStr,
    p.category  || "其他",
    p.subCat    || "",
    p.desc      || "",
    parseFloat(p.amount) || 0,
    p.type      || "支出",
    parseFloat(p.calories) || 0,
    p.inputMode || "文字",
    p.raw       || "",
  ]);

  const row   = s.getLastRow();
  const color = row % 2 === 0 ? "#1E1E30" : "#16162A";
  s.getRange(row,1,1,10).setBackground(color).setFontColor("#E2E8F0");

  // 金額欄高亮
  const amtColor = p.type==="收入" ? "#BBF7D0" : "#FCA5A5";
  s.getRange(row,6).setFontColor(amtColor).setFontWeight("bold");

  updateSummary(dateStr.slice(0,7));
  return { ok:true };
}

// ─── 摘要 ─────────────────────────────────────────────────────────
function getSummary() {
  const s    = getSheet(SHEET_SUMMARY);
  const rows = s.getDataRange().getValues();
  if (rows.length <= 1) return { ok:true, data:[] };
  return {
    ok: true,
    data: rows.slice(1).map(r => ({
      month: r[0], income: r[1], expense: r[2],
      net: r[3], foodExpense: r[4], totalCal: r[5],
    }))
  };
}

function updateSummary(month) {
  const records = getSheet(SHEET_RECORDS);
  const rows    = records.getDataRange().getValues().slice(1);
  const monthly = rows.filter(r => fmtD(r[0]).startsWith(month));

  const income  = monthly.filter(r=>r[6]==="收入").reduce((s,r)=>s+(parseFloat(r[5])||0),0);
  const expense = monthly.filter(r=>r[6]==="支出").reduce((s,r)=>s+(parseFloat(r[5])||0),0);
  const foodExp = monthly.filter(r=>r[6]==="支出"&&(r[2]==="餐飲"||r[2]==="食物")).reduce((s,r)=>s+(parseFloat(r[5])||0),0);
  const cal     = monthly.reduce((s,r)=>s+(parseFloat(r[7])||0),0);

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sumSh = ss.getSheetByName(SHEET_SUMMARY);
  const srows = sumSh.getDataRange().getValues();
  let found   = false;
  for (let i=1; i<srows.length; i++) {
    if (String(srows[i][0]) === month) {
      sumSh.getRange(i+1,1,1,6).setValues([[month, income, expense, income-expense, foodExp, cal]]);
      found = true; break;
    }
  }
  if (!found) sumSh.appendRow([month, income, expense, income-expense, foodExp, cal]);
}

// ─── 工具 ────────────────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s    = ss.getSheetByName(name);
  if (!s) { initSheets(); s = ss.getSheetByName(name); }
  return s;
}

function fmtD(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "Asia/Taipei", "yyyy-MM-dd");
  return String(val);
}
