/** แอปบันทึกรายรับ-รายจ่าย — Google Apps Script
 *  ข้อมูลเก็บใน Google Sheet, ไฟล์แนบเก็บในโฟลเดอร์ Google Drive (สร้างให้อัตโนมัติ) */
const APP_NAME = 'บันทึกรายรับ-รายจ่าย';
const HEADERS = ['id', 'วันที่', 'หมวดหมู่', 'รายละเอียด', 'จำนวนเงิน', 'ไฟล์แนบ', 'บันทึกเมื่อ', 'ประเภท'];
const CATEGORIES = ['อาหาร', 'เดินทาง', 'ที่พัก', 'ช้อปปิ้ง', 'บิล/สาธารณูปโภค', 'สุขภาพ', 'บันเทิง', 'อื่นๆ'];
const INCOME_CATEGORIES = ['เงินเดือน', 'โบนัส', 'ขายของ/ธุรกิจ', 'ดอกเบี้ย/ปันผล', 'อื่นๆ'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function props_() { return PropertiesService.getUserProperties(); }
function newId_() { return Utilities.getUuid().slice(0, 8); }

function setup_() {
  const p = props_();
  let ss = null, folder = null;
  try { if (p.getProperty('SS_ID')) ss = SpreadsheetApp.openById(p.getProperty('SS_ID')); } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.create(APP_NAME);
    ss.getSheets()[0].setName('Expenses');
    p.setProperty('SS_ID', ss.getId());
  }
  try { if (p.getProperty('FOLDER_ID')) folder = DriveApp.getFolderById(p.getProperty('FOLDER_ID')); } catch (e) {}
  if (!folder) {
    folder = DriveApp.createFolder(APP_NAME + ' - ใบเสร็จ');
    p.setProperty('FOLDER_ID', folder.getId());
  }
  return { ss: ss, folder: folder };
}

function sheet_(ss) {
  let sh = ss.getSheetByName('Expenses') || ss.insertSheet('Expenses');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('#,##0.00');
  }
  if (sh.getRange(1, HEADERS.length).getValue() === '') sh.getRange(1, HEADERS.length).setValue('ประเภท').setFontWeight('bold');
  return sh;
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).trim().slice(0, 10);
}

function getBootstrap() {
  const s = setup_();
  const sh = sheet_(s.ss);
  const last = sh.getLastRow();
  let rows = [];
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    let dirty = false;
    vals.forEach(function (r) {
      if (!r[0] && (r[1] || r[4])) { r[0] = newId_(); dirty = true; }
    });
    if (dirty) sh.getRange(2, 1, vals.length, 1).setValues(vals.map(function (r) { return [r[0]]; }));
    rows = vals
      .filter(function (r) { return r[0] && /^\d{4}-\d{2}-\d{2}/.test(dateStr_(r[1])); })
      .map(function (r) {
        return { id: String(r[0]), date: dateStr_(r[1]), category: String(r[2] || 'อื่นๆ'),
                 note: String(r[3] || ''), amount: Number(r[4]) || 0, file: String(r[5] || ''),
                 type: String(r[7]) === 'รายรับ' ? 'income' : 'expense' };
      });
  }
  return { rows: rows, cats: CATEGORIES, incomeCats: INCOME_CATEGORIES, sheetUrl: s.ss.getUrl(), folderUrl: s.folder.getUrl() };
}

function addEntry(e) {
  const amount = Number(e.amount);
  if (!(amount > 0)) throw new Error('กรอกจำนวนเงินมากกว่า 0');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(e.date)) throw new Error('วันที่ไม่ถูกต้อง');
  const lock = LockService.getUserLock();
  lock.waitLock(20000);
  try {
    const s = setup_();
    const sh = sheet_(s.ss);
    let url = '';
    if (e.file) {
      const blob = Utilities.newBlob(Utilities.base64Decode(e.file.data), e.file.type, e.date + '_' + e.file.name);
      url = s.folder.createFile(blob).getUrl();
    }
    const isIncome = e.type === 'income';
    const row = [newId_(), e.date, e.category || 'อื่นๆ', e.note || '', amount, url, new Date(), isIncome ? 'รายรับ' : 'รายจ่าย'];
    sh.getRange(sh.getLastRow() + 1, 2).setNumberFormat('@');
    sh.appendRow(row);
    return { id: row[0], date: e.date, category: row[2], note: row[3], amount: amount, file: url, type: isIncome ? 'income' : 'expense' };
  } finally {
    lock.releaseLock();
  }
}

function deleteExpense(id) {
  const lock = LockService.getUserLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(setup_().ss);
    const hit = sh.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
    if (!hit || hit.getRow() < 2) throw new Error('ไม่พบรายการนี้');
    sh.deleteRow(hit.getRow());
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** เชื่อมกับ Google Sheet ที่มีอยู่แล้ว (วางลิงก์หรือ ID) */
function linkSheet(input) {
  const m = String(input).match(/[-\w]{25,}/);
  if (!m) throw new Error('ลิงก์ Google Sheet ไม่ถูกต้อง');
  const ss = SpreadsheetApp.openById(m[0]);
  sheet_(ss);
  props_().setProperty('SS_ID', ss.getId());
  return getBootstrap();
}
