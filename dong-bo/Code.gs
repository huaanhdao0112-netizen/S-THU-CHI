/**
 * SỔ THU CHI — cổng đồng bộ qua Google Sheets
 * Dán toàn bộ file này vào Apps Script của file Google Sheet (Tiện ích mở rộng > Apps Script).
 * Chỉ cần đổi MÃ BÍ MẬT ở dòng dưới, rồi Triển khai > Ứng dụng web (xem HUONG-DAN-DONG-BO.txt).
 */
const MA_BI_MAT = 'DOI-MA-BI-MAT-NAY';

const SHEET_GD = 'GiaoDich';
const SHEET_CD = 'CaiDat';
const COLS = 14;
// id | ngày | loại | số tiền | danh mục | hũ | nội dung | người | hình thức | ghi chú | mã danh mục | ts | đã xóa | cập nhật (máy chủ)
const HEAD = ['id', 'ngay', 'loai', 'so_tien', 'danh_muc', 'hu', 'noi_dung', 'nguoi', 'hinh_thuc', 'ghi_chu', 'ma_dm', 'ts', 'da_xoa', 'cap_nhat'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  const bad = kiemTra_(p.token);
  if (bad) return out_(bad);
  const sh = sheetGD_();
  return out_({ ok: true, rows: Math.max(0, sh.getLastRow() - 1), now: Date.now() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'json' }); }
  const bad = kiemTra_(req && req.token);
  if (bad) return out_(bad);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return out_(xuLy_(req)); } finally { lock.releaseLock(); }
}

function kiemTra_(token) {
  if (MA_BI_MAT === 'DOI-MA-BI-MAT-NAY') return { ok: false, error: 'chua-doi-ma' };
  if (String(token || '') !== MA_BI_MAT) return { ok: false, error: 'sai-ma' };
  return null;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheetGD_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_GD);
  if (!sh) {
    sh = ss.insertSheet(SHEET_GD);
    sh.getRange(1, 1, 1, COLS).setValues([HEAD]);
    sh.setFrozenRows(1);
    // cột id, ngày, mã danh mục là văn bản để Sheets không tự đổi thành ngày/số
    [1, 2, 11].forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  }
  return sh;
}

function sheetCD_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_CD);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CD);
    sh.getRange(1, 1, 1, 2).setValues([['json', 'cap_nhat']]);
    sh.getRange(2, 1, 1, 1).setNumberFormat('@');
  }
  return sh;
}

function ngay_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v == null ? '' : v);
}

function toRow_(it, now) {
  return [
    String(it.id), String(it.date || ''), it.type === 'thu' ? 'Thu' : 'Chi', Number(it.amount) || 0,
    String(it.catName || ''), String(it.jar || ''), String(it.note || ''), String(it.who || ''),
    String(it.pay || ''), String(it.memo || ''), String(it.cat || ''), Number(it.ts) || 0,
    it.deleted ? 1 : 0, now
  ];
}

function fromRow_(v, tz) {
  return {
    id: String(v[0]), date: ngay_(v[1], tz), type: String(v[2]).toLowerCase() === 'thu' ? 'thu' : 'chi',
    amount: Number(v[3]) || 0, catName: String(v[4] || ''), jar: String(v[5] || ''), note: String(v[6] || ''),
    who: String(v[7] || ''), pay: String(v[8] || ''), memo: String(v[9] || ''), cat: String(v[10] || ''),
    ts: Number(v[11]) || 0, deleted: Number(v[12]) === 1 ? 1 : 0, upd: Number(v[13]) || 0
  };
}

function xuLy_(req) {
  const now = Date.now();
  const sh = sheetGD_();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const last = sh.getLastRow();
  const rowOf = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { rowOf[String(r[0])] = i + 2; });

  // 1) ghi: giao dịch thêm mới / xóa (đánh dấu da_xoa=1, không xóa dòng)
  const pushed = {};
  const append = [];
  (req.items || []).forEach(function (it) {
    if (!it || !it.id) return;
    const id = String(it.id);
    if (pushed[id]) return;
    pushed[id] = 1;
    const row = toRow_(it, now);
    if (rowOf[id]) {
      if (it.deleted) sh.getRange(rowOf[id], 13, 1, 2).setValues([[1, now]]);
      else sh.getRange(rowOf[id], 1, 1, COLS).setValues([row]);
    } else {
      append.push(row);
    }
  });
  if (append.length) sh.getRange(last + 1, 1, append.length, COLS).setValues(append);

  // 2) đọc: các dòng mới hơn mốc `since` của máy gửi (trừ dòng máy đó vừa gửi)
  const since = Number(req.since) || 0;
  const items = [];
  const total = sh.getLastRow();
  if (total > 1) {
    sh.getRange(2, 1, total - 1, COLS).getValues().forEach(function (v) {
      if (pushed[String(v[0])]) return;
      if ((Number(v[13]) || 0) >= since) items.push(fromRow_(v, tz));
    });
  }

  // 3) cài đặt (danh mục, hạn mức): máy nào gửi sau thì thắng
  const cd = sheetCD_();
  if (req.settings && typeof req.settings.json === 'string') cd.getRange(2, 1, 1, 2).setValues([[req.settings.json, now]]);
  const cur = cd.getRange(2, 1, 1, 2).getValues()[0];
  const curUpd = Number(cur[1]) || 0;
  let settings = null;
  if (!req.settings && cur[0] && curUpd > (Number(req.sinceSettings) || 0)) settings = { json: String(cur[0]), upd: curUpd };

  return { ok: true, now: now, items: items, settings: settings, settingsUpd: curUpd, total: Math.max(0, sh.getLastRow() - 1) };
}
