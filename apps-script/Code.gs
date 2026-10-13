/**
 * YUSEN — Google Apps Script backend starter
 * Run setupYusen() once from the editor.
 *
 * Production note: for sensitive deployments, prefer an authenticated POST
 * API or serve the UI through Apps Script and use google.script.run.
 * JSONP is used to support a static GitHub Pages frontend.
 */
const SHEETS = {
  Users: ['userId', 'name', 'department', 'qrToken', 'active'],
  Attendance: ['timestamp', 'date', 'time', 'userId', 'name', 'action', 'status'],
  LeaveRequests: ['timestamp', 'userId', 'name', 'type', 'date', 'note', 'status'],
  Settings: ['key', 'value']
};

function setupYusen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Google Sheet target terlebih dahulu.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold');
    sh.autoResizeColumns(1, SHEETS[name].length);
  });
  const settings = ss.getSheetByName('Settings');
  if (settings.getLastRow() < 2) {
    settings.appendRow(['organizationName', 'YUSEN']);
    settings.appendRow(['workStart', '08:00']);
    settings.appendRow(['lateAfter', '08:15']);
  }
  return 'YUSEN sheets are ready.';
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = /^[a-zA-Z_$][0-9a-zA-Z_$\.]{0,80}$/.test(p.callback || '') ? p.callback : '';
  let result;
  try {
    const action = String(p.action || 'summary');
    const data = p.payload ? JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p.payload)).getDataAsString()) : p;
    result = route_(action, data);
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  const body = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function route_(action, data) {
  if (action === 'summary') return summary_();
  if (action === 'attendance') return getAttendance_(data || {});
  if (action === 'checkin') return recordAttendance_(data || {}, 'CHECK-IN');
  if (action === 'checkout') return recordAttendance_(data || {}, 'CHECK-OUT');
  if (action === 'requestLeave') return requestLeave_(data || {});
  if (action === 'adminUsers') {
    requireAdmin_(data.adminKey);
    return { ok: true, users: getRows_('Users') };
  }
  if (action === 'createUser') {
    requireAdmin_(data.adminKey);
    return createUser_(data);
  }
  if (action === 'createQr') {
    requireAdmin_(data.adminKey);
    return createQr_(data);
  }
  if (action === 'leaveRequests') {
    requireAdmin_(data.adminKey);
    return { ok: true, requests: getRows_('LeaveRequests') };
  }
  if (action === 'updateLeave') {
    requireAdmin_(data.adminKey);
    return updateLeave_(data);
  }
  return { ok: false, error: 'Unknown action.' };
}

function summary_() {
  const today = dateKey_(new Date());
  const users = getRows_('Users').filter(u => String(u.active).toLowerCase() !== 'false' && u.userId);
  const rows = getRows_('Attendance').filter(r => String(r.date) === today);
  const checked = new Set(rows.filter(r => r.action === 'CHECK-IN').map(r => String(r.userId)));
  const late = rows.filter(r => r.action === 'CHECK-IN' && r.status === 'Terlambat').length;
  return { ok: true, date: today, totalUsers: users.length, present: checked.size, late: late,
    absent: Math.max(0, users.length - checked.size), recent: rows.slice(-8).reverse() };
}

function getAttendance_(data) {
  const rows = getRows_('Attendance').reverse();
  const date = String(data.date || '');
  const q = String(data.q || '').toLowerCase();
  return { ok: true, rows: rows.filter(r => (!date || String(r.date) === date) &&
    (!q || String(r.name).toLowerCase().includes(q) || String(r.userId).toLowerCase().includes(q))).slice(0, 500) };
}

function recordAttendance_(data, action) {
  const token = String(data.qrToken || data.token || '');
  if (!token) return { ok: false, error: 'QR token is required.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const user = getRows_('Users').find(u => String(u.qrToken) === token && String(u.active).toLowerCase() !== 'false');
    if (!user) return { ok: false, error: 'QR tidak dikenal atau pengguna nonaktif.' };
    const now = new Date();
    const date = dateKey_(now);
    const rows = getRows_('Attendance').filter(r => String(r.userId) === String(user.userId) && String(r.date) === date);
    if (action === 'CHECK-IN' && rows.some(r => r.action === 'CHECK-IN')) return { ok: false, error: 'Check-in hari ini sudah tercatat.' };
    if (action === 'CHECK-OUT' && (!rows.some(r => r.action === 'CHECK-IN') || rows.some(r => r.action === 'CHECK-OUT'))) {
      return { ok: false, error: 'Check-in belum ada atau check-out sudah tercatat.' };
    }
    const lateAfter = getSetting_('lateAfter', '08:15');
    const time = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    const status = action === 'CHECK-IN' && time > lateAfter + ':00' ? 'Terlambat' : 'Tepat waktu';
    sheet_('Attendance').appendRow([now, date, time, user.userId, user.name, action, status]);
    return { ok: true, message: action === 'CHECK-IN' ? 'Check-in berhasil.' : 'Check-out berhasil.',
      name: user.name, time: time, status: status, action: action };
  } finally {
    lock.releaseLock();
  }
}

function requestLeave_(data) {
  const token = String(data.qrToken || data.token || '');
  const user = getRows_('Users').find(u => String(u.qrToken) === token && String(u.active).toLowerCase() !== 'false');
  if (!user) return { ok: false, error: 'Pengguna tidak ditemukan.' };
  const type = String(data.type || '');
  const date = String(data.date || '');
  if (!['Izin', 'Sakit'].includes(type) || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'Jenis atau tanggal izin tidak valid.' };
  sheet_('LeaveRequests').appendRow([new Date(), user.userId, user.name, type, date, String(data.note || '').slice(0, 500), 'Menunggu']);
  return { ok: true, message: 'Permohonan berhasil dikirim.' };
}

function createUser_(data) {
  const name = String(data.name || '').trim();
  if (!name) return { ok: false, error: 'Nama wajib diisi.' };
  const id = String(data.userId || Utilities.getUuid().slice(0, 8)).trim();
  const dept = String(data.department || '').trim();
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  sheet_('Users').appendRow([id, name, dept, token, true]);
  return { ok: true, user: { userId: id, name: name, department: dept, qrToken: token, active: true } };
}

function createQr_(data) {
  const id = String(data.userId || '');
  const sh = sheet_('Users');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: false, error: 'Belum ada pengguna.' };
  const headers = values[0];
  const ix = headers.indexOf('userId'), tx = headers.indexOf('qrToken');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][ix]) === id) {
      const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
      sh.getRange(i + 1, tx + 1).setValue(token);
      return { ok: true, qrToken: token };
    }
  }
  return { ok: false, error: 'Pengguna tidak ditemukan.' };
}

function updateLeave_(data) {
  const rowNum = Number(data.row);
  const status = String(data.status || '');
  if (!Number.isInteger(rowNum) || rowNum < 2 || !['Disetujui', 'Ditolak', 'Menunggu'].includes(status)) {
    return { ok: false, error: 'Data perubahan tidak valid.' };
  }
  sheet_('LeaveRequests').getRange(rowNum, 7).setValue(status);
  return { ok: true };
}

function requireAdmin_(key) {
  const saved = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!saved || !key || key !== saved) throw new Error('Admin key salah atau belum dikonfigurasi.');
}

function getRows_(name) {
  const sh = sheet_(name), values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(row => row.some(v => v !== '')).map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => obj[h] = row[j] instanceof Date ? row[j].toISOString() : row[j]);
    return obj;
  });
}
function getSetting_(key, fallback) {
  const row = getRows_('Settings').find(r => String(r.key) === key);
  return row ? String(row.value) : fallback;
}
function sheet_(name) {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Jalankan setupYusen() terlebih dahulu.');
  const ss = SpreadsheetApp.openById(id);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + name);
  return sh;
}
function dateKey_(date) { return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
