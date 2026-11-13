/**
 * WEBAPP ĐĂNG KÝ PHÒNG MÁY TÍNH  –  Google Apps Script (backend)
 * ------------------------------------------------------------------
 * - Đọc thời khóa biểu từ trang tính "TKB"
 * - Đọc / ghi phiếu đăng ký vào trang tính "DangKy"
 * - Kiểm tra trùng lịch (trùng tiết TKB hoặc trùng người mượn khác)
 * - Gửi email xác nhận cho người mượn (và cho quản trị nếu cấu hình)
 *
 * Lần chạy đầu tiên sẽ TỰ TẠO 2 trang tính "TKB" và "DangKy".
 */

// =================== CẤU HÌNH – thầy/cô sửa ở đây ===================
const CONFIG = {
  SCHOOL_NAME: 'Trường Tiểu học và THCS FPT Cầu Giấy', // Tên trường hiển thị trên web
  ADMIN_EMAIL: '',                         // Email nhận thông báo mỗi khi có người đăng ký ('' = không gửi)
  SEND_CONFIRM_EMAIL: true,                // Gửi email xác nhận cho người mượn
  MAX_DAYS_AHEAD: 60,                      // Chỉ cho đăng ký trước tối đa bao nhiêu ngày
  SPREADSHEET_ID: '',                      // Để trống nếu dán code trong Apps Script của chính file Sheet
  TIMEZONE: 'Asia/Ho_Chi_Minh',
  SHEET_TKB: 'TKB',
  SHEET_BOOKINGS: 'DangKy'
};

const ROOMS = [
  { id: '117T', name: 'Phòng máy 1', short: 'P.117T' },
  { id: '119T', name: 'Phòng máy 2', short: 'P.119T' }
];

const SUBJECTS = { THOC1: 'Tin học', CNGH1: 'Công nghệ', CODI: 'Coding' };

const DAYS = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6'];

// Thời khóa biểu gốc (từ file TKB-phong-may-tong-hop.xlsx) – chỉ dùng để tạo trang "TKB" lần đầu.
// Mỗi dòng: [Tiết, Bắt đầu, Kết thúc, T2-117T, T2-119T, T3-117T, T3-119T, T4-117T, T4-119T, T5-117T, T5-119T, T6-117T, T6-119T]
const DEFAULT_TKB = [
  [1, '08:00', '08:35', '', '', '', '', '', '', '', '', '3A2–THOC1', '4A6–CODI'],
  [2, '08:40', '09:15', '5A2–THOC1', '3A6–THOC1', '', '4A1–THOC1', '3A3–THOC1', '', '5A6–CODI', '', '3A2–CNGH1', '4A4–THOC1'],
  [3, '09:20', '09:55', '5A2–CNGH1', '3A6–CNGH1', '', '4A1–CNGH1', '3A3–CNGH1', '', '5A5–CODI', '', '', '4A4–CNGH1'],
  [4, '10:10', '10:45', '5A3–THOC1', '4A2–THOC1', '5A6–THOC1', '3A4–THOC1', '3A1–THOC1', '5A1–THOC1', '', '4A2–CODI', '', '4A3–THOC1'],
  [5, '10:50', '11:25', '5A3–CNGH1', '4A2–CNGH1', '5A6–CNGH1', '3A4–CNGH1', '3A1–CNGH1', '5A1–CNGH1', '', '4A1–CODI', '', '4A3–CNGH1'],
  [6, '13:35', '14:10', '5A4–THOC1', '3A5–THOC1', '5A5–THOC1', '4A5–CODI', '', '4A6–THOC1', '5A2–CODI', '', '', '4A5–THOC1'],
  [7, '14:15', '14:50', '5A4–CNGH1', '3A5–CNGH1', '5A5–CNGH1', '4A3–CODI', '', '4A6–CNGH1', '5A3–CODI', '', '', '4A5–CNGH1'],
  [8, '15:10', '15:45', '', '5A1–CODI', '', '4A4–CODI', '', '', '5A4–CODI', '', '', '']
];

// Khung NGOÀI GIỜ HỌC: giáo viên được mượn và tự ghi giờ cụ thể (VD 16:00–17:00).
// Ô trong khung này có thể ghi hoạt động cố định, VD "CLB" (chiếm cả khung) hoặc "CLB Robotics 16:00–17:00".
const DEFAULT_WINDOWS = [
  ['Nghỉ trưa', '11:25', '13:35', '', '', '', '', '', '', '', '', '', ''],
  ['Cuối giờ học', '15:45', '17:30', '', '', '', '', 'CLB', '', 'CLB', 'CLB', '', '']
];

const BOOKING_HEADERS = ['Thời điểm gửi', 'Mã ĐK', 'Họ và tên', 'Email', 'Số điện thoại', 'Phòng',
  'Ngày mượn', 'Thứ', 'Tiết', 'Khung giờ', 'Mục đích sử dụng', 'Số người dùng', 'Cam kết', 'Trạng thái', 'Ghi chú'];
const STATUS = { NEW: 'Đã đăng ký', OK: 'Đã duyệt', CANCEL: 'Đã hủy' };

// =================== WEB ===================
function doGet(e) {
  // Trang trên GitHub Pages gọi ?action=data để lấy dữ liệu dạng JSON
  if (e && e.parameter && e.parameter.action === 'data') return json_(getData());
  ensureSetup_(getSS_());
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Đăng ký phòng máy tính')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Trang trên GitHub Pages gửi phiếu đăng ký bằng POST (JSON). */
function doPost(e) {
  let f = {};
  try { f = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, type: 'invalid', message: 'Dữ liệu gửi lên không hợp lệ.' }); }
  return json_(submitBooking(f));
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Chạy hàm này 1 lần trong trình soạn thảo để tạo sẵn các trang tính + cấp quyền. */
function setup() {
  ensureSetup_(getSS_());
  Logger.log('Đã tạo xong trang "TKB" và "DangKy".');
}

/** Dữ liệu cho trang web: TKB + các lượt mượn còn hiệu lực (ẩn email/SĐT). */
function getData() {
  const ss = getSS_();
  ensureSetup_(ss);
  const tkb = readTkb_(ss);
  const from = addDays_(todayStr_(), -14);
  const bookings = readBookings_(ss)
    .filter(b => !isCancelled_(b.status) && b.date >= from)
    .map(b => ({ code: b.code, name: b.name, room: b.room, date: b.date, periods: b.periods,
      win: b.win, start: b.start, end: b.end, purpose: b.purpose, status: b.status }));
  return {
    config: { school: CONFIG.SCHOOL_NAME, maxDaysAhead: CONFIG.MAX_DAYS_AHEAD },
    rooms: ROOMS, subjects: SUBJECTS, days: DAYS,
    periods: tkb.periods, timetable: tkb.timetable, windows: tkb.windows, winTT: tkb.winTT, bookings: bookings
  };
}

/** Nhận phiếu đăng ký từ web. */
function submitBooking(f) {
  f = f || {};
  const data = {
    name: clean_(f.name, 80), email: clean_(f.email, 120).toLowerCase(),
    phone: clean_(f.phone, 20).replace(/[\s.\-]/g, ''), room: clean_(f.room, 10),
    date: clean_(f.date, 10), purpose: clean_(f.purpose, 500),
    students: clean_(f.students, 5), agree: f.agree === true,
    periods: (Array.isArray(f.periods) ? f.periods : []).map(Number)
      .filter(p => p >= 1 && p <= 20).filter((p, i, a) => a.indexOf(p) === i).sort((a, b) => a - b),
    win: clean_(f.win, 40), start: normTime_(clean_(f.start, 5)), end: normTime_(clean_(f.end, 5))
  };
  if (data.win) data.periods = []; else { data.start = ''; data.end = ''; }

  const err = validate_(data);
  if (err) return { ok: false, type: 'invalid', message: err };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {
    return { ok: false, type: 'busy', message: 'Hệ thống đang bận, vui lòng thử lại sau vài giây.' };
  }
  let booking;
  try {
    const ss = getSS_();
    ensureSetup_(ss);
    const tkb = readTkb_(ss);
    const err2 = validateTime_(data, tkb);
    if (err2) return { ok: false, type: 'invalid', message: err2 };

    const conflicts = findConflicts_(tkb, readBookings_(ss), data);
    if (conflicts.length) {
      return { ok: false, type: 'conflict', message: 'Thời gian bạn chọn đã có người sử dụng phòng.', conflicts: conflicts };
    }

    const code = 'PM' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd') + '-' + randomCode_(4);
    const timeRange = data.win ? data.start + '–' + data.end : periodRange_(tkb.periods, data.periods);
    const timeLabel = data.win ? data.win + ': ' + timeRange : 'Tiết ' + data.periods.join(', ') + ' (' + timeRange + ')';
    const dayName = DAYS[dayIdx_(data.date)];
    const sh = ss.getSheetByName(CONFIG.SHEET_BOOKINGS);
    sh.appendRow([
      Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss'), code, data.name, data.email,
      data.phone, roomLabel_(data.room), data.date, dayName, data.win || data.periods.join(','), timeRange,
      data.purpose, data.students, 'Đã cam kết', STATUS.NEW, ''
    ].map(noFormula_));
    SpreadsheetApp.flush();
    booking = { code: code, name: data.name, email: data.email, room: data.room, roomLabel: roomLabel_(data.room),
      date: data.date, dayName: dayName, periods: data.periods, win: data.win, start: data.start, end: data.end,
      timeRange: timeRange, timeLabel: timeLabel, purpose: data.purpose };
  } finally {
    lock.releaseLock();
  }

  booking.emailSent = sendEmails_(booking, data);
  return { ok: true, booking: booking };
}

// =================== KIỂM TRA ===================
function validate_(d) {
  if (d.name.length < 2) return 'Vui lòng nhập họ và tên người mượn.';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(d.email)) return 'Email không hợp lệ.';
  if (!/^(0|\+84)\d{9,10}$/.test(d.phone)) return 'Số điện thoại không hợp lệ (VD: 0912345678).';
  if (!ROOMS.some(r => r.id === d.room)) return 'Vui lòng chọn phòng máy.';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date)) return 'Vui lòng chọn ngày mượn.';
  if (dayIdx_(d.date) < 0) return 'Phòng máy chỉ cho mượn từ Thứ 2 đến Thứ 6.';
  const today = todayStr_();
  if (d.date < today) return 'Không thể đăng ký cho ngày đã qua.';
  if (d.date > addDays_(today, CONFIG.MAX_DAYS_AHEAD)) return 'Chỉ được đăng ký trước tối đa ' + CONFIG.MAX_DAYS_AHEAD + ' ngày.';
  if (!d.periods.length && !d.win) return 'Vui lòng chọn ít nhất 1 tiết hoặc một khung ngoài giờ.';
  if (d.purpose.length < 3) return 'Vui lòng ghi mục đích sử dụng.';
  if (!d.agree) return 'Bạn cần tích vào ô cam kết trước khi gửi.';
  return '';
}

function validateTime_(d, tkb) {
  const nowMin = d.date === todayStr_() ? toMin_(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm')) : -1;
  if (d.win) {
    const w = tkb.windows.filter(x => x.name === d.win)[0];
    if (!w) return 'Khung giờ không hợp lệ.';
    if (!/^\d\d:\d\d$/.test(d.start) || !/^\d\d:\d\d$/.test(d.end)) return 'Vui lòng nhập giờ bắt đầu và giờ kết thúc.';
    if (toMin_(d.start) >= toMin_(d.end)) return 'Giờ kết thúc phải sau giờ bắt đầu.';
    if (toMin_(d.start) < toMin_(w.start) || toMin_(d.end) > toMin_(w.end))
      return 'Thời gian phải nằm trong khung ' + w.name + ' (' + w.start + '–' + w.end + ').';
    if (toMin_(d.end) <= nowMin) return 'Khung giờ này hôm nay đã qua.';
    return '';
  }
  const valid = tkb.periods.map(p => p.p);
  if (d.periods.some(p => valid.indexOf(p) < 0)) return 'Tiết không hợp lệ.';
  const past = d.periods.filter(p => toMin_(tkb.periods.filter(x => x.p === p)[0].end) <= nowMin);
  if (past.length) return 'Tiết ' + past.join(', ') + ' hôm nay đã kết thúc.';
  return '';
}

/** Tất cả hoạt động đã chiếm phòng trong ngày: tiết học, CLB/hoạt động cố định, lượt mượn. */
function occupancy_(tkb, bookings, room, date) {
  const d = dayIdx_(date), out = [];
  if (d < 0) return out;
  tkb.periods.forEach(p => {
    const c = (tkb.timetable[room][d] || [])[p.p - 1];
    if (c) out.push({ start: p.start, end: p.end, label: 'Lớp ' + classLabel_(c) });
  });
  tkb.windows.forEach(w => {
    const c = ((tkb.winTT[room] || [])[d] || {})[w.name];
    if (c) { const a = parseAct_(c, w); out.push({ start: a.start, end: a.end, label: a.label }); }
  });
  bookings.forEach(b => {
    if (isCancelled_(b.status) || b.room !== room || b.date !== date) return;
    bookingTimes_(tkb, b).forEach(t => out.push({ start: t.start, end: t.end, label: 'lượt mượn của ' + b.name }));
  });
  return out;
}

function bookingTimes_(tkb, b) {
  if (b.periods && b.periods.length) {
    return b.periods.map(p => tkb.periods.filter(x => x.p === p)[0]).filter(Boolean).map(x => ({ start: x.start, end: x.end, p: x.p }));
  }
  return (b.start && b.end) ? [{ start: b.start, end: b.end, win: b.win }] : [];
}

function findConflicts_(tkb, bookings, d) {
  const occ = occupancy_(tkb, bookings, d.room, d.date), out = [];
  bookingTimes_(tkb, d).forEach(r => occ.forEach(o => {
    if (toMin_(r.start) < toMin_(o.end) && toMin_(o.start) < toMin_(r.end)) {
      out.push({ when: r.p ? 'Tiết ' + r.p : (r.win + ' ' + r.start + '–' + r.end), time: o.start + '–' + o.end, label: o.label });
    }
  }));
  return out;
}

/** "CLB Robotics 16:00–17:00" -> {label, start, end}; không ghi giờ thì chiếm cả khung. */
function parseAct_(cell, w) {
  const m = String(cell).match(/(\d{1,2})[:h](\d{2})\s*[–\-]\s*(\d{1,2})[:h](\d{2})/);
  const label = (m ? cell.replace(m[0], '') : cell).replace(/[()\s]+$/, '').trim() || 'Hoạt động';
  return m ? { label: label, start: ('0' + m[1]).slice(-2) + ':' + m[2], end: ('0' + m[3]).slice(-2) + ':' + m[4] }
           : { label: label, start: w.start, end: w.end };
}

// =================== ĐỌC DỮ LIỆU ===================
function readTkb_(ss) {
  const values = ss.getSheetByName(CONFIG.SHEET_TKB).getDataRange().getDisplayValues();
  const periods = [], windows = [];
  const timetable = {}, winTT = {};
  ROOMS.forEach(r => { timetable[r.id] = DAYS.map(() => []); winTT[r.id] = DAYS.map(() => ({})); });
  const cell = (row, d, ri) => String(row[3 + d * ROOMS.length + ri] || '').trim();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const p = Number(row[0]);
    const start = normTime_(row[1]), end = normTime_(row[2]);
    if (p) {
      periods.push({ p: p, start: start, end: end });
      DAYS.forEach((_, d) => ROOMS.forEach((r, ri) => { timetable[r.id][d][p - 1] = cell(row, d, ri); }));
    } else if (String(row[0]).trim() && /^\d\d:\d\d$/.test(start) && /^\d\d:\d\d$/.test(end)) {
      const name = String(row[0]).trim();
      windows.push({ name: name, start: start, end: end });
      DAYS.forEach((_, d) => ROOMS.forEach((r, ri) => { const c = cell(row, d, ri); if (c) winTT[r.id][d][name] = c; }));
    }
  }
  return { periods: periods, timetable: timetable, windows: windows, winTT: winTT };
}

function readBookings_(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEET_BOOKINGS);
  if (sh.getLastRow() < 2) return [];
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, BOOKING_HEADERS.length).getValues();
  return v.filter(r => r[1] || r[2]).map(r => ({
    code: String(r[1]), name: String(r[2]), room: roomId_(r[5]), date: normDate_(r[6]),
    periods: /^[\d,;\s]+$/.test(String(r[8])) ? String(r[8]).split(/[,;\s]+/).map(Number).filter(Boolean) : [],
    win: /^[\d,;\s]*$/.test(String(r[8])) ? '' : String(r[8]).trim(),
    start: timeAt_(r[9], 1), end: timeAt_(r[9], 3),
    purpose: String(r[10]), status: String(r[13])
  }));
}
/** Lấy giờ từ ô "16:00–17:00" (i=1: bắt đầu, i=3: kết thúc). */
function timeAt_(v, i) {
  const m = String(v).match(/(\d{1,2})[:h](\d{2})\s*[–\-]\s*(\d{1,2})[:h](\d{2})/);
  return m ? ('0' + m[i]).slice(-2) + ':' + m[i + 1] : '';
}

// =================== TẠO TRANG TÍNH ===================
function ensureSetup_(ss) {
  if (!ss.getSheetByName(CONFIG.SHEET_TKB)) {
    const sh = ss.insertSheet(CONFIG.SHEET_TKB);
    const head = ['Tiết', 'Bắt đầu', 'Kết thúc'];
    DAYS.forEach(d => ROOMS.forEach(r => head.push(d + ' | ' + r.id)));
    const rows = DEFAULT_TKB.slice(0, 5).concat([DEFAULT_WINDOWS[0]], DEFAULT_TKB.slice(5), [DEFAULT_WINDOWS[1]]);
    sh.getRange(1, 1, rows.length + 1, head.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#cfe8fc').setWrap(true);
    sh.getRange(2, 1, rows.length, head.length).setValues(rows.map(r => r.map(String)));
    rows.forEach((r, i) => { if (!Number(r[0])) sh.getRange(i + 2, 1, 1, head.length).setBackground('#fff4d6'); });
    sh.setFrozenRows(1);
    sh.setFrozenColumns(3);
    sh.setColumnWidths(4, head.length - 3, 95);
    sh.getRange(rows.length + 3, 1).setValue(NOTE_TKB_);
  } else {
    migrateWindows_(ss.getSheetByName(CONFIG.SHEET_TKB));
  }
  if (!ss.getSheetByName(CONFIG.SHEET_BOOKINGS)) {
    const sh = ss.insertSheet(CONFIG.SHEET_BOOKINGS);
    sh.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS])
      .setFontWeight('bold').setBackground('#ffe08a');
    sh.setFrozenRows(1);
    sh.getRange('E:E').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@');
    sh.getRange('I:I').setNumberFormat('@');
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([STATUS.NEW, STATUS.OK, STATUS.CANCEL], true).build();
    sh.getRange(2, 14, 999, 1).setDataValidation(rule);
    sh.setColumnWidth(3, 170); sh.setColumnWidth(4, 200); sh.setColumnWidth(11, 260);
  }
}

const NOTE_TKB_ = 'Ghi chú: sửa trực tiếp tên lớp trong ô (dạng 5A2–THOC1). Xóa trống ô = phòng trống. ' +
  'Dòng "Nghỉ trưa", "Cuối giờ học" là khung ngoài giờ cho mượn theo giờ; ghi "CLB" (cả khung) hoặc "CLB Robotics 16:00–17:00". Web cập nhật ngay.';

/** Trang TKB tạo từ bản cũ chưa có khung ngoài giờ: tự chèn thêm "Nghỉ trưa" và "Cuối giờ học". */
function migrateWindows_(sh) {
  const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getDisplayValues().map(r => String(r[0]).trim());
  const width = 3 + DAYS.length * ROOMS.length;
  if (col.indexOf(DEFAULT_WINDOWS[0][0]) < 0) {
    const after = col.indexOf('5') + 1;                      // dòng của tiết 5
    const at = after > 0 ? after : sh.getLastRow();
    sh.insertRowAfter(at);
    sh.getRange(at + 1, 1, 1, width).setNumberFormat('@').setValues([DEFAULT_WINDOWS[0].map(String)]).setBackground('#fff4d6');
    col.splice(at, 0, DEFAULT_WINDOWS[0][0]);
  }
  if (col.indexOf(DEFAULT_WINDOWS[1][0]) < 0) {
    const after = col.indexOf('8') + 1;                      // dòng của tiết 8
    const at = after > 0 ? after : sh.getLastRow();
    sh.insertRowAfter(at);
    sh.getRange(at + 1, 1, 1, width).setNumberFormat('@').setValues([DEFAULT_WINDOWS[1].map(String)]).setBackground('#fff4d6');
  }
}

// =================== EMAIL ===================
function sendEmails_(b, d) {
  let sent = false;
  const dateVN = b.date.split('-').reverse().join('/');
  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#23324a">' +
    '<h2 style="color:#1f6fa0">Xác nhận đăng ký phòng máy</h2>' +
    '<p>Xin chào <b>' + esc_(b.name) + '</b>, phiếu đăng ký của bạn đã được ghi nhận.</p>' +
    '<table cellpadding="6" style="border-collapse:collapse;border:1px solid #cfe8fc">' +
    row_('Mã đăng ký', b.code) + row_('Phòng', b.roomLabel) + row_('Ngày', b.dayName + ', ' + dateVN) +
    row_('Thời gian', b.timeLabel) + row_('Mục đích', b.purpose) + '</table>' +
    '<p>Bạn đã cam kết sử dụng phòng máy đúng mục đích, đúng thời gian và tuân thủ nội quy phòng máy.</p>' +
    '<p style="color:#888">Muốn hủy lịch, vui lòng liên hệ người quản lý phòng máy và cung cấp mã đăng ký.</p></div>';
  const subject = '[Phòng máy] Xác nhận mượn ' + b.roomLabel + ' – ' + dateVN + ' – Mã ' + b.code;
  try {
    if (CONFIG.SEND_CONFIRM_EMAIL) { MailApp.sendEmail({ to: b.email, subject: subject, htmlBody: html }); sent = true; }
    if (CONFIG.ADMIN_EMAIL) {
      MailApp.sendEmail({ to: CONFIG.ADMIN_EMAIL, subject: '[Phòng máy] Đăng ký mới: ' + b.name + ' – ' + dateVN,
        htmlBody: html + '<p>SĐT: ' + esc_(d.phone) + ' – Email: ' + esc_(d.email) + '</p>' });
    }
  } catch (e) { Logger.log('Lỗi gửi mail: ' + e); }
  return sent;
}
function row_(k, v) { return '<tr><td style="background:#eef7fe"><b>' + k + '</b></td><td>' + esc_(v) + '</td></tr>'; }

// =================== TIỆN ÍCH ===================
function getSS_() { return CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActive(); }
function todayStr_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function addDays_(s, n) { const p = s.split('-').map(Number); const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)); return d.toISOString().slice(0, 10); }
function dayIdx_(s) { const p = s.split('-').map(Number); const w = new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay(); return (w >= 1 && w <= 5) ? w - 1 : -1; }
function toMin_(t) { const p = String(t).split(':').map(Number); return p[0] * 60 + (p[1] || 0); }
function normTime_(t) { const m = String(t).match(/(\d{1,2})[:h](\d{2})/); return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : String(t); }
function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : s;
}
function roomLabel_(id) { const r = ROOMS.filter(x => x.id === id)[0]; return r ? r.name + ' (' + r.id + ')' : id; }
function roomId_(v) { const s = String(v); const r = ROOMS.filter(x => s.indexOf(x.id) >= 0)[0]; return r ? r.id : s; }
function isCancelled_(s) { return /h[uủ]y/i.test(String(s)); }
function classLabel_(c) { const p = c.split(/[–\-]/); return p[0] + (p[1] ? ' – ' + (SUBJECTS[p[1].trim()] || p[1]) : ''); }
function periodRange_(periods, list) {
  const sel = periods.filter(p => list.indexOf(p.p) >= 0);
  return sel.map(p => p.start + '–' + p.end).join(', ');
}
function clean_(v, max) { return String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, max); }
function noFormula_(v) { return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v; }
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function randomCode_(n) { const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; let s = ''; for (let i = 0; i < n; i++) s += c[Math.floor(Math.random() * c.length)]; return s; }
