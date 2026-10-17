/**
 * SISTEM MANAJEMEN MoU & PERJANJIAN KERJA SAMA
 * Google Apps Script + Google Sheets
 *
 * Peran:
 *  - admin  : lihat, tambah, ubah, hapus MoU + kelola akun
 *  - mitra  : hanya lihat MoU milik perusahaannya sendiri
 */

// ====================== KONFIGURASI ======================
const CONFIG = {
  SHEET_MOU: 'MoU',
  SHEET_USERS: 'Users',
  SHEET_LOG: 'LogNotifikasi',

  // Akun admin pertama (dibuat saat setup() dijalankan).
  // WAJIB: ganti ADMIN_EMAIL dengan email Anda SEBELUM menjalankan setup().
  // ADMIN_PASSWORD dibiarkan kosong = dibuat acak, lihat di Execution log setelah setup().
  // Jangan menulis password asli di sini jika kode diunggah ke GitHub.
  ADMIN_EMAIL: 'admin@contoh.com',
  ADMIN_PASSWORD: '',
  ADMIN_NAMA: 'Administrator',

  // Pengingat dikirim pada H-sekian sebelum kontrak berakhir (0 = hari berakhir)
  REMINDER_DAYS: [90, 30, 14, 7, 1, 0],
  REMINDER_HOUR: 8, // jam pengiriman otomatis (waktu sesuai timezone script)

  NAMA_ORGANISASI: 'Perusahaan Anda',
  SESSION_SECONDS: 6 * 60 * 60 // 6 jam
};

const MOU_HEADERS = [
  'ID', 'NomorDokumen', 'Judul', 'Jenis', 'PihakPertama', 'PihakKedua',
  'EmailPihakPertama', 'WAPihakPertama', 'EmailPihakKedua', 'WAPihakKedua',
  'TanggalMulai', 'TanggalBerakhir', 'Deskripsi', 'LinkDokumen',
  'CatatanInternal', 'ReminderTerkirim', 'DibuatPada', 'DiubahPada'
];
const USER_HEADERS = ['Email', 'Nama', 'Role', 'NamaMitra', 'Salt', 'PasswordHash', 'Aktif'];
const LOG_HEADERS = ['Waktu', 'IDMoU', 'Nomor', 'Kanal', 'Tujuan', 'Status', 'Keterangan'];

// ====================== SETUP (jalankan sekali) ======================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, CONFIG.SHEET_MOU, MOU_HEADERS);
  ensureSheet_(ss, CONFIG.SHEET_USERS, USER_HEADERS);
  ensureSheet_(ss, CONFIG.SHEET_LOG, LOG_HEADERS);

  // Kolom tanggal & nomor telepon sebagai teks agar tidak berubah format
  const mou = ss.getSheetByName(CONFIG.SHEET_MOU);
  ['NomorDokumen', 'WAPihakPertama', 'WAPihakKedua', 'TanggalMulai', 'TanggalBerakhir', 'ReminderTerkirim']
    .forEach(function (h) {
      const c = MOU_HEADERS.indexOf(h) + 1;
      mou.getRange(2, c, mou.getMaxRows() - 1, 1).setNumberFormat('@');
    });

  // Buat admin pertama
  const users = getUsers_();
  let pesanAdmin = 'Akun admin sudah ada, tidak diubah.';
  if (!users.some(function (u) { return u.Email === CONFIG.ADMIN_EMAIL.toLowerCase(); })) {
    const pw = CONFIG.ADMIN_PASSWORD || Utilities.getUuid().slice(0, 12);
    upsertUser_({
      Email: CONFIG.ADMIN_EMAIL, Nama: CONFIG.ADMIN_NAMA, Role: 'admin',
      NamaMitra: '', Aktif: 'TRUE'
    }, pw);
    pesanAdmin = 'Password awal admin: ' + pw + ' (segera ganti setelah login)';
  }
  installDailyTrigger();
  Logger.log('Setup selesai. Login admin: ' + CONFIG.ADMIN_EMAIL);
  Logger.log(pesanAdmin);
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkExpiryAndNotify') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkExpiryAndNotify')
    .timeBased().everyDays(1).atHour(CONFIG.REMINDER_HOUR).create();
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#e8eef3');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ====================== WEB APP ======================
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Perjanjian Kerja Sama & MoU')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ====================== AUTH ======================
function hash_(salt, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function login(email, password) {
  email = String(email || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + email;
  if (Number(cache.get(failKey) || 0) >= 5) {
    return { ok: false, error: 'Terlalu banyak percobaan gagal. Coba lagi dalam 15 menit.' };
  }
  const user = getUsers_().filter(function (u) { return u.Email === email; })[0];
  const valid = user && String(user.Aktif).toUpperCase() === 'TRUE' &&
    hash_(user.Salt, String(password || '')) === user.PasswordHash;
  if (!valid) {
    cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 900);
    return { ok: false, error: 'Email atau password salah.' };
  }
  cache.remove(failKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  const session = { email: user.Email, nama: user.Nama, role: user.Role, mitra: user.NamaMitra };
  cache.put('s_' + token, JSON.stringify(session), CONFIG.SESSION_SECONDS);
  return { ok: true, token: token, user: session };
}

function logout(token) {
  CacheService.getScriptCache().remove('s_' + token);
  return true;
}

function getSession_(token) {
  const raw = CacheService.getScriptCache().get('s_' + token);
  if (!raw) throw new Error('SESI_HABIS');
  return JSON.parse(raw);
}

function requireAdmin_(token) {
  const s = getSession_(token);
  if (s.role !== 'admin') throw new Error('Akses ditolak: hanya admin.');
  return s;
}

function checkSession(token) {
  try { return { ok: true, user: getSession_(token) }; } catch (e) { return { ok: false }; }
}

function changePassword(token, oldPass, newPass) {
  const s = getSession_(token);
  if (!newPass || String(newPass).length < 8) throw new Error('Password baru minimal 8 karakter.');
  const user = getUsers_().filter(function (u) { return u.Email === s.email; })[0];
  if (hash_(user.Salt, String(oldPass)) !== user.PasswordHash) throw new Error('Password lama salah.');
  upsertUser_(user, newPass);
  return true;
}

// ====================== DATA MoU ======================
function getMous(token) {
  const s = getSession_(token);
  const isAdmin = s.role === 'admin';
  return readMous_()
    .filter(function (m) {
      return isAdmin || String(m.PihakKedua).trim().toLowerCase() === String(s.mitra).trim().toLowerCase();
    })
    .map(function (m) {
      const info = statusInfo_(m.TanggalBerakhir);
      m.SisaHari = info.days;
      m.Status = info.status;
      if (!isAdmin) { // data internal tidak dikirim ke mitra
        delete m.CatatanInternal; delete m.ReminderTerkirim;
        delete m.WAPihakPertama; delete m.WAPihakKedua;
        delete m.EmailPihakPertama; delete m.EmailPihakKedua;
      }
      return m;
    });
}

function saveMou(token, data) {
  requireAdmin_(token);
  const required = ['NomorDokumen', 'Judul', 'Jenis', 'PihakPertama', 'PihakKedua', 'TanggalMulai', 'TanggalBerakhir'];
  required.forEach(function (f) {
    if (!String(data[f] || '').trim()) throw new Error('Kolom wajib belum diisi: ' + f);
  });
  if (data.TanggalBerakhir < data.TanggalMulai) throw new Error('Tanggal berakhir tidak boleh sebelum tanggal mulai.');
  if (data.LinkDokumen && !/^https?:\/\//i.test(data.LinkDokumen)) throw new Error('Link dokumen harus diawali http:// atau https://');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MOU);
    const now = nowStr_();
    const rows = sh.getDataRange().getValues();
    let rowIndex = -1;
    if (data.ID) {
      for (let i = 1; i < rows.length; i++) if (String(rows[i][0]) === String(data.ID)) { rowIndex = i + 1; break; }
    }
    const rec = {};
    MOU_HEADERS.forEach(function (h) { rec[h] = data[h] !== undefined ? String(data[h]).trim() : ''; });

    if (rowIndex === -1) {
      rec.ID = 'MOU-' + Utilities.getUuid().slice(0, 8).toUpperCase();
      rec.DibuatPada = now; rec.DiubahPada = now; rec.ReminderTerkirim = '';
      sh.appendRow(MOU_HEADERS.map(function (h) { return rec[h]; }));
    } else {
      const old = rows[rowIndex - 1];
      rec.ID = String(old[0]);
      rec.DibuatPada = String(old[MOU_HEADERS.indexOf('DibuatPada')]);
      rec.DiubahPada = now;
      // jika tanggal berakhir diubah (perpanjangan), reset pengingat
      const oldEnd = toDateStr_(old[MOU_HEADERS.indexOf('TanggalBerakhir')]);
      rec.ReminderTerkirim = oldEnd === rec.TanggalBerakhir ? String(old[MOU_HEADERS.indexOf('ReminderTerkirim')]) : '';
      sh.getRange(rowIndex, 1, 1, MOU_HEADERS.length).setNumberFormat('@')
        .setValues([MOU_HEADERS.map(function (h) { return rec[h]; })]);
    }
    return rec.ID;
  } finally { lock.releaseLock(); }
}

function deleteMou(token, id) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MOU);
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) { sh.deleteRow(i + 1); return true; }
    }
    throw new Error('Data tidak ditemukan.');
  } finally { lock.releaseLock(); }
}

function readMous_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MOU);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, MOU_HEADERS.length).getValues();
  return rows.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const o = {};
    MOU_HEADERS.forEach(function (h, i) {
      o[h] = (h === 'TanggalMulai' || h === 'TanggalBerakhir') ? toDateStr_(r[i]) : String(r[i] === null ? '' : r[i]);
    });
    return o;
  });
}

// ====================== MANAJEMEN AKUN (admin) ======================
function getUsers_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_USERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, USER_HEADERS.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      const o = {};
      USER_HEADERS.forEach(function (h, i) { o[h] = String(r[i]); });
      o.Email = o.Email.toLowerCase();
      return o;
    });
}

function upsertUser_(u, newPassword) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_USERS);
  const email = String(u.Email).trim().toLowerCase();
  const rows = sh.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) if (String(rows[i][0]).toLowerCase() === email) { rowIndex = i + 1; break; }

  let salt = u.Salt, pwHash = u.PasswordHash;
  if (newPassword) { salt = Utilities.getUuid(); pwHash = hash_(salt, newPassword); }
  const rec = [email, u.Nama || '', u.Role || 'mitra', u.NamaMitra || '', salt, pwHash, String(u.Aktif || 'TRUE').toUpperCase()];
  if (rowIndex === -1) sh.appendRow(rec);
  else sh.getRange(rowIndex, 1, 1, USER_HEADERS.length).setValues([rec]);
}

function listUsers(token) {
  requireAdmin_(token);
  return getUsers_().map(function (u) {
    return { Email: u.Email, Nama: u.Nama, Role: u.Role, NamaMitra: u.NamaMitra, Aktif: u.Aktif };
  });
}

function saveUser(token, data) {
  requireAdmin_(token);
  const email = String(data.Email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Format email tidak valid.');
  if (data.Role === 'mitra' && !String(data.NamaMitra || '').trim()) throw new Error('Nama mitra wajib diisi untuk akun mitra (harus sama persis dengan kolom Pihak Kedua).');
  const existing = getUsers_().filter(function (u) { return u.Email === email; })[0];
  if (!existing && (!data.Password || data.Password.length < 8)) throw new Error('Password minimal 8 karakter untuk akun baru.');
  if (data.Password && data.Password.length < 8) throw new Error('Password minimal 8 karakter.');
  const base = existing || {};
  upsertUser_({
    Email: email, Nama: data.Nama, Role: data.Role === 'admin' ? 'admin' : 'mitra',
    NamaMitra: data.NamaMitra, Aktif: data.Aktif === 'FALSE' ? 'FALSE' : 'TRUE',
    Salt: base.Salt, PasswordHash: base.PasswordHash
  }, data.Password || null);
  return true;
}

function deleteUser(token, email) {
  const s = requireAdmin_(token);
  email = String(email).toLowerCase();
  if (email === s.email) throw new Error('Tidak bisa menghapus akun yang sedang dipakai.');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_USERS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === email) { sh.deleteRow(i + 1); return true; }
  }
  throw new Error('Akun tidak ditemukan.');
}

// ====================== NOTIFIKASI ======================
/** Dipanggil trigger harian. Admin juga bisa menjalankannya manual dari aplikasi. */
function checkExpiryAndNotify() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MOU);
  if (!sh || sh.getLastRow() < 2) return 0;
  const col = MOU_HEADERS.indexOf('ReminderTerkirim') + 1;
  const thresholds = CONFIG.REMINDER_DAYS.slice().sort(function (a, b) { return a - b; });
  const mous = readMous_();
  let sent = 0;

  mous.forEach(function (m, idx) {
    const days = statusInfo_(m.TanggalBerakhir).days;
    if (days === null) return;
    let key = null;
    if (days < 0) key = 0;
    else for (let i = 0; i < thresholds.length; i++) { if (days <= thresholds[i]) { key = thresholds[i]; break; } }
    if (key === null) return;

    const done = m.ReminderTerkirim ? m.ReminderTerkirim.split(',') : [];
    if (done.indexOf(String(key)) !== -1) return;
    if (days < 0 && done.indexOf('0') !== -1) return;

    notifyParties_(m, days);
    done.push(String(key));
    sh.getRange(idx + 2, col).setNumberFormat('@').setValue(done.join(','));
    sent++;
  });
  return sent;
}

function runReminderNow(token) {
  requireAdmin_(token);
  return checkExpiryAndNotify();
}

function notifyParties_(m, days) {
  const sisa = days < 0 ? 'telah berakhir ' + Math.abs(days) + ' hari yang lalu'
    : days === 0 ? 'berakhir HARI INI' : 'akan berakhir dalam ' + days + ' hari';
  const subject = '[Pengingat] ' + m.Jenis + ' ' + m.NomorDokumen + ' ' + sisa;
  const text =
    'Yth. ' + m.PihakPertama + ' dan ' + m.PihakKedua + ',\n\n' +
    'Kami mengingatkan bahwa ' + m.Jenis + ' berikut ' + sisa + ':\n\n' +
    'Judul   : ' + m.Judul + '\n' +
    'Nomor   : ' + m.NomorDokumen + '\n' +
    'Periode : ' + fmtDate_(m.TanggalMulai) + ' s.d. ' + fmtDate_(m.TanggalBerakhir) + '\n' +
    (m.LinkDokumen ? 'Dokumen : ' + m.LinkDokumen + '\n' : '') +
    '\nMohon segera menindaklanjuti (perpanjangan atau penutupan kerja sama).\n\n' +
    'Terima kasih,\n' + CONFIG.NAMA_ORGANISASI;

  // Email ke kedua pihak
  [m.EmailPihakPertama, m.EmailPihakKedua].forEach(function (to) {
    if (!to) return;
    try {
      MailApp.sendEmail({ to: to, subject: subject, body: text, name: CONFIG.NAMA_ORGANISASI });
      log_(m, 'Email', to, 'OK', '');
    } catch (e) { log_(m, 'Email', to, 'GAGAL', e.message); }
  });

  // WhatsApp ke kedua pihak
  [m.WAPihakPertama, m.WAPihakKedua].forEach(function (no) {
    if (!no) return;
    sendWhatsApp_(no, text, m);
  });
}

/**
 * WhatsApp via Fonnte (https://fonnte.com). Simpan token di:
 * Project Settings > Script Properties > FONNTE_TOKEN
 * Untuk provider lain (Wablas, Qontak, WA Cloud API), ganti isi fungsi ini.
 */
function sendWhatsApp_(number, message, m) {
  const token = PropertiesService.getScriptProperties().getProperty('FONNTE_TOKEN');
  if (!token) { log_(m, 'WhatsApp', number, 'DILEWATI', 'FONNTE_TOKEN belum diisi'); return; }
  try {
    const res = UrlFetchApp.fetch('https://api.fonnte.com/send', {
      method: 'post',
      headers: { Authorization: token },
      payload: { target: normalizePhone_(number), message: message, countryCode: '62' },
      muteHttpExceptions: true
    });
    const body = res.getContentText();
    const ok = res.getResponseCode() === 200 && /"status"\s*:\s*true/.test(body);
    log_(m, 'WhatsApp', number, ok ? 'OK' : 'GAGAL', ok ? '' : body.slice(0, 200));
  } catch (e) { log_(m, 'WhatsApp', number, 'GAGAL', e.message); }
}

function normalizePhone_(n) {
  n = String(n).replace(/[^\d]/g, '');
  if (n.indexOf('0') === 0) n = '62' + n.slice(1);
  return n;
}

function log_(m, kanal, tujuan, status, ket) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LOG)
    .appendRow([nowStr_(), m.ID, m.NomorDokumen, kanal, "'" + tujuan, status, ket]);
}

function getLogs(token) {
  requireAdmin_(token);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LOG);
  if (sh.getLastRow() < 2) return [];
  const n = Math.min(100, sh.getLastRow() - 1);
  return sh.getRange(sh.getLastRow() - n + 1, 1, n, LOG_HEADERS.length).getDisplayValues().reverse();
}

// ====================== UTIL ======================
function tz_() { return Session.getScriptTimeZone(); }
function nowStr_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'); }

function toDateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return String(v || '').trim();
}

function parseDate_(s) {
  const p = String(s || '').split('-');
  if (p.length !== 3) return null;
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function fmtDate_(s) {
  const d = parseDate_(s);
  return d ? Utilities.formatDate(d, tz_(), 'dd MMMM yyyy') : '-';
}

function statusInfo_(endStr) {
  const end = parseDate_(endStr);
  if (!end) return { days: null, status: 'Tidak valid' };
  const today = parseDate_(Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'));
  const days = Math.round((end - today) / 86400000);
  const status = days < 0 ? 'Berakhir' : days <= 90 ? 'Segera berakhir' : 'Aktif';
  return { days: days, status: status };
}
