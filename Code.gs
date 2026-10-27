/* =====================================================================
 * AGENDAGURU — BACKEND (Google Apps Script + Google Sheets)
 *
 * Tampilan aplikasi (frontend) di-host di GitHub Pages dan selalu versi terbaru.
 * File ini hanya API + database. Cara pasang (untuk tiap guru):
 *  1. Buat Google Spreadsheet baru -> Ekstensi -> Apps Script.
 *  2. Tempel seluruh isi file ini ke Code.gs, simpan.
 *  3. Terapkan -> Deployment baru -> Jenis: Aplikasi web.
 *     Jalankan sebagai: Saya.  Yang memiliki akses: Siapa saja.
 *  4. Buka URL Aplikasi web, klik "Buka AgendaGuru", lalu daftar akun.
 * ===================================================================== */

// ---- PENGATURAN PEMILIK APLIKASI (isi sekali sebelum file ini dibagikan) ----
const FRONTEND_URL = 'https://restinovika123-stack.github.io/agendaguru_/';
// Halaman panduan pemasangan untuk pembeli (landing page di Vercel)
const GUIDE_URL = 'https://agendaguru-id.vercel.app';
// Nomor WhatsApp penjual untuk meminta kode lisensi & bantuan (format internasional, tanpa +)
const SUPPORT_WA = '6285788408882';
// Naikkan angka ini jika fungsi API berubah, supaya frontend bisa memberi tahu guru
// yang masih memakai Code.gs lama untuk memperbaruinya.
// v2: koleksi tambahan otomatis (sheet "X_..."), dipakai Pengolah Nilai.
// v3: lisensi (kode lisensi diverifikasi ke server lisensi penjual).
const API_VERSION = 3;
// Alamat Aplikasi Web server lisensi milik penjual (berakhiran /exec).
const LICENSE_SERVER_URL = 'https://script.google.com/macros/s/AKfycbzyQBF32jM2Gwr5ShNV1Ga2O3X3XqFed9OdPt8jgtparnbIMVozwZW2PJ2RiAJcS-7uTg/exec';

/** Halaman pembuka: mengarahkan guru ke frontend GitHub sambil membawa alamat API ini. */
/* =====================================================================
 * SALINAN TEMPLATE
 * Pembeli memakai "Buat salinan" dari template Spreadsheet milik penjual.
 * Jika properti akun/sesi/lisensi ikut tersalin dari Spreadsheet lain,
 * hapus agar pembeli memulai dengan akun & lisensinya sendiri.
 * ===================================================================== */
function guardCopiedProject_() {
  const p = PropertiesService.getScriptProperties();
  const sid = ScriptApp.getScriptId();
  const tercatat = p.getProperty('PROJECT_SCRIPT_ID');
  if (tercatat === sid) return;
  if (tercatat) {
    const all = p.getProperties();
    Object.keys(all).forEach(k => { if (/^(AUTH_|SESS_|LICENSE_)/.test(k)) p.deleteProperty(k); });
  }
  p.setProperty('PROJECT_SCRIPT_ID', sid);
}

/** Menu "AgendaGuru" di Spreadsheet agar guru mudah membuka aplikasi. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('AgendaGuru')
    .addItem('Buka Aplikasi', 'menuBukaAplikasi')
    .addItem('Cara Mengaktifkan Aplikasi', 'menuCaraAktifkan')
    .addToUi();
}

function menuBukaAplikasi() {
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { }
  if (!url || !/\/exec$/.test(url)) return menuCaraAktifkan();
  const esc = s => String(s).replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,sans-serif;color:#172B4D">' +
    '<p style="margin-top:0">Klik tombol di bawah untuk membuka AgendaGuru. Simpan (bookmark) halamannya agar mudah dibuka lagi.</p>' +
    '<p><a href="' + esc(url) + '" target="_blank" style="display:inline-block;background:#159A9C;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:700">Buka AgendaGuru ↗</a></p>' +
    '<p style="font-size:12px;color:#5E6C84">Alamat aplikasi Anda (buka juga di HP/laptop lain):</p>' +
    '<input readonly value="' + esc(url) + '" onclick="this.select()" style="width:100%;padding:8px;border:1px solid #ccd;border-radius:6px;font-size:12px"></div>'
  ).setWidth(460).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'AgendaGuru');
}

function menuCaraAktifkan() {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,sans-serif;color:#172B4D;font-size:14px;line-height:1.6">' +
    '<p style="margin-top:0">Aplikasi belum diaktifkan. Lakukan sekali saja:</p><ol style="padding-left:18px">' +
    '<li>Klik menu <b>Ekstensi → Apps Script</b>.</li>' +
    '<li>Klik tombol biru <b>Terapkan → Deployment baru</b>.</li>' +
    '<li>Klik ⚙ di samping "Pilih jenis", pilih <b>Aplikasi web</b>.</li>' +
    '<li>Jalankan sebagai: <b>Saya</b>. Yang memiliki akses: <b>Siapa saja</b>. Klik <b>Terapkan</b>.</li>' +
    '<li>Izinkan akses (jika muncul peringatan: <b>Lanjutan → Buka (tidak aman) → Izinkan</b>).</li>' +
    '<li>Kembali ke Spreadsheet ini, klik <b>AgendaGuru → Buka Aplikasi</b>.</li></ol>' +
    '<p style="font-size:12px;color:#5E6C84">Panduan lengkap: <a href="' + GUIDE_URL + '" target="_blank">' + GUIDE_URL + '</a></p></div>'
  ).setWidth(480).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cara Mengaktifkan AgendaGuru');
}

/**
 * KHUSUS PENJUAL: jalankan sekali di Spreadsheet TEMPLATE sebelum dibagikan.
 * Menghapus semua akun, sesi, lisensi, dan isi data, sehingga setiap salinan
 * pembeli bersih. Jangan dijalankan di Spreadsheet yang dipakai mengajar.
 */
function bersihkanTemplate() {
  const p = PropertiesService.getScriptProperties();
  p.deleteAllProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const keep = ss.insertSheet('Petunjuk_' + Date.now());
  sheets.forEach(sh => ss.deleteSheet(sh));
  keep.setName('Petunjuk');
  keep.getRange(1, 1, 4, 1).setValues([
    ['AgendaGuru — salinan Anda sendiri'],
    ['1. Klik menu AgendaGuru → Cara Mengaktifkan Aplikasi (muncul beberapa detik setelah Spreadsheet dibuka).'],
    ['2. Setelah aktif, buka aplikasi lewat menu AgendaGuru → Buka Aplikasi.'],
    ['Data kelas, siswa, dan nilai akan tersimpan otomatis di Spreadsheet ini.']
  ]);
  keep.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  keep.setColumnWidth(1, 720);
  console.log('Template bersih. Bagikan dengan akses "Siapa saja yang memiliki link: Pelihat" lalu kirim link .../copy ke pembeli.');
}

/**
 * Mengambil tampilan aplikasi terbaru (index.html di GitHub) di sisi server, sehingga
 * link /exec milik guru langsung membuka aplikasi. Disimpan di cache 5 menit agar cepat.
 * Mengembalikan null bila gagal (halaman pembuka biasa dipakai sebagai cadangan).
 */
function loadFrontendHtml_() {
  const cache = CacheService.getScriptCache();
  const CHUNK = 90000, KEY = 'FE_HTML_';
  try {
    const n = Number(cache.get(KEY + 'N'));
    if (n > 0) {
      const parts = cache.getAll(Array.from({ length: n }, (_, i) => KEY + i));
      const html = Array.from({ length: n }, (_, i) => parts[KEY + i]).join('');
      if (html && html.indexOf('</html>') > 0) return html;
    }
  } catch (e) { }
  try {
    const stamp = Math.floor(Date.now() / 300000);
    const res = UrlFetchApp.fetch(FRONTEND_URL + 'index.html?t=' + stamp, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return null;
    let html = res.getContentText('UTF-8');
    if (html.indexOf('api.js') < 0 || html.indexOf('app.js') < 0) return null;
    // File pendukung (api.js, app.js) dimuat dari alamat frontend
    html = html.replace('<head>', '<head><base href="' + FRONTEND_URL + '" target="_self">');
    try {
      const items = {}; let i = 0;
      for (let p = 0; p < html.length; p += CHUNK, i++) items[KEY + i] = html.slice(p, p + CHUNK);
      items[KEY + 'N'] = String(i);
      cache.putAll(items, 300);
    } catch (e) { }
    return html;
  } catch (e) {
    return null;
  }
}

function doGet(e) {
  guardCopiedProject_();
  // Link /exec langsung membuka aplikasi. Tambahkan ?halaman=info untuk halaman pembuka lama.
  const minta = (e && e.parameter && e.parameter.halaman) || '';
  if (minta !== 'info' && FRONTEND_URL.indexOf('USERNAME') === -1) {
    const app = loadFrontendHtml_();
    if (app) {
      return HtmlService.createHtmlOutput(app)
        .setTitle('AgendaGuru - Manajemen Kelas Digital')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
  }
  const apiUrl = ScriptApp.getService().getUrl();
  const configured = FRONTEND_URL.indexOf('USERNAME') === -1;
  const target = FRONTEND_URL + (FRONTEND_URL.indexOf('?') >= 0 ? '&' : '?') + 'api=' + encodeURIComponent(apiUrl);
  const esc = s => String(s).replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
  // Pemberitahuan kode lisensi hanya untuk Spreadsheet yang belum berlisensi
  const belumBerlisensi = !PropertiesService.getScriptProperties().getProperty('LICENSE_KEY');
  const waText = 'Halo, saya sudah memasang AgendaGuru dan ingin meminta Kode Lisensi.\nNama: \nEmail Google: ';
  const waUrl = 'https://wa.me/' + SUPPORT_WA + '?text=' + encodeURIComponent(waText);
  const noticeLisensi = belumBerlisensi
    ? '<div class="lic">' +
      '<div class="lic-title">🔑 Siapkan Kode Lisensi</div>' +
      '<p>Saat membuat akun, aplikasi akan meminta <b>Kode Lisensi</b> (format <code>AG-XXXX-XXXX-XXXX</code>). ' +
      'Belum punya atau belum menerima kodenya? Minta ke penjual lewat WhatsApp.</p>' +
      '<a class="btn-wa" href="' + esc(waUrl) + '" target="_blank" rel="noopener">💬 Minta Kode Lisensi via WhatsApp</a>' +
      '</div>'
    : '';
  const body = configured
    ? '<h2>AgendaGuru siap dipakai</h2>' +
      '<p>Database Anda ada di Google Spreadsheet ini. Klik tombol di bawah untuk membuka aplikasi.</p>' +
      noticeLisensi +
      '<p><a class="btn" href="' + esc(target) + '" target="_top">Buka AgendaGuru</a></p>' +
      '<p class="muted">Simpan (bookmark) halaman ini. Jika diminta alamat Apps Script di aplikasi, tempel alamat berikut:</p>' +
      '<input readonly value="' + esc(apiUrl) + '" onclick="this.select()">'
    : '<h2>Pengaturan belum lengkap</h2><p>Pemilik aplikasi perlu mengisi <b>FRONTEND_URL</b> di bagian atas Code.gs.</p>';
  const html = '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:system-ui,sans-serif;max-width:520px;margin:40px auto;padding:0 16px;color:#172B4D}' +
    '.btn{display:inline-block;background:#159A9C;color:#fff;padding:12px 22px;border-radius:10px;text-decoration:none;font-weight:700}' +
    '.muted{color:#6b778c;font-size:14px}input{width:100%;padding:10px;border:1px solid #ccd;border-radius:8px;font-size:13px}' +
    '.lic{background:#FFF7E6;border:1px solid #F5D9A8;border-radius:12px;padding:14px 16px;margin:18px 0}' +
    '.lic-title{font-weight:800;margin-bottom:4px}.lic p{margin:0 0 12px;font-size:14px;line-height:1.5}' +
    '.lic code{background:#fff;padding:1px 6px;border-radius:6px;border:1px solid #F5D9A8;white-space:nowrap}' +
    '.btn-wa{display:inline-block;background:#25D366;color:#fff;padding:10px 18px;border-radius:10px;text-decoration:none;font-weight:700}</style>' +
    '</head><body>' + body + '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('AgendaGuru')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * API untuk frontend GitHub. Permintaan: POST body JSON {fn, args}.
 * Hanya fungsi di daftar ini yang bisa dipanggil dari luar.
 */
function doPost(e) {
  const api = {
    getServerInfo: getServerInfo,
    getAuthStatus: getAuthStatus,
    registerAccount: registerAccount,
    loginAccount: loginAccount,
    logoutAccount: logoutAccount,
    changePassword: changePassword,
    getAllData: getAllData,
    saveAllData: saveAllData,
    uploadFileToDrive: uploadFileToDrive,
    syncToCalendar: syncToCalendar,
    activateLicense: activateLicense
  };
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = Object.prototype.hasOwnProperty.call(api, req.fn) ? api[req.fn] : null;
    if (!fn) throw new Error('Fungsi tidak dikenal: ' + req.fn);
    guardCopiedProject_();
    ensureDatabase_();
    out = { ok: true, result: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function getServerInfo() {
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { }
  return { app: 'AgendaGuru', apiVersion: API_VERSION, registered: isRegistered_(), scriptId: ScriptApp.getScriptId(), url: url };
}

/* =====================================================================
 * LISENSI
 * Kode lisensi diverifikasi ke server lisensi penjual. Server membalas "izin"
 * bertanda tangan digital (berlaku 14 hari) yang diperiksa oleh aplikasi.
 * Izin diperbarui otomatis sebelum habis selama kodenya masih aktif.
 * ===================================================================== */
const LICENSE_RENEW_BEFORE_MS = 7 * 24 * 3600 * 1000; // perbarui jika sisa < 7 hari

function verifyLicenseRemote_(key, email) {
  if (!LICENSE_SERVER_URL) return { ok: false, network: true, message: 'Server lisensi belum diatur oleh penjual.' };
  try {
    const res = UrlFetchApp.fetch(LICENSE_SERVER_URL, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true, followRedirects: true,
      payload: JSON.stringify({
        action: 'verify', key: String(key || '').trim().toUpperCase(), scriptId: ScriptApp.getScriptId(),
        email: email || props_().getProperty('AUTH_EMAIL') || '', app: 'AgendaGuru', version: API_VERSION
      })
    });
    const out = JSON.parse(res.getContentText());
    return out && out.ok ? out : { ok: false, network: true, message: 'Server lisensi tidak merespons dengan benar.' };
  } catch (e) {
    return { ok: false, network: true, message: 'Tidak dapat menghubungi server lisensi: ' + e.message };
  }
}

function storeLicense_(key, license) {
  props_().setProperties({ LICENSE_KEY: String(key).trim().toUpperCase(), LICENSE_PAYLOAD: license.payload, LICENSE_SIG: license.sig });
}

function clearLicense_() {
  ['LICENSE_PAYLOAD', 'LICENSE_SIG'].forEach(k => props_().deleteProperty(k));
}

/** Izin lisensi saat ini, diperbarui otomatis bila hampir habis. */
function currentLicense_() {
  const p = props_();
  const key = p.getProperty('LICENSE_KEY');
  let payload = p.getProperty('LICENSE_PAYLOAD'), sig = p.getProperty('LICENSE_SIG');
  let exp = 0;
  try { exp = Number(JSON.parse(payload || '{}').exp) || 0; } catch (e) { }
  if (key && (!payload || exp - Date.now() < LICENSE_RENEW_BEFORE_MS)) {
    const r = verifyLicenseRemote_(key);
    if (r.ok && r.valid) { storeLicense_(key, r.license); payload = r.license.payload; sig = r.license.sig; }
    else if (r.ok && !r.valid) { clearLicense_(); payload = null; sig = null; return { payload: null, sig: null, scriptId: ScriptApp.getScriptId(), reason: r.reason, message: r.message }; }
    // gangguan jaringan: tetap pakai izin lama sampai masa berlakunya habis
  }
  return { payload: payload || null, sig: sig || null, scriptId: ScriptApp.getScriptId(), hasKey: !!key };
}

function withLicense_(data) {
  data._license = currentLicense_();
  return data;
}

/** Mengaktifkan kode lisensi untuk akun yang sudah login (mis. instalasi lama). */
function activateLicense(token, key) {
  if (!isValidToken_(token)) return { ok: false, authError: true };
  const r = verifyLicenseRemote_(key);
  if (!r.ok) return { ok: false, message: r.message || 'Tidak dapat menghubungi server lisensi.' };
  if (!r.valid) return { ok: false, reason: r.reason, message: r.message || 'Kode lisensi tidak valid.' };
  storeLicense_(key, r.license);
  return { ok: true, license: currentLicense_() };
}

/** Buat sheet otomatis saat pertama dipakai, jadi guru tidak perlu menjalankan setupDatabase. */
function ensureDatabase_() {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config')) setupDatabase();
}

// Kolom TahunPelajaran supaya absensi/nilai bulan yang sama di tahun berbeda tidak saling menimpa.
// Kolom Semester di Grades sebelumnya tidak ikut tersimpan ke sheet, sekarang disimpan.
const ATTENDANCE_HEADERS = ['RecordID', 'ClassID', 'MonthYear', 'DatesData', 'AttendanceData', 'TahunPelajaran'];
const GRADES_HEADERS = ['RecordID', 'ClassID', 'Subject', 'Type', 'MonthYear', 'DatesData', 'Cakupan', 'GradesData', 'Semester', 'TahunPelajaran'];

/* =====================================================================
 * AUTENTIKASI (dicek di server, bukan di browser)
 * - Password disimpan sebagai hash + salt di Script Properties, tidak pernah
 *   dikirim ke browser dan tidak lagi ditulis di sheet Config.
 * - Setelah login, browser memegang token sesi. Semua fungsi data mewajibkan token.
 * ===================================================================== */
const SESSION_DAYS = 30;
const MAX_LOGIN_FAILS = 5;       // batas percobaan login gagal ...
const LOGIN_LOCK_SECONDS = 600;  // ... dalam 10 menit

function props_() { return PropertiesService.getScriptProperties(); }

function hashPassword_(password, salt) {
  let h = salt + '|' + password;
  for (let i = 0; i < 500; i++) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8);
    h = bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  }
  return h;
}

function newRandomToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

function setPassword_(email, password) {
  const salt = newRandomToken_();
  props_().setProperties({
    AUTH_EMAIL: String(email).trim().toLowerCase(),
    AUTH_SALT: salt,
    AUTH_HASH: hashPassword_(password, salt)
  });
}

/** Pindahkan password lama (teks biasa di sheet Config) ke hash, lalu hapus dari sheet. */
function migrateLegacyAuth_() {
  if (props_().getProperty('AUTH_HASH')) return;
  const config = {};
  cleanData('Config').forEach(row => { config[row.Key] = row.Value; });
  if (!config.registeredEmail || !config.registeredPassword) return;
  setPassword_(config.registeredEmail, String(config.registeredPassword));
  delete config.registeredPassword;
  const arr = Object.keys(config).map(k => ({Key: k, Value: config[k]}));
  writeSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'Config', ['Key', 'Value'], arr);
}

function isRegistered_() {
  migrateLegacyAuth_();
  return !!props_().getProperty('AUTH_HASH');
}

function createSession_() {
  const p = props_();
  const now = Date.now();
  // Bersihkan sesi kedaluwarsa
  const all = p.getProperties();
  Object.keys(all).forEach(k => {
    if (k.indexOf('SESS_') === 0 && Number(all[k]) < now) p.deleteProperty(k);
  });
  const token = newRandomToken_();
  p.setProperty('SESS_' + token, String(now + SESSION_DAYS * 24 * 3600 * 1000));
  return token;
}

function isValidToken_(token) {
  if (!token || typeof token !== 'string' || !/^[a-f0-9]{64}$/.test(token)) return false;
  const exp = Number(props_().getProperty('SESS_' + token));
  return !!exp && exp > Date.now();
}

/** Status awal untuk layar login: sudah ada akun atau belum. */
function getAuthStatus() {
  return { registered: isRegistered_() };
}

function registerAccount(name, email, password, licenseKey) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (isRegistered_()) return { ok: false, message: 'Aplikasi ini sudah memiliki akun. Silakan masuk.' };
    email = String(email || '').trim().toLowerCase();
    name = String(name || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, message: 'Format email tidak valid.' };
    if (String(password || '').length < 6) return { ok: false, message: 'Password minimal 6 karakter.' };
    // Kode lisensi wajib dan dicek ke server lisensi sebelum akun dibuat
    if (!String(licenseKey || '').trim()) return { ok: false, message: 'Kode lisensi wajib diisi.' };
    const lic = verifyLicenseRemote_(licenseKey, email);
    if (!lic.ok) return { ok: false, message: lic.message || 'Tidak dapat menghubungi server lisensi. Coba lagi.' };
    if (!lic.valid) return { ok: false, message: lic.message || 'Kode lisensi tidak valid.' };
    storeLicense_(licenseKey, lic.license);
    setPassword_(email, String(password));

    const config = {};
    cleanData('Config').forEach(row => { config[row.Key] = row.Value; });
    delete config.registeredPassword;
    config.registeredEmail = email;
    config.userName = name || config.userName || 'Guru';
    config.userSchool = config.userSchool || 'Sekolah Saya';
    writeSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'Config', ['Key', 'Value'],
      Object.keys(config).map(k => ({Key: k, Value: config[k]})));
  } finally {
    lock.releaseLock();
  }
  return { ok: true, token: createSession_(), data: withLicense_(readAllData_()) };
}

function loginAccount(email, password) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('LOGIN_FAILS')) || 0;
  if (fails >= MAX_LOGIN_FAILS) {
    return { ok: false, message: 'Terlalu banyak percobaan gagal. Coba lagi dalam 10 menit.' };
  }
  if (!isRegistered_()) return { ok: false, notRegistered: true, message: 'Belum ada akun terdaftar.' };

  const p = props_();
  const emailOk = String(email || '').trim().toLowerCase() === p.getProperty('AUTH_EMAIL');
  const passOk = hashPassword_(String(password || ''), p.getProperty('AUTH_SALT')) === p.getProperty('AUTH_HASH');
  if (!emailOk || !passOk) {
    cache.put('LOGIN_FAILS', String(fails + 1), LOGIN_LOCK_SECONDS);
    return { ok: false, message: 'Email atau password salah.' };
  }
  cache.remove('LOGIN_FAILS');
  return { ok: true, token: createSession_(), data: withLicense_(readAllData_()) };
}

function logoutAccount(token) {
  if (token && typeof token === 'string') props_().deleteProperty('SESS_' + token);
  return true;
}

/** Ganti password dari menu profil (wajib password lama). */
function changePassword(token, oldPassword, newPassword) {
  if (!isValidToken_(token)) return { ok: false, authError: true };
  const p = props_();
  if (hashPassword_(String(oldPassword || ''), p.getProperty('AUTH_SALT')) !== p.getProperty('AUTH_HASH')) {
    return { ok: false, message: 'Password lama salah.' };
  }
  if (String(newPassword || '').length < 6) return { ok: false, message: 'Password baru minimal 6 karakter.' };
  setPassword_(p.getProperty('AUTH_EMAIL'), String(newPassword));
  return { ok: true };
}

/**
 * LUPA PASSWORD: jalankan fungsi ini sekali dari editor Apps Script (pilih lalu klik Run).
 * Akun login dan semua sesi dihapus, DATA TIDAK DIHAPUS. Setelah itu buka aplikasi
 * dan daftar ulang dengan email/password baru.
 */
function resetAkunLogin() {
  const p = props_();
  const all = p.getProperties();
  Object.keys(all).forEach(k => {
    if (k.indexOf('SESS_') === 0 || k.indexOf('AUTH_') === 0) p.deleteProperty(k);
  });
  // Pastikan password lama (format lama) tidak ikut termigrasi ulang
  const config = {};
  cleanData('Config').forEach(row => { config[row.Key] = row.Value; });
  delete config.registeredPassword;
  delete config.registeredEmail;
  writeSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'Config', ['Key', 'Value'],
    Object.keys(config).map(k => ({Key: k, Value: config[k]})));
  console.log('Akun login direset. Buka aplikasi lalu daftar ulang. Data kelas/siswa/nilai tetap ada.');
}

/**
 * Jalankan fungsi ini pertama kali untuk membuat struktur database (Sheets).
 * Aman dijalankan ulang kapan saja: hanya membuat sheet yang belum ada,
 * dan memastikan seluruh kolom data diformat sebagai Teks (lihat catatan di writeSheet_).
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    {name: 'Config', headers: ['Key', 'Value']},
    {name: 'Classes', headers: ['ClassID', 'ClassName']},
    {name: 'Students', headers: ['StudentID', 'ClassID', 'StudentName', 'NIS', 'NISN', 'TempatLahir', 'TanggalLahir', 'JenisKelamin', 'Agama', 'Alamat', 'NamaAyah', 'NamaIbu', 'NamaWali', 'ParentPhone']},
    {name: 'Attendance', headers: ATTENDANCE_HEADERS},
    {name: 'Grades', headers: GRADES_HEADERS},
    {name: 'Koku', headers: ['RecordID', 'ClassID', 'Tanggal', 'Jenis', 'Projek', 'Deskripsi', 'Dimensi', 'KokuData']},
    {name: 'Ekskul', headers: ['EkskulID', 'ClassID', 'StudentID', 'StudentName', 'Semester', 'TahunPelajaran', 'NamaEkskul', 'Predikat', 'Keterangan']},
    {name: 'Journal', headers: ['RecordID', 'Date', 'ClassID', 'Subject', 'Content', 'AttachmentUrl']},
    {name: 'Schedule', headers: ['RecordID', 'Day', 'TimeStart', 'TimeEnd', 'Subject', 'ClassID']},
    {name: 'Subjects', headers: ['SubjectName']},
    {name: 'Poin', headers: ['RecordID', 'StudentID', 'ClassID', 'Tanggal', 'PoinReward', 'PoinPelanggaran', 'Keterangan']},
    {name: 'TPBank', headers: ['TPID', 'Subject', 'GradeLevel', 'Code', 'Desc']},
    {name: 'RaporNotes', headers: ['RecordID', 'StudentID', 'ClassID', 'Semester', 'TahunPelajaran', 'CatatanWali', 'Keputusan']}
  ];

  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight('bold').setBackground('#f4f7f6');
      sheet.setFrozenRows(1);
    }
    // PENTING: paksa seluruh kolom data (baris di bawah header) sebagai format Teks.
    // Tanpa ini, Google Sheets otomatis mengubah tipe data saat ditulis lewat setValues(),
    // misalnya string "07" (bulan Juli) menjadi angka 7, atau "2026-07-11" menjadi objek Date.
    // Akibatnya pencocokan data di aplikasi (yang membandingkan string persis, mis. filter bulan)
    // gagal menemukan data tersebut walau datanya benar-benar ada di spreadsheet.
    const maxRows = sheet.getMaxRows();
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, s.headers.length).setNumberFormat('@');
    }
  });
}

/**
 * Helper function untuk membaca Sheet menjadi array of object
 */
function cleanData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Kosong atau hanya header

  const headers = data[0];
  const result = [];

  for (let i = 1; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      // Jaga-jaga jika ada sel lama yang masih tersimpan sebagai Date/Number akibat bug lama
      // (sebelum kolom diformat Teks) supaya tetap terbaca benar oleh frontend.
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return result;
}

/**
 * Mengambil semua data dari spreadsheet ke frontend. Wajib token sesi yang valid.
 */
function getAllData(token) {
  if (!isValidToken_(token)) return { authError: true, registered: isRegistered_() };
  return withLicense_(readAllData_());
}

function readAllData_() {
  const configData = cleanData('Config');
  let configObj = {};
  configData.forEach(row => { configObj[row.Key] = row.Value; });
  delete configObj.registeredPassword; // jangan pernah kirim password ke browser
  configObj.registeredEmail = props_().getProperty('AUTH_EMAIL') || configObj.registeredEmail || '';
  configObj._rev = Number(configObj._rev) || 0;

  const subjectsData = cleanData('Subjects').map(row => row.SubjectName);

  const result = {
    config: configObj,
    classes: cleanData('Classes'),
    students: cleanData('Students'),
    attendance: cleanData('Attendance'),
    grades: cleanData('Grades'),
    koku: cleanData('Koku'),
    ekskul: cleanData('Ekskul'),
    journal: cleanData('Journal'),
    schedule: cleanData('Schedule'),
    subjects: subjectsData,
    poin: cleanData('Poin'),
    tpBank: cleanData('TPBank'),
    raporNotes: cleanData('RaporNotes')
  };
  // Koleksi tambahan dari fitur baru (sheet berawalan "X_"), mis. X_nilaiAkhir -> result.nilaiAkhir
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sh => {
    const name = sh.getName();
    if (name.indexOf(EXTRA_PREFIX) === 0) result[name.slice(EXTRA_PREFIX.length)] = cleanData(name);
  });
  return result;
}

/* =====================================================================
 * KOLEKSI TAMBAHAN (agar fitur baru di frontend tidak perlu update Code.gs lagi)
 * Setiap array data di luar daftar inti disimpan otomatis ke sheet "X_<nama>".
 * ===================================================================== */
const EXTRA_PREFIX = 'X_';
const CORE_COLLECTIONS = ['config', 'classes', 'students', 'attendance', 'grades', 'koku', 'ekskul',
  'journal', 'schedule', 'subjects', 'poin', 'tpBank', 'raporNotes'];

function writeExtraCollections_(ss, db) {
  Object.keys(db).forEach(key => {
    if (CORE_COLLECTIONS.indexOf(key) >= 0) return;
    if (!/^[A-Za-z][A-Za-z0-9]{0,40}$/.test(key) || !Array.isArray(db[key])) return;
    const rows = db[key].filter(r => r && typeof r === 'object' && !Array.isArray(r));
    const headers = [];
    rows.forEach(r => Object.keys(r).forEach(h => { if (headers.indexOf(h) < 0) headers.push(h); }));
    const name = EXTRA_PREFIX + key;
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clear(); // kolom bisa berubah antar versi, jadi tulis ulang header + isi
    if (headers.length === 0) return;
    const values = [headers].concat(rows.map(r => headers.map(h => {
      const v = r[h];
      if (v === undefined || v === null) return '';
      return typeof v === 'object' ? JSON.stringify(v) : String(v);
    })));
    const range = sheet.getRange(1, 1, values.length, headers.length);
    range.setNumberFormat('@');
    range.setValues(values);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f4f7f6');
    sheet.setFrozenRows(1);
  });
}

function writeSheet_(ss, sheetName, headers, dataArray) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f4f7f6');
    sheet.setFrozenRows(1);
  } else {
    // Pastikan header selalu mencakup seluruh kolom terbaru jika ada penambahan kolom
    const lastCol = sheet.getLastColumn();
    if (lastCol < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f4f7f6');
      sheet.setFrozenRows(1);
    }
  }

  // Hapus data lama, pertahankan header
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), headers.length)).clearContent();
  }

  if (dataArray && dataArray.length > 0) {
    const rows = dataArray.map(obj => headers.map(h => {
      let val = obj[h];
      if (val === undefined || val === null) return '';
      // Jika nilai adalah array/objek (seperti data nilai/absen), ubah ke string
      if (typeof val === 'object') return JSON.stringify(val);
      return String(val);
    }));
    const range = sheet.getRange(2, 1, rows.length, headers.length);
    // PENTING: format sebagai Teks SEBELUM menulis nilai. Ini mencegah Google Sheets
    // otomatis mengonversi string seperti "07" menjadi angka 7, atau "2026-07-11"
    // menjadi objek Date — akar penyebab data "hilang" dari aplikasi meski ada di spreadsheet.
    range.setNumberFormat('@');
    range.setValues(rows);
  }
}

/**
 * Menyimpan seluruh data JSON dari frontend ke masing-masing Sheet.
 * Dibungkus LockService agar dua proses simpan tidak berjalan bersamaan dan saling
 * menimpa/merusak data (mis. dua tab/perangkat menyimpan pada waktu yang hampir sama).
 */
function saveAllData(dbString, token, baseRev, force) {
  if (!isValidToken_(token)) return { ok: false, authError: true };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // tunggu maks 30 detik jika ada proses simpan lain yang sedang berjalan
  } catch (e) {
    throw new Error('Server sedang sibuk menyimpan data lain, silakan coba lagi beberapa saat.');
  }

  try {
    const db = JSON.parse(dbString);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // PENTING untuk keamanan multi-perangkat (HP/Tablet/Laptop dengan akun sama):
    // nomor revisi (_rev) TIDAK dipercaya dari client, melainkan SELALU dibaca ulang dan
    // dinaikkan di sini, di server, sehingga menjadi satu penghitung tunggal yang sama-sama
    // dipakai oleh semua perangkat — tidak peduli perangkat mana yang terakhir menyimpan.
    const currentConfig = {};
    cleanData('Config').forEach(row => { currentConfig[row.Key] = row.Value; });
    const currentRev = Number(currentConfig._rev) || 0;

    // CEGAH DATA TERTIMPA: perangkat hanya boleh menyimpan jika data yang dia ubah
    // berasal dari revisi terbaru di server. Jika perangkat lain sudah menyimpan lebih dulu,
    // tolak dan minta perangkat ini memuat data terbaru (kecuali user memilih menimpa).
    if (!force && Number(baseRev) !== currentRev) {
      return { ok: false, conflict: true, rev: currentRev };
    }

    const newRev = currentRev + 1;
    db.config = db.config || {};
    db.config._rev = newRev;
    delete db.config.registeredPassword;
    db.config.registeredEmail = props_().getProperty('AUTH_EMAIL') || db.config.registeredEmail || '';

    const configArr = [];
    for (let key in db.config) {
      configArr.push({Key: key, Value: db.config[key]});
    }

    writeSheet_(ss, 'Config', ['Key', 'Value'], configArr);
    writeSheet_(ss, 'Classes', ['ClassID', 'ClassName'], db.classes);
    writeSheet_(ss, 'Students', ['StudentID', 'ClassID', 'StudentName', 'NIS', 'NISN', 'TempatLahir', 'TanggalLahir', 'JenisKelamin', 'Agama', 'Alamat', 'NamaAyah', 'NamaIbu', 'NamaWali', 'ParentPhone'], db.students);
    writeSheet_(ss, 'Attendance', ATTENDANCE_HEADERS, db.attendance);
    writeSheet_(ss, 'Grades', GRADES_HEADERS, db.grades);
    writeSheet_(ss, 'Koku', ['RecordID', 'ClassID', 'Tanggal', 'Jenis', 'Projek', 'Deskripsi', 'Dimensi', 'KokuData'], db.koku);
    writeSheet_(ss, 'Ekskul', ['EkskulID', 'ClassID', 'StudentID', 'StudentName', 'Semester', 'TahunPelajaran', 'NamaEkskul', 'Predikat', 'Keterangan'], db.ekskul || []);
    writeSheet_(ss, 'Journal', ['RecordID', 'Date', 'ClassID', 'Subject', 'Content', 'AttachmentUrl'], db.journal);
    writeSheet_(ss, 'Schedule', ['RecordID', 'Day', 'TimeStart', 'TimeEnd', 'Subject', 'ClassID'], db.schedule);
    writeSheet_(ss, 'Poin', ['RecordID', 'StudentID', 'ClassID', 'Tanggal', 'PoinReward', 'PoinPelanggaran', 'Keterangan'], db.poin);
    writeSheet_(ss, 'RaporNotes', ['RecordID', 'StudentID', 'ClassID', 'Semester', 'TahunPelajaran', 'CatatanWali', 'Keputusan'], db.raporNotes || []);

    const subjArr = (db.subjects || []).map(s => ({SubjectName: s}));
    writeSheet_(ss, 'Subjects', ['SubjectName'], subjArr);
    writeSheet_(ss, 'TPBank', ['TPID', 'Subject', 'GradeLevel', 'Code', 'Desc'], db.tpBank || []);
    writeExtraCollections_(ss, db);

    return { ok: true, rev: newRev };
  } finally {
    lock.releaseLock();
  }
}

/**
 * JALANKAN SEKALI SAJA secara manual dari editor Apps Script (pilih fungsi ini lalu klik Run)
 * untuk memperbaiki data yang SUDAH TERLANJUR rusak akibat bug lama:
 *  - Kolom MonthYear di Attendance/Grades yang berubah dari "07" menjadi angka 7 (kehilangan
 *    angka nol di depan), sehingga data absensi/nilai bulan Jan-Sep tidak muncul di aplikasi.
 *  - Kolom Date di Journal yang berubah dari teks "2026-07-11" menjadi objek Date.
 * Aman dijalankan berkali-kali (idempotent) — tidak akan merusak data yang sudah benar.
 *
 * SETELAH menjalankan ini, buka aplikasi lalu klik "Segarkan Data dari Spreadsheet"
 * (di menu profil) di SETIAP perangkat/browser yang pernah dipakai, supaya cache lama di
 * perangkat tersebut tidak menimpa balik data yang baru saja diperbaiki ini.
 */
function repairLegacyDataTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function fixMonthYearColumn(sheetName, colIndex) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const range = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1);
    const values = range.getValues();
    const fixed = values.map(row => {
      let v = row[0];
      if (v === '' || v === null || v === undefined) return [v];
      let num = Number(v);
      if (!isNaN(num) && String(v).trim() !== '') return [String(num).padStart(2, '0')];
      return [v];
    });
    range.setNumberFormat('@');
    range.setValues(fixed);
  }

  function fixDateColumn(sheetName, colIndex) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const range = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1);
    const values = range.getValues();
    const fixed = values.map(row => {
      let v = row[0];
      if (v instanceof Date) return [Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')];
      return [v];
    });
    range.setNumberFormat('@');
    range.setValues(fixed);
  }

  fixMonthYearColumn('Attendance', 3); // kolom C = MonthYear
  fixMonthYearColumn('Grades', 5);     // kolom E = MonthYear
  fixDateColumn('Journal', 2);         // kolom B = Date

  SpreadsheetApp.getUi().alert('Perbaikan data selesai. Sekarang buka aplikasi lalu klik "Segarkan Data dari Spreadsheet" di menu profil pada setiap perangkat yang pernah dipakai.');
}

/**
 * JALANKAN SEKALI SAJA secara manual dari editor Apps Script untuk memperbaiki fitur
 * Nilai Kokurikuler: sheet 'Koku' versi lama hanya punya kolom [Siswa, Kategori, Narasi]
 * tanpa ClassID sama sekali, sehingga aplikasi tidak pernah bisa menampilkan ulang nilai
 * yang sudah diinput per kelas. Fungsi ini mengubah skema sheet 'Koku' menjadi
 * [RecordID, ClassID, Tanggal, Jenis, Projek, Deskripsi, Dimensi, KokuData], memetakan
 * data lama ke kelasnya masing-masing (berdasarkan nama siswa yang cocok dengan sheet
 * Students saat ini), dan mencadangkan data lama apa adanya ke sheet 'Koku_Backup_Lama'
 * sebelum ditimpa. Aman dijalankan berkali-kali (idempotent) — kalau skema sudah baru,
 * fungsi ini tidak melakukan apa-apa.
 *
 * SETELAH menjalankan ini, buka aplikasi lalu klik "Segarkan Data dari Spreadsheet"
 * (di menu profil) di setiap perangkat yang pernah dipakai.
 */
function migrateKokuSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Koku');
  if (!sheet) { setupDatabase(); return; }

  const newHeaders = ['RecordID', 'ClassID', 'Tanggal', 'Jenis', 'Projek', 'Deskripsi', 'Dimensi', 'KokuData'];
  const currentHeaders = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];

  if (currentHeaders.length === newHeaders.length && currentHeaders.every((h, i) => h === newHeaders[i])) {
    SpreadsheetApp.getUi().alert('Sheet Koku sudah memakai skema baru. Tidak ada yang perlu dimigrasikan.');
    return;
  }

  // Baca data lama (format lama: Siswa, Kategori, Narasi) SEBELUM header ditimpa
  const oldRows = cleanData('Koku');

  // Cadangkan data lama apa adanya, untuk jaga-jaga
  if (oldRows.length > 0) {
    let backupSheet = ss.getSheetByName('Koku_Backup_Lama');
    if (!backupSheet) backupSheet = ss.insertSheet('Koku_Backup_Lama');
    backupSheet.clearContents();
    const oldHeaders = Object.keys(oldRows[0]);
    backupSheet.appendRow(oldHeaders);
    backupSheet.getRange(2, 1, oldRows.length, oldHeaders.length)
      .setValues(oldRows.map(r => oldHeaders.map(h => r[h])));
  }

  // Petakan nama siswa -> ClassID memakai data Siswa saat ini, lalu kelompokkan nilai lama per kelas
  const students = cleanData('Students');
  const nameToStudent = {};
  students.forEach(s => { nameToStudent[s.StudentName] = s; });

  const grouped = {}; // ClassID -> { StudentID: {Kategori, Narasi} }
  let skipped = 0;
  oldRows.forEach(row => {
    const student = nameToStudent[row.Siswa];
    if (!student) { skipped++; return; }
    if (!grouped[student.ClassID]) grouped[student.ClassID] = {};
    grouped[student.ClassID][student.StudentID] = { Kategori: row.Kategori, Narasi: row.Narasi };
  });

  const newRows = Object.keys(grouped).map((cid, i) => ({
    RecordID: 'K' + Date.now() + i,
    ClassID: cid, Tanggal: '', Jenis: 'LintasMapel',
    Projek: 'Data Lama (Migrasi Otomatis)', Deskripsi: '', Dimensi: '[]',
    KokuData: JSON.stringify(grouped[cid])
  }));

  sheet.clear();
  sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold').setBackground('#f4f7f6');
  sheet.setFrozenRows(1);
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) sheet.getRange(2, 1, maxRows - 1, newHeaders.length).setNumberFormat('@');

  if (newRows.length > 0) {
    const rows = newRows.map(r => newHeaders.map(h => r[h]));
    sheet.getRange(2, 1, rows.length, newHeaders.length).setValues(rows);
  }

  SpreadsheetApp.getUi().alert(
    'Migrasi sheet Koku selesai.\n' +
    newRows.length + ' kelas berhasil dimigrasikan ke format baru.\n' +
    (skipped > 0 ? skipped + ' data lama dilewati karena nama siswanya tidak ditemukan di data Siswa saat ini (sudah dicadangkan di sheet "Koku_Backup_Lama").' : 'Semua data berhasil dimigrasikan.')
  );
}

/**
 * JALANKAN SEKALI dari editor Apps Script untuk menutup akses publik pada folder bukti lama
 * beserta file di dalamnya (yang dulu dibagikan "siapa saja yang memiliki link").
 */
function jadikanFolderBuktiPrivat() {
  var folders = DriveApp.getFoldersByName("AgendaGuru_Bukti");
  var count = 0;
  while (folders.hasNext()) {
    var folder = folders.next();
    folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      if (f.getSharingAccess() !== DriveApp.Access.PRIVATE) {
        f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
        count++;
      }
    }
  }
  console.log('Selesai. ' + count + ' file bukti diubah menjadi privat.');
}

/**
 * --- FITUR BARU: UPLOAD KE GOOGLE DRIVE ---
 */
function uploadFileToDrive(base64Data, filename, mimeType, token) {
  if (!isValidToken_(token)) return "";
  try {
    var folder;
    var folders = DriveApp.getFoldersByName("AgendaGuru_Bukti");
    if (folders.hasNext()) { folder = folders.next(); }
    else { folder = DriveApp.createFolder("AgendaGuru_Bukti"); }
    // Folder bukti berisi foto siswa: harus PRIVAT (hanya pemilik akun).
    // Ini juga menutup akses publik pada folder lama yang dulu dibuat "siapa saja yang punya link".
    try {
      if (folder.getSharingAccess() !== DriveApp.Access.PRIVATE) {
        folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      }
    } catch (shareErr) { /* mis. folder di Shared Drive: biarkan, upload tetap jalan */ }

    var base64 = base64Data.split(',')[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, filename);
    var file = folder.createFile(blob);
    return file.getUrl();
  } catch (e) {
    return "";
  }
}

/**
 * --- FITUR BARU: SINKRONISASI KE GOOGLE CALENDAR ---
 */
function syncToCalendar(type, data, token) {
  if (!isValidToken_(token)) return false;
  try {
    var cal = CalendarApp.getDefaultCalendar();

    if(type === 'jurnal') {
      // Buat event 1 hari penuh untuk Jurnal Harian
      var title = "Tugas Mengajar: " + data.subject + " (" + data.className + ")";
      cal.createAllDayEvent(title, new Date(data.date), {description: data.content});
      return true;

    } else if (type === 'schedule') {
      // Buat event rutinan mingguan (recurring)
      var daysMap = { 'Minggu': CalendarApp.Weekday.SUNDAY, 'Senin': CalendarApp.Weekday.MONDAY, 'Selasa': CalendarApp.Weekday.TUESDAY, 'Rabu': CalendarApp.Weekday.WEDNESDAY, 'Kamis': CalendarApp.Weekday.THURSDAY, 'Jumat': CalendarApp.Weekday.FRIDAY, 'Sabtu': CalendarApp.Weekday.SATURDAY };
      var targetDayObj = {'Minggu': 0, 'Senin': 1, 'Selasa': 2, 'Rabu': 3, 'Kamis': 4, 'Jumat': 5, 'Sabtu': 6};
      var targetDay = targetDayObj[data.day];

      // Cari tanggal terdekat dari hari yang dipilih
      var d = new Date();
      d.setDate(d.getDate() + (targetDay + 7 - d.getDay()) % 7);

      var startT = data.start.split(':');
      var endT = data.end.split(':');
      var startDateTime = new Date(d.getFullYear(), d.getMonth(), d.getDate(), startT[0], startT[1]);
      var endDateTime = new Date(d.getFullYear(), d.getMonth(), d.getDate(), endT[0], endT[1]);

      var recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(daysMap[data.day]);
      var titleSch = "Jadwal: " + data.subject + " (" + data.className + ")";
      cal.createEventSeries(titleSch, startDateTime, endDateTime, recurrence);
      return true;
    }
  } catch (e) {
    return false;
  }
}
