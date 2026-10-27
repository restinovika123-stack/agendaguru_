/* AgendaGuru — penghubung frontend (GitHub Pages) ke backend Apps Script milik tiap guru.
 *
 * File ini membuat tiruan `google.script.run` yang mengirim panggilan lewat fetch ke
 * URL Aplikasi Web Apps Script. Dengan begitu kode aplikasi (app.js) tetap sama seperti
 * saat dijalankan di dalam Apps Script.
 *
 * Alamat Apps Script disimpan di perangkat. Cara mengisinya:
 *  - otomatis: buka URL Apps Script -> klik "Buka AgendaGuru" (membawa ?api=...)
 *  - manual: tempel alamat di layar "Hubungkan ke Apps Script".
 */
(function () {
    'use strict';
    var API_KEY = 'agendaGuru_apiUrl';
    var MIN_API_VERSION = 3; // versi Code.gs minimum (v2: Pengolah Nilai, v3: lisensi)
    // Data milik koneksi lama yang harus dibersihkan saat pindah ke Apps Script lain
    var PER_ACCOUNT_KEYS = ['agendaGuru_token', 'agendaGuru_unsynced', 'agendaGuru_isLogged', 'AGENDA_GURU_MASTER_DB'];

    function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
    function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) { } }
    function lsDel(k) { try { localStorage.removeItem(k); } catch (e) { } }

    function normalizeUrl(u) { return String(u || '').trim().replace(/\?.*$/, ''); }
    function isValidApiUrl(u) {
        return /^https:\/\/script\.google\.com\/(a\/macros\/[^\/]+|macros)\/s\/[A-Za-z0-9_-]+\/exec$/.test(u) ||
            /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?\/[^?#]*$/.test(u); // untuk pengujian lokal
    }

    function clearAccountData() {
        PER_ACCOUNT_KEYS.forEach(lsDel);
        try { if (window.indexedDB) indexedDB.deleteDatabase('localforage'); } catch (e) { }
    }

    function saveApiUrl(u) {
        var old = lsGet(API_KEY);
        if (old !== u) clearAccountData(); // termasuk data demo, agar tidak tercampur
        lsSet(API_KEY, u);
    }

    // 1. Ambil ?api=... dari alamat (datang dari halaman pembuka Apps Script), lalu hapus dari URL
    try {
        var params = new URLSearchParams(location.search);
        var fromQuery = normalizeUrl(params.get('api'));
        if (fromQuery && isValidApiUrl(fromQuery)) saveApiUrl(fromQuery);
        if (params.has('api')) {
            params.delete('api');
            var rest = params.toString();
            history.replaceState(null, '', location.pathname + (rest ? '?' + rest : '') + location.hash);
        }
    } catch (e) { }

    var apiUrl = lsGet(API_KEY);
    window.AGENDA_API_URL = apiUrl || '';

    function rawCall(url, name, args) {
        return fetch(url, {
            method: 'POST',
            // text/plain = "simple request", tanpa preflight CORS (Apps Script tidak melayani OPTIONS)
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ fn: name, args: args }),
            redirect: 'follow'
        }).then(function (r) {
            if (!r.ok) throw new Error('Server membalas HTTP ' + r.status);
            return r.text();
        }).then(function (t) {
            var j;
            try { j = JSON.parse(t); } catch (e) {
                throw new Error('Balasan server tidak valid. Pastikan Aplikasi Web di-deploy dengan akses "Siapa saja" dan alamatnya benar.');
            }
            if (!j.ok) throw new Error(j.error || 'Terjadi kesalahan di server.');
            return j.result;
        });
    }

    // 2. Tiruan google.script.run
    if (apiUrl) {
        var makeRunner = function (ok, fail) {
            return new Proxy({}, {
                get: function (_, name) {
                    if (name === 'withSuccessHandler') return function (f) { return makeRunner(f, fail); };
                    if (name === 'withFailureHandler') return function (f) { return makeRunner(ok, f); };
                    if (name === 'withUserObject') return function () { return makeRunner(ok, fail); };
                    if (typeof name !== 'string' || name === 'then') return undefined;
                    return function () {
                        var args = Array.prototype.slice.call(arguments);
                        rawCall(apiUrl, name, args).then(function (res) {
                            // Panggil handler di luar rantai promise agar error di handler tidak dianggap gagal kirim
                            setTimeout(function () { if (ok) ok(res); }, 0);
                        }, function (err) {
                            setTimeout(function () { if (fail) fail(err); else console.error(err); }, 0);
                        });
                    };
                }
            });
        };
        window.google = { script: { get run() { return makeRunner(null, null); } } };

        // Cek versi Code.gs di latar belakang
        rawCall(apiUrl, 'getServerInfo', []).then(function (info) {
            window.AGENDA_SERVER_INFO = info || null; // dipakai fitur yang butuh Code.gs versi tertentu
            if (info && Number(info.apiVersion || 0) < MIN_API_VERSION) {
                showBanner('Code.gs di Spreadsheet Anda versi lama. Salin Code.gs terbaru ke Apps Script lalu buat deployment versi baru.');
            }
        }, function () { /* offline: diabaikan, aplikasi memakai cache */ });
    }

    function onReady(fn) {
        if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', fn); else fn();
    }

    function showBanner(text) {
        onReady(function () {
            var b = document.createElement('div');
            b.setAttribute('role', 'alert');
            b.style.cssText = 'position:fixed;left:0;right:0;bottom:0;z-index:3000;background:#B7791F;color:#fff;padding:10px 16px;font:600 14px system-ui,sans-serif;text-align:center';
            b.textContent = text;
            document.body.appendChild(b);
        });
    }

    // 3. Layar "Hubungkan ke Apps Script"
    function showConnect(opts) {
        opts = opts || {};
        onReady(function () {
            var old = document.getElementById('ag-connect'); if (old) old.remove();
            var wrap = document.createElement('div');
            wrap.id = 'ag-connect';
            wrap.style.cssText = 'position:fixed;inset:0;z-index:4000;background:rgba(23,43,77,.96);display:flex;align-items:center;justify-content:center;padding:16px;overflow:auto';
            wrap.innerHTML =
                '<div style="background:#fff;border-radius:16px;max-width:480px;width:100%;padding:28px;font-family:system-ui,sans-serif;color:#172B4D">' +
                '<h4 style="font-weight:800;margin:0 0 6px">Hubungkan ke Apps Script</h4>' +
                '<p style="color:#6b778c;font-size:14px;margin:0 0 16px">Data Anda disimpan di Google Spreadsheet milik Anda sendiri. ' +
                'Tempel alamat <b>Aplikasi Web</b> dari Apps Script Anda (berakhiran <code>/exec</code>).</p>' +
                '<label for="ag-api-input" style="font-weight:700;font-size:13px">Alamat Aplikasi Web</label>' +
                '<input id="ag-api-input" type="url" class="form-control" placeholder="https://script.google.com/macros/s/.../exec" style="margin:6px 0 6px">' +
                '<div id="ag-api-msg" style="font-size:13px;min-height:20px;color:#C44747"></div>' +
                '<button id="ag-api-btn" class="btn btn-primary w-100" style="margin-top:6px">Hubungkan</button>' +
                (opts.allowDemo !== false ? '<button id="ag-demo-btn" class="btn btn-link w-100" style="margin-top:4px">Coba mode demo (data contoh, tidak tersimpan)</button>' : '') +
                (opts.allowClose ? '<button id="ag-close-btn" class="btn btn-link w-100 text-secondary">Batal</button>' : '') +
                '<p style="color:#6b778c;font-size:13px;margin:12px 0 0">Belum punya alamat Aplikasi Web? <a href="panduan.html" target="_blank" rel="noopener" style="color:#159A9C;font-weight:700">Lihat Panduan Pemasangan (5–10 menit)</a></p>' +
                '</div>';
            document.body.appendChild(wrap);
            var input = document.getElementById('ag-api-input');
            var msg = document.getElementById('ag-api-msg');
            var btn = document.getElementById('ag-api-btn');
            input.value = lsGet(API_KEY) || '';
            btn.onclick = function () {
                var u = normalizeUrl(input.value);
                if (!isValidApiUrl(u)) { msg.textContent = 'Alamat tidak valid. Contoh: https://script.google.com/macros/s/XXXX/exec'; return; }
                msg.style.color = '#6b778c'; msg.textContent = 'Menguji koneksi...'; btn.disabled = true;
                rawCall(u, 'getServerInfo', []).then(function (info) {
                    if (!info || info.app !== 'AgendaGuru') throw new Error('Alamat ini bukan backend AgendaGuru.');
                    saveApiUrl(u);
                    location.reload();
                }).catch(function (err) {
                    btn.disabled = false; msg.style.color = '#C44747';
                    msg.textContent = 'Gagal terhubung: ' + (err && err.message ? err.message : err);
                });
            };
            var demo = document.getElementById('ag-demo-btn'); if (demo) demo.onclick = function () { wrap.remove(); };
            var close = document.getElementById('ag-close-btn'); if (close) close.onclick = function () { wrap.remove(); };
        });
    }
    window.agendaGuruShowConnect = function () { showConnect({ allowClose: true, allowDemo: false }); };
    // Tampilkan tautan "Ganti alamat Apps Script" (disembunyikan jika api.js tidak dipakai)
    onReady(function () {
        document.querySelectorAll('.ag-connect-link').forEach(function (el) { el.style.display = ''; });
    });

    if (!apiUrl) showConnect({ allowDemo: true });
})();
