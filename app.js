window.AGENDA_GURU_LOADED = true;
    // --- Lazy-load library besar (xlsx, jsPDF) hanya saat dibutuhkan ---
    // PENTING: library ini SENGAJA tidak dimuat otomatis lewat <script src> di head lagi.
    // Salah satu dari library besar ini memicu pelanggaran Content Security Policy sandbox
    // Google Apps Script (CSP memblokir eval()), yang membuat SELURUH script aplikasi ini
    // gagal berjalan sama sekali (halaman jadi blank putih). Dengan dimuat belakangan
    // (baru saat tombol Export Excel/PDF benar-benar diklik), aplikasi utama tetap jalan
    // normal terlepas dari masalah itu.
    const _loadedScripts = {};
    function loadScriptOnce(url) {
        if (_loadedScripts[url]) return _loadedScripts[url];
        _loadedScripts[url] = new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = url;
            s.onload = () => resolve();
            s.onerror = () => { delete _loadedScripts[url]; reject(new Error('Gagal memuat ' + url)); };
            document.head.appendChild(s);
        });
        return _loadedScripts[url];
    }
    let _xlsxLoaded = false, _pdfLoaded = false;
    async function ensureXLSX() {
        if (!_xlsxLoaded) { await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'); _xlsxLoaded = true; }
    }
    async function ensurePDF() {
        if (!_pdfLoaded) {
            await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
            await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js');
            _pdfLoaded = true;
        }
    }

    const MASTER_DB_ID = 'AGENDA_GURU_MASTER_DB';

    // Dummy/Fallback Data for Simulation
    let db = { classes: [], students: [], attendance: [], grades: [], koku: [], journal: [], schedule: [], poin: [], subjects: [], tpBank: [], config: {} };
    let chartHadir, chartPrestasi, chartAbsensiObj;
    let editingScheduleId = null;
    let editingJournalId = null;
    let editingKokuId = null;
    let editingPoinId = null;

    // Tanggal hari ini menurut jam perangkat (bukan UTC). toISOString() memberi tanggal
    // kemarin jika dipakai sebelum pukul 07.00 WIB.
    function tanggalHariIni() {
        const d = new Date();
        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }

    function toggleSidebar() {
        document.getElementById('sidebar-menu').classList.toggle('show');
        document.getElementById('sidebar-overlay').classList.toggle('show');
    }

    // --- Mode & Sesi Login ---
    // IS_GAS: aplikasi dibuka lewat Google Apps Script (ada server). Jika tidak, berjalan
    // sebagai mode simulasi/demo lokal di browser.
    const IS_GAS = (typeof google !== 'undefined' && !!google.script && !!google.script.run);
    const TOKEN_KEY = 'agendaGuru_token';      // token sesi dari server (bukan password)
    const DIRTY_KEY = 'agendaGuru_unsynced';   // '1' jika ada perubahan yang belum diterima server
    function getToken() { try { return localStorage.getItem(TOKEN_KEY) || ''; } catch (e) { return ''; } }
    function setToken(t) { try { if (t) localStorage.setItem(TOKEN_KEY, t); else localStorage.removeItem(TOKEN_KEY); } catch (e) { } }
    function setDirty(v) { try { if (v) localStorage.setItem(DIRTY_KEY, '1'); else localStorage.removeItem(DIRTY_KEY); } catch (e) { } }
    function isDirty() { try { return localStorage.getItem(DIRTY_KEY) === '1'; } catch (e) { return false; } }

    // --- Tahun Pelajaran (TP) ---
    // Absensi & nilai disimpan per bulan; tanpa tahun pelajaran, bulan Juli tahun ini dan
    // Juli tahun depan akan saling menimpa. Setiap record kini membawa TahunPelajaran.
    function getDefaultTapel(date) {
        const d = date || new Date(); const y = d.getFullYear();
        return (d.getMonth() + 1) >= 7 ? `${y}/${y + 1}` : `${y - 1}/${y}`;
    }
    function getActiveTapel() {
        let t = ''; try { t = localStorage.getItem('agendaGuru_activeTapel') || ''; } catch (e) { }
        return /^\d{4}\/\d{4}$/.test(t) ? t : getDefaultTapel();
    }
    function setActiveTapel(t) {
        try { localStorage.setItem('agendaGuru_activeTapel', t); } catch (e) { }
        const sel = document.getElementById('global-tapel');
        if (sel && sel.value !== t) { if (![...sel.options].some(o => o.value === t)) populateTapelSelect(); sel.value = t; }
        const raporTapel = document.getElementById('rapor-tapel');
        if (raporTapel) raporTapel.value = t;
        const ekskulTapel = document.getElementById('ekskul-tapel');
        if (ekskulTapel) ekskulTapel.value = t;
        updateTapelSelectStyle();
        const active = document.querySelector('.view-section.active');
        if (active) switchView(active.id.replace('view-', ''));
    }
    // Pilihan TP diberi warna peringatan jika yang dibuka bukan TP berjalan,
    // supaya guru sadar sedang mengisi data untuk tahun pelajaran lain.
    function updateTapelSelectStyle() {
        const el = document.getElementById('global-tapel'); if (!el) return;
        const current = getDefaultTapel();
        const other = getActiveTapel() !== current;
        ['border-warning', 'bg-warning-subtle', 'text-warning-emphasis'].forEach(c => el.classList.toggle(c, other));
        el.title = other
            ? `Perhatian: Anda sedang membuka TP ${getActiveTapel()}, bukan TP berjalan (${current}).`
            : 'Tahun Pelajaran aktif untuk absensi, nilai, dan rapor';
    }

    // --- Peringatan tanggal vs TP ---
    // Saat guru mengisi tanggal di kolom absensi/nilai yang termasuk TP lain dari TP yang
    // sedang dibuka, tampilkan peringatan dan tawarkan memindahkan data ke TP yang benar.
    const tapelWarnDismissed = {};
    function findTableRecordInTapel(kind, tapel) {
        const same = r => String(r.TahunPelajaran || '').replace(/\s+/g, '') === String(tapel).replace(/\s+/g, '');
        if (kind === 'absen') {
            const cid = document.getElementById('absen-kelas').value, month = document.getElementById('absen-bulan').value;
            return db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(month) && same(a));
        }
        const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value,
            jenis = document.getElementById('nilai-jenis').value, bulan = document.getElementById('nilai-bulan').value;
        if (jenis !== 'Formatif') return null;
        return db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) &&
            String(g.Type) === 'Formatif' && String(g.MonthYear) === String(bulan) && same(g));
    }
    function moveTableToTapel(kind, from, to) {
        const rec = findTableRecordInTapel(kind, from);
        if (!rec) return;
        if (findTableRecordInTapel(kind, to)) {
            return Swal.fire('Tidak Bisa Dipindahkan', `TP ${to} sudah memiliki data untuk kelas dan bulan yang sama. Data di TP ${from} tidak diubah. Buka TP ${to} untuk memeriksanya.`, 'info');
        }
        rec.TahunPelajaran = to;
        setActiveTapel(to);
        syncDB();
        Swal.fire({ icon: 'success', title: 'Dipindahkan', text: `Data dipindahkan ke TP ${to}.`, timer: 1800, showConfirmButton: false });
    }
    function warnTapelMismatch(kind, dateStr) {
        const dateTapel = inferTapelFromDates([dateStr]);
        const active = getActiveTapel();
        if (!dateTapel || dateTapel === active || tapelWarnDismissed[active + '>' + dateTapel]) return;
        const tgl = typeof formatTanggalIndo === 'function' ? formatTanggalIndo(dateStr) : dateStr;
        Swal.fire({
            icon: 'warning',
            title: 'Tahun Pelajaran tidak cocok',
            html: `Tanggal <b>${escapeHtml(tgl)}</b> termasuk <b>TP ${dateTapel}</b>, tetapi Anda sedang membuka <b>TP ${active}</b>.<br><br>` +
                `Jika tetap di TP ${active}, data ini tidak akan muncul di rapor TP ${dateTapel}.`,
            showCancelButton: true,
            confirmButtonText: `Pindahkan ke TP ${dateTapel}`,
            cancelButtonText: `Tetap di TP ${active}`
        }).then(r => {
            if (r.isConfirmed) moveTableToTapel(kind, active, dateTapel);
            else tapelWarnDismissed[active + '>' + dateTapel] = true; // jangan tanya lagi di sesi ini
        });
    }
    // Didaftarkan di document: berjalan SETELAH onchange kolom (yang menyimpan datanya lebih dulu)
    document.addEventListener('change', e => {
        const el = e.target;
        if (!el || !el.classList || !el.value) return;
        if (el.classList.contains('date-absen-col')) warnTapelMismatch('absen', el.value);
        else if (el.classList.contains('date-nilai-col')) warnTapelMismatch('nilai', el.value);
    });
    function inferTapelFromDates(datesData) {
        let dates = [];
        try { dates = typeof datesData === 'string' ? JSON.parse(datesData) : (datesData || []); } catch (e) { }
        const first = (Array.isArray(dates) ? dates : []).find(d => /^\d{4}-\d{2}-\d{2}$/.test(String(d)));
        if (!first) return '';
        const parts = String(first).split('-').map(Number);
        return parts[1] >= 7 ? `${parts[0]}/${parts[0] + 1}` : `${parts[0] - 1}/${parts[0]}`;
    }
    // Record tanpa TahunPelajaran (data lama yang belum dimigrasi) dianggap cocok dengan TP mana pun.
    // tapel kosong = tanpa filter.
    function matchTapel(rec, tapel) {
        if (!tapel) return true;
        const norm = v => String(v).replace(/\s+/g, '');
        const rt = rec && rec.TahunPelajaran ? norm(rec.TahunPelajaran) : '';
        return !rt || rt === norm(tapel);
    }
    // Beri TahunPelajaran pada data lama: dari tanggal di kolomnya bila ada; jika tidak ada,
    // pakai TP saat migrasi pertama (disimpan di config agar konsisten di semua perangkat).
    function migrateTapelFields() {
        let changed = false;
        const legacyTapel = () => {
            if (!db.config.legacyTapel) db.config.legacyTapel = getDefaultTapel();
            return db.config.legacyTapel;
        };
        [db.attendance || [], db.grades || []].forEach(list => list.forEach(r => {
            if (!r.TahunPelajaran) { r.TahunPelajaran = inferTapelFromDates(r.DatesData) || legacyTapel(); changed = true; }
        }));
        return changed;
    }
    function populateTapelSelect() {
        const el = document.getElementById('global-tapel'); if (!el) return;
        const active = getActiveTapel();
        const def = Number(getDefaultTapel().split('/')[0]);
        const years = new Set([active, `${def - 1}/${def}`, `${def}/${def + 1}`, `${def + 1}/${def + 2}`]);
        [...(db.attendance || []), ...(db.grades || [])].forEach(r => { if (/^\d{4}\/\d{4}$/.test(String(r.TahunPelajaran || ''))) years.add(String(r.TahunPelajaran)); });
        getCustomTapels().forEach(t => years.add(t));
        const current = getDefaultTapel();
        el.innerHTML = [...years].sort().reverse().map(t => `<option value="${t}" ${t === active ? 'selected' : ''}>TP ${t}${t === current ? ' (berjalan)' : ''}</option>`).join('') +
            '<option value="__add__">+ Tambah TP lain...</option>';
        updateTapelSelectStyle();
    }

    // TP tambahan buatan guru disimpan di config (ikut tersinkron ke semua perangkat)
    function normalizeTapel(v) {
        const m = /^\s*(\d{4})\s*\/\s*(\d{4})\s*$/.exec(String(v || ''));
        return m && Number(m[2]) === Number(m[1]) + 1 ? `${m[1]}/${m[2]}` : '';
    }
    function getCustomTapels() {
        return String((db.config && db.config.customTapels) || '').split(',').map(normalizeTapel).filter(Boolean);
    }
    async function onTapelSelectChange(sel) {
        if (sel.value !== '__add__') return setActiveTapel(sel.value);
        sel.value = getActiveTapel(); // kembalikan pilihan selama dialog terbuka
        const start = Number(getDefaultTapel().split('/')[0]) + 2;
        const r = await Swal.fire({
            title: 'Tambah Tahun Pelajaran',
            input: 'text',
            inputLabel: 'Format: TTTT/TTTT, contoh 2028/2029',
            inputValue: `${start}/${start + 1}`,
            showCancelButton: true, confirmButtonText: 'Tambah & Buka', cancelButtonText: 'Batal',
            inputValidator: v => {
                if (!/^\s*\d{4}\s*\/\s*\d{4}\s*$/.test(v || '')) return 'Gunakan format TTTT/TTTT, contoh 2028/2029';
                if (!normalizeTapel(v)) return 'Tahun kedua harus satu tahun setelah tahun pertama (contoh 2028/2029)';
                const y = Number(String(v).trim().slice(0, 4));
                if (y < 2000 || y > 2100) return 'Tahun di luar rentang yang wajar';
            }
        });
        if (!r.isConfirmed) return;
        const t = normalizeTapel(r.value);
        const list = getCustomTapels();
        if (!list.includes(t)) { list.push(t); db.config.customTapels = list.join(','); syncDB(); }
        try { localStorage.setItem('agendaGuru_activeTapel', t); } catch (e) { }
        populateTapelSelect();
        setActiveTapel(t);
    }

    // --- Core Sync Logic ---
    // Antrean simpan: pengiriman ke Apps Script berurutan (tidak tumpang tindih) dan hanya
    // mengirim versi data terbaru. Server menolak simpanan yang berasal dari data lama
    // (perangkat lain sudah menyimpan lebih dulu) supaya data tidak saling menimpa.
    let syncQueue = Promise.resolve();
    let pendingSyncCount = 0;
    let lastSaveFailed = false;
    // true jika data yang ditampilkan berasal dari cache lokal (perangkat offline).
    let isOfflineMode = false;
    let serverRev = 0;        // revisi server yang menjadi dasar data di perangkat ini
    let localChangeSeq = 0;   // bertambah setiap ada perubahan lokal
    let savedChangeSeq = 0;   // perubahan terakhir yang sudah diterima server
    let conflictPending = false;
    let retryTimer = null;

    function updateSyncIndicator() {
        const el = document.getElementById('sync-indicator');
        if (!el) return;
        el.onclick = null; el.style.cursor = '';
        if (conflictPending) {
            el.style.display = 'inline-flex'; el.style.cursor = 'pointer'; el.onclick = showConflictDialog;
            el.innerHTML = '<i class="bi bi-exclamation-triangle-fill text-danger"></i> <span class="text-danger">Data bentrok, klik untuk memilih</span>';
        } else if (pendingSyncCount > 0) {
            el.style.display = 'inline-flex';
            el.innerHTML = '<i class="bi bi-cloud-arrow-up-fill text-primary"></i> <span class="text-primary">Menyimpan...</span>';
        } else if (lastSaveFailed) {
            el.style.display = 'inline-flex';
            el.innerHTML = '<i class="bi bi-cloud-slash-fill text-danger"></i> <span class="text-danger">Gagal sinkron, mencoba lagi...</span>';
        } else if (isOfflineMode) {
            el.style.display = 'inline-flex';
            el.innerHTML = '<i class="bi bi-wifi-off text-warning"></i> <span class="text-warning">Mode Offline (data cache)</span>';
        } else {
            el.style.display = 'none';
        }
    }

    async function persistLocal() {
        const s = JSON.stringify(db);
        try { if (typeof localforage !== 'undefined' && localforage) await localforage.setItem(MASTER_DB_ID, s); } catch (e) { console.error("Gagal menyimpan ke memori lokal", e); }
        try { localStorage.setItem(MASTER_DB_ID, s); } catch (e) { }
    }

    async function syncDB() {
        db.config._revTime = Date.now();
        localChangeSeq++;
        if (IS_GAS) setDirty(true);
        await persistLocal();
        if (IS_GAS) queueServerSave(false);
    }

    function scheduleRetry() {
        if (retryTimer) return;
        retryTimer = setTimeout(() => { retryTimer = null; queueServerSave(false); }, 8000);
    }

    function queueServerSave(force) {
        if (!IS_GAS) return;
        if (retryTimer) { clearTimeout(retryTimer); retryTimer = null; }
        pendingSyncCount++; updateSyncIndicator();
        syncQueue = syncQueue.then(() => new Promise(resolve => {
            const done = () => { pendingSyncCount = Math.max(0, pendingSyncCount - 1); updateSyncIndicator(); resolve(); };
            // Lewati jika sedang bentrok (menunggu pilihan user) atau versi terbaru sudah terkirim
            if ((conflictPending && !force) || (!force && localChangeSeq === savedChangeSeq)) return done();
            const seq = localChangeSeq;
            db.config._rev = serverRev;
            const payload = JSON.stringify(db);
            google.script.run
                .withSuccessHandler(result => {
                    if (result && result.authError) { done(); handleSessionExpired(); return; }
                    if (result && result.conflict) { conflictPending = true; done(); showConflictDialog(); return; }
                    if (!result || !result.ok) { lastSaveFailed = true; done(); scheduleRetry(); return; }
                    serverRev = Number(result.rev) || serverRev;
                    db.config._rev = serverRev;
                    savedChangeSeq = seq;
                    lastSaveFailed = false; isOfflineMode = false; conflictPending = false;
                    if (savedChangeSeq === localChangeSeq) setDirty(false);
                    persistLocal();
                    done();
                })
                .withFailureHandler(err => {
                    console.error("Gagal menyimpan ke spreadsheet:", err);
                    lastSaveFailed = true; done(); scheduleRetry();
                })
                .saveAllData(payload, getToken(), serverRev, !!force);
        }));
    }

    let conflictDialogOpen = false;
    function showConflictDialog() {
        if (conflictDialogOpen || typeof Swal === 'undefined') return;
        conflictDialogOpen = true;
        Swal.fire({
            icon: 'warning',
            title: 'Data diubah di perangkat lain',
            html: 'Data di Spreadsheet sudah diperbarui dari perangkat atau tab lain.<br><br>Supaya tidak saling menimpa, perubahan dari perangkat ini <b>belum dikirim</b>. Pilih tindakan:',
            showDenyButton: true, showCancelButton: true,
            confirmButtonText: 'Muat data terbaru',
            denyButtonText: 'Timpa dengan data perangkat ini',
            cancelButtonText: 'Unduh cadangan dulu',
            allowOutsideClick: false, allowEscapeKey: false
        }).then(r => {
            conflictDialogOpen = false;
            if (r.isConfirmed) {
                forceRefreshFromServer();
            } else if (r.isDenied) {
                Swal.fire({
                    icon: 'warning', title: 'Yakin menimpa?',
                    text: 'Perubahan yang dibuat di perangkat lain akan hilang dan diganti dengan data dari perangkat ini.',
                    showCancelButton: true, confirmButtonText: 'Ya, timpa', cancelButtonText: 'Batal', confirmButtonColor: '#C44747'
                }).then(c => {
                    if (c.isConfirmed) { conflictPending = false; queueServerSave(true); }
                    else showConflictDialog();
                });
            } else {
                downloadLocalBackup();
                setTimeout(showConflictDialog, 800);
            }
        });
    }

    function downloadLocalBackup() {
        const blob = new Blob([JSON.stringify(db, null, 2)], { type: 'application/json' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'AgendaGuru_cadangan_' + new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-') + '.json';
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(a.href), 3000);
    }

    function handleSessionExpired() {
        setToken('');
        showAuthScreen(true);
        Swal.fire({ icon: 'info', title: 'Sesi berakhir', text: 'Silakan masuk kembali. Perubahan yang belum terkirim tetap tersimpan di perangkat ini dan akan dikirim setelah Anda masuk.' });
    }

    // Peringatkan user jika masih ada data yang belum tersimpan ke server saat menutup halaman.
    window.addEventListener('beforeunload', (e) => {
        if (pendingSyncCount > 0 || conflictPending || (IS_GAS && isDirty())) {
            e.preventDefault();
            e.returnValue = '';
        }
    });

    // Ambil ulang data langsung dari Spreadsheet dan TIMPA data di perangkat ini.
    function forceRefreshFromServer() {
        if (!IS_GAS) { return Swal.fire('Info', 'Fitur ini hanya berjalan saat aplikasi diakses melalui Google Apps Script.', 'info'); }
        Swal.fire({ title: 'Menyegarkan data...', text: 'Mengambil data terbaru dari Spreadsheet.', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
        google.script.run
            .withSuccessHandler(async res => {
                if (res && res.authError) { Swal.close(); handleSessionExpired(); return; }
                if (res && res.config) {
                    delete res._license;
                    db = res;
                    setDirty(false); conflictPending = false;
                    await persistLocal();
                    try { closeModal('modalProfile'); } catch (e) { }
                    Swal.fire('Berhasil', 'Data terbaru dari Spreadsheet berhasil dimuat.', 'success').then(() => location.reload());
                } else {
                    Swal.fire('Gagal', 'Tidak bisa mengambil data dari server.', 'error');
                }
            })
            .withFailureHandler(err => Swal.fire('Gagal', 'Tidak bisa terhubung ke server: ' + err, 'error'))
            .getAllData(getToken());
    }

    async function getLocalFallbackDb() {
        let localDataStr = null;
        try {
            if (typeof localforage !== 'undefined' && localforage) {
                localDataStr = await localforage.getItem(MASTER_DB_ID);
            }
        } catch (e) {
            console.warn("localforage tidak dapat diakses:", e);
        }
        if (!localDataStr) {
            try {
                localDataStr = localStorage.getItem(MASTER_DB_ID);
            } catch (e) {
                console.warn("localStorage tidak dapat diakses:", e);
            }
        }
        try {
            return localDataStr ? JSON.parse(localDataStr) : null;
        } catch (e) {
            return null;
        }
    }

    // Data dari server diterima (setelah buka aplikasi atau login).
    // Jika perangkat ini masih punya perubahan yang belum terkirim, jangan dibuang:
    // - server belum berubah sejak itu  -> pakai data perangkat ini lalu kirim;
    // - server sudah berubah            -> tampilkan pilihan (bentrok).
    async function applyServerData(res) {
        isOfflineMode = false;
        const lic = res ? res._license : null;
        if (res) delete res._license; // bukan bagian dari data guru
        if (!(await licenseGate(lic, false))) return;
        const localDb = isDirty() ? await getLocalFallbackDb() : null;
        if (localDb && localDb.config) {
            processLoadedData(localDb);
            if (Number(localDb.config._rev || 0) === Number((res.config && res.config._rev) || 0)) {
                localChangeSeq++;
                queueServerSave(false);
            } else {
                conflictPending = true;
                updateSyncIndicator();
                showConflictDialog();
            }
            return;
        }
        db = res;
        await persistLocal();
        processLoadedData(res);
    }

    async function useOfflineCache(err) {
        console.error("Gagal mengambil data dari server, memakai cache lokal:", err);
        const localDb = await getLocalFallbackDb();
        if (getToken() && localDb && localDb.config && localDb.config.registeredEmail) {
            if (!(await licenseGate(null, true))) return;
            isOfflineMode = true;
            processLoadedData(localDb);
            updateSyncIndicator();
        } else {
            showAuthScreen(true);
            Swal.fire('Tidak dapat terhubung', 'Periksa koneksi internet, lalu muat ulang halaman.', 'error');
        }
    }

    function showAuthScreen(registered) {
        const ov = document.getElementById('loading-overlay'); if (ov) ov.style.display = 'none';
        document.getElementById('app-view').style.display = 'none';
        document.getElementById('login-view').style.display = 'block';
        document.getElementById('form-login').style.display = registered ? 'block' : 'none';
        document.getElementById('form-register').style.display = registered ? 'none' : 'block';
    }

    // --- Initial Load Logic ---
    // Mode GAS: data hanya diambil dari server dengan token sesi yang valid.
    // Cache lokal hanya dipakai jika perangkat offline (dan pernah login di perangkat ini).
    async function initAgendaGuru() {
        if (window._agendaGuruHasInitialized) return;
        window._agendaGuruHasInitialized = true;

        let initDone = false;
        const finish = fn => { if (initDone) return; initDone = true; fn(); };

        // Safety timeout — jangan biarkan loading overlay menggantung jika ada kendala
        setTimeout(() => {
            let loadingEl = document.getElementById('loading-overlay');
            if (!loadingEl || loadingEl.style.display === 'none') return;
            if (IS_GAS) {
                finish(() => { if (getToken()) useOfflineCache('timeout'); else showAuthScreen(true); });
                return;
            }
            loadingEl.style.display = 'none';
            let isLogged = false;
            try { isLogged = localStorage.getItem('agendaGuru_isLogged') === 'true'; } catch (e) { }
            if (isLogged) {
                showApp();
            } else {
                let lv = document.getElementById('login-view');
                if (lv) lv.style.display = 'block';
            }
        }, IS_GAS ? 20000 : 3000);

        try {
            let topbarDate = document.getElementById('topbar-date');
            if (topbarDate) {
                topbarDate.innerText = new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
            }
        } catch (e) { }

        let loadingOverlay = document.getElementById('loading-overlay');
        if (loadingOverlay) loadingOverlay.style.display = 'flex';

        if (IS_GAS) {
            const token = getToken();
            if (!token) {
                google.script.run
                    .withSuccessHandler(st => finish(() => showAuthScreen(!!(st && st.registered))))
                    .withFailureHandler(err => finish(() => {
                        showAuthScreen(true);
                        Swal.fire({
                            icon: 'error', title: 'Tidak dapat terhubung ke server',
                            html: 'Periksa koneksi internet. Jika masih gagal, pastikan Aplikasi Web Apps Script di-deploy dengan akses <b>Siapa saja</b>.' +
                                (typeof agendaGuruShowConnect === 'function' ? '<br><br><a href="javascript:void(0)" onclick="Swal.close();agendaGuruShowConnect()">Ganti alamat Apps Script</a>' : ''),
                            footer: '<small class="text-muted">' + escapeHtml((err && err.message) || err) + '</small>'
                        });
                    }))
                    .getAuthStatus();
                return;
            }
            google.script.run
                .withSuccessHandler(res => finish(() => {
                    if (res && res.authError) { setToken(''); showAuthScreen(!!res.registered); return; }
                    applyServerData(res);
                }))
                .withFailureHandler(err => finish(() => useOfflineCache(err)))
                .getAllData(token);
        } else {
            // Simulation Mode (Canvas / Browser Local) — tidak ada server GAS untuk dihubungi
            try {
                const localDb = await getLocalFallbackDb();
                if (localDb && localDb.config && localDb.config.registeredEmail) {
                    processLoadedData(localDb);
                } else {
                    const dummyDB = {
                        config: { registeredEmail: "guru@sekolah.id", registeredPassword: "123", userName: "Ahmad Guru, M.Pd.", userSchool: "SDN 1 Dummy" },
                        classes: [{ ClassID: "C1", ClassName: "Kelas 5A" }, { ClassID: "C2", ClassName: "Kelas 5B" }],
                        students: [
                            { StudentID: "S1", ClassID: "C1", StudentName: "Andi Saputra", ParentPhone: "08123456789" },
                            { StudentID: "S2", ClassID: "C1", StudentName: "Budi Santoso", ParentPhone: "08129876543" },
                            { StudentID: "S3", ClassID: "C1", StudentName: "Citra Kirana", ParentPhone: "08120000111" }
                        ],
                        subjects: ["Matematika", "IPA", "Bahasa Indonesia"],
                        attendance: [], grades: [], koku: [], journal: [], schedule: [], poin: []
                    };
                    let dummyStr = JSON.stringify(dummyDB);
                    try { if (typeof localforage !== 'undefined') await localforage.setItem(MASTER_DB_ID, dummyStr); } catch (e) { }
                    try { localStorage.setItem(MASTER_DB_ID, dummyStr); } catch (e) { }
                    processLoadedData(dummyDB);
                }
            } catch (err) {
                console.error("Simulation mode error:", err);
                processLoadedData(null);
            }
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initAgendaGuru);
    } else {
        initAgendaGuru();
    }

    function processLoadedData(res) {
        document.getElementById('loading-overlay').style.display = 'none';
        if (res && res.config && res.config.registeredEmail) {
            db = res;
            db.config = db.config || {};
            db.subjects = db.subjects || [];
            if (!db.attendance) db.attendance = [];
            if (!db.grades) db.grades = [];
            if (!db.schedule) db.schedule = [];
            if (!db.poin) db.poin = [];
            if (!db.ekskul) db.ekskul = [];
            if (!db.raporNotes) db.raporNotes = [];
            if (!db.nilaiAkhir) db.nilaiAkhir = []; // hasil Pengolah Nilai
            if (!db.students) db.students = [];

            // Pastikan atribut identitas siswa selalu terdefinisi dengan aman (dan bertipe teks)
            db.students.forEach(s => {
                s.StudentName = s.StudentName !== undefined && s.StudentName !== null ? String(s.StudentName) : '';
                s.NIS = s.NIS !== undefined && s.NIS !== null ? String(s.NIS) : '';
                s.NISN = s.NISN !== undefined && s.NISN !== null ? String(s.NISN) : '';
                s.TempatLahir = s.TempatLahir || '';
                s.TanggalLahir = s.TanggalLahir || '';
                s.JenisKelamin = s.JenisKelamin || 'L';
                s.Agama = s.Agama || 'Islam';
                s.Alamat = s.Alamat || '';
                s.NamaAyah = s.NamaAyah || '';
                s.NamaIbu = s.NamaIbu || '';
                s.NamaWali = s.NamaWali || '';
                s.ParentPhone = s.ParentPhone ? String(s.ParentPhone) : '';
            });

            if (!db.tpBank || !Array.isArray(db.tpBank) || db.tpBank.length === 0) {
                db.tpBank = JSON.parse(JSON.stringify(DEFAULT_TP_BANK));
            }

            serverRev = Number(db.config._rev) || 0;
            const migrated = migrateTapelFields() | bersihkanDataContohRapor();

            let isLogged = IS_GAS; // mode GAS: data hanya sampai ke sini jika token sesi valid
            if (!IS_GAS) { try { isLogged = localStorage.getItem('agendaGuru_isLogged') === 'true'; } catch (e) { } }
            if (isLogged) {
                showApp();
            } else {
                document.getElementById('login-view').style.display = 'block';
                document.getElementById('form-login').style.display = 'block';
                document.getElementById('form-register').style.display = 'none';
            }
            if (migrated) syncDB();
        } else {
            db = { classes: [], students: [], attendance: [], grades: [], koku: [], ekskul: [], journal: [], schedule: [], poin: [], subjects: [], tpBank: JSON.parse(JSON.stringify(DEFAULT_TP_BANK)), raporNotes: [], config: {} };
            document.getElementById('login-view').style.display = 'block';
            document.getElementById('form-login').style.display = 'none';
            document.getElementById('form-register').style.display = 'block';
        }
    }

    // --- Auth Logic ---
    function handleAuth(e, mode) {
        e.preventDefault();
        if (IS_GAS) return handleAuthServer(mode);

        // Mode simulasi/demo lokal (tanpa server)
        if (mode === 'login') {
            let email = document.getElementById('loginUser').value.trim();
            let pass = document.getElementById('loginPass').value;
            if (!db.config.registeredEmail) return Swal.fire('Akses Ditolak', 'Belum ada akun terdaftar.', 'warning');
            if (email !== db.config.registeredEmail || pass !== db.config.registeredPassword) return Swal.fire('Gagal Masuk', 'Email atau Password salah.', 'error');
            localStorage.setItem('agendaGuru_isLogged', 'true');
            showApp();
        } else if (mode === 'register') {
            let name = document.getElementById('regName').value.trim();
            let email = document.getElementById('regEmail').value.trim();
            let pass = document.getElementById('regPass').value;
            if (db.config.registeredEmail) return Swal.fire('Pendaftaran Ditolak', 'Aplikasi ini sudah memiliki pengguna.', 'error');

            db.config.registeredEmail = email; db.config.registeredPassword = pass; db.config.userName = name; db.config.userSchool = "Sekolah Saya";
            syncDB(); localStorage.setItem('agendaGuru_isLogged', 'true');
            Swal.fire({ icon: 'success', title: 'Berhasil', text: 'Akun berhasil dibuat!', timer: 1500, showConfirmButton: false });
            showApp();
        }
    }

    // Login/daftar di mode GAS: password dicek di server, browser hanya menerima token sesi.
    function setAuthBusy(busy) {
        document.querySelectorAll('#form-login button[type=submit], #form-register button[type=submit]').forEach(b => { b.disabled = busy; });
    }
    function handleAuthServer(mode) {
        const onResult = res => {
            setAuthBusy(false);
            if (!res || !res.ok) {
                if (res && res.notRegistered) toggleAuthMode('register');
                return Swal.fire(mode === 'login' ? 'Gagal Masuk' : 'Pendaftaran Gagal', (res && res.message) || 'Terjadi kesalahan.', 'error');
            }
            setToken(res.token);
            ['loginPass', 'regPass'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
            applyServerData(res.data);
            if (mode === 'register') Swal.fire({ icon: 'success', title: 'Berhasil', text: 'Akun berhasil dibuat!', timer: 1500, showConfirmButton: false });
        };
        const onFail = err => {
            setAuthBusy(false);
            Swal.fire('Gagal', 'Tidak bisa terhubung ke server: ' + ((err && err.message) || err), 'error');
        };
        setAuthBusy(true);
        if (mode === 'login') {
            const email = document.getElementById('loginUser').value.trim();
            const pass = document.getElementById('loginPass').value;
            google.script.run.withSuccessHandler(onResult).withFailureHandler(onFail).loginAccount(email, pass);
        } else {
            const name = document.getElementById('regName').value.trim();
            const email = document.getElementById('regEmail').value.trim();
            const pass = document.getElementById('regPass').value;
            if (pass.length < 6) { setAuthBusy(false); return Swal.fire('Password Terlalu Pendek', 'Password minimal 6 karakter.', 'warning'); }
            const lisensiEl = document.getElementById('regLicense');
            const licenseKey = normalizeLicenseKey(lisensiEl ? lisensiEl.value : '');
            if (!LICENSE_KEY_RE.test(licenseKey)) { setAuthBusy(false); return Swal.fire('Kode Lisensi', 'Masukkan kode lisensi dari penjual (format AG-XXXX-XXXX-XXXX).', 'warning'); }
            google.script.run.withSuccessHandler(onResult).withFailureHandler(onFail).registerAccount(name, email, pass, licenseKey);
        }
    }

    // Ganti password (mode GAS) dari menu profil
    async function changePasswordPrompt() {
        if (!IS_GAS) return Swal.fire('Info', 'Fitur ini hanya berjalan saat aplikasi diakses melalui Google Apps Script.', 'info');
        const r = await Swal.fire({
            title: 'Ganti Password',
            html: '<input type="password" id="cp-old" class="swal2-input" placeholder="Password lama">' +
                '<input type="password" id="cp-new" class="swal2-input" placeholder="Password baru (min. 6 karakter)">' +
                '<input type="password" id="cp-new2" class="swal2-input" placeholder="Ulangi password baru">',
            showCancelButton: true, confirmButtonText: 'Simpan', cancelButtonText: 'Batal', focusConfirm: false,
            preConfirm: () => {
                const o = document.getElementById('cp-old').value, n = document.getElementById('cp-new').value, n2 = document.getElementById('cp-new2').value;
                if (!o || !n) { Swal.showValidationMessage('Semua kolom wajib diisi'); return false; }
                if (n.length < 6) { Swal.showValidationMessage('Password baru minimal 6 karakter'); return false; }
                if (n !== n2) { Swal.showValidationMessage('Ulangan password baru tidak sama'); return false; }
                return { o, n };
            }
        });
        if (!r.isConfirmed) return;
        Swal.fire({ title: 'Menyimpan...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
        google.script.run
            .withSuccessHandler(res => {
                if (res && res.authError) return handleSessionExpired();
                if (res && res.ok) Swal.fire('Berhasil', 'Password berhasil diganti.', 'success');
                else Swal.fire('Gagal', (res && res.message) || 'Password gagal diganti.', 'error');
            })
            .withFailureHandler(err => Swal.fire('Gagal', 'Tidak bisa terhubung ke server: ' + err, 'error'))
            .changePassword(getToken(), r.value.o, r.value.n);
    }

    function showApp() {
        document.getElementById('login-view').style.display = 'none';
        document.getElementById('app-view').style.display = 'block';
        populateTapelSelect();
        populateClassSelects();
        populateMapelSelects();
        updateProfileUI();
        switchView('beranda');
    }

    function toggleAuthMode(mode) {
        if (mode === 'register') {
            document.getElementById('form-login').style.display = 'none'; document.getElementById('form-register').style.display = 'block';
        } else {
            document.getElementById('form-login').style.display = 'block'; document.getElementById('form-register').style.display = 'none';
        }
    }

    function logout() {
        if (IS_GAS) {
            if (pendingSyncCount > 0 || conflictPending || isDirty()) {
                return Swal.fire('Tunggu Sebentar', 'Masih ada perubahan yang belum tersimpan ke Spreadsheet. Pastikan internet tersambung, tunggu indikator "Menyimpan..." hilang, lalu coba keluar lagi.', 'warning');
            }
            const t = getToken();
            setToken('');
            // Hapus salinan data di perangkat ini (penting untuk komputer/HP yang dipakai bersama)
            const clearAndReload = async () => {
                try { localStorage.removeItem(MASTER_DB_ID); } catch (e) { }
                try { if (typeof localforage !== 'undefined') await localforage.removeItem(MASTER_DB_ID); } catch (e) { }
                try { localStorage.removeItem(LICENSE_CACHE_KEY); } catch (e) { }
                location.reload();
            };
            google.script.run.withSuccessHandler(clearAndReload).withFailureHandler(clearAndReload).logoutAccount(t);
            return;
        }

        localStorage.removeItem('agendaGuru_isLogged');
        document.getElementById('app-view').style.display = 'none'; document.getElementById('login-view').style.display = 'block';
        document.getElementById('form-login').style.display = 'block'; document.getElementById('form-register').style.display = 'none';
        document.getElementById('loginUser').value = ''; document.getElementById('loginPass').value = '';

        // Mereset UI kembali ke beranda tanpa melakukan reload paksa
        switchView('beranda');
    }

    // --- Navigation ---
    function switchView(viewId) {
        document.querySelectorAll('.view-section').forEach(el => el.classList.remove('active'));
        const targetView = document.getElementById('view-' + viewId);
        if (targetView) targetView.classList.add('active');

        document.querySelectorAll('.nav-link').forEach(el => el.classList.remove('active'));
        const activeLink = document.querySelector(`.nav-link[onclick*="switchView('${viewId}')"]`);
        if (activeLink) activeLink.classList.add('active');

        let titles = { beranda: "Beranda", kelas: "Manajemen Kelas", absensi: "Absensi Siswa", nilai: "Data Nilai", kokurikuler: "Nilai Kokurikuler", ekstrakurikuler: "Nilai Ekstrakurikuler", jurnal: "Jurnal Harian", jadwal: "Jadwal Pelajaran", poin: "Poin Siswa", pengolah: "Pengolah Nilai", rapor: "Cetak Rapor Siswa", rekap: "Rekap Laporan" };
        document.getElementById('topbar-title').innerText = titles[viewId] || "AgendaGuru";

        if (viewId === 'beranda') renderBeranda();
        if (viewId === 'kelas') switchKelasTab('siswa');
        if (viewId === 'absensi') loadAbsensiData();
        if (viewId === 'nilai') loadNilaiData();
        if (viewId === 'kokurikuler') { switchKokuTab('input'); renderKokuDimensi(); loadKokuHistory(); }
        if (viewId === 'ekstrakurikuler') { switchEkskulTab('input'); loadEkskulData(); }
        if (viewId === 'jurnal') { switchJurnalTab('input'); renderJurnal(); if (document.getElementById('jurnal-date')) document.getElementById('jurnal-date').value = tanggalHariIni(); }
        if (viewId === 'jadwal') renderSchedule();
        if (viewId === 'poin') { switchPoinTab('input'); loadPoinData(); if (document.getElementById('poin-tanggal')) document.getElementById('poin-tanggal').value = tanggalHariIni(); }
        if (viewId === 'rapor') loadRaporViewData();
        if (viewId === 'pengolah') loadPengolahView();
        if (viewId === 'rekap') loadRekapView();

        if (window.innerWidth < 992) { document.getElementById('sidebar-menu').classList.remove('show'); document.getElementById('sidebar-overlay').classList.remove('show'); }
    }

    // --- Data Population Helpers ---
    function populateClassSelects() {
        // Untuk pilihan Manajemen Kelas yang hanya menampilkan kelas spesifik (tanpa "Semua Kelas")
        let optsManajemen = '<option value="">-- Pilih Kelas --</option>';
        db.classes.forEach(c => optsManajemen += `<option value="${escapeHtml(c.ClassID)}">${escapeHtml(c.ClassName)}</option>`);
        const elManajemen = document.getElementById('select-kelas-manajemen');
        if (elManajemen) {
            let currentVal = elManajemen.value;
            elManajemen.innerHTML = optsManajemen;
            if (currentVal && currentVal !== 'All') elManajemen.value = currentVal;
        }

        // Pilihan kelas untuk Cetak Rapor Siswa (hanya kelas nyata tanpa "Semua Kelas")
        const elRapor = document.getElementById('rapor-kelas');
        if (elRapor) {
            let curRaporVal = elRapor.value;
            elRapor.innerHTML = db.classes.map(c => `<option value="${c.ClassID}">${escapeHtml(c.ClassName)}</option>`).join('');
            if (curRaporVal && db.classes.some(c => String(c.ClassID) === String(curRaporVal))) {
                elRapor.value = curRaporVal;
            } else if (db.classes[0]) {
                elRapor.value = db.classes[0].ClassID;
            }
        }

        // Pilihan kelas untuk form Ekstrakurikuler (hanya kelas nyata)
        const elEkskul = document.getElementById('ekskul-kelas');
        if (elEkskul) {
            let curEkskulVal = elEkskul.value;
            elEkskul.innerHTML = '<option value="">-- Pilih Kelas --</option>' + db.classes.map(c => `<option value="${c.ClassID}">${escapeHtml(c.ClassName)}</option>`).join('');
            if (curEkskulVal && db.classes.some(c => String(c.ClassID) === String(curEkskulVal))) {
                elEkskul.value = curEkskulVal;
            } else if (db.classes[0]) {
                elEkskul.value = db.classes[0].ClassID;
            }
        }

        // Untuk pilihan filter di menu lainnya (dengan tambahan opsi "Semua Kelas")
        const selects = ['absen-kelas', 'nilai-kelas', 'koku-kelas', 'koku-riwayat-kelas', 'sched-class', 'filter-sched-class', 'jurnal-kelas', 'filter-jurnal-kelas', 'poin-kelas', 'filter-poin-kelas', 'ekskul-riwayat-kelas'];
        let optsAll = '<option value="">-- Pilih Kelas --</option><option value="All">Semua Kelas</option>';
        db.classes.forEach(c => optsAll += `<option value="${escapeHtml(c.ClassID)}">${escapeHtml(c.ClassName)}</option>`);
        selects.forEach(id => {
            let el = document.getElementById(id);
            if (el) {
                let val = el.value;
                el.innerHTML = optsAll;
                if (val) el.value = val;
            }
        });
    }

    function populateMapelSelects() {
        const selects = ['nilai-mapel', 'jurnal-mapel', 'sched-subject'];
        let opts = '<option value="">-- Pilih Mapel --</option>';
        if (db.subjects) {
            db.subjects.forEach(m => opts += `<option value="${m}">${m}</option>`);
        }
        selects.forEach(id => {
            let el = document.getElementById(id);
            if (el) {
                let val = el.value;
                el.innerHTML = opts;
                if (val && db.subjects.includes(val)) el.value = val;
            }
        });

        // Sinkronkan juga filter dropdown mata pelajaran di Modal Bank TP
        const elBankTP = document.getElementById('bank-tp-filter-mapel');
        if (elBankTP && db.subjects && db.subjects.length > 0) {
            let curVal = elBankTP.value;
            let tpOpts = db.subjects.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
            elBankTP.innerHTML = tpOpts;
            if (curVal && db.subjects.includes(curVal)) {
                elBankTP.value = curVal;
            } else {
                elBankTP.value = db.subjects[0];
            }
        }
    }

    function showAddMapelModal() {
        Swal.fire({
            title: 'Tambah Mata Pelajaran', input: 'text', inputPlaceholder: 'Cth: Matematika / IPAS',
            showCancelButton: true, confirmButtonText: 'Simpan', cancelButtonText: 'Batal'
        }).then((result) => {
            if (result.isConfirmed && result.value) {
                let newMapel = result.value.trim();
                if (!db.subjects) db.subjects = [];
                if (!db.subjects.includes(newMapel)) {
                    db.subjects.push(newMapel); populateMapelSelects(); syncDB();
                    Swal.fire('Tersimpan!', `Mata Pelajaran ${newMapel} berhasil ditambahkan.`, 'success');
                } else { Swal.fire('Info', 'Mata Pelajaran tersebut sudah ada di daftar.', 'info'); }
            }
        });
    }

    function escapeHtml(str) {
        return String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    }

    function showManageMapelModal() {
        if (!db.subjects) db.subjects = [];
        let listHtml;
        if (db.subjects.length === 0) {
            listHtml = '<p class="text-muted text-center mb-0">Belum ada mata pelajaran.</p>';
        } else {
            listHtml = '<ul class="list-group text-start">' + db.subjects.map(m => `
                    <li class="list-group-item d-flex justify-content-between align-items-center">
                        <span>${escapeHtml(m)}</span>
                        <span class="text-nowrap">
                            <button class="btn btn-sm btn-outline-primary me-1" onclick="renameMapel('${encodeURIComponent(m)}')" title="Edit Nama"><i class="bi bi-pencil-square"></i></button>
                            <button class="btn btn-sm btn-outline-danger" onclick="deleteMapel('${encodeURIComponent(m)}')" title="Hapus"><i class="bi bi-trash"></i></button>
                        </span>
                    </li>`).join('') + '</ul>';
        }
        Swal.fire({
            title: 'Kelola Mata Pelajaran',
            html: `<div style="max-height:320px; overflow-y:auto;">${listHtml}</div>`,
            showCancelButton: false,
            confirmButtonText: 'Tutup'
        });
    }

    function renameMapel(encodedOldName) {
        const oldName = decodeURIComponent(encodedOldName);
        Swal.fire({
            title: 'Edit Nama Mata Pelajaran', input: 'text', inputValue: oldName,
            showCancelButton: true, confirmButtonText: 'Simpan', cancelButtonText: 'Batal'
        }).then((result) => {
            if (!result.isConfirmed || !result.value) return;
            const newName = result.value.trim();
            if (!newName || newName === oldName) return showManageMapelModal();
            if (db.subjects.includes(newName)) { Swal.fire('Gagal', 'Nama mata pelajaran tersebut sudah ada.', 'error').then(showManageMapelModal); return; }

            const idx = db.subjects.indexOf(oldName);
            if (idx > -1) db.subjects[idx] = newName;

            // Perbarui juga seluruh data Nilai, Jurnal, dan Jadwal yang sudah memakai nama lama,
            // supaya riwayat data yang sudah diinput tidak terlepas dari mata pelajaran barunya.
            db.grades.forEach(g => { if (g.Subject === oldName) g.Subject = newName; });
            db.journal.forEach(j => { if (j.Subject === oldName) j.Subject = newName; });
            db.schedule.forEach(s => { if (s.Subject === oldName) s.Subject = newName; });

            populateMapelSelects(); syncDB();
            renderManajemenMapelTable();

            // Render ulang tampilan yang mungkin sedang menampilkan data mapel ini
            if (document.getElementById('view-nilai').classList.contains('active')) loadNilaiData();
            if (document.getElementById('view-jurnal').classList.contains('active')) renderJurnal();
            if (document.getElementById('view-jadwal').classList.contains('active')) renderSchedule();

            Swal.fire('Tersimpan!', `Mata Pelajaran berhasil diubah menjadi "${newName}".`, 'success');
        });
    }

    function deleteMapel(encodedName) {
        const name = decodeURIComponent(encodedName);
        const usageCount = db.grades.filter(g => g.Subject === name).length
            + db.journal.filter(j => j.Subject === name).length
            + db.schedule.filter(s => s.Subject === name).length;
        const warningText = usageCount > 0
            ? `Mata pelajaran ini masih dipakai di ${usageCount} data Nilai/Jurnal/Jadwal. Data tersebut TIDAK akan terhapus, hanya mata pelajarannya yang tidak akan muncul lagi di pilihan dropdown.`
            : 'Mata pelajaran ini akan dihapus dari daftar.';

        Swal.fire({
            title: `Hapus "${escapeHtml(name)}"?`, text: warningText, icon: 'warning',
            showCancelButton: true, confirmButtonColor: '#C44747', confirmButtonText: 'Ya, Hapus!', cancelButtonText: 'Batal'
        }).then((result) => {
            if (result.isConfirmed) {
                db.subjects = db.subjects.filter(m => m !== name);
                populateMapelSelects(); syncDB();
                renderManajemenMapelTable();
                Swal.fire('Terhapus!', `Mata Pelajaran "${name}" berhasil dihapus.`, 'success');
            }
        });
    }

    // --- Dashboard / Beranda ---
    let dashboardTimer;
    function renderBeranda() {
        document.getElementById('stat-siswa').innerText = db.students.length;
        document.getElementById('stat-kelas').innerText = db.classes.length;
        let tH = 0, tS = 0, tI = 0, tA = 0;
        db.attendance.filter(rec => matchTapel(rec, getActiveTapel())).forEach(rec => {
            try {
                let attData = typeof rec.AttendanceData === 'string' ? JSON.parse(rec.AttendanceData) : rec.AttendanceData;
                for (let key in attData) { attData[key].forEach(val => { if (val === 'H') tH++; else if (val === 'S') tS++; else if (val === 'I') tI++; else if (val === 'A') tA++; }); }
            } catch (e) { }
        });
        let totalAtt = tH + tS + tI + tA;
        document.getElementById('stat-hadir').innerText = (totalAtt > 0 ? Math.round((tH / totalAtt) * 100) : 0) + "%";

        function updateLiveSchedule() {
            const daysIndo = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'], todayStr = daysIndo[new Date().getDay()];
            let todaySchedules = db.schedule.filter(s => s.Day === todayStr).sort((a, b) => a.TimeStart.localeCompare(b.TimeStart));
            const dashboardJadwal = document.getElementById('dashboard-jadwal');
            if (todaySchedules.length === 0) { dashboardJadwal.innerHTML = `<li class="list-group-item text-center text-muted py-3">Tidak ada jadwal hari ini (${todayStr})</li>`; }
            else {
                let htmlJadwal = ''; const now = new Date(), currentTime = now.getHours().toString().padStart(2, '0') + ':' + now.getMinutes().toString().padStart(2, '0');
                todaySchedules.forEach(s => {
                    let className = s.ClassID === 'All' ? 'Semua Kelas' : (db.classes.find(c => c.ClassID === s.ClassID)?.ClassName || s.ClassID);
                    let statusBadge = currentTime > s.TimeEnd ? `<span class="badge bg-secondary rounded-pill">Selesai</span>` : (currentTime >= s.TimeStart && currentTime <= s.TimeEnd ? `<span class="badge bg-success rounded-pill placeholder-glow"><span class="placeholder col-12 px-2">Berlangsung</span></span>` : `<span class="badge bg-primary rounded-pill">Akan Datang</span>`);
                    htmlJadwal += `<li class="list-group-item d-flex justify-content-between align-items-center py-3"><div><span class="text-warning fw-bold me-2 fs-5"><i class="bi bi-clock"></i> ${s.TimeStart}</span> <span class="fw-bold">${escapeHtml(s.Subject)}</span> <small class="text-muted ms-1">(${escapeHtml(className)})</small></div>${statusBadge}</li>`;
                });
                dashboardJadwal.innerHTML = htmlJadwal;
            }
        }
        updateLiveSchedule();
        if (dashboardTimer) clearInterval(dashboardTimer);
        dashboardTimer = setInterval(updateLiveSchedule, 60000);

        if (chartHadir) chartHadir.destroy();
        chartHadir = new Chart(document.getElementById('chartKehadiran').getContext('2d'), {
            type: 'doughnut', data: { labels: ['Hadir', 'Sakit', 'Izin', 'Alpha'], datasets: [{ data: totalAtt > 0 ? [tH, tS, tI, tA] : [1, 0, 0, 0], backgroundColor: ['#21865B', '#315E91', '#B7791F', '#C44747'] }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right' } } }
        });
    }

    // --- Class & Student Management ---
    function showAddClassModal() {
        Swal.fire({ title: 'Tambah Kelas Baru', input: 'text', inputPlaceholder: 'Nama Kelas (cth: Kelas 10-C)', showCancelButton: true, confirmButtonText: 'Simpan', cancelButtonText: 'Batal' }).then((result) => {
            if (result.isConfirmed && result.value) {
                // ID harus unik: setelah ada kelas yang dihapus, "C" + jumlah kelas bisa bentrok dengan ID yang masih dipakai
                let n = db.classes.length + 1;
                while (db.classes.some(c => String(c.ClassID) === 'C' + n)) n++;
                db.classes.push({ ClassID: "C" + n, ClassName: result.value }); populateClassSelects(); syncDB();
                Swal.fire('Tersimpan!', `Kelas ${result.value} berhasil ditambahkan.`, 'success');
            }
        });
    }

    // Fungsi baru untuk menghapus Kelas beserta isinya
    function deleteClass() {
        const cid = document.getElementById('select-kelas-manajemen').value;
        if (!cid || cid === 'All') return Swal.fire('Peringatan', 'Pilih kelas yang valid untuk dihapus.', 'warning');

        const className = db.classes.find(c => c.ClassID === cid)?.ClassName || cid;

        Swal.fire({
            title: `Hapus Kelas ${className}?`,
            text: "Semua data siswa, absensi, dan nilai di kelas ini akan ikut terhapus secara permanen!",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#C44747',
            confirmButtonText: 'Ya, Hapus Kelas!'
        }).then((result) => {
            if (result.isConfirmed) {
                // Hapus data kelas
                db.classes = db.classes.filter(c => c.ClassID !== cid);
                // Hapus semua siswa di kelas ini
                db.students = db.students.filter(s => s.ClassID !== cid);

                // Membersihkan data turunan secara permanen
                db.attendance = db.attendance.filter(a => a.ClassID !== cid);
                db.grades = db.grades.filter(g => g.ClassID !== cid);
                db.journal = db.journal.filter(j => j.ClassID !== cid);
                db.schedule = db.schedule.filter(s => s.ClassID !== cid);
                db.koku = db.koku.filter(k => k.ClassID !== cid);
                db.poin = db.poin.filter(p => p.ClassID !== cid);

                syncDB();
                populateClassSelects();

                // Reset tabel
                document.getElementById('select-kelas-manajemen').value = '';
                renderStudentsTable();

                Swal.fire('Terhapus!', `Kelas ${className} beserta data siswanya telah berhasil dihapus.`, 'success');
            }
        });
    }

    let activeStudentDetailId = null;

    function showAddStudentModal() {
        const cid = document.getElementById('select-kelas-manajemen').value;
        if (!cid || cid === 'All') return Swal.fire('Peringatan', 'Pilih kelas terlebih dahulu.', 'warning');

        const classSelect = document.getElementById('student-form-class');
        if (classSelect) {
            classSelect.innerHTML = db.classes.map(c => `<option value="${c.ClassID}">${escapeHtml(c.ClassName)}</option>`).join('');
            classSelect.value = cid;
        }

        document.getElementById('modalStudentTitle').innerText = 'Tambah Siswa Baru';
        document.getElementById('student-form-id').value = '';
        document.getElementById('student-form-name').value = '';
        document.getElementById('student-form-nis').value = '';
        document.getElementById('student-form-nisn').value = '';
        document.getElementById('student-form-jk').value = 'L';
        document.getElementById('student-form-agama').value = 'Islam';
        document.getElementById('student-form-tempat-lahir').value = '';
        document.getElementById('student-form-tanggal-lahir').value = '';
        document.getElementById('student-form-alamat').value = '';
        document.getElementById('student-form-ayah').value = '';
        document.getElementById('student-form-ibu').value = '';
        document.getElementById('student-form-wali').value = '';
        document.getElementById('student-form-phone').value = '';

        const modalEl = document.getElementById('modalStudent');
        let modalInstance = bootstrap.Modal.getInstance(modalEl);
        if (!modalInstance) modalInstance = new bootstrap.Modal(modalEl);
        modalInstance.show();
    }

    function editStudent(studentId) {
        const student = db.students.find(s => s.StudentID === studentId);
        if (!student) return;

        const classSelect = document.getElementById('student-form-class');
        if (classSelect) {
            classSelect.innerHTML = db.classes.map(c => `<option value="${c.ClassID}">${escapeHtml(c.ClassName)}</option>`).join('');
            classSelect.value = student.ClassID;
        }

        document.getElementById('modalStudentTitle').innerText = `Edit Data Siswa - ${student.StudentName}`;
        document.getElementById('student-form-id').value = student.StudentID;
        document.getElementById('student-form-name').value = student.StudentName || '';
        document.getElementById('student-form-nis').value = student.NIS !== undefined && student.NIS !== null ? String(student.NIS) : '';
        document.getElementById('student-form-nisn').value = student.NISN !== undefined && student.NISN !== null ? String(student.NISN) : '';
        document.getElementById('student-form-jk').value = student.JenisKelamin || 'L';
        document.getElementById('student-form-agama').value = student.Agama || 'Islam';
        document.getElementById('student-form-tempat-lahir').value = student.TempatLahir || '';
        document.getElementById('student-form-tanggal-lahir').value = student.TanggalLahir || '';
        document.getElementById('student-form-alamat').value = student.Alamat || '';
        document.getElementById('student-form-ayah').value = student.NamaAyah || '';
        document.getElementById('student-form-ibu').value = student.NamaIbu || '';
        document.getElementById('student-form-wali').value = student.NamaWali || '';
        document.getElementById('student-form-phone').value = student.ParentPhone || '';

        const modalEl = document.getElementById('modalStudent');
        let modalInstance = bootstrap.Modal.getInstance(modalEl);
        if (!modalInstance) modalInstance = new bootstrap.Modal(modalEl);
        modalInstance.show();
    }

    function handleSaveStudent(e) {
        e.preventDefault();
        const studentId = document.getElementById('student-form-id').value;
        const cid = document.getElementById('student-form-class').value;
        const name = document.getElementById('student-form-name').value.trim();
        const nis = String(document.getElementById('student-form-nis').value || '').trim();
        const nisn = String(document.getElementById('student-form-nisn').value || '').trim();
        const jk = document.getElementById('student-form-jk').value;
        const agama = document.getElementById('student-form-agama').value;
        const tempatLahir = document.getElementById('student-form-tempat-lahir').value.trim();
        const tglLahir = document.getElementById('student-form-tanggal-lahir').value;
        const alamat = document.getElementById('student-form-alamat').value.trim();
        const ayah = document.getElementById('student-form-ayah').value.trim();
        const ibu = document.getElementById('student-form-ibu').value.trim();
        const wali = document.getElementById('student-form-wali').value.trim();
        let phone = document.getElementById('student-form-phone').value.trim();

        if (!name) return Swal.fire('Peringatan', 'Nama siswa wajib diisi.', 'warning');
        if (!cid) return Swal.fire('Peringatan', 'Kelas wajib dipilih.', 'warning');

        if (phone && !phone.startsWith('0') && !phone.startsWith('+') && !phone.startsWith('62')) {
            phone = '0' + phone;
        }

        if (studentId) {
            const idx = db.students.findIndex(s => s.StudentID === studentId);
            if (idx > -1) {
                db.students[idx].ClassID = cid;
                db.students[idx].StudentName = name;
                db.students[idx].NIS = nis;
                db.students[idx].NISN = nisn;
                db.students[idx].JenisKelamin = jk;
                db.students[idx].Agama = agama;
                db.students[idx].TempatLahir = tempatLahir;
                db.students[idx].TanggalLahir = tglLahir;
                db.students[idx].Alamat = alamat;
                db.students[idx].NamaAyah = ayah;
                db.students[idx].NamaIbu = ibu;
                db.students[idx].NamaWali = wali;
                db.students[idx].ParentPhone = phone;
            }
        } else {
            const newId = "S" + Date.now() + Math.floor(Math.random() * 1000);
            db.students.push({
                StudentID: newId,
                ClassID: cid,
                StudentName: name,
                NIS: nis,
                NISN: nisn,
                JenisKelamin: jk,
                Agama: agama,
                TempatLahir: tempatLahir,
                TanggalLahir: tglLahir,
                Alamat: alamat,
                NamaAyah: ayah,
                NamaIbu: ibu,
                NamaWali: wali,
                ParentPhone: phone
            });
        }

        syncDB();
        renderStudentsTable();

        // Render ulang tabel modul lain jika sedang aktif
        if (document.getElementById('view-absensi')?.classList.contains('active')) loadAbsensiData();
        if (document.getElementById('view-nilai')?.classList.contains('active')) loadNilaiData();
        if (document.getElementById('view-rapor')?.classList.contains('active')) {
            onRaporKelasChange();
            previewRaporSiswa();
        }

        const modalEl = document.getElementById('modalStudent');
        if (modalEl) {
            let modalInstance = bootstrap.Modal.getInstance(modalEl);
            if (modalInstance) modalInstance.hide();
        }

        Swal.fire({
            icon: 'success',
            title: 'Berhasil!',
            text: studentId ? 'Data siswa berhasil diperbarui.' : 'Siswa baru berhasil ditambahkan.',
            timer: 1500,
            showConfirmButton: false
        });
    }

    function viewStudentDetail(studentId) {
        const s = db.students.find(x => x.StudentID === studentId);
        if (!s) return;
        activeStudentDetailId = studentId;

        const className = db.classes.find(c => String(c.ClassID) === String(s.ClassID))?.ClassName || '-';
        const jkText = s.JenisKelamin === 'L' ? 'Laki-laki (L)' : (s.JenisKelamin === 'P' ? 'Perempuan (P)' : (s.JenisKelamin || '-'));
        const tglText = s.TanggalLahir ? formatTanggalIndo(s.TanggalLahir) : '-';
        const ttl = (s.TempatLahir ? s.TempatLahir + ', ' : '') + tglText;

        const bodyEl = document.getElementById('modalDetailStudentBody');
        if (bodyEl) {
            bodyEl.innerHTML = `
                <div class="text-center mb-3">
                    <div class="avatar-circle mx-auto mb-2" style="width: 60px; height: 60px; border-radius: 50%; background: var(--teal-light); color: var(--teal-accent); display: flex; align-items: center; justify-content: center; font-size: 24px; font-weight: bold;">
                        ${escapeHtml((s.StudentName || '?').charAt(0).toUpperCase())}
                    </div>
                    <h5 class="fw-bold mb-0 text-dark">${escapeHtml(s.StudentName)}</h5>
                    <span class="badge bg-primary mt-1">Kelas ${escapeHtml(className)}</span>
                </div>
                <div class="table-responsive">
                    <table class="table table-sm table-borderless mb-0" style="font-size: 0.9rem;">
                        <tbody>
                            <tr><td class="text-muted" width="40%">ID Internal:</td><td class="fw-semibold font-monospace">${escapeHtml(s.StudentID)}</td></tr>
                            <tr><td class="text-muted">NIS:</td><td class="fw-semibold">${escapeHtml(s.NIS || '-')}</td></tr>
                            <tr><td class="text-muted">NISN:</td><td class="fw-semibold">${escapeHtml(s.NISN || '-')}</td></tr>
                            <tr><td class="text-muted">Jenis Kelamin:</td><td class="fw-semibold">${escapeHtml(jkText)}</td></tr>
                            <tr><td class="text-muted">Agama:</td><td class="fw-semibold">${escapeHtml(s.Agama || '-')}</td></tr>
                            <tr><td class="text-muted">Tempat, Tgl Lahir:</td><td class="fw-semibold">${escapeHtml(ttl)}</td></tr>
                            <tr><td class="text-muted">Alamat:</td><td class="fw-semibold">${escapeHtml(s.Alamat || '-')}</td></tr>
                            <tr class="border-top"><td class="text-muted pt-2">Nama Ayah:</td><td class="fw-semibold pt-2">${escapeHtml(s.NamaAyah || '-')}</td></tr>
                            <tr><td class="text-muted">Nama Ibu:</td><td class="fw-semibold">${escapeHtml(s.NamaIbu || '-')}</td></tr>
                            <tr><td class="text-muted">Nama Wali:</td><td class="fw-semibold">${escapeHtml(s.NamaWali || '-')}</td></tr>
                            <tr><td class="text-muted">No. WhatsApp Ortu:</td><td class="fw-semibold text-success">${s.ParentPhone ? `<i class="bi bi-whatsapp"></i> ${escapeHtml(s.ParentPhone)}` : '<span class="text-muted">-</span>'}</td></tr>
                        </tbody>
                    </table>
                </div>
            `;
        }

        const modalEl = document.getElementById('modalDetailStudent');
        let modalInstance = bootstrap.Modal.getInstance(modalEl);
        if (!modalInstance) modalInstance = new bootstrap.Modal(modalEl);
        modalInstance.show();
    }

    function openRaporFromDetail() {
        if (!activeStudentDetailId) return;
        const modalEl = document.getElementById('modalDetailStudent');
        if (modalEl) {
            let modalInstance = bootstrap.Modal.getInstance(modalEl);
            if (modalInstance) modalInstance.hide();
        }
        openRaporForStudent(activeStudentDetailId);
    }

    async function downloadTemplateExcel(btn) {
        let originalHtml = '';
        if (btn) {
            originalHtml = btn.innerHTML;
            btn.disabled = true;
            btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1" role="status" aria-hidden="true"></span> Menyiapkan...';
        }
        try {
            await ensureXLSX();
            const wb = XLSX.utils.book_new();
            const templateData = [
                ["No", "Nama Siswa", "NIS", "NISN", "Tempat Lahir", "Tanggal Lahir (YYYY-MM-DD)", "Jenis Kelamin (L/P)", "Agama", "Alamat", "Nama Ayah", "Nama Ibu", "Nama Wali", "No WhatsApp Ortu"],
                [1, "Ahmad Pratama", "2024001", "0081234567", "Bandung", "2010-05-14", "L", "Islam", "Jl. Merdeka No. 12", "Heri Pratama", "Siti Aminah", "", "081234567890"],
                [2, "Siti Nurhaliza", "2024002", "0082345678", "Jakarta", "2010-08-20", "P", "Islam", "Jl. Mawar No. 45", "Rahmat Hidayat", "Nurhasanah", "", "081298765432"],
                [3, "Budi Santoso", "2024003", "0083456789", "Surabaya", "2010-02-11", "L", "Kristen", "Jl. Melati No. 8", "Santoso Wijaya", "Maria Ulfah", "", "081311223344"]
            ];
            const ws = XLSX.utils.aoa_to_sheet(templateData);

            // Set lebar kolom agar sangat rapi dibuka di Excel
            ws['!cols'] = [
                { wch: 6 },
                { wch: 28 },
                { wch: 14 },
                { wch: 16 },
                { wch: 18 },
                { wch: 22 },
                { wch: 18 },
                { wch: 14 },
                { wch: 32 },
                { wch: 22 },
                { wch: 22 },
                { wch: 20 },
                { wch: 22 }
            ];

            // Pastikan seluruh sel NIS, NISN, Phone, dan Teks bertipe string agar angka 0 di depan tidak terpotong
            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = 1; R <= range.e.r; ++R) {
                for (let C = 1; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                    if (ws[cellRef]) {
                        ws[cellRef].t = 's';
                        ws[cellRef].z = '@';
                    }
                }
            }

            XLSX.utils.book_append_sheet(wb, ws, "Data Siswa");
            XLSX.writeFile(wb, "Template_Data_Siswa_Lengkap.xlsx");

            if (btn) {
                btn.innerHTML = '<i class="bi bi-check-circle me-1"></i> Terunduh!';
                btn.classList.remove('btn-outline-success');
                btn.classList.add('btn-success');
                setTimeout(() => {
                    btn.disabled = false;
                    btn.classList.remove('btn-success');
                    btn.classList.add('btn-outline-success');
                    btn.innerHTML = originalHtml;
                }, 2000);
            }
        } catch (err) {
            console.error('Error saat download template Excel:', err);
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = originalHtml;
            }
            alert('Gagal mengunduh template Excel: ' + (err.message || err));
        }
    }

    function showUploadExcelModal() {
        const cid = document.getElementById('select-kelas-manajemen').value;
        if (!cid || cid === 'All') {
            return Swal.fire({
                icon: 'warning',
                title: 'Pilih Kelas Terlebih Dahulu',
                html: `
                    <div class="text-start">
                        <p class="text-muted small mb-3">Silakan pilih kelas terlebih dahulu sebelum mengupload data siswa.</p>
                        <div class="p-3 bg-light rounded border text-center">
                            <div class="small text-muted mb-2">Belum memiliki format file Excel? Unduh template di bawah:</div>
                            <button type="button" class="btn btn-sm btn-outline-success fw-bold" onclick="downloadTemplateExcel(this)">
                                <i class="bi bi-file-earmark-arrow-down me-1"></i> Download Template Excel Lengkap
                            </button>
                        </div>
                    </div>
                `,
                confirmButtonText: 'Tutup'
            });
        }
        const className = db.classes.find(c => c.ClassID === cid)?.ClassName || cid;
        Swal.fire({
            title: `Upload Siswa ke ${className}`,
            html: `
                <div class="text-start">
                    <div class="p-3 mb-3 bg-light rounded border">
                        <div class="d-flex align-items-center justify-content-between flex-wrap gap-2">
                            <div>
                                <div class="fw-bold small text-dark"><i class="bi bi-file-earmark-excel text-success me-1"></i> Template Excel Siswa</div>
                                <div class="text-muted" style="font-size:0.75rem;">Unduh template lengkap resmi agar pemetaan kolom otomatis & akurat.</div>
                            </div>
                            <button type="button" class="btn btn-sm btn-outline-success fw-bold" onclick="downloadTemplateExcel(this)">
                                <i class="bi bi-file-earmark-arrow-down me-1"></i> Unduh Template
                            </button>
                        </div>
                    </div>
                    <div class="mb-3">
                        <p class="text-muted small mb-1">
                            <b>Ketentuan Kolom:</b><br>
                            &bull; Kolom wajib: <b>Nama Siswa</b>.<br>
                            &bull; Kolom pendukung: <b>NIS</b>, <b>NISN</b>, <b>Tempat Lahir</b>, <b>Tanggal Lahir</b>, <b>Jenis Kelamin</b>, <b>Agama</b>, <b>Alamat</b>, <b>Nama Ayah</b>, <b>Nama Ibu</b>, <b>Nama Wali</b>, <b>No WhatsApp</b>.<br>
                            &bull; Kompatibel dengan template versi lama maupun format Excel kustom.
                        </p>
                    </div>
                    <label class="form-label small fw-bold">Pilih File Excel (.xlsx, .xls):</label>
                    <input type="file" id="excel-file-upload" class="form-control" accept=".xlsx, .xls">
                </div>
            `,
            showCancelButton: true,
            confirmButtonText: 'Proses Upload',
            cancelButtonText: 'Batal',
            preConfirm: () => {
                const file = document.getElementById('excel-file-upload').files[0];
                if (!file) Swal.showValidationMessage('Pilih file Excel terlebih dahulu');
                return file;
            }
        }).then((result) => {
            if (result.isConfirmed && result.value) processExcelUpload(result.value, cid);
        });
    }

    async function processExcelUpload(file, classId) {
        await ensureXLSX();
        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { raw: false });
                if (json.length === 0) return Swal.fire('Error', 'File kosong.', 'error');

                const keys = Object.keys(json[0]);
                const findKey = (patterns) => keys.find(k => patterns.some(p => k.toLowerCase().includes(p.toLowerCase())));

                const nameKey = findKey(['nama']);
                if (!nameKey) return Swal.fire('Error', 'Kolom "Nama Siswa" tidak ditemukan di file Excel.', 'error');

                const nisKey = findKey(['nis', 'nomor induk']);
                const nisnKey = findKey(['nisn']);
                const tmptKey = findKey(['tempat lahir', 'tempat_lahir', 'tmplahir', 'tempat']);
                const tglKey = findKey(['tanggal lahir', 'tanggal_lahir', 'tgllahir', 'tgl_lahir']);
                const jkKey = findKey(['jenis kelamin', 'gender', 'jk', 'kelamin']);
                const agamaKey = findKey(['agama']);
                const alamatKey = findKey(['alamat']);
                const ayahKey = findKey(['ayah', 'bapak']);
                const ibuKey = findKey(['ibu']);
                const waliKey = findKey(['wali']);
                const phoneKey = findKey(['hp', 'wa', 'telepon', 'whatsapp', 'phone', 'kontak']);

                let countAdded = 0;
                json.forEach((row, i) => {
                    const rawName = row[nameKey];
                    if (rawName && String(rawName).trim() !== '') {
                        let phoneVal = phoneKey && row[phoneKey] ? String(row[phoneKey]).trim() : '';
                        if (phoneVal && !phoneVal.startsWith('0') && !phoneVal.startsWith('+') && !phoneVal.startsWith('62')) {
                            phoneVal = '0' + phoneVal;
                        }

                        let rawJK = jkKey && row[jkKey] ? String(row[jkKey]).trim().toUpperCase() : 'L';
                        let jkVal = rawJK.startsWith('P') || rawJK.includes('PEREMPUAN') ? 'P' : 'L';

                        let rawTgl = tglKey && row[tglKey] ? String(row[tglKey]).trim() : '';
                        // Format YYYY-MM-DD jika tanggal numerik dari excel
                        if (/^\d{4}-\d{2}-\d{2}$/.test(rawTgl)) {
                            // Valid ISO format
                        } else if (/^\d{2}\/\d{2}\/\d{4}$/.test(rawTgl)) {
                            const parts = rawTgl.split('/');
                            rawTgl = `${parts[2]}-${parts[1]}-${parts[0]}`;
                        }

                        db.students.push({
                            StudentID: "S" + Date.now() + countAdded + i,
                            ClassID: classId,
                            StudentName: String(rawName).trim(),
                            NIS: nisKey && row[nisKey] ? String(row[nisKey]).trim() : '',
                            NISN: nisnKey && row[nisnKey] ? String(row[nisnKey]).trim() : '',
                            TempatLahir: tmptKey && row[tmptKey] ? String(row[tmptKey]).trim() : '',
                            TanggalLahir: rawTgl,
                            JenisKelamin: jkVal,
                            Agama: agamaKey && row[agamaKey] ? String(row[agamaKey]).trim() : 'Islam',
                            Alamat: alamatKey && row[alamatKey] ? String(row[alamatKey]).trim() : '',
                            NamaAyah: ayahKey && row[ayahKey] ? String(row[ayahKey]).trim() : '',
                            NamaIbu: ibuKey && row[ibuKey] ? String(row[ibuKey]).trim() : '',
                            NamaWali: waliKey && row[waliKey] ? String(row[waliKey]).trim() : '',
                            ParentPhone: phoneVal
                        });
                        countAdded++;
                    }
                });

                syncDB();
                renderStudentsTable();
                Swal.fire('Berhasil!', `${countAdded} data siswa berhasil diimpor ke sistem.`, 'success');
            } catch (error) {
                console.error("Gagal membaca Excel:", error);
                Swal.fire('Error', 'Format file Excel tidak sesuai: ' + error.message, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function renderStudentsTable() {
        const cid = document.getElementById('select-kelas-manajemen').value;
        const tbody = document.getElementById('table-students-body');
        const btnDeleteClass = document.getElementById('btn-hapus-kelas');
        const searchInput = document.getElementById('search-student-manajemen');
        const query = (searchInput?.value || '').toLowerCase().trim();

        if (!cid || cid === 'All') {
            if (btnDeleteClass) btnDeleteClass.style.display = 'none';
            return tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-4">Pilih kelas terlebih dahulu</td></tr>';
        }
        if (btnDeleteClass) btnDeleteClass.style.display = 'inline-block';

        let stds = db.students.filter(s => String(s.ClassID) === String(cid));

        if (query) {
            stds = stds.filter(s => 
                (s.StudentName && s.StudentName.toLowerCase().includes(query)) ||
                (s.NIS && String(s.NIS).toLowerCase().includes(query)) ||
                (s.NISN && String(s.NISN).toLowerCase().includes(query)) ||
                (s.ParentPhone && String(s.ParentPhone).includes(query))
            );
        }

        if (stds.length === 0) {
            return tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">${query ? 'Tidak ada siswa yang cocok dengan pencarian.' : 'Belum ada siswa di kelas ini.'}</td></tr>`;
        }

        let html = '';
        stds.forEach((s, i) => {
            const nis = s.NIS || '-';
            const nisn = s.NISN || '-';
            const jkBadge = s.JenisKelamin === 'P' ? '<span class="badge bg-danger-subtle text-danger border border-danger-subtle">P</span>' : '<span class="badge bg-primary-subtle text-primary border border-primary-subtle">L</span>';
            const phoneText = s.ParentPhone ? `<a href="https://wa.me/${encodeURIComponent(String(s.ParentPhone).replace(/\D/g, '').replace(/^0/, '62'))}" target="_blank" rel="noopener" class="text-decoration-none text-success small"><i class="bi bi-whatsapp"></i> ${escapeHtml(s.ParentPhone)}</a>` : '<small class="text-muted fst-italic">Belum ada no. WA</small>';
            const ortuName = s.NamaAyah || s.NamaIbu || s.NamaWali || '-';
            const chkDisabled = s.ParentPhone ? '' : 'disabled';

            html += `<tr>
                <td class="text-center align-middle"><input type="checkbox" class="chk-std form-check-input" value="${s.StudentID}" ${chkDisabled}></td>
                <td class="text-center align-middle">${i + 1}</td>
                <td class="align-middle">
                    <span class="fw-bold text-dark">${escapeHtml(s.StudentName)}</span>
                    ${s.TempatLahir ? `<br><small class="text-muted" style="font-size:0.75rem;"><i class="bi bi-geo-alt"></i> ${escapeHtml(s.TempatLahir)}</small>` : ''}
                </td>
                <td class="align-middle small">
                    <div><span class="text-muted">NIS:</span> <b>${escapeHtml(nis)}</b></div>
                    <div><span class="text-muted">NISN:</span> <b>${escapeHtml(nisn)}</b></div>
                </td>
                <td class="text-center align-middle">${jkBadge}</td>
                <td class="align-middle small">
                    <div class="text-dark fw-semibold">${escapeHtml(ortuName)}</div>
                    <div>${phoneText}</div>
                </td>
                <td class="text-center align-middle text-nowrap">
                    <button class="btn btn-sm btn-outline-info me-1 mb-1" onclick="viewStudentDetail('${s.StudentID}')" title="Detail Siswa"><i class="bi bi-eye"></i></button>
                    <button class="btn btn-sm btn-outline-primary me-1 mb-1" onclick="editStudent('${s.StudentID}')" title="Edit Siswa"><i class="bi bi-pencil-square"></i></button>
                    <button class="btn btn-sm btn-outline-success me-1 mb-1" onclick="openRaporForStudent('${s.StudentID}')" title="Buka Rapor"><i class="bi bi-printer"></i></button>
                    <button class="btn btn-sm btn-outline-danger mb-1" onclick="deleteStudent('${s.StudentID}')" title="Hapus Siswa"><i class="bi bi-trash"></i></button>
                </td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    function deleteStudent(studentId) {
        Swal.fire({ title: 'Hapus Siswa?', text: "Yakin ingin menghapus?", icon: 'warning', showCancelButton: true, confirmButtonColor: '#C44747', confirmButtonText: 'Ya, Hapus!' }).then((result) => {
            if (result.isConfirmed) { db.students = db.students.filter(s => s.StudentID !== studentId); syncDB(); renderStudentsTable(); Swal.fire('Terhapus!', 'Data dihapus.', 'success'); }
        });
    }

    // --- Absensi ---
    function loadAbsensiData() {
        const cid = document.getElementById('absen-kelas').value, month = document.getElementById('absen-bulan').value;
        if (!cid) { document.getElementById('absensi-chart-container').style.display = 'none'; return; }

        const stds = db.students.filter(s => String(s.ClassID) === String(cid));
        const record = db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(month) && matchTapel(a, getActiveTapel()));

        let savedDates = [], savedAttendance = {};
        if (record) { try { savedDates = JSON.parse(record.DatesData); savedAttendance = JSON.parse(record.AttendanceData); } catch (e) { } }

        if (savedDates.length === 0) savedDates = [""];
        let colCount = savedDates.length;

        let thead = '<th>Nama Siswa</th>';
        for (let i = 0; i < colCount; i++) {
            thead += `<th><input type="date" class="date-input-header date-absen-col" data-idx="${i}" value="${savedDates[i] || ""}" onchange="saveAbsensi(true)">
                    ${colCount > 1 ? `<button type="button" class="btn btn-sm btn-outline-danger mt-1 py-0 px-1" onclick="deleteAbsenColumn(${i})" title="Hapus Kolom Ini"><i class="bi bi-trash"></i></button>` : ''}
                </th>`;
        }
        document.getElementById('absen-thead').innerHTML = thead;

        let tbody = '';
        stds.forEach(s => {
            tbody += `<tr><td class="fw-bold d-flex align-items-center justify-content-between">${escapeHtml(s.StudentName)}
                <button class="btn btn-sm btn-outline-success py-0 px-1" onclick="sendWAAbsen('${s.StudentID}')" title="Kirim WA Orang Tua"><i class="bi bi-whatsapp"></i></button>
                </td>`;
            let sAtt = savedAttendance[s.StudentID] || Array(colCount).fill('-');
            for (let i = 0; i < colCount; i++) {
                let val = sAtt[i] || '-', color = val !== '-' ? `status-${val}` : '';
                tbody += `<td><select class="status-select select-absen-val ${color}" data-sid="${s.StudentID}" data-idx="${i}" onchange="updateSelectColor(this)">
                        <option value="H" ${val === 'H' ? 'selected' : ''}>H</option><option value="S" ${val === 'S' ? 'selected' : ''}>S</option>
                        <option value="I" ${val === 'I' ? 'selected' : ''}>I</option><option value="A" ${val === 'A' ? 'selected' : ''}>A</option><option value="-" ${val === '-' ? 'selected' : ''}>-</option>
                    </select></td>`;
            }
            tbody += `</tr>`;
        });
        document.getElementById('absen-tbody').innerHTML = tbody; document.getElementById('absensi-chart-container').style.display = 'block'; renderChartAbsensi();
    }

    function addAbsenColumn() {
        const cid = document.getElementById('absen-kelas').value, month = document.getElementById('absen-bulan').value;
        if (!cid) return Swal.fire('Error', 'Pilih kelas terlebih dahulu.', 'error');
        saveAbsensi(true);
        let record = db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(month) && matchTapel(a, getActiveTapel()));
        if (record) {
            let dates = JSON.parse(record.DatesData); dates.push(""); record.DatesData = JSON.stringify(dates);
            let att = JSON.parse(record.AttendanceData); for (let sid in att) { att[sid].push("-"); } record.AttendanceData = JSON.stringify(att); syncDB();
        }
        loadAbsensiData();
    }

    function deleteAbsenColumn(idx) {
        const cid = document.getElementById('absen-kelas').value, month = document.getElementById('absen-bulan').value;
        if (!cid) return;
        saveAbsensi(true); // pastikan perubahan yang belum tersimpan tidak ikut hilang
        let record = db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(month) && matchTapel(a, getActiveTapel()));
        if (!record) return;
        let dates = []; try { dates = JSON.parse(record.DatesData); } catch (e) { }
        if (dates.length <= 1) return Swal.fire('Tidak Bisa', 'Minimal harus ada 1 kolom tanggal.', 'warning');

        Swal.fire({
            title: 'Hapus Kolom Ini?', text: 'Seluruh data absensi pada kolom tanggal ini akan dihapus.', icon: 'warning',
            showCancelButton: true, confirmButtonColor: '#C44747', confirmButtonText: 'Ya, Hapus!', cancelButtonText: 'Batal'
        }).then((result) => {
            if (!result.isConfirmed) return;
            dates.splice(idx, 1);
            record.DatesData = JSON.stringify(dates);
            let att = {}; try { att = JSON.parse(record.AttendanceData); } catch (e) { }
            for (let sid in att) { att[sid].splice(idx, 1); }
            record.AttendanceData = JSON.stringify(att);
            syncDB(); loadAbsensiData();
            Swal.fire('Terhapus!', 'Kolom berhasil dihapus.', 'success');
        });
    }

    function updateSelectColor(sel) { sel.className = `status-select select-absen-val status-${sel.value}`; renderChartAbsensi(); saveAbsensi(true); }

    function renderChartAbsensi() {
        let tH = 0, tS = 0, tI = 0, tA = 0;
        document.querySelectorAll('.select-absen-val').forEach(sel => { if (sel.value === 'H') tH++; else if (sel.value === 'S') tS++; else if (sel.value === 'I') tI++; else if (sel.value === 'A') tA++; });
        if (chartAbsensiObj) chartAbsensiObj.destroy();
        chartAbsensiObj = new Chart(document.getElementById('chartAbsensi').getContext('2d'), {
            type: 'bar', data: { labels: ['Hadir', 'Sakit', 'Izin', 'Alpha'], datasets: [{ label: 'Kehadiran', data: [tH, tS, tI, tA], backgroundColor: ['#21865B', '#315E91', '#B7791F', '#C44747'], borderRadius: 5 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });
    }

    function saveAbsensi(isSilent = false) {
        const cid = document.getElementById('absen-kelas').value, month = document.getElementById('absen-bulan').value;
        if (!cid) { if (!isSilent) Swal.fire('Error', 'Pilih kelas', 'error'); return; }
        let customDates = []; document.querySelectorAll('.date-absen-col').forEach(i => customDates.push(i.value));
        let attData = {};
        db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => { let st = []; document.querySelectorAll(`.select-absen-val[data-sid="${s.StudentID}"]`).forEach(sel => st.push(sel.value)); attData[s.StudentID] = st; });

        let idx = db.attendance.findIndex(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(month) && matchTapel(a, getActiveTapel()));
        if (idx >= 0) { db.attendance[idx].DatesData = JSON.stringify(customDates); db.attendance[idx].AttendanceData = JSON.stringify(attData); db.attendance[idx].TahunPelajaran = db.attendance[idx].TahunPelajaran || getActiveTapel(); }
        else db.attendance.push({ RecordID: "A" + Date.now(), ClassID: cid, MonthYear: month, TahunPelajaran: getActiveTapel(), DatesData: JSON.stringify(customDates), AttendanceData: JSON.stringify(attData) });
        syncDB(); if (!isSilent) Swal.fire({ icon: 'success', title: 'Tersimpan', text: 'Data Absensi disimpan', timer: 1500, showConfirmButton: false });
    }

    async function downloadAbsenExcel(semester) {
        const cid = document.getElementById('absen-kelas').value;
        if (!cid) return Swal.fire('Error', 'Pilih kelas', 'error');
        saveAbsensi(true); const className = db.classes.find(c => c.ClassID === cid)?.ClassName || cid;
        Swal.fire({ icon: 'info', title: 'Mengekspor', text: `Menyiapkan Excel...`, timer: 1500 });
        await ensureXLSX();
        const wb = XLSX.utils.book_new(), stds = db.students.filter(s => String(s.ClassID) === String(cid));
        const targetMonths = semester === 'ganjil' ? ['07', '08', '09', '10', '11', '12'] : ['01', '02', '03', '04', '05', '06'];
        const monthNames = { '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'Mei', '06': 'Jun', '07': 'Jul', '08': 'Agu', '09': 'Sep', '10': 'Okt', '11': 'Nov', '12': 'Des' };
        let rekapSemester = {}; stds.forEach(s => rekapSemester[s.StudentID] = { Name: s.StudentName, H: 0, S: 0, I: 0, A: 0 });

        targetMonths.forEach(m => {
            const record = db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(m) && matchTapel(a, getActiveTapel()));
            let sDates = [], sAtt = {};
            if (record) { try { sDates = JSON.parse(record.DatesData); sAtt = JSON.parse(record.AttendanceData); } catch (e) { } }
            let maxCols = sDates.length > 0 ? sDates.length : 1;
            let headDates = Array.from({ length: maxCols }, (_, i) => sDates[i] || `K-${i + 1}`);
            let ws_data = [["Nama Siswa", ...headDates, "Total H", "Total S", "Total I", "Total A"]];

            stds.forEach(s => {
                let row = [s.StudentName], att = sAtt[s.StudentID] || Array(maxCols).fill('-');
                let mH = 0, mS = 0, mI = 0, mA = 0;
                for (let i = 0; i < maxCols; i++) {
                    let val = att[i] || '-';
                    row.push(val);
                    if (val === 'H') { mH++; rekapSemester[s.StudentID].H++; }
                    if (val === 'S') { mS++; rekapSemester[s.StudentID].S++; }
                    if (val === 'I') { mI++; rekapSemester[s.StudentID].I++; }
                    if (val === 'A') { mA++; rekapSemester[s.StudentID].A++; }
                }
                row.push(mH, mS, mI, mA); ws_data.push(row);
            });
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ws_data), monthNames[m]);
        });

        let rekap_data = [["Nama Siswa", "Total Hadir (H)", "Total Sakit (S)", "Total Izin (I)", "Total Alpha (A)"]];
        stds.forEach(s => { let r = rekapSemester[s.StudentID]; rekap_data.push([r.Name, r.H, r.S, r.I, r.A]); });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rekap_data), "Rekap"); XLSX.writeFile(wb, `Absensi_${className}_${semester}.xlsx`);
    }

    // --- Data Nilai ---
    function isGradeMatchingSemester(g, targetSemester) {
        if (!g) return false;
        const sem = String(targetSemester || 'ganjil').toLowerCase();
        const targetMonths = sem === 'ganjil' ? ['07', '08', '09', '10', '11', '12'] : ['01', '02', '03', '04', '05', '06'];

        // 1. Cek properti Semester eksplisit jika ada
        if (g.Semester) {
            return String(g.Semester).toLowerCase() === sem;
        }
        // 2. Cek MonthYear jika ada (misal '08' atau '03')
        if (g.MonthYear) {
            return targetMonths.includes(String(g.MonthYear));
        }
        // 3. Fallback untuk data lama tanpa atribut semester:
        // Record dialokasikan untuk semester Ganjil secara aman agar tidak bocor ke semester Genap
        return sem === 'ganjil';
    }

    function loadNilaiData() {
        const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value, jenis = document.getElementById('nilai-jenis').value, bulan = document.getElementById('nilai-bulan').value, container = document.getElementById('container-table-nilai');
        if (!cid || !mapel) { container.innerHTML = '<div class="alert alert-secondary text-center">Silakan pilih Kelas dan Mata Pelajaran.</div>'; document.getElementById('nilai-stats-container').style.display = 'none'; document.getElementById('btn-add-nilai-col').style.display = 'none'; return; }

        const currentSemester = ['07', '08', '09', '10', '11', '12'].includes(String(bulan)) ? 'ganjil' : 'genap';
        const stds = db.students.filter(s => String(s.ClassID) === String(cid));
        const record = db.grades.slice().reverse().find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === String(jenis) && (jenis === 'Formatif' ? String(g.MonthYear) === String(bulan) : isGradeMatchingSemester(g, currentSemester)) && matchTapel(g, getActiveTapel()));

        let sDates = [], sGrades = {}, sCakupan = [];
        if (record) { try { sDates = JSON.parse(record.DatesData) || []; sGrades = JSON.parse(record.GradesData) || {}; sCakupan = JSON.parse(record.Cakupan) || []; } catch (e) { } }

        let html = '';
        if (jenis === 'Formatif') {
            document.getElementById('btn-add-nilai-col').style.display = 'inline-block';
            if (sDates.length === 0) { sDates = [""]; sCakupan = [""]; }
            let colCount = sDates.length;

            html += '<table class="table table-bordered table-dynamic" id="table-nilai-input"><thead><tr class="table-success"><th style="min-width:210px; vertical-align:middle;">Nama Siswa</th>';
            for (let i = 0; i < colCount; i++) {
                html += `<th><input type="date" class="date-input-header date-nilai-col mb-1" value="${sDates[i] || ""}" onchange="saveNilai(true)"><input type="text" class="form-control form-control-sm cakupan-nilai-col" style="font-size:0.75rem; width:125px; margin:0 auto;" placeholder="Materi" value="${sCakupan[i] || ""}" onchange="saveNilai(true)">
                        ${colCount > 1 ? `<button type="button" class="btn btn-sm btn-outline-danger mt-1 py-0 px-1" onclick="deleteNilaiColumn(${i})" title="Hapus Kolom Ini"><i class="bi bi-trash"></i></button>` : ''}
                    </th>`;
            }
            html += '</tr></thead><tbody>';
            stds.forEach(s => {
                html += `<tr><td class="fw-bold d-flex align-items-center justify-content-between">${escapeHtml(s.StudentName)}
                    <button class="btn btn-sm btn-outline-success py-0 px-1" onclick="sendWANilai('${s.StudentID}')" title="Kirim WA Orang Tua"><i class="bi bi-whatsapp"></i></button>
                    </td>`;
                let grades = sGrades[s.StudentID] || Array(colCount).fill('');
                for (let i = 0; i < colCount; i++) html += `<td><input type="number" class="form-control form-control-sm px-1 text-center input-nilai-val" style="width:60px; margin:0 auto;" data-sid="${s.StudentID}" value="${grades[i] || ''}" oninput="calculateStatsNilai()"></td>`;
                html += '</tr>';
            });
            html += '</tbody></table>'; document.getElementById('nilai-bulan').disabled = false;
        } else {
            document.getElementById('btn-add-nilai-col').style.display = 'none';
            const jenisLabel = jenis === 'SumatifLM' ? 'Penilaian Tengah Semester' : (jenis === 'SumatifAS' ? 'Penilaian Akhir Semester' : jenis);
            html += `<table class="table table-bordered" id="table-nilai-input"><thead class="table-primary"><tr><th width="30%">Nama Siswa</th><th>Nilai ${jenisLabel}</th></tr></thead><tbody>`;
            stds.forEach(s => {
                html += `<tr><td class="fw-bold d-flex align-items-center justify-content-between">${escapeHtml(s.StudentName)}
                    <button class="btn btn-sm btn-outline-success py-0 px-1" onclick="sendWANilai('${s.StudentID}')" title="Kirim WA Orang Tua"><i class="bi bi-whatsapp"></i></button>
                    </td><td><input type="number" class="form-control w-50 input-nilai-val" data-sid="${s.StudentID}" value="${sGrades[s.StudentID] ? sGrades[s.StudentID][0] : ''}" oninput="calculateStatsNilai()"></td></tr>`;
            });
            html += '</tbody></table>'; document.getElementById('nilai-bulan').disabled = true;
        }
        container.innerHTML = html; document.getElementById('nilai-stats-container').style.display = 'flex'; document.getElementById('nilai-chart-container').style.display = 'block'; calculateStatsNilai(true);
    }

    function addNilaiColumn() {
        const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value, jenis = document.getElementById('nilai-jenis').value, bulan = document.getElementById('nilai-bulan').value;
        if (!cid || !mapel) return Swal.fire('Error', 'Pilih kelas & mapel', 'error');
        if (jenis !== 'Formatif') return;
        saveNilai(true);
        let record = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === String(jenis) && String(g.MonthYear) === String(bulan) && matchTapel(g, getActiveTapel()));
        if (record) {
            let dates = JSON.parse(record.DatesData); dates.push(""); record.DatesData = JSON.stringify(dates);
            let caku = JSON.parse(record.Cakupan); caku.push(""); record.Cakupan = JSON.stringify(caku);
            let gd = JSON.parse(record.GradesData); for (let sid in gd) { gd[sid].push(""); } record.GradesData = JSON.stringify(gd); syncDB();
        }
        loadNilaiData();
    }

    function deleteNilaiColumn(idx) {
        const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value, jenis = document.getElementById('nilai-jenis').value, bulan = document.getElementById('nilai-bulan').value;
        if (!cid || !mapel || jenis !== 'Formatif') return;
        saveNilai(true); // pastikan perubahan yang belum tersimpan tidak ikut hilang
        let record = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === String(jenis) && String(g.MonthYear) === String(bulan) && matchTapel(g, getActiveTapel()));
        if (!record) return;
        let dates = []; try { dates = JSON.parse(record.DatesData); } catch (e) { }
        if (dates.length <= 1) return Swal.fire('Tidak Bisa', 'Minimal harus ada 1 kolom penilaian.', 'warning');

        Swal.fire({
            title: 'Hapus Kolom Ini?', text: 'Seluruh nilai siswa pada kolom ini akan dihapus.', icon: 'warning',
            showCancelButton: true, confirmButtonColor: '#C44747', confirmButtonText: 'Ya, Hapus!', cancelButtonText: 'Batal'
        }).then((result) => {
            if (!result.isConfirmed) return;
            dates.splice(idx, 1);
            record.DatesData = JSON.stringify(dates);
            let caku = []; try { caku = JSON.parse(record.Cakupan); } catch (e) { }
            caku.splice(idx, 1);
            record.Cakupan = JSON.stringify(caku);
            let gd = {}; try { gd = JSON.parse(record.GradesData); } catch (e) { }
            for (let sid in gd) { gd[sid].splice(idx, 1); }
            record.GradesData = JSON.stringify(gd);
            syncDB(); loadNilaiData();
            Swal.fire('Terhapus!', 'Kolom berhasil dihapus.', 'success');
        });
    }

    function calculateStatsNilai(isInit = false) {
        const cid = document.getElementById('nilai-kelas').value, jenis = document.getElementById('nilai-jenis').value, kkmValue = parseFloat(document.getElementById('input-kkm').value) || 75;
        let stdAvgs = [], names = [], scores = [], kkmNames = [];
        db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => {
            let sum = 0, count = 0; document.querySelectorAll(`.input-nilai-val[data-sid="${s.StudentID}"]`).forEach(inp => { let val = parseFloat(inp.value); if (!isNaN(val)) { sum += val; count++; } });
            let avg = count > 0 ? (sum / count) : 0; stdAvgs.push({ name: s.StudentName, avg: avg }); names.push(s.StudentName.split(' ')[0]); scores.push(avg.toFixed(1));
            if (count > 0 && avg < kkmValue) kkmNames.push(s.StudentName);
        });
        let highest = stdAvgs.reduce((max, obj) => (obj.avg > max.avg ? obj : max), { name: '-', avg: 0 });
        document.getElementById('stat-nilai-max-name').innerText = highest.avg > 0 ? highest.name : '-'; document.getElementById('stat-nilai-max-val').innerText = highest.avg > 0 ? `Nilai: ${highest.avg.toFixed(1)}` : 'Nilai: 0';
        document.getElementById('stat-nilai-kkm-count').innerText = `${kkmNames.length} Siswa`; document.getElementById('stat-nilai-kkm-names').innerText = kkmNames.length > 0 ? kkmNames.join(', ') : 'Semua tuntas!';
        if (chartPrestasi) chartPrestasi.destroy();
        chartPrestasi = new Chart(document.getElementById('chartNilai').getContext('2d'), { type: 'bar', data: { labels: names, datasets: [{ label: `Rata-rata ${jenis}`, data: scores, backgroundColor: 'rgba(49, 94, 145, 0.85)', borderColor: '#172B4D', borderWidth: 1, borderRadius: 6 }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 100 } } } });
        if (!isInit) saveNilai(true);
    }

    function saveNilai(isSilent = false) {
        const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value, jenis = document.getElementById('nilai-jenis').value, bulan = document.getElementById('nilai-bulan').value;
        if (!cid || !mapel) { if (!isSilent) Swal.fire('Error', 'Pilih Kelas & Mapel', 'error'); return; }
        const currentSemester = ['07', '08', '09', '10', '11', '12'].includes(String(bulan)) ? 'ganjil' : 'genap';
        let customDates = [], customCakupan = [], gradesData = {};
        if (jenis === 'Formatif') {
            document.querySelectorAll('.date-nilai-col').forEach(input => customDates.push(input.value)); document.querySelectorAll('.cakupan-nilai-col').forEach(input => customCakupan.push(input.value));
            db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => { let gArray = []; document.querySelectorAll(`.input-nilai-val[data-sid="${s.StudentID}"]`).forEach(inp => gArray.push(inp.value)); gradesData[s.StudentID] = gArray; });
        } else { db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => { let inp = document.querySelector(`.input-nilai-val[data-sid="${s.StudentID}"]`); gradesData[s.StudentID] = [inp ? inp.value : '']; }); }

        let idx = db.grades.findIndex(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === String(jenis) && (jenis === 'Formatif' ? String(g.MonthYear) === String(bulan) : isGradeMatchingSemester(g, currentSemester)) && matchTapel(g, getActiveTapel()));
        if (idx >= 0) { 
            db.grades[idx].Cakupan = JSON.stringify(customCakupan); 
            db.grades[idx].DatesData = JSON.stringify(customDates); 
            db.grades[idx].GradesData = JSON.stringify(gradesData); 
            db.grades[idx].TahunPelajaran = db.grades[idx].TahunPelajaran || getActiveTapel();
            if (jenis !== 'Formatif') {
                db.grades[idx].Semester = currentSemester;
                if (!db.grades[idx].MonthYear) db.grades[idx].MonthYear = currentSemester === 'ganjil' ? '12' : '06';
            }
        } else {
            db.grades.push({ 
                RecordID: "G" + Date.now(), 
                ClassID: cid, 
                Subject: mapel, 
                Cakupan: JSON.stringify(customCakupan), 
                Type: jenis, 
                Semester: currentSemester,
                TahunPelajaran: getActiveTapel(),
                MonthYear: jenis === 'Formatif' ? bulan : (currentSemester === 'ganjil' ? '12' : '06'), 
                DatesData: JSON.stringify(customDates), 
                GradesData: JSON.stringify(gradesData) 
            });
        }
        syncDB(); if (!isSilent) Swal.fire({ icon: 'success', title: 'Tersimpan', text: 'Data disimpan!', timer: 1500, showConfirmButton: false });
    }

    function toggleDownloadNilaiDropdown(e) {
        if (e) {
            e.preventDefault();
            e.stopPropagation();
        }
        const menu = document.getElementById('menuDropdownDownloadNilai');
        if (!menu) return;
        const isShown = menu.classList.contains('show');
        document.querySelectorAll('.dropdown-menu.show').forEach(el => el.classList.remove('show'));
        if (!isShown) {
            menu.classList.add('show');
        }
    }

    function closeDownloadNilaiDropdown() {
        const menu = document.getElementById('menuDropdownDownloadNilai');
        if (menu) menu.classList.remove('show');
    }

    document.addEventListener('click', function(e) {
        const menu = document.getElementById('menuDropdownDownloadNilai');
        const btn = document.getElementById('btnDropdownDownloadNilai');
        if (menu && menu.classList.contains('show')) {
            if (!menu.contains(e.target) && e.target !== btn && !btn.contains(e.target)) {
                menu.classList.remove('show');
            }
        }
    });

    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            closeDownloadNilaiDropdown();
        }
    });

    async function downloadNilaiExcel(semester) {
        saveNilai(true); const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value;
        if (!cid || !mapel) return Swal.fire('Error', 'Pilih kelas & mapel', 'error');
        const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
        Swal.fire({ icon: 'info', title: 'Mengekspor', text: `Menyiapkan Excel...`, timer: 1500 });
        await ensureXLSX();
        const wb = XLSX.utils.book_new(), stds = db.students.filter(s => String(s.ClassID) === String(cid));
        const targetMonths = semester === 'ganjil' ? ['07', '08', '09', '10', '11', '12'] : ['01', '02', '03', '04', '05', '06'];
        const monthNames = { '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'Mei', '06': 'Jun', '07': 'Jul', '08': 'Agu', '09': 'Sep', '10': 'Okt', '11': 'Nov', '12': 'Des' };
        let rekapSemester = {}; stds.forEach(s => rekapSemester[s.StudentID] = { Name: s.StudentName, SumFormatif: 0, CountFormatif: 0, SLM: 0, SAS: 0 });

        targetMonths.forEach(m => {
            const record = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'Formatif' && String(g.MonthYear) === String(m) && matchTapel(g, getActiveTapel()));
            let sDates = [], sGrades = {}, sCakupan = [];
            if (record) { try { sDates = JSON.parse(record.DatesData) || []; sGrades = JSON.parse(record.GradesData) || {}; sCakupan = JSON.parse(record.Cakupan) || []; } catch (e) { } }
            let maxCols = sDates.length > 0 ? sDates.length : 1;
            let headDates = Array.from({ length: maxCols }, (_, i) => sDates[i] || `N-${i + 1}`);
            let headCakupan = Array.from({ length: maxCols }, (_, i) => sCakupan[i] || '-');
            let ws_data = [[`Mapel: ${mapel}`], [], ["Tanggal", ...headDates, ""], ["Cakupan", ...headCakupan, "Rata-rata"]];
            stds.forEach(s => {
                let row = [s.StudentName], grades = sGrades[s.StudentID] || Array(maxCols).fill(''), sum = 0, count = 0;
                for (let i = 0; i < maxCols; i++) {
                    let gradeVal = grades[i] || '';
                    row.push(gradeVal);
                    if (gradeVal !== "" && !isNaN(parseFloat(gradeVal))) { let num = parseFloat(gradeVal); sum += num; count++; rekapSemester[s.StudentID].SumFormatif += num; rekapSemester[s.StudentID].CountFormatif++; }
                }
                row.push(count > 0 ? (sum / count).toFixed(1) : 0); ws_data.push(row);
            });
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ws_data), monthNames[m]);
        });

        const recSLM = db.grades.slice().reverse().find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifLM' && isGradeMatchingSemester(g, semester) && matchTapel(g, getActiveTapel())); let slmData = {}; if (recSLM) try { slmData = JSON.parse(recSLM.GradesData); } catch (e) { }
        const recSAS = db.grades.slice().reverse().find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifAS' && isGradeMatchingSemester(g, semester) && matchTapel(g, getActiveTapel())); let sasData = {}; if (recSAS) try { sasData = JSON.parse(recSAS.GradesData); } catch (e) { }
        let rekap_data = [[`Rapor - ${mapel}`], [], ["Nama Siswa", "Rata Penilaian Harian", "Penilaian Tengah Semester (PTS)", "Penilaian Akhir Semester (PAS)", "Nilai Rapor"]];
        stds.forEach(s => {
            let r = rekapSemester[s.StudentID], avg = r.CountFormatif > 0 ? (r.SumFormatif / r.CountFormatif) : 0;
            let slm = parseFloat(slmData[s.StudentID]?.[0] || 0), sas = parseFloat(sasData[s.StudentID]?.[0] || 0);
            rekap_data.push([r.Name, avg.toFixed(1), slm.toFixed(1), sas.toFixed(1), ((avg + slm + sas) / 3).toFixed(1)]);
        });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rekap_data), "Rekap"); XLSX.writeFile(wb, `Nilai_${mapel}_${semester}.xlsx`);
    }

    async function downloadNilaiPDF() {
        saveNilai(true); const cid = document.getElementById('nilai-kelas').value; const mapel = document.getElementById('nilai-mapel').value;
        if (!cid || !mapel) return Swal.fire('Error', 'Pilih kelas & mapel', 'error');
        await ensurePDF();
        const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
        let rekapSemester = {}; db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => rekapSemester[s.StudentID] = { Name: s.StudentName, SumFormatif: 0, CountFormatif: 0 });

        ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'].forEach(m => {
            const record = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'Formatif' && String(g.MonthYear) === String(m) && matchTapel(g, getActiveTapel()));
            if (record) {
                let sG = {}; try { sG = JSON.parse(record.GradesData); } catch (e) { }
                db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => {
                    let grades = sG[s.StudentID] || [];
                    for (let i = 0; i < grades.length; i++) {
                        if (grades[i] !== "" && !isNaN(parseFloat(grades[i]))) {
                            rekapSemester[s.StudentID].SumFormatif += parseFloat(grades[i]); rekapSemester[s.StudentID].CountFormatif++;
                        }
                    }
                });
            }
        });

        const bulanVal = document.getElementById('nilai-bulan')?.value || '07';
        const activeSem = ['07', '08', '09', '10', '11', '12'].includes(String(bulanVal)) ? 'ganjil' : 'genap';
        const recSLM = db.grades.slice().reverse().find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifLM' && isGradeMatchingSemester(g, activeSem) && matchTapel(g, getActiveTapel())); let slmD = {}; if (recSLM) try { slmD = JSON.parse(recSLM.GradesData); } catch (e) { }
        const recSAS = db.grades.slice().reverse().find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifAS' && isGradeMatchingSemester(g, activeSem) && matchTapel(g, getActiveTapel())); let sasD = {}; if (recSAS) try { sasD = JSON.parse(recSAS.GradesData); } catch (e) { }

        const { jsPDF } = window.jspdf; const doc = new jsPDF();
        doc.setFontSize(14); doc.setFont("helvetica", "bold"); doc.text(`Rekapitulasi Nilai - ${mapel}`, 14, 16); doc.setFontSize(11); doc.setFont("helvetica", "normal"); doc.text(`Kelas: ${className}`, 14, 22);
        const tableData = db.students.filter(s => String(s.ClassID) === String(cid)).map(s => {
            let r = rekapSemester[s.StudentID], avg = r.CountFormatif > 0 ? (r.SumFormatif / r.CountFormatif) : 0;
            let slm = parseFloat(slmD[s.StudentID]?.[0] || 0), sas = parseFloat(sasD[s.StudentID]?.[0] || 0);
            return [s.StudentName, avg.toFixed(1), slm.toFixed(1), sas.toFixed(1), ((avg + slm + sas) / 3).toFixed(1)];
        });
        doc.autoTable({ head: [['Nama Siswa', 'Rata PH', 'PTS', 'PAS', 'Nilai Rapor']], body: tableData, startY: 28, headStyles: { fillColor: [23, 43, 77] } });
        doc.save(`Rekap_${mapel}.pdf`);
    }

    // --- Bank Data Standar Capaian & Tujuan Pembelajaran (TP) Kurikulum Merdeka ---
    const DEFAULT_TP_BANK = [
        // Pendidikan Pancasila
        { id: "TP-PKN-01", subject: "Pendidikan Pancasila", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "memahami makna sila-sila Pancasila dan penerapannya dalam kehidupan sehari-hari" },
        { id: "TP-PKN-02", subject: "Pendidikan Pancasila", gradeLevel: "Umum", code: "TP 2", semester: "ganjil", desc: "mengidentifikasi norma, hak, dan kewajiban sebagai anggota keluarga dan warga sekolah" },
        { id: "TP-PKN-03", subject: "Pendidikan Pancasila", gradeLevel: "Umum", code: "TP 3", semester: "genap", desc: "menghargai keragaman identitas suku bangsa, budaya, dan agama di Indonesia" },
        { id: "TP-PKN-04", subject: "Pendidikan Pancasila", gradeLevel: "Umum", code: "TP 4", semester: "genap", desc: "menunjukkan sikap gotong royong dan musyawarah dalam mufakat" },
        // Bahasa Indonesia
        { id: "TP-BIN-01", subject: "Bahasa Indonesia", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "menemukan dan menjelaskan ide pokok serta informasi penting pada teks narasi/deskripsi" },
        { id: "TP-BIN-02", subject: "Bahasa Indonesia", gradeLevel: "Umum", code: "TP 2", semester: "ganjil", desc: "menulis teks laporan hasil pengamatan atau pengalaman dengan kaidah bahasa yang tepat" },
        { id: "TP-BIN-03", subject: "Bahasa Indonesia", gradeLevel: "Umum", code: "TP 3", semester: "genap", desc: "mempresentasikan pendapat dengan intonasi jelas dan runtut" },
        { id: "TP-BIN-04", subject: "Bahasa Indonesia", gradeLevel: "Umum", code: "TP 4", semester: "genap", desc: "memaknai kosakata baru dan ungkapan kiasan dalam teks bacaan" },
        // Matematika
        { id: "TP-MAT-01", subject: "Matematika", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "memahami konsep bilangan cacah dan melakukan operasi hitung penjumlahan, pengurangan, perkalian, serta pembagian" },
        { id: "TP-MAT-02", subject: "Matematika", gradeLevel: "Umum", code: "TP 2", semester: "ganjil", desc: "menentukan dan membandingkan pecahan senilai serta melakukan operasi hitung pecahan sederhana" },
        { id: "TP-MAT-03", subject: "Matematika", gradeLevel: "Umum", code: "TP 3", semester: "genap", desc: "menghitung keliling dan luas bangun datar sederhana (persegi, persegi panjang, segitiga)" },
        { id: "TP-MAT-04", subject: "Matematika", gradeLevel: "Umum", code: "TP 4", semester: "genap", desc: "mengumpulkan, menyajikan, dan menafsirkan data dalam bentuk diagram batang dan tabel" },
        // IPAS / IPA
        { id: "TP-IPA-01", subject: "IPAS", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "menganalisis hubungan bentuk serta fungsi bagian tubuh pada tumbuhan dan hewan" },
        { id: "TP-IPA-02", subject: "IPAS", gradeLevel: "Umum", code: "TP 2", semester: "ganjil", desc: "menjelaskan proses siklus air dan pengaruhnya terhadap peristiwa di bumi" },
        { id: "TP-IPA-03", subject: "IPAS", gradeLevel: "Umum", code: "TP 3", semester: "genap", desc: "mengidentifikasi ragam bentang alam serta interaksi manusia dengan lingkungan sekitar" },
        { id: "TP-IPA-04", subject: "IPAS", gradeLevel: "Umum", code: "TP 4", semester: "genap", desc: "menganalisis bentuk perubahan wujud zat dan transformasi energi dalam kehidupan sehari-hari" },
        { id: "TP-IPA-05", subject: "IPA", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "memahami sistem organ pernapasan, pencernaan, dan peredaran darah manusia" },
        { id: "TP-IPA-06", subject: "IPA", gradeLevel: "Umum", code: "TP 2", semester: "genap", desc: "menyelidiki sifat-sifat cahaya, bunyi, magnet, dan listrik melalui percobaan sederhana" },
        // IPS
        { id: "TP-IPS-01", subject: "IPS", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "memahami kondisi geografis wilayah dan kegiatan pemanfaatan sumber daya alam" },
        { id: "TP-IPS-02", subject: "IPS", gradeLevel: "Umum", code: "TP 2", semester: "genap", desc: "mengidentifikasi peninggalan sejarah dan nilai-nilai perjuangan bangsa" },
        // Pendidikan Agama Islam (PAI)
        { id: "TP-PAI-01", subject: "Pendidikan Agama Islam", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "membaca, menulis, dan memahami pesan pokok surat-surat pendek Al-Qur'an" },
        { id: "TP-PAI-02", subject: "Pendidikan Agama Islam", gradeLevel: "Umum", code: "TP 2", semester: "ganjil", desc: "mempraktikkan tata cara shalat, bersuci (thaharah), dan doa harian" },
        { id: "TP-PAI-03", subject: "Pendidikan Agama Islam", gradeLevel: "Umum", code: "TP 3", semester: "genap", desc: "menunjukkan perilaku terpuji, amanah, dan berbakti kepada orang tua serta guru" },
        // Pendidikan Agama dan Budi Pekerti
        { id: "TP-AGM-01", subject: "Pendidikan Agama dan Budi Pekerti", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "memahami ajaran pokok agama dan menerapkannya dalam ibadah sehari-hari" },
        { id: "TP-AGM-02", subject: "Pendidikan Agama dan Budi Pekerti", gradeLevel: "Umum", code: "TP 2", semester: "genap", desc: "menunjukkan akhlak mulia, toleransi antarumat beragama, dan kepedulian terhadap sesama" },
        // Bahasa Inggris
        { id: "TP-ENG-01", subject: "Bahasa Inggris", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "mengidentifikasi kosakata dan frasa sederhana tentang aktivitas harian dan benda di sekitar" },
        { id: "TP-ENG-02", subject: "Bahasa Inggris", gradeLevel: "Umum", code: "TP 2", semester: "genap", desc: "merespons instruksi lisan serta berkomunikasi menggunakan kalimat sederhana" },
        // PJOK
        { id: "TP-PJK-01", subject: "PJOK", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "mempraktikkan variasi pola gerak dasar lokomotor, non-lokomotor, dan manipulatif dalam olahraga" },
        { id: "TP-PJK-02", subject: "PJOK", gradeLevel: "Umum", code: "TP 2", semester: "genap", desc: "menerapkan perilaku hidup bersih dan sehat serta menjaga kebersihan lingkungan" },
        // Seni Budaya / Seni Rupa / Seni Musik
        { id: "TP-SNB-01", subject: "Seni Rupa", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "mengeksplorasi unsur visual (garis, bentuk, warna) dalam membuat karya seni rupa dua dan tiga dimensi" },
        { id: "TP-SNM-01", subject: "Seni Musik", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "mengenal dan menyanyikan lagu-lagu wajib nasional serta daerah dengan nada dan tempo yang tepat" },
        // Informatika
        { id: "TP-INF-01", subject: "Informatika", gradeLevel: "Umum", code: "TP 1", semester: "ganjil", desc: "mengenal perangkat teknologi informasi dan memanfaatkan aplikasi dasar untuk belajar" },
        { id: "TP-INF-02", subject: "Informatika", gradeLevel: "Umum", code: "TP 2", semester: "genap", desc: "menerapkan berpikir komputasional dalam menyelesaikan persoalan sehari-hari secara logis" }
    ];

    // Koleksi Teks Capaian Pembelajaran (CP) 1 Tahun Pelajaran Kurikulum Merdeka
    const DEFAULT_CP_MAP = {
        'Matematika': "Peserta didik mampu memahami konsep bilangan cacah dan melakukan operasi hitung penjumlahan, pengurangan, perkalian, serta pembagian. Peserta didik dapat menentukan dan membandingkan pecahan senilai serta melakukan operasi hitung pecahan sederhana. Peserta didik mampu menghitung keliling dan luas bangun datar sederhana serta menyelesaikan masalah kontekstual. Peserta didik dapat mengumpulkan, menyajikan, dan menafsirkan data dalam bentuk diagram batang dan tabel frekuensi.",
        'Bahasa Indonesia': "Peserta didik mampu menemukan dan menjelaskan ide pokok serta informasi penting pada teks narasi dan deskripsi. Peserta didik dapat menulis teks laporan hasil pengamatan atau pengalaman dengan kaidah bahasa yang tepat. Peserta didik mampu mempresentasikan pendapat dengan intonasi jelas dan struktur kalimat yang runtut. Peserta didik mampu memaknai kosakata baru dan ungkapan kiasan dalam teks bacaan fiksi dan nonfiksi.",
        'Pendidikan Pancasila': "Peserta didik mampu memahami makna sila-sila Pancasila dan penerapannya dalam kehidupan sehari-hari di rumah dan sekolah. Peserta didik dapat mengidentifikasi norma, hak, dan kewajiban sebagai anggota keluarga dan warga sekolah. Peserta didik menghargai keragaman identitas suku bangsa, budaya, dan agama di Indonesia. Peserta didik menunjukkan sikap gotong royong dan musyawarah dalam mufakat untuk menyelesaikan persoalan bersama.",
        'Ilmu Pengetahuan Alam': "Peserta didik mampu menganalisis hubungan bentuk serta fungsi bagian tubuh pada tumbuhan dan hewan. Peserta didik dapat menjelaskan proses siklus air dan pengaruhnya terhadap peristiwa di bumi. Peserta didik mampu mengidentifikasi ragam bentang alam serta interaksi manusia dengan lingkungan sekitar. Peserta didik mampu menganalisis bentuk perubahan wujud zat dan transformasi energi dalam kehidupan sehari-hari.",
        'IPAS': "Peserta didik mampu menganalisis hubungan bentuk serta fungsi bagian tubuh pada tumbuhan dan hewan. Peserta didik dapat menjelaskan proses siklus air dan pengaruhnya terhadap peristiwa di bumi. Peserta didik mampu mengidentifikasi ragam bentang alam serta interaksi manusia dengan lingkungan sekitar. Peserta didik mampu menganalisis bentuk perubahan wujud zat dan transformasi energi dalam kehidupan sehari-hari.",
        'Ilmu Pengetahuan Sosial': "Peserta didik mampu memahami kondisi geografis wilayah dan kegiatan pemanfaatan sumber daya alam secara berkelanjutan. Peserta didik dapat mengidentifikasi peninggalan sejarah dan nilai-nilai perjuangan para pahlawan bangsa. Peserta didik mampu menganalisis interaksi sosial, budaya, dan ekonomi masyarakat dalam kehidupan berbangsa.",
        'Pendidikan Agama Islam': "Peserta didik mampu membaca, menulis, dan memahami pesan pokok surat-surat pendek Al-Qur'an. Peserta didik dapat mempraktikkan tata cara shalat, bersuci (thaharah), dan doa harian dengan benar. Peserta didik menunjukkan perilaku terpuji, jujur, amanah, dan berbakti kepada orang tua serta guru.",
        'Pendidikan Agama dan Budi Pekerti': "Peserta didik mampu memahami ajaran pokok agama dan menerapkannya dalam ibadah sehari-hari. Peserta didik dapat mempraktikkan doa, tata ibadah, dan pembiasaan religius dengan penuh penghayatan. Peserta didik menunjukkan akhlak mulia, toleransi antarumat beragama, dan kepedulian terhadap sesama.",
        'Bahasa Inggris': "Peserta didik mampu memahami instruksi lisan dan teks deskriptif sederhana tentang lingkungan sekitar. Peserta didik dapat memperkenalkan diri, berinteraksi, dan mengungkapkan perasaan menggunakan ungkapan komunikatif sederhana. Peserta didik mampu membaca cerita pendek bergambar dan menulis kalimat sederhana secara tepat.",
        'Pendidikan Jasmani, Olahraga, dan Kesehatan': "Peserta didik mampu mempraktikkan variasi pola gerak dasar lokomotor, non-lokomotor, dan manipulatif dalam olahraga. Peserta didik dapat menerapkan perilaku hidup bersih dan sehat serta menjaga kebersihan lingkungan. Peserta didik menunjukkan sikap sportivitas, disiplin, dan kerja sama dalam permainan aktivitas fisik.",
        'PJOK': "Peserta didik mampu mempraktikkan variasi pola gerak dasar lokomotor, non-lokomotor, dan manipulatif dalam olahraga. Peserta didik dapat menerapkan perilaku hidup bersih dan sehat serta menjaga kebersihan lingkungan. Peserta didik menunjukkan sikap sportivitas, disiplin, dan kerja sama dalam permainan aktivitas fisik.",
        'Informatika': "Peserta didik mampu mengenal perangkat teknologi informasi dan memanfaatkan aplikasi pengolah data dasar untuk belajar. Peserta didik dapat menerapkan berpikir komputasional dalam menyelesaikan persoalan sederhana secara logis dan runtut.",
        'Seni Budaya': "Peserta didik mampu mengeksplorasi unsur visual, garis, bentuk, dan warna dalam karya seni rupa. Peserta didik dapat mengenal dan menyanyikan lagu-lagu wajib nasional serta daerah dengan nada dan tempo yang tepat."
    };

    function ensureTPBank() {
        if (!db.tpBank || !Array.isArray(db.tpBank) || db.tpBank.length === 0) {
            db.tpBank = JSON.parse(JSON.stringify(DEFAULT_TP_BANK));
        }
        if (!db.config) db.config = {};
        if (!db.config.cpData) db.config.cpData = {};

        // Pastikan setiap record TP memiliki semester default jika belum ada
        db.tpBank.forEach((t, i) => {
            if (!t.semester) {
                t.semester = (i % 2 === 0) ? 'ganjil' : 'genap';
            }
        });
    }

    // Fungsi pemecah teks Capaian Pembelajaran (CP) menjadi butir-butir TP
    function generateTPFromCPText(mapel, cpText, targetSemesterMode = 'split') {
        if (!cpText || !cpText.trim()) return [];

        const rawParts = cpText
            .split(/(?:\r?\n+|;\s*|\.\s+|\.$)/g)
            .map(s => s.trim())
            .filter(s => s.length >= 8);

        if (rawParts.length === 0) return [];

        let generated = [];
        const midpoint = Math.ceil(rawParts.length / 2);

        rawParts.forEach((part, idx) => {
            let clean = part
                .replace(/^\s*(?:[0-9]+|[a-z])[.\)]\s*/i, '')
                .replace(/^[•\-\*\s]+/, '')
                .replace(/^(peserta didik|siswa)\s+(mampu|dapat|memiliki kemampuan untuk|diharapkan dapat|terampil)\s+/i, '')
                .replace(/[\.\;]+$/, '')
                .trim();
            if (!clean) return;

            clean = clean.charAt(0).toLowerCase() + clean.slice(1);

            let sem = 'ganjil';
            if (targetSemesterMode === 'split') {
                sem = idx < midpoint ? 'ganjil' : 'genap';
            } else if (targetSemesterMode === 'genap') {
                sem = 'genap';
            } else if (targetSemesterMode === 'ganjil') {
                sem = 'ganjil';
            } else {
                sem = 'semua';
            }

            generated.push({
                id: `TP-${Date.now()}-${idx + 1}`,
                subject: mapel,
                gradeLevel: "Umum",
                code: `TP ${idx + 1}`,
                desc: clean,
                semester: sem,
                source: 'cp'
            });
        });

        return generated;
    }

    // Sinkronisasi Materi Formatif dari Data Nilai menjadi Butir TP
    function syncTPFromDataNilai(mapel) {
        if (!mapel) return { added: 0, ganjil: 0, genap: 0, totalGanjilMateri: 0, totalGenapMateri: 0 };

        let ganjilMateri = new Set();
        let genapMateri = new Set();

        db.grades.forEach(g => {
            if (String(g.Subject).toLowerCase() === String(mapel).toLowerCase() && String(g.Type) === 'Formatif' && g.Cakupan) {
                try {
                    let cakupan = JSON.parse(g.Cakupan);
                    if (Array.isArray(cakupan)) {
                        cakupan.forEach(c => {
                            let text = String(c || '').trim();
                            if (text && text !== '-' && text.length >= 2) {
                                const month = String(g.MonthYear || '');
                                if (['07', '08', '09', '10', '11', '12'].includes(month)) {
                                    ganjilMateri.add(text);
                                } else if (['01', '02', '03', '04', '05', '06'].includes(month)) {
                                    genapMateri.add(text);
                                } else {
                                    ganjilMateri.add(text);
                                }
                            }
                        });
                    }
                } catch (e) { }
            }
        });

        let addedCount = 0;
        let ganjilCount = 0;
        let genapCount = 0;

        Array.from(ganjilMateri).forEach(materi => {
            const desc = `memahami dan menguasai materi ${materi}`;
            const exists = db.tpBank.some(t => String(t.subject).toLowerCase() === String(mapel).toLowerCase() && (t.desc.toLowerCase() === desc.toLowerCase() || t.desc.toLowerCase().includes(materi.toLowerCase())));
            if (!exists) {
                const nextIdx = db.tpBank.filter(t => String(t.subject).toLowerCase() === String(mapel).toLowerCase()).length + 1;
                db.tpBank.push({
                    id: `TP-DN-${Date.now()}-${nextIdx}`,
                    subject: mapel,
                    gradeLevel: "Umum",
                    code: `TP ${nextIdx}`,
                    desc: desc,
                    semester: 'ganjil',
                    source: 'materi'
                });
                addedCount++;
                ganjilCount++;
            }
        });

        Array.from(genapMateri).forEach(materi => {
            const desc = `memahami dan menguasai materi ${materi}`;
            const exists = db.tpBank.some(t => String(t.subject).toLowerCase() === String(mapel).toLowerCase() && (t.desc.toLowerCase() === desc.toLowerCase() || t.desc.toLowerCase().includes(materi.toLowerCase())));
            if (!exists) {
                const nextIdx = db.tpBank.filter(t => String(t.subject).toLowerCase() === String(mapel).toLowerCase()).length + 1;
                db.tpBank.push({
                    id: `TP-DN-${Date.now()}-${nextIdx}`,
                    subject: mapel,
                    gradeLevel: "Umum",
                    code: `TP ${nextIdx}`,
                    desc: desc,
                    semester: 'genap',
                    source: 'materi'
                });
                addedCount++;
                genapCount++;
            }
        });

        if (addedCount > 0) syncDB();
        return { added: addedCount, ganjil: ganjilCount, genap: genapCount, totalGanjilMateri: ganjilMateri.size, totalGenapMateri: genapMateri.size };
    }

    // Event Handler saat Pilihan Mapel di Modal TP Berubah
    function onBankTPMapelChanged() {
        ensureTPBank();
        const mapel = document.getElementById('bank-tp-filter-mapel')?.value || '';
        const cpTextarea = document.getElementById('bank-cp-textarea');

        if (cpTextarea) {
            const savedCP = db.config?.cpData?.[mapel];
            if (savedCP) {
                cpTextarea.value = savedCP;
            } else if (DEFAULT_CP_MAP[mapel]) {
                cpTextarea.value = DEFAULT_CP_MAP[mapel];
            } else {
                cpTextarea.value = '';
            }
        }

        cancelEditTP();
        renderBankTPTable();
    }

    function loadDefaultCPExample() {
        const mapel = document.getElementById('bank-tp-filter-mapel')?.value || '';
        const cpTextarea = document.getElementById('bank-cp-textarea');
        if (!cpTextarea) return;

        let example = DEFAULT_CP_MAP[mapel];
        if (!example) {
            for (let k in DEFAULT_CP_MAP) {
                if (mapel.toLowerCase().includes(k.toLowerCase()) || k.toLowerCase().includes(mapel.toLowerCase())) {
                    example = DEFAULT_CP_MAP[k];
                    break;
                }
            }
        }

        if (example) {
            cpTextarea.value = example;
            Swal.fire({ icon: 'info', title: 'Contoh CP Dimuat', text: `Teks Capaian Pembelajaran standar 1 tahun untuk "${mapel}" berhasil dimuat ke editor.`, timer: 1500, showConfirmButton: false });
        } else {
            Swal.fire('Info', `Silakan ketik atau tempelkan teks Capaian Pembelajaran (CP) untuk "${mapel}".`, 'info');
        }
    }

    function saveCPOnly() {
        ensureTPBank();
        const mapel = document.getElementById('bank-tp-filter-mapel')?.value;
        const text = document.getElementById('bank-cp-textarea')?.value.trim();
        if (!mapel) return Swal.fire('Peringatan', 'Silakan pilih mata pelajaran terlebih dahulu.', 'warning');
        if (!text) return Swal.fire('Peringatan', 'Teks Capaian Pembelajaran (CP) masih kosong.', 'warning');

        if (!db.config.cpData) db.config.cpData = {};
        db.config.cpData[mapel] = text;
        syncDB();

        Swal.fire({
            icon: 'success',
            title: 'Tersimpan!',
            text: `Teks Capaian Pembelajaran (CP) 1 tahun untuk "${mapel}" berhasil disimpan.`,
            timer: 1500,
            showConfirmButton: false
        });
    }

    function handleGenerateTPFromCP() {
        ensureTPBank();
        const mapel = document.getElementById('bank-tp-filter-mapel')?.value;
        const cpText = document.getElementById('bank-cp-textarea')?.value.trim();
        const splitMode = document.getElementById('bank-cp-semester-split')?.value || 'split';

        if (!mapel) return Swal.fire('Peringatan', 'Pilih mata pelajaran terlebih dahulu.', 'warning');
        if (!cpText) return Swal.fire('Peringatan', 'Teks Capaian Pembelajaran (CP) tidak boleh kosong. Silakan tempelkan CP 1 tahun atau klik "Muat Contoh CP Kemendikbud".', 'warning');

        const generatedTPs = generateTPFromCPText(mapel, cpText, splitMode);
        if (generatedTPs.length === 0) {
            return Swal.fire('Gagal Memecah TP', 'Teks Capaian Pembelajaran belum dapat diurai menjadi butir TP. Pastikan terdapat kalimat atau tanda baca pemisah (. atau ;).', 'error');
        }

        if (!db.config.cpData) db.config.cpData = {};
        db.config.cpData[mapel] = cpText;

        const existingCount = db.tpBank.filter(t => String(t.subject).toLowerCase() === String(mapel).toLowerCase()).length;

        const applyGenerated = () => {
            db.tpBank = db.tpBank.filter(t => String(t.subject).toLowerCase() !== String(mapel).toLowerCase());
            generatedTPs.forEach(t => db.tpBank.push(t));

            syncDB();
            renderBankTPTable();

            const ganjilCount = generatedTPs.filter(t => t.semester === 'ganjil').length;
            const genapCount = generatedTPs.filter(t => t.semester === 'genap').length;

            Swal.fire({
                icon: 'success',
                title: 'TP Berhasil Di-generate!',
                html: `Berhasil mengekstrak <b>${generatedTPs.length}</b> Tujuan Pembelajaran untuk <b>${escapeHtml(mapel)}</b>:<br>
                       <span class="badge bg-primary mt-2 me-1">Semester Ganjil: ${ganjilCount} TP</span>
                       <span class="badge bg-success mt-2">Semester Genap: ${genapCount} TP</span><br>
                       <small class="text-muted d-block mt-2">TP ini kini siap digunakan untuk deskripsi capaian rapor.</small>`
            });
        };

        if (existingCount > 0) {
            Swal.fire({
                title: 'Perbarui TP dari CP?',
                text: `Mata pelajaran "${mapel}" sudah memiliki ${existingCount} TP di Bank Data. Apakah Anda ingin memperbarui dan menggantinya dengan hasil generate CP baru ini?`,
                icon: 'question',
                showCancelButton: true,
                confirmButtonText: 'Ya, Perbarui',
                cancelButtonText: 'Batal'
            }).then(res => {
                if (res.isConfirmed) applyGenerated();
            });
        } else {
            applyGenerated();
        }
    }

    function handleSyncTPFromDataNilai() {
        ensureTPBank();
        const mapel = document.getElementById('bank-tp-filter-mapel')?.value;
        if (!mapel) return Swal.fire('Peringatan', 'Silakan pilih mata pelajaran terlebih dahulu.', 'warning');

        const result = syncTPFromDataNilai(mapel);
        renderBankTPTable();

        if (result.added > 0) {
            Swal.fire({
                icon: 'success',
                title: 'Sinkronisasi Materi Berhasil!',
                html: `Berhasil menambahkan <b>${result.added}</b> TP baru dari topik materi Data Nilai:<br>
                       <span class="badge bg-primary mt-2 me-1">Semester Ganjil: ${result.ganjil} TP</span>
                       <span class="badge bg-success mt-2">Semester Genap: ${result.genap} TP</span>`
            });
        } else if (result.totalGanjilMateri > 0 || result.totalGenapMateri > 0) {
            Swal.fire('Info', `Semua materi (${result.totalGanjilMateri + result.totalGenapMateri} topik) dari Data Nilai "${mapel}" sudah ada di Bank TP.`, 'info');
        } else {
            Swal.fire({
                title: 'Belum Ada Materi Tercatat',
                text: `Belum ditemukan input nama materi pada kolom Formatif di menu Data Nilai untuk mapel "${mapel}". Anda dapat memasukkan materi terlebih dahulu di Data Nilai atau men-generate TP dari Capaian Pembelajaran (CP).`,
                icon: 'info'
            });
        }
    }

    // --- Modal Bank TP (Tujuan Pembelajaran) ---
    function openBankTPModal(defaultSubject) {
        ensureTPBank();
        const selectMapel = document.getElementById('bank-tp-filter-mapel');
        if (selectMapel) {
            let mapelList = (db.subjects && db.subjects.length > 0) ? [...db.subjects] : [
                'Pendidikan Agama dan Budi Pekerti',
                'Pendidikan Pancasila',
                'Bahasa Indonesia',
                'Matematika',
                'Ilmu Pengetahuan Alam',
                'Ilmu Pengetahuan Sosial',
                'Bahasa Inggris',
                'Pendidikan Jasmani, Olahraga, dan Kesehatan',
                'Informatika',
                'Seni Budaya'
            ];

            selectMapel.innerHTML = mapelList.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');

            const activeMapel = defaultSubject || document.getElementById('nilai-mapel')?.value || mapelList[0] || '';
            if (activeMapel && mapelList.includes(activeMapel)) {
                selectMapel.value = activeMapel;
            } else if (mapelList.length > 0) {
                selectMapel.value = mapelList[0];
            }
        }

        onBankTPMapelChanged();

        let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalBankTP'));
        if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalBankTP'));
        modalInstance.show();
    }

    function renderBankTPTable() {
        ensureTPBank();
        const curMapel = document.getElementById('bank-tp-filter-mapel')?.value || '';
        const filterSem = document.getElementById('bank-tp-table-filter-sem')?.value || 'all';
        const tbody = document.getElementById('bank-tp-tbody');
        const countBadge = document.getElementById('bank-tp-count-badge');
        if (!tbody) return;

        const curLower = curMapel.toLowerCase().trim();
        let filtered = db.tpBank.filter(t => {
            if (!t.subject) return false;
            const sLower = t.subject.toLowerCase().trim();
            if (sLower === curLower) return true;
            if (sLower === 'ipa' && (curLower.includes('ilmu pengetahuan alam') || curLower.includes('ipa'))) return true;
            if (sLower === 'ips' && (curLower.includes('ilmu pengetahuan sosial') || curLower.includes('ips'))) return true;
            if (sLower === 'pjok' && (curLower.includes('penjas') || curLower.includes('pjok') || curLower.includes('olahraga'))) return true;
            if (sLower === 'pendidikan agama islam' && (curLower.includes('agama') || curLower.includes('pai'))) return true;
            return false;
        });

        if (countBadge) countBadge.innerText = `${filtered.length} TP`;

        if (filterSem !== 'all') {
            filtered = filtered.filter(t => (t.semester && t.semester.toLowerCase() === filterSem) || t.semester === 'semua');
        }

        if (filtered.length === 0) {
            tbody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-4">Belum ada Tujuan Pembelajaran untuk mata pelajaran <b>${escapeHtml(curMapel)}</b> ${filterSem !== 'all' ? `(Semester ${filterSem})` : ''}. Silakan gunakan tombol <b>Generate TP Otomatis dari CP</b> di atas.</td></tr>`;
            return;
        }

        let html = '';
        filtered.forEach((t, idx) => {
            const sem = t.semester || 'ganjil';
            const semBadge = sem === 'ganjil' 
                ? '<span class="badge bg-primary">Ganjil</span>' 
                : (sem === 'genap' ? '<span class="badge bg-success">Genap</span>' : '<span class="badge bg-secondary">Semua</span>');

            html += `<tr>
                <td class="text-center fw-bold">${idx + 1}</td>
                <td class="text-center">${semBadge}</td>
                <td class="text-center"><span class="badge bg-light text-dark border px-2 py-1">${escapeHtml(t.code || `TP ${idx + 1}`)}</span></td>
                <td><span style="color: var(--navy-deep);">${escapeHtml(t.desc)}</span></td>
                <td class="text-center text-nowrap">
                    <button class="btn btn-sm btn-outline-warning me-1" onclick="editTPFromBank('${t.id}')" title="Edit TP"><i class="bi bi-pencil"></i></button>
                    <button class="btn btn-sm btn-outline-danger" onclick="deleteTPFromBank('${t.id}')" title="Hapus TP"><i class="bi bi-trash"></i></button>
                </td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    function saveTPToBank() {
        ensureTPBank();
        const curMapel = document.getElementById('bank-tp-filter-mapel')?.value;
        const codeInput = document.getElementById('bank-tp-input-code');
        const semInput = document.getElementById('bank-tp-input-semester');
        const descInput = document.getElementById('bank-tp-input-desc');
        const editIdInput = document.getElementById('bank-tp-edit-id');

        if (!curMapel) return Swal.fire('Peringatan', 'Silakan pilih mata pelajaran terlebih dahulu.', 'warning');
        const desc = (descInput?.value || '').trim();
        if (!desc) return Swal.fire('Peringatan', 'Deskripsi Tujuan Pembelajaran (TP) tidak boleh kosong.', 'warning');
        const code = (codeInput?.value || '').trim() || `TP ${(db.tpBank.filter(t => t.subject === curMapel).length + 1)}`;
        const sem = semInput ? semInput.value : 'ganjil';

        const editId = editIdInput ? editIdInput.value : '';
        if (editId) {
            const idx = db.tpBank.findIndex(t => t.id === editId);
            if (idx > -1) {
                db.tpBank[idx].code = code;
                db.tpBank[idx].semester = sem;
                db.tpBank[idx].desc = desc;
                db.tpBank[idx].subject = curMapel;
            }
            Swal.fire({ icon: 'success', title: 'Berhasil', text: 'Tujuan Pembelajaran berhasil diperbarui.', timer: 1200, showConfirmButton: false });
        } else {
            const newId = 'TP-' + Date.now();
            db.tpBank.push({
                id: newId,
                subject: curMapel,
                gradeLevel: 'Umum',
                code: code,
                semester: sem,
                desc: desc,
                source: 'manual'
            });
            Swal.fire({ icon: 'success', title: 'Berhasil', text: 'Tujuan Pembelajaran berhasil ditambahkan ke bank data.', timer: 1200, showConfirmButton: false });
        }

        syncDB();
        cancelEditTP();
        renderBankTPTable();

        const descMapel = document.getElementById('nilai-mapel')?.value;
        if (descMapel && descMapel.toLowerCase() === curMapel.toLowerCase()) {
            renderTPCheckboxesForMapel(descMapel);
        }
    }

    function editTPFromBank(id) {
        const item = db.tpBank.find(t => t.id === id);
        if (!item) return;

        const manualTabBtn = document.getElementById('tab-manual-btn');
        if (manualTabBtn) {
            let tabInst = bootstrap.Tab.getOrCreateInstance(manualTabBtn);
            tabInst.show();
        }

        document.getElementById('bank-tp-edit-id').value = item.id;
        document.getElementById('bank-tp-input-code').value = item.code || '';
        document.getElementById('bank-tp-input-desc').value = item.desc || '';
        if (document.getElementById('bank-tp-input-semester')) {
            document.getElementById('bank-tp-input-semester').value = item.semester || 'ganjil';
        }
        document.getElementById('bank-tp-form-title').innerHTML = '<i class="bi bi-pencil-square me-1 text-warning"></i> Edit Tujuan Pembelajaran (TP)';
        document.getElementById('btn-cancel-edit-tp').style.display = 'inline-block';
        document.getElementById('bank-tp-input-desc').focus();
    }

    function cancelEditTP() {
        const editIdInput = document.getElementById('bank-tp-edit-id');
        if (editIdInput) editIdInput.value = '';
        const codeInput = document.getElementById('bank-tp-input-code');
        if (codeInput) codeInput.value = '';
        const descInput = document.getElementById('bank-tp-input-desc');
        if (descInput) descInput.value = '';
        const semInput = document.getElementById('bank-tp-input-semester');
        if (semInput) semInput.value = 'ganjil';
        const titleEl = document.getElementById('bank-tp-form-title');
        if (titleEl) titleEl.innerHTML = '<i class="bi bi-plus-circle me-1"></i> Tambah / Edit Butir Tujuan Pembelajaran (TP)';
        const btnCancel = document.getElementById('btn-cancel-edit-tp');
        if (btnCancel) btnCancel.style.display = 'none';
    }

    function deleteTPFromBank(id) {
        Swal.fire({
            title: 'Hapus TP ini?',
            text: 'Tujuan Pembelajaran ini akan dihapus dari Bank Data.',
            icon: 'warning',
            showCancelButton: true,
            confirmButtonText: 'Ya, Hapus',
            cancelButtonText: 'Batal',
            confirmButtonColor: '#d33'
        }).then(res => {
            if (res.isConfirmed) {
                db.tpBank = db.tpBank.filter(t => t.id !== id);
                syncDB();
                renderBankTPTable();
                const descMapel = document.getElementById('nilai-mapel')?.value;
                if (descMapel) renderTPCheckboxesForMapel(descMapel);
                Swal.fire({ icon: 'success', title: 'Terhapus', text: 'TP berhasil dihapus.', timer: 1200, showConfirmButton: false });
            }
        });
    }

    function resetDefaultTPBank() {
        Swal.fire({
            title: 'Pulihkan TP Default?',
            text: 'Tujuan Pembelajaran standar Kemendikbudristek untuk semua mapel umum akan dimuat ulang ke Bank Data.',
            icon: 'question',
            showCancelButton: true,
            confirmButtonText: 'Ya, Pulihkan',
            cancelButtonText: 'Batal'
        }).then(res => {
            if (res.isConfirmed) {
                db.tpBank = JSON.parse(JSON.stringify(DEFAULT_TP_BANK));
                syncDB();
                renderBankTPTable();
                const descMapel = document.getElementById('nilai-mapel')?.value;
                if (descMapel) renderTPCheckboxesForMapel(descMapel);
                Swal.fire('Berhasil!', 'Bank Data TP standar Kurikulum Merdeka berhasil dipulihkan.', 'success');
            }
        });
    }

    // --- Integrasi TP di Modal Deskripsi Rapor ---
    function renderTPCheckboxesForMapel(mapel) {
        ensureTPBank();
        const container = document.getElementById('tp-checkboxes-container');
        if (!container) return;

        const matches = db.tpBank.filter(t => 
            t.subject && (
                t.subject.toLowerCase() === mapel.toLowerCase() || 
                mapel.toLowerCase().includes(t.subject.toLowerCase()) || 
                t.subject.toLowerCase().includes(mapel.toLowerCase())
            )
        );

        if (matches.length === 0) {
            container.innerHTML = `
                <div class="small text-muted py-1">
                    <i class="bi bi-info-circle me-1"></i> Belum ada TP untuk mata pelajaran <b>${escapeHtml(mapel)}</b> di Bank TP.
                    <button type="button" class="btn btn-outline-primary btn-xs ms-2 py-0 px-2" style="font-size:0.75rem;" onclick="openBankTPModal('${escapeHtml(mapel)}')">
                        <i class="bi bi-plus-circle me-1"></i> Tambah TP Mapel Ini
                    </button>
                </div>`;
            return;
        }

        let html = '';
        matches.forEach((t, idx) => {
            html += `
                <div class="form-check form-check-inline bg-white px-3 py-1 border rounded shadow-sm mb-1" style="cursor: pointer;">
                    <input class="form-check-input tp-checkbox" type="checkbox" id="tp-chk-${t.id}" value="${t.id}" checked onchange="onTPCheckboxChange()">
                    <label class="form-check-label small" for="tp-chk-${t.id}" style="cursor: pointer;">
                        <span class="badge bg-primary me-1">${escapeHtml(t.code || `TP ${idx + 1}`)}</span>
                        ${escapeHtml(t.desc)}
                    </label>
                </div>`;
        });
        container.innerHTML = html;
        updateMateriUtamaFromCheckedTP();
    }

    function selectAllModalTP(selectAll) {
        const checkboxes = document.querySelectorAll('.tp-checkbox');
        checkboxes.forEach(cb => cb.checked = selectAll);
        onTPCheckboxChange();
    }

    function onTPCheckboxChange() {
        updateMateriUtamaFromCheckedTP();
        generateDeskripsiRapor();
    }

    function updateMateriUtamaFromCheckedTP() {
        const checkedTPs = Array.from(document.querySelectorAll('.tp-checkbox:checked')).map(cb => {
            return db.tpBank.find(t => t.id === cb.value);
        }).filter(Boolean);

        const materiInput = document.getElementById('desc-materi-utama');
        if (!materiInput) return;

        if (checkedTPs.length > 0) {
            materiInput.value = checkedTPs.map(t => t.desc).join('; ');
        }
    }

    // --- Generator Deskripsi Capaian Rapor Otomatis (Kurikulum Merdeka) ---
    let currentDeskripsiData = [];

    function openDeskripsiNilaiModal() {
        ensureTPBank();
        const cid = document.getElementById('nilai-kelas').value;
        const mapel = document.getElementById('nilai-mapel').value;

        if (!cid || !mapel) {
            return Swal.fire('Peringatan', 'Silakan pilih Kelas dan Mata Pelajaran terlebih dahulu pada menu Data Nilai.', 'warning');
        }

        const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
        const subtitleEl = document.getElementById('modal-desc-subtitle');
        if (subtitleEl) subtitleEl.innerText = `${mapel} - ${className}`;

        renderTPCheckboxesForMapel(mapel);

        let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalDeskripsiNilai'));
        if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalDeskripsiNilai'));
        modalInstance.show();

        generateDeskripsiRapor();
    }

    function generateDeskripsiRapor() {
        ensureTPBank();
        const cid = document.getElementById('nilai-kelas').value;
        const mapel = document.getElementById('nilai-mapel').value;
        const materiTopic = (document.getElementById('desc-materi-utama')?.value || `materi ${mapel}`).trim();
        const kriteriaA = parseFloat(document.getElementById('desc-kriteria-a')?.value) || 85;
        const kriteriaB = parseFloat(document.getElementById('desc-kriteria-b')?.value) || 75;

        const tbody = document.getElementById('modal-desc-tbody');
        if (!tbody) return;

        const stds = db.students.filter(s => String(s.ClassID) === String(cid));
        if (stds.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted py-4">Belum ada siswa di kelas ini.</td></tr>';
            return;
        }

        // Ambil TP yang dicentang
        const checkedTPs = Array.from(document.querySelectorAll('.tp-checkbox:checked')).map(cb => {
            return db.tpBank.find(t => t.id === cb.value);
        }).filter(Boolean);

        // Hitung nilai akhir/rata-rata siswa untuk mapel ini
        const targetMonths = ['07', '08', '09', '10', '11', '12', '01', '02', '03', '04', '05', '06'];
        let stdStats = {};
        stds.forEach(s => { stdStats[s.StudentID] = { sum: 0, count: 0 }; });

        targetMonths.forEach(m => {
            const rec = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'Formatif' && String(g.MonthYear) === String(m) && matchTapel(g, getActiveTapel()));
            if (rec && rec.GradesData) {
                try {
                    const gd = JSON.parse(rec.GradesData);
                    stds.forEach(s => {
                        const vals = gd[s.StudentID] || [];
                        vals.forEach(val => {
                            const num = parseFloat(val);
                            if (!isNaN(num)) {
                                stdStats[s.StudentID].sum += num;
                                stdStats[s.StudentID].count++;
                            }
                        });
                    });
                } catch (e) { }
            }
        });

        // Ambil nilai PTS & PAS jika ada
        const recSLM = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifLM' && matchTapel(g, getActiveTapel()));
        let slmD = {}; if (recSLM) try { slmD = JSON.parse(recSLM.GradesData); } catch (e) { }
        const recSAS = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifAS' && matchTapel(g, getActiveTapel()));
        let sasD = {}; if (recSAS) try { sasD = JSON.parse(recSAS.GradesData); } catch (e) { }

        currentDeskripsiData = [];
        let html = '';

        stds.forEach((s, idx) => {
            const r = stdStats[s.StudentID];
            const avgPH = r.count > 0 ? (r.sum / r.count) : 0;
            const ptsVal = parseFloat(slmD[s.StudentID]?.[0] || 0);
            const pasVal = parseFloat(sasD[s.StudentID]?.[0] || 0);

            let finalScore = avgPH;
            if (ptsVal > 0 || pasVal > 0) {
                let divisor = 1;
                let total = avgPH;
                if (ptsVal > 0) { total += ptsVal; divisor++; }
                if (pasVal > 0) { total += pasVal; divisor++; }
                finalScore = total / divisor;
            }

            finalScore = Math.round(finalScore * 10) / 10;

            let kategori = '';
            let badgeClass = '';
            let deskripsi = '';

            // Format deskripsi capaian rapor Kurikulum Merdeka berbasis TP
            if (checkedTPs.length >= 2) {
                const tp1 = checkedTPs[0].desc;
                const tp2 = checkedTPs[1].desc;
                if (finalScore >= kriteriaA) {
                    kategori = 'Sangat Baik (A)';
                    badgeClass = 'bg-success text-white';
                    deskripsi = `Ananda ${s.StudentName} menunjukkan penguasaan sangat baik dalam ${tp1} serta ${tp2}.`;
                } else if (finalScore >= kriteriaB) {
                    kategori = 'Baik (B)';
                    badgeClass = 'bg-primary text-white';
                    deskripsi = `Ananda ${s.StudentName} menunjukkan penguasaan yang baik dalam ${tp1}, namun masih memerlukan pendampingan dalam ${tp2}.`;
                } else if (finalScore > 0) {
                    kategori = 'Perlu Bimbingan (C)';
                    badgeClass = 'bg-warning text-dark';
                    deskripsi = `Ananda ${s.StudentName} perlu bimbingan dan pendampingan lebih intensif dalam ${tp1} serta penguatan pemahaman pada ${tp2}.`;
                } else {
                    kategori = 'Belum Dinilai';
                    badgeClass = 'bg-secondary text-white';
                    deskripsi = `Belum ada data penilaian yang mencukupi untuk menguraikan capaian pembelajaran.`;
                }
            } else if (checkedTPs.length === 1) {
                const tpSingle = checkedTPs[0].desc;
                if (finalScore >= kriteriaA) {
                    kategori = 'Sangat Baik (A)';
                    badgeClass = 'bg-success text-white';
                    deskripsi = `Ananda ${s.StudentName} menunjukkan penguasaan sangat baik dalam ${tpSingle}.`;
                } else if (finalScore >= kriteriaB) {
                    kategori = 'Baik (B)';
                    badgeClass = 'bg-primary text-white';
                    deskripsi = `Ananda ${s.StudentName} menunjukkan penguasaan yang baik dalam ${tpSingle}.`;
                } else if (finalScore > 0) {
                    kategori = 'Perlu Bimbingan (C)';
                    badgeClass = 'bg-warning text-dark';
                    deskripsi = `Ananda ${s.StudentName} cukup memahami materi dasar, namun perlu bimbingan lebih intensif dalam ${tpSingle}.`;
                } else {
                    kategori = 'Belum Dinilai';
                    badgeClass = 'bg-secondary text-white';
                    deskripsi = `Belum ada data penilaian yang mencukupi.`;
                }
            } else {
                if (finalScore >= kriteriaA) {
                    kategori = 'Sangat Baik (A)';
                    badgeClass = 'bg-success text-white';
                    deskripsi = `Menunjukkan penguasaan yang sangat baik dalam memahami dan menerapkan ${materiTopic}, serta sangat aktif menyelesaikan setiap tugas pembelajaran.`;
                } else if (finalScore >= kriteriaB) {
                    kategori = 'Baik (B)';
                    badgeClass = 'bg-primary text-white';
                    deskripsi = `Menunjukkan penguasaan yang baik dan mampu memahami ${materiTopic} secara tepat sesuai dengan tujuan pembelajaran yang ditetapkan.`;
                } else if (finalScore > 0) {
                    kategori = 'Perlu Bimbingan (C)';
                    badgeClass = 'bg-warning text-dark';
                    deskripsi = `Cukup menguasai materi dasar, namun masih memerlukan pendampingan dan bimbingan lebih lanjut dalam mendalami ${materiTopic}.`;
                } else {
                    kategori = 'Belum Dinilai';
                    badgeClass = 'bg-secondary text-white';
                    deskripsi = `Belum ada data penilaian yang mencukupi untuk menguraikan deskripsi capaian pembelajaran.`;
                }
            }

            currentDeskripsiData.push({
                no: idx + 1,
                studentId: s.StudentID,
                nama: s.StudentName,
                nilai: finalScore,
                kategori: kategori,
                deskripsi: deskripsi
            });

            html += `<tr>
                <td class="text-center fw-bold">${idx + 1}</td>
                <td>
                    <span class="fw-bold" style="color: var(--navy-deep);">${escapeHtml(s.StudentName)}</span>
                </td>
                <td class="text-center fw-bold fs-6">${finalScore > 0 ? finalScore : '-'}</td>
                <td class="text-center">
                    <span class="badge ${badgeClass} px-2 py-1" style="font-size:0.75rem;">${kategori}</span>
                </td>
                <td>
                    <textarea class="form-control form-control-sm desc-text-row" data-idx="${idx}" rows="2" style="font-size:0.83rem;">${escapeHtml(deskripsi)}</textarea>
                </td>
                <td class="text-center text-nowrap">
                    <button class="btn btn-sm btn-outline-primary me-1" onclick="copyDeskripsiText(${idx})" title="Salin Deskripsi Siswa Ini">
                        <i class="bi bi-clipboard"></i> Salin
                    </button>
                    <button class="btn btn-sm btn-outline-danger" onclick="openRaporForStudent('${s.StudentID}')" title="Cetak Rapor Siswa Ini">
                        <i class="bi bi-printer"></i>
                    </button>
                </td>
            </tr>`;
        });

        tbody.innerHTML = html;
    }

    function copyDeskripsiText(idx) {
        const textareas = document.querySelectorAll('.desc-text-row');
        let textToCopy = '';
        if (textareas[idx]) {
            textToCopy = textareas[idx].value;
        } else if (currentDeskripsiData[idx]) {
            textToCopy = currentDeskripsiData[idx].deskripsi;
        }

        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(textToCopy).then(() => {
                Swal.fire({
                    icon: 'success',
                    title: 'Tersalin!',
                    text: 'Deskripsi capaian rapor berhasil disalin ke clipboard.',
                    timer: 1200,
                    showConfirmButton: false
                });
            });
        } else {
            const temp = document.createElement('textarea');
            temp.value = textToCopy;
            document.body.appendChild(temp);
            temp.select();
            document.execCommand('copy');
            document.body.removeChild(temp);
            Swal.fire({
                icon: 'success',
                title: 'Tersalin!',
                text: 'Deskripsi berhasil disalin.',
                timer: 1200,
                showConfirmButton: false
            });
        }
    }

    async function exportDeskripsiExcel() {
        await ensureXLSX();
        const cid = document.getElementById('nilai-kelas').value;
        const mapel = document.getElementById('nilai-mapel').value;
        const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;

        const textareas = document.querySelectorAll('.desc-text-row');
        textareas.forEach((ta, idx) => {
            if (currentDeskripsiData[idx]) {
                currentDeskripsiData[idx].deskripsi = ta.value;
            }
        });

        const wb = XLSX.utils.book_new();
        const header = [
            [`DESKRIPSI CAPAIAN NILAI RAPOR - KURIKULUM MERDEKA`],
            [`Mata Pelajaran: ${mapel} | Kelas: ${className}`],
            [],
            ['No', 'Nama Siswa', 'Nilai Akhir', 'Kategori Predikat', 'Deskripsi Capaian Pembelajaran Rapor']
        ];

        const rows = currentDeskripsiData.map(d => [
            d.no,
            d.nama,
            d.nilai,
            d.kategori,
            d.deskripsi
        ]);

        const wsData = [...header, ...rows];
        const ws = XLSX.utils.aoa_to_sheet(wsData);

        ws['!cols'] = [
            { wch: 6 },
            { wch: 28 },
            { wch: 12 },
            { wch: 22 },
            { wch: 75 }
        ];

        XLSX.utils.book_append_sheet(wb, ws, "Deskripsi Rapor");
        const safeMapel = mapel.replace(/[^a-zA-Z0-9_-]/g, '_');
        const safeClass = className.replace(/[^a-zA-Z0-9_-]/g, '_');
        XLSX.writeFile(wb, `Deskripsi_Rapor_${safeMapel}_${safeClass}.xlsx`);

        Swal.fire('Berhasil!', 'File Excel Deskripsi Rapor berhasil diunduh.', 'success');
    }

    // --- Cetak Laporan Capaian Belajar (Rapor Siswa) Kurikulum Merdeka (Standar Kertas A4 Resmi) ---
    function loadRaporViewData(targetClassId, targetStudentId) {
        ensureTPBank();
        const selectKelas = document.getElementById('rapor-kelas');
        if (selectKelas) {
            const curVal = selectKelas.value;
            selectKelas.innerHTML = db.classes.map(c => `<option value="${c.ClassID}">${escapeHtml(c.ClassName)}</option>`).join('');
            if (targetClassId && db.classes.some(c => String(c.ClassID) === String(targetClassId))) {
                selectKelas.value = targetClassId;
            } else if (curVal && db.classes.some(c => String(c.ClassID) === String(curVal))) {
                selectKelas.value = curVal;
            } else if (db.classes[0]) {
                selectKelas.value = db.classes[0].ClassID;
            }
        }

        const curDate = new Date();
        const curYear = curDate.getFullYear();
        const curMonth = curDate.getMonth() + 1;
        const academicYear = curMonth >= 7 ? `${curYear}/${curYear + 1}` : `${curYear - 1}/${curYear}`;
        const defaultSem = curMonth >= 7 ? 'ganjil' : 'genap';

        const elTapel = document.getElementById('rapor-tapel');
        if (elTapel && !elTapel.value) elTapel.value = getActiveTapel();

        // HTML memberi nilai bawaan "genap", jadi pakai penanda pilihan guru (bukan cek nilai kosong)
        const elSem = document.getElementById('rapor-semester');
        if (elSem && !elSem.dataset.userSet) elSem.value = defaultSem;

        // Data sekolah/penandatangan TIDAK diisi otomatis: guru mengisinya sendiri di Pengaturan Sekolah.
        if (!db.config.raporFase) db.config.raporFase = 'D';

        onRaporKelasChange(targetStudentId);
    }

    function openRaporForStudent(targetStudentId) {
        const modalEl = document.getElementById('modalDeskripsiNilai');
        if (modalEl) {
            let modalInstance = bootstrap.Modal.getInstance(modalEl);
            if (modalInstance) modalInstance.hide();
        }

        const student = db.students.find(s => String(s.StudentID) === String(targetStudentId));
        const targetClassId = student ? student.ClassID : (document.getElementById('nilai-kelas')?.value || (db.classes[0] ? db.classes[0].ClassID : ''));

        switchView('rapor');
        loadRaporViewData(targetClassId, targetStudentId);
    }

    function openCetakRaporModal(targetStudentId) {
        openRaporForStudent(targetStudentId);
    }

    function onRaporKelasChange(targetStudentId) {
        const cid = document.getElementById('rapor-kelas')?.value;
        const selectSiswa = document.getElementById('rapor-siswa');
        if (!selectSiswa) return;

        const stds = db.students.filter(s => String(s.ClassID) === String(cid));
        let html = '<option value="ALL">-- Cetak Seluruh Siswa di Kelas Ini (Massal / Batch) --</option>';
        stds.forEach(s => {
            const nisInfo = s.NIS ? ` [NIS: ${s.NIS}]` : (s.NISN ? ` [NISN: ${s.NISN}]` : '');
            html += `<option value="${s.StudentID}">${escapeHtml(s.StudentName)}${nisInfo}</option>`;
        });
        selectSiswa.innerHTML = html;

        if (targetStudentId && stds.some(s => String(s.StudentID) === String(targetStudentId))) {
            selectSiswa.value = targetStudentId;
        } else if (stds[0]) {
            selectSiswa.value = stds[0].StudentID;
        }

        onRaporStudentChange();
    }

    function onRaporStudentChange() {
        previewRaporSiswa();
    }

    function onRaporPeriodeChange() {
        const tapel = document.getElementById('rapor-tapel')?.value.trim();
        if (tapel) {
            db.config.academicYear = tapel;
            syncDB();
        }
        previewRaporSiswa();
    }

    function focusPilihSiswa() {
        const el = document.getElementById('rapor-siswa');
        if (el) {
            el.focus();
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    }

    // Modal Pengaturan Sekolah & Titimangsa Rapor
    function openModalPengaturanRapor() {
        document.getElementById('setting-rapor-sekolah').value = db.config.userSchool || '';
        document.getElementById('setting-rapor-alamat').value = db.config.schoolAddress || '';
        document.getElementById('setting-rapor-kota').value = db.config.schoolCity || '';
        document.getElementById('setting-rapor-fase').value = db.config.raporFase || 'D';
        document.getElementById('setting-rapor-tanggal').value = db.config.raporTanggal || tanggalHariIni();
        document.getElementById('setting-rapor-wali-nama').value = db.config.userName || '';
        document.getElementById('setting-rapor-wali-nip').value = db.config.userNIP || '';
        document.getElementById('setting-rapor-kepsek-nama').value = db.config.headmasterName || '';
        document.getElementById('setting-rapor-kepsek-nip').value = db.config.headmasterNIP || '';

        const modalEl = document.getElementById('modalPengaturanRapor');
        let modalInstance = bootstrap.Modal.getInstance(modalEl);
        if (!modalInstance) modalInstance = new bootstrap.Modal(modalEl);
        modalInstance.show();
    }

    function handleSavePengaturanRapor(e) {
        e.preventDefault();
        db.config.userSchool = document.getElementById('setting-rapor-sekolah').value.trim();
        db.config.schoolAddress = document.getElementById('setting-rapor-alamat').value.trim();
        db.config.schoolCity = document.getElementById('setting-rapor-kota').value.trim();
        db.config.raporFase = document.getElementById('setting-rapor-fase').value;
        db.config.raporTanggal = document.getElementById('setting-rapor-tanggal').value;
        db.config.userName = document.getElementById('setting-rapor-wali-nama').value.trim();
        db.config.userNIP = document.getElementById('setting-rapor-wali-nip').value.trim();
        db.config.headmasterName = document.getElementById('setting-rapor-kepsek-nama').value.trim();
        db.config.headmasterNIP = document.getElementById('setting-rapor-kepsek-nip').value.trim();

        syncDB();

        const modalEl = document.getElementById('modalPengaturanRapor');
        if (modalEl) {
            let modalInstance = bootstrap.Modal.getInstance(modalEl);
            if (modalInstance) modalInstance.hide();
        }

        previewRaporSiswa();
        Swal.fire({
            icon: 'success',
            title: 'Pengaturan Disimpan!',
            text: 'Identitas sekolah dan titimangsa berhasil diperbarui untuk seluruh rapor.',
            timer: 1500,
            showConfirmButton: false
        });
    }

    // Modal Catatan Wali Kelas & Keputusan
    function openModalCatatanWali(targetStudentId) {
        const cid = document.getElementById('rapor-kelas')?.value;
        const currentSid = targetStudentId || document.getElementById('rapor-siswa')?.value;
        const semester = document.getElementById('rapor-semester')?.value || 'genap';
        const tapel = document.getElementById('rapor-tapel')?.value.trim() || getActiveTapel();

        const stds = db.students.filter(s => String(s.ClassID) === String(cid));
        if (stds.length === 0) return Swal.fire('Informasi', 'Belum ada siswa di kelas ini.', 'info');

        const selectStudent = document.getElementById('modal-catatan-select-siswa');
        if (selectStudent) {
            selectStudent.innerHTML = stds.map(s => `<option value="${s.StudentID}">${escapeHtml(s.StudentName)}</option>`).join('');
            if (currentSid && currentSid !== 'ALL' && stds.some(s => String(s.StudentID) === String(currentSid))) {
                selectStudent.value = currentSid;
            } else {
                selectStudent.value = stds[0].StudentID;
            }
        }

        // Tampilkan informasi semester pada form keputusan
        const isGenap = semester === 'genap';
        const badgeSem = document.getElementById('badge-keputusan-sem');
        if (badgeSem) {
            badgeSem.innerText = isGenap ? 'Semester 2 (Genap) - Aktif' : 'Semester 1 (Ganjil) - Nonaktif';
            badgeSem.className = isGenap ? 'badge bg-success-subtle text-success border border-success-subtle' : 'badge bg-secondary-subtle text-muted border';
        }
        const inputKeputusan = document.getElementById('modal-keputusan-text');
        if (inputKeputusan) {
            inputKeputusan.disabled = !isGenap;
            if (!isGenap) inputKeputusan.placeholder = 'Keputusan tidak dicantumkan pada rapor semester Ganjil';
        }

        onModalCatatanSiswaChange();

        const modalEl = document.getElementById('modalCatatanWali');
        let modalInstance = bootstrap.Modal.getInstance(modalEl);
        if (!modalInstance) modalInstance = new bootstrap.Modal(modalEl);
        modalInstance.show();
    }

    function onModalCatatanSiswaChange() {
        const studentId = document.getElementById('modal-catatan-select-siswa')?.value;
        const cid = document.getElementById('rapor-kelas')?.value;
        const semester = document.getElementById('rapor-semester')?.value || 'genap';
        const tapel = document.getElementById('rapor-tapel')?.value.trim() || getActiveTapel();
        const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || '';

        const notes = getStudentRaporNotes(studentId, semester, tapel, className);
        document.getElementById('modal-catatan-text').value = notes.catatanWali || '';
        document.getElementById('modal-keputusan-text').value = notes.keputusan || '';
    }

    function handleSaveCatatanWali(e) {
        e.preventDefault();
        const studentId = document.getElementById('modal-catatan-select-siswa')?.value;
        const cid = document.getElementById('rapor-kelas')?.value;
        const semester = document.getElementById('rapor-semester')?.value || 'genap';
        const tapel = document.getElementById('rapor-tapel')?.value.trim() || getActiveTapel();
        const catatan = document.getElementById('modal-catatan-text')?.value.trim() || '';
        const keputusan = document.getElementById('modal-keputusan-text')?.value.trim() || '';

        if (!studentId) return;

        if (!db.raporNotes) db.raporNotes = [];
        const idx = db.raporNotes.findIndex(n => 
            String(n.StudentID) === String(studentId) && 
            String(n.Semester) === String(semester) && 
            String(n.TahunPelajaran) === String(tapel)
        );

        if (idx > -1) {
            db.raporNotes[idx].CatatanWali = catatan;
            db.raporNotes[idx].Keputusan = keputusan;
            db.raporNotes[idx].ClassID = cid;
        } else {
            db.raporNotes.push({
                RecordID: 'RN' + Date.now(),
                StudentID: studentId,
                ClassID: cid,
                Semester: semester,
                TahunPelajaran: tapel,
                CatatanWali: catatan,
                Keputusan: keputusan
            });
        }

        syncDB();

        const modalEl = document.getElementById('modalCatatanWali');
        if (modalEl) {
            let modalInstance = bootstrap.Modal.getInstance(modalEl);
            if (modalInstance) modalInstance.hide();
        }

        previewRaporSiswa();
        Swal.fire({
            icon: 'success',
            title: 'Tersimpan!',
            text: 'Catatan wali kelas dan keputusan siswa berhasil disimpan.',
            timer: 1500,
            showConfirmButton: false
        });
    }

    function getStudentRaporNotes(studentId, semester, tapel, className) {
        if (!db.raporNotes) db.raporNotes = [];
        const found = db.raporNotes.find(n => 
            String(n.StudentID) === String(studentId) && 
            String(n.Semester) === String(semester) && 
            String(n.TahunPelajaran) === String(tapel)
        );

        if (found) {
            return {
                catatanWali: found.CatatanWali || '',
                keputusan: found.Keputusan || ''
            };
        }

        // Rekomendasi default jika belum pernah disimpan guru
        let defKeputusan = '';
        if (semester === 'genap') {
            const cName = className || '';
            let nextClass = "KELAS 8";
            if (/VII\b|7\b/i.test(cName)) nextClass = "KELAS 8";
            else if (/VIII\b|8\b/i.test(cName)) nextClass = "KELAS 9";
            else if (/IX\b|9\b/i.test(cName)) nextClass = "LULUS";
            else if (/1\b/i.test(cName)) nextClass = "KELAS 2";
            else if (/2\b/i.test(cName)) nextClass = "KELAS 3";
            else if (/3\b/i.test(cName)) nextClass = "KELAS 4";
            else if (/4\b/i.test(cName)) nextClass = "KELAS 5";
            else if (/5\b/i.test(cName)) nextClass = "KELAS 6";
            else if (/6\b/i.test(cName)) nextClass = "LULUS";
            else nextClass = "KELAS BERIKUTNYA";

            defKeputusan = nextClass === 'LULUS' ? "LULUS" : `NAIK KE ${nextClass}`;
        }

        return {
            catatanWali: "Menunjukkan perilaku dan motivasi belajar yang baik, pertahankan dan terus tingkatkan prestasimu!",
            keputusan: defKeputusan
        };
    }

    function getSubjectCategory(subjName) {
        const s = (subjName || '').toLowerCase().trim();
        if (s.includes('seni tari') || s.includes('seni musik') || s.includes('seni rupa') || s.includes('seni teater') || s.includes('pilihan') || s.includes('prakarya')) {
            return 'Pilihan';
        }
        if (s.includes('muatan lokal') || s.includes('mulok') || s.includes('sunda') || s.includes('jawa') || s.includes('bali') || s.includes('madura') || s.includes('daerah') || s.includes('informatika') || s.includes('inggris') || s.includes('plhj') || s.includes('bahasa arab') || s.includes('lampung') || s.includes('banjar')) {
            return 'Muatan Lokal';
        }
        return 'Umum';
    }

    function getStudentAllGrades(cid, studentId, semester, tapel) {
        ensureTPBank();
        const targetMonths = semester === 'ganjil' ? ['07', '08', '09', '10', '11', '12'] : ['01', '02', '03', '04', '05', '06'];

        let subjectList = [...(db.subjects || [])];
        if (subjectList.length === 0) {
            subjectList = [
                'Pendidikan Agama dan Budi Pekerti',
                'Pendidikan Jasmani, Olahraga, dan Kesehatan',
                'Pendidikan Pancasila',
                'Bahasa Indonesia',
                'Matematika',
                'Ilmu Pengetahuan Alam',
                'Ilmu Pengetahuan Sosial',
                'Seni Budaya',
                'Informatika',
                'Bahasa Inggris'
            ];
        }

        db.grades.filter(g => String(g.ClassID) === String(cid)).forEach(g => {
            if (g.Subject && !subjectList.includes(g.Subject)) subjectList.push(g.Subject);
        });

        let results = [];
        subjectList.forEach(subj => {
            let sumF = 0, countF = 0;
            targetMonths.forEach(m => {
                const rec = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(subj) && String(g.Type) === 'Formatif' && String(g.MonthYear) === String(m) && matchTapel(g, tapel));
                if (rec && rec.GradesData) {
                    try {
                        const gd = JSON.parse(rec.GradesData);
                        const vals = gd[studentId] || [];
                        vals.forEach(v => {
                            if (v !== '' && v !== null && v !== undefined && !isNaN(parseFloat(v))) {
                                sumF += parseFloat(v);
                                countF++;
                            }
                        });
                    } catch (e) { }
                }
            });

            const avgPH = countF > 0 ? (sumF / countF) : 0;
            const recSLM = db.grades.slice().reverse().find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(subj) && String(g.Type) === 'SumatifLM' && isGradeMatchingSemester(g, semester) && matchTapel(g, tapel));
            let pts = 0;
            let hasPts = false;
            if (recSLM) {
                try {
                    const rawVal = JSON.parse(recSLM.GradesData)[studentId]?.[0];
                    if (rawVal !== '' && rawVal !== null && rawVal !== undefined && !isNaN(parseFloat(rawVal))) {
                        pts = parseFloat(rawVal);
                        hasPts = true;
                    }
                } catch (e) { }
            }

            const recSAS = db.grades.slice().reverse().find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(subj) && String(g.Type) === 'SumatifAS' && isGradeMatchingSemester(g, semester) && matchTapel(g, tapel));
            let pas = 0;
            let hasPas = false;
            if (recSAS) {
                try {
                    const rawVal = JSON.parse(recSAS.GradesData)[studentId]?.[0];
                    if (rawVal !== '' && rawVal !== null && rawVal !== undefined && !isNaN(parseFloat(rawVal))) {
                        pas = parseFloat(rawVal);
                        hasPas = true;
                    }
                } catch (e) { }
            }

            // Nilai akhir dari menu Pengolah Nilai (jika sudah disimpan) menggantikan rata-rata biasa
            const nilaiOlahan = typeof getNilaiOlahanSiswa === 'function' ? getNilaiOlahanSiswa(cid, subj, semester, tapel, studentId) : null;
            const hasAnyScore = countF > 0 || hasPts || hasPas || nilaiOlahan !== null;
            let finalScore = null;
            let capaian = '';

            if (hasAnyScore) {
                let divisor = 0;
                let total = 0;
                if (countF > 0) { total += avgPH; divisor++; }
                if (hasPts) { total += pts; divisor++; }
                if (hasPas) { total += pas; divisor++; }
                finalScore = divisor > 0 ? Math.round((total / divisor) * 10) / 10 : null;
                if (nilaiOlahan !== null) finalScore = nilaiOlahan;

                // Ambil Tujuan Pembelajaran (TP) yang sesuai dengan mata pelajaran dan semester rapor saat ini
                let tps = db.tpBank.filter(t => 
                    t.subject && (
                        t.subject.toLowerCase() === subj.toLowerCase() || 
                        subj.toLowerCase().includes(t.subject.toLowerCase()) || 
                        t.subject.toLowerCase().includes(subj.toLowerCase())
                    ) && (
                        !t.semester || t.semester === 'semua' || t.semester.toLowerCase() === String(semester).toLowerCase()
                    )
                );

                // Fallback jika belum ada TP khusus semester tersebut
                if (tps.length === 0) {
                    tps = db.tpBank.filter(t => 
                        t.subject && (
                            t.subject.toLowerCase() === subj.toLowerCase() || 
                            subj.toLowerCase().includes(t.subject.toLowerCase()) || 
                            t.subject.toLowerCase().includes(subj.toLowerCase())
                        )
                    );
                }

                if (finalScore !== null) {
                    if (finalScore >= 85) {
                        let bestTP = tps[0]?.desc || `memahami materi pokok ${subj}`;
                        let needTP = tps[1]?.desc || '';
                        capaian = `Mencapai Kompetensi dengan sangat baik dalam hal ${bestTP}`;
                        if (needTP && tps.length > 1) {
                            capaian += `\nPerlu peningkatan dalam hal ${needTP}`;
                        }
                    } else if (finalScore >= 75) {
                        let bestTP = tps[0]?.desc || `memahami materi pembelajaran ${subj}`;
                        let needTP = tps[1]?.desc || '';
                        capaian = `Mencapai Kompetensi dengan baik dalam hal ${bestTP}`;
                        if (needTP && tps.length > 1) {
                            capaian += `\nPerlu peningkatan dalam hal ${needTP}`;
                        }
                    } else if (finalScore > 0) {
                        let tpDesc = tps[0]?.desc || `materi ${subj}`;
                        capaian = `Cukup mencapai kompetensi dalam hal ${tpDesc}\nPerlu pendampingan dan latihan intensif untuk penguatan pemahaman`;
                    }
                }
            } else {
                capaian = 'Belum ada data penilaian';
            }

            results.push({
                subject: subj,
                finalScore: finalScore,
                hasScore: hasAnyScore,
                capaian: capaian,
                category: getSubjectCategory(subj)
            });
        });

        return results;
    }

    function getStudentAttendanceSummary(cid, studentId, semester, tapel) {
        const targetMonths = semester === 'ganjil' ? ['07', '08', '09', '10', '11', '12'] : ['01', '02', '03', '04', '05', '06'];
        let H = 0, S = 0, I = 0, A = 0;
        let recordsCount = 0;

        targetMonths.forEach(m => {
            const rec = db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(m) && matchTapel(a, tapel));
            if (rec && rec.AttendanceData) {
                try {
                    const ad = JSON.parse(rec.AttendanceData);
                    const list = ad[studentId] || [];
                    if (list.length > 0) recordsCount++;
                    list.forEach(v => {
                        if (v === 'H') H++;
                        else if (v === 'S') S++;
                        else if (v === 'I') I++;
                        else if (v === 'A') A++;
                    });
                } catch (e) { }
            }
        });

        return { H, S, I, A, hasRecords: recordsCount > 0 };
    }

    function getStudentKokuSummary(cid, studentId, semester) {
        if (!db.koku || db.koku.length === 0) {
            return { hasData: false, items: [], text: 'Belum ada data penilaian kokurikuler untuk periode ini.' };
        }

        const targetMonths = semester === 'ganjil' ? ['07', '08', '09', '10', '11', '12'] : ['01', '02', '03', '04', '05', '06'];
        const records = db.koku.filter(k => {
            if (String(k.ClassID) !== String(cid)) return false;
            if (k.Tanggal) {
                const month = k.Tanggal.split('-')[1];
                if (month && !targetMonths.includes(month)) return false;
            }
            return true;
        });

        let studentKokuList = [];
        records.forEach(rec => {
            if (!rec.KokuData) return;
            try {
                const kd = JSON.parse(rec.KokuData);
                const sData = kd[studentId];
                if (sData && sData.Kategori) {
                    let cleanNarasi = (sData.Narasi || '').replace(/<[^>]*>/g, '').trim();
                    studentKokuList.push({
                        projek: rec.Projek || 'Kokurikuler',
                        dimensi: rec.Dimensi ? JSON.parse(rec.Dimensi || '[]') : [],
                        kategori: sData.Kategori,
                        narasi: cleanNarasi
                    });
                }
            } catch (e) { }
        });

        if (studentKokuList.length === 0) {
            return { hasData: false, items: [], text: 'Belum ada data penilaian kokurikuler untuk periode ini.' };
        }

        let combinedText = studentKokuList.map(item => {
            let dimStr = item.dimensi && item.dimensi.length > 0 ? ` (Dimensi: ${item.dimensi.join(', ')})` : '';
            return `[${item.projek}]${dimStr} - Capaian: ${item.kategori}. ${item.narasi}`;
        }).join('\n\n');

        return { hasData: true, items: studentKokuList, text: combinedText };
    }

    function getStudentEkskulSummary(studentId, semester, tapel) {
        if (!db.ekskul) return [];
        return db.ekskul.filter(e => 
            String(e.StudentID) === String(studentId) && 
            (!semester || String(e.Semester).toLowerCase() === String(semester).toLowerCase()) &&
            (!tapel || !e.TahunPelajaran || String(e.TahunPelajaran).trim() === String(tapel).trim())
        );
    }

    function previewRaporSiswa() {
        const cid = document.getElementById('rapor-kelas')?.value;
        const studentId = document.getElementById('rapor-siswa')?.value;
        const semester = document.getElementById('rapor-semester')?.value || 'genap';
        const tapel = document.getElementById('rapor-tapel')?.value.trim() || getActiveTapel();

        const fase = db.config.raporFase || 'D';
        const kota = db.config.schoolCity || RAPOR_KOSONG;
        const tglStr = db.config.raporTanggal || tanggalHariIni();
        const sekolah = db.config.userSchool || RAPOR_KOSONG;
        const alamat = db.config.schoolAddress || RAPOR_KOSONG;
        const waliNama = db.config.userName || RAPOR_KOSONG;
        const waliNip = db.config.userNIP || '';
        const kepsekNama = db.config.headmasterName || RAPOR_KOSONG;
        const kepsekNip = db.config.headmasterNIP || '';

        const badge = document.getElementById('rapor-preview-target-badge');
        const container = document.getElementById('rapor-preview-container');
        if (!container) return;

        const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || '-';
        const stds = db.students.filter(s => String(s.ClassID) === String(cid));

        if (!cid) {
            container.innerHTML = '<div class="text-center text-muted py-5"><i class="bi bi-diagram-3 fs-1 text-primary d-block mb-2"></i>Pilih Kelas / Rombel terlebih dahulu untuk menampilkan rapor.</div>';
            return;
        }

        if (stds.length === 0) {
            container.innerHTML = '<div class="text-center text-muted py-5"><i class="bi bi-person-x fs-1 text-muted d-block mb-2"></i>Belum ada data siswa di kelas ini. Silakan tambahkan siswa di menu Manajemen Kelas.</div>';
            return;
        }

        if (studentId === 'ALL') {
            if (badge) badge.innerText = `Cetak Massal: ${stds.length} Siswa`;
            let massalHtml = `
                <div class="alert alert-info d-flex align-items-center mb-4 no-print">
                    <i class="bi bi-people-fill fs-3 me-3 text-primary"></i>
                    <div>
                        <h6 class="fw-bold mb-1">Mode Cetak Massal Seluruh Siswa (${escapeHtml(className)})</h6>
                        <p class="mb-0 small text-muted">Sistem siap mencetak Laporan Hasil Belajar untuk <b>${stds.length} peserta didik</b> sekaligus dalam standar kertas A4 (2 halaman per siswa). Klik tombol <b>Cetak Dokumen</b> atau <b>Unduh PDF Rapor</b> di atas.</p>
                    </div>
                </div>
                <div class="text-muted small mb-3 no-print"><i class="bi bi-eye me-1"></i> Pratinjau Dokumen Seluruh Peserta Didik:</div>`;

            stds.forEach((std, idx) => {
                massalHtml += `<div class="mb-4">${generateRaporHtmlSheet(std, className, cid, semester, tapel, fase, kota, tglStr, sekolah, alamat, waliNama, waliNip, kepsekNama, kepsekNip)}</div>`;
            });
            container.innerHTML = massalHtml;
            return;
        }

        const student = stds.find(s => String(s.StudentID) === String(studentId));
        if (!student) {
            container.innerHTML = '<div class="text-center text-muted py-4">Pilih siswa untuk melihat pratinjau dokumen rapor.</div>';
            return;
        }

        if (badge) badge.innerText = student.StudentName;
        container.innerHTML = generateRaporHtmlSheet(student, className, cid, semester, tapel, fase, kota, tglStr, sekolah, alamat, waliNama, waliNip, kepsekNama, kepsekNip);
    }

    function generateRaporHtmlSheet(student, className, cid, semester, tapel, fase, kota, tglStr, sekolah, alamat, waliNama, waliNip, kepsekNama, kepsekNip) {
        const nis = student.NIS || '-';
        const nisn = student.NISN || '-';
        const nisText = `${nis} / ${nisn}`;
        const semText = semester === 'ganjil' ? '1 (Ganjil)' : '2 (Genap)';
        const tglFormatted = formatTanggalIndo(tglStr).replace(/^[A-Za-z]+,\s*/, '');

        const allGrades = getStudentAllGrades(cid, student.StudentID, semester, tapel);
        const attendance = getStudentAttendanceSummary(cid, student.StudentID, semester, tapel);
        const kokuSummary = getStudentKokuSummary(cid, student.StudentID, semester);
        const ekskulList = getStudentEkskulSummary(student.StudentID, semester, tapel);
        const notes = getStudentRaporNotes(student.StudentID, semester, tapel, className);

        // Grouping Mapel
        const catOrder = ['Mata Pelajaran Umum', 'Mata Pelajaran Pilihan', 'Muatan Lokal'];
        const catMap = {
            'Mata Pelajaran Umum': allGrades.filter(g => g.category === 'Umum'),
            'Mata Pelajaran Pilihan': allGrades.filter(g => g.category === 'Pilihan'),
            'Muatan Lokal': allGrades.filter(g => g.category === 'Muatan Lokal')
        };

        let gradesTableRows = '';
        let totalListed = 0;
        catOrder.forEach(cat => {
            const list = catMap[cat];
            if (list && list.length > 0) {
                gradesTableRows += `<tr style="background:#f8fafc;"><td colspan="4" style="border:1px solid #000; padding:5px 8px; font-weight:bold; font-size:11.5px;">${cat}</td></tr>`;
                list.forEach((g, idx) => {
                    totalListed++;
                    const scoreText = g.hasScore && g.finalScore !== null ? g.finalScore.toString() : '-';
                    const formattedCapaian = g.capaian ? escapeHtml(g.capaian).replace(/\n/g, '<br>') : '-';
                    gradesTableRows += `
                    <tr>
                        <td style="border:1px solid #000; padding:5px; text-align:center; vertical-align:top; font-size:11px;">${totalListed}</td>
                        <td style="border:1px solid #000; padding:5px 7px; vertical-align:top; font-size:11px;">${escapeHtml(g.subject)}</td>
                        <td style="border:1px solid #000; padding:5px; text-align:center; vertical-align:top; font-weight:bold; font-size:11.5px;">${scoreText}</td>
                        <td style="border:1px solid #000; padding:5px 7px; vertical-align:top; font-size:10.5px; line-height:1.35;">${formattedCapaian}</td>
                    </tr>`;
                });
            }
        });

        if (totalListed === 0) {
            gradesTableRows = `<tr><td colspan="4" style="border:1px solid #000; padding:15px; text-align:center; color:#64748b; font-style:italic;">Belum ada daftar mata pelajaran atau nilai tercatat.</td></tr>`;
        }

        // Ekstrakurikuler rows
        let ekskulRows = '';
        if (ekskulList.length > 0) {
            ekskulList.forEach((e, idx) => {
                ekskulRows += `
                <tr>
                    <td style="border:1px solid #000; padding:5px; text-align:center; font-size:11px;">${idx + 1}</td>
                    <td style="border:1px solid #000; padding:5px 8px; font-size:11px;">${escapeHtml(e.NamaEkskul)}</td>
                    <td style="border:1px solid #000; padding:5px; text-align:center; font-size:11px; font-weight:semibold;">${escapeHtml(e.Predikat || '-')}</td>
                    <td style="border:1px solid #000; padding:5px 8px; font-size:10.5px; line-height:1.3;">${escapeHtml(e.Keterangan || '-')}</td>
                </tr>`;
            });
        } else {
            ekskulRows = `
            <tr>
                <td colspan="4" style="border:1px solid #000; padding:6px 10px; text-align:center; font-size:11px; color:#64748b; font-style:italic;">
                    Tidak mengikuti kegiatan ekstrakurikuler / Belum ada data
                </td>
            </tr>`;
        }

        // Box Keputusan (HANYA DITAMPILKAN PADA SEMESTER GENAP)
        const isGenap = semester === 'genap';
        let keputusanHtml = '';
        if (isGenap) {
            keputusanHtml = `
            <table style="width:100%; border-collapse:collapse; border:1px solid #000; margin-bottom:12px;">
                <thead>
                    <tr style="background:#f2f2f2;">
                        <th style="border:1px solid #000; padding:5px 8px; text-align:center; font-size:11.5px;">Keputusan Kenaikan / Kelulusan</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="border:1px solid #000; padding:7px 10px; font-size:11px;">
                            Berdasarkan pencapaian seluruh tujuan pembelajaran, Peserta Didik ditetapkan: <b>${escapeHtml(notes.keputusan || 'NAIK KELAS')}</b>
                        </td>
                    </tr>
                </tbody>
            </table>`;
        }

        // Header Rapor
        const headerHtml = `
            <div style="display:flex; justify-content:space-between; font-size:11px; line-height:1.45;">
                <div style="width:52%;">
                    <div style="display:flex;"><span style="width:115px; font-weight:bold;">Nama Peserta Didik</span><span>: ${escapeHtml(student.StudentName)}</span></div>
                    <div style="display:flex;"><span style="width:115px;">NIS / NISN</span><span>: ${escapeHtml(nisText)}</span></div>
                    <div style="display:flex;"><span style="width:115px;">Nama Sekolah</span><span>: ${escapeHtml(sekolah)}</span></div>
                    <div style="display:flex;"><span style="width:115px;">Alamat Sekolah</span><span>: ${escapeHtml(alamat)}</span></div>
                </div>
                <div style="width:46%;">
                    <div style="display:flex;"><span style="width:110px;">Kelas / Rombel</span><span>: ${escapeHtml(className)}</span></div>
                    <div style="display:flex;"><span style="width:110px;">Fase</span><span>: ${escapeHtml(fase)}</span></div>
                    <div style="display:flex;"><span style="width:110px;">Semester</span><span>: ${escapeHtml(semText)}</span></div>
                    <div style="display:flex;"><span style="width:110px;">Tahun Pelajaran</span><span>: ${escapeHtml(tapel)}</span></div>
                </div>
            </div>
            <div style="border-bottom:1.5px solid #000; margin:8px 0 14px 0;"></div>`;

        return `
        <!-- Lembar 1: Laporan Hasil Belajar -->
        <div class="rapor-paper-sheet">
            <div class="text-end text-muted small mb-2 no-print" style="font-family:sans-serif; font-size:9px;">[ Lembar 1: Capaian Hasil Belajar - Standar Kertas A4 ]</div>
            ${headerHtml}
            <div style="text-align:center; font-weight:bold; font-size:13px; margin-bottom:12px; letter-spacing:0.5px;">
                LAPORAN HASIL BELAJAR
            </div>
            <table style="width:100%; border-collapse:collapse; border:1px solid #000;">
                <thead>
                    <tr style="background:#f2f2f2;">
                        <th style="border:1px solid #000; padding:6px 4px; text-align:center; width:6%; font-size:11px;">No</th>
                        <th style="border:1px solid #000; padding:6px; width:28%; font-size:11px;">Mata Pelajaran</th>
                        <th style="border:1px solid #000; padding:6px; text-align:center; width:12%; font-size:11px;">Nilai Akhir</th>
                        <th style="border:1px solid #000; padding:6px; text-align:center; width:54%; font-size:11px;">Capaian Kompetensi</th>
                    </tr>
                </thead>
                <tbody>
                    ${gradesTableRows}
                </tbody>
            </table>
        </div>

        <!-- Lembar 2: Kokurikuler, Ekstrakurikuler, Presensi & Pengesahan -->
        <div class="rapor-paper-sheet">
            <div class="text-end text-muted small mb-2 no-print" style="font-family:sans-serif; font-size:9px;">[ Lembar 2: Kokurikuler, Ekskul & Pengesahan - Standar Kertas A4 ]</div>
            ${headerHtml}
            
            <!-- Box Kokurikuler -->
            <table style="width:100%; border-collapse:collapse; border:1px solid #000; margin-bottom:12px;">
                <thead>
                    <tr style="background:#f2f2f2;">
                        <th style="border:1px solid #000; padding:5px 8px; text-align:center; font-size:11.5px;">Kokurikuler</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="border:1px solid #000; padding:7px 10px; font-size:11px; line-height:1.4; color: ${kokuSummary.hasData ? '#000' : '#64748b'}; font-style: ${kokuSummary.hasData ? 'normal' : 'italic'};">
                            ${escapeHtml(kokuSummary.text).replace(/\n\n/g, '<br><br>')}
                        </td>
                    </tr>
                </tbody>
            </table>

            <!-- Table Ekstrakurikuler -->
            <table style="width:100%; border-collapse:collapse; border:1px solid #000; margin-bottom:12px;">
                <thead>
                    <tr style="background:#f2f2f2;">
                        <th style="border:1px solid #000; padding:5px 4px; text-align:center; width:6%; font-size:11px;">No</th>
                        <th style="border:1px solid #000; padding:5px 8px; width:28%; font-size:11px;">Kegiatan Ekstrakurikuler</th>
                        <th style="border:1px solid #000; padding:5px; text-align:center; width:16%; font-size:11px;">Predikat</th>
                        <th style="border:1px solid #000; padding:5px 8px; width:50%; font-size:11px;">Keterangan</th>
                    </tr>
                </thead>
                <tbody>
                    ${ekskulRows}
                </tbody>
            </table>

            <!-- Side-by-side Ketidakhadiran & Catatan Wali Kelas -->
            <div style="display:flex; justify-content:space-between; gap:12px; margin-bottom:12px;">
                <div style="width:42%;">
                    <table style="width:100%; border-collapse:collapse; border:1px solid #000; height:100%;">
                        <thead>
                            <tr style="background:#f2f2f2;">
                                <th colspan="2" style="border:1px solid #000; padding:5px 8px; text-align:center; font-size:11.5px;">Ketidakhadiran</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td style="border:1px solid #000; padding:5px 8px; font-size:11px; width:55%;">Sakit</td>
                                <td style="border:1px solid #000; padding:5px 8px; font-size:11px;">: ${attendance.hasRecords ? attendance.S + ' hari' : '-'}</td>
                            </tr>
                            <tr>
                                <td style="border:1px solid #000; padding:5px 8px; font-size:11px;">Izin</td>
                                <td style="border:1px solid #000; padding:5px 8px; font-size:11px;">: ${attendance.hasRecords ? attendance.I + ' hari' : '-'}</td>
                            </tr>
                            <tr>
                                <td style="border:1px solid #000; padding:5px 8px; font-size:11px;">Tanpa Keterangan</td>
                                <td style="border:1px solid #000; padding:5px 8px; font-size:11px;">: ${attendance.hasRecords ? attendance.A + ' hari' : '-'}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <div style="width:58%;">
                    <table style="width:100%; border-collapse:collapse; border:1px solid #000; height:100%;">
                        <thead>
                            <tr style="background:#f2f2f2;">
                                <th style="border:1px solid #000; padding:5px 8px; text-align:center; font-size:11.5px;">Catatan Wali Kelas</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td style="border:1px solid #000; padding:7px 10px; font-size:11px; line-height:1.4; vertical-align:top; min-height:55px;">
                                    ${escapeHtml(notes.catatanWali)}
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Box Keputusan Kenaikan (HANYA PADA SEMESTER GENAP) -->
            ${keputusanHtml}

            <!-- Box Tanggapan Orang Tua/Wali Murid -->
            <table style="width:100%; border-collapse:collapse; border:1px solid #000; margin-bottom:20px;">
                <thead>
                    <tr style="background:#f2f2f2;">
                        <th style="border:1px solid #000; padding:5px 8px; text-align:center; font-size:11.5px;">Tanggapan Orang Tua / Wali Murid</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="border:1px solid #000; padding:18px 10px; font-size:11px;">&nbsp;</td>
                    </tr>
                </tbody>
            </table>

            <!-- Tanda Tangan 3 Pihak -->
            <div style="font-size:11px; line-height:1.35; margin-top:10px;">
                <!-- Baris 1: Orang Tua / Wali Murid sejajar dengan Guru / Wali Kelas -->
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:18px;">
                    <div style="width:45%; text-align:center;">
                        <div style="visibility:hidden; height:16px;">${escapeHtml(kota)}, ${escapeHtml(tglFormatted)}</div>
                        <div>Orang Tua / Wali Murid,</div>
                        <div style="height:50px;"></div>
                        <div>..................................................</div>
                    </div>
                    <div style="width:45%; text-align:center;">
                        <div style="height:16px;">${escapeHtml(kota)}, ${escapeHtml(tglFormatted)}</div>
                        <div>Wali Kelas,</div>
                        <div style="height:50px;"></div>
                        <div><b>${escapeHtml(waliNama)}</b></div>
                        <div style="font-size:10px; color:#222;">${waliNip && waliNip !== '-' ? `NIP. ${escapeHtml(waliNip)}` : '&nbsp;'}</div>
                    </div>
                </div>

                <!-- Baris 2: Kepala Sekolah di Bawah Rata Tengah -->
                <div style="text-align:center; margin:0 auto; width:55%;">
                    <div>Mengetahui,</div>
                    <div>Kepala ${escapeHtml(sekolah)}</div>
                    <div style="height:50px;"></div>
                    <div><b>${escapeHtml(kepsekNama)}</b></div>
                    <div style="font-size:10px; color:#222;">${kepsekNip && kepsekNip !== '-' ? `NIP. ${escapeHtml(kepsekNip)}` : '&nbsp;'}</div>
                </div>
            </div>
        </div>`;
    }

    function cetakRaporBrowser(lewatiCekPengaturan) {
        const cid = document.getElementById('rapor-kelas')?.value;
        if (!cid) return Swal.fire('Peringatan', 'Silakan pilih kelas terlebih dahulu.', 'warning');
        if (lewatiCekPengaturan !== true) return cekPengaturanRapor(() => cetakRaporBrowser(true));

        // Selalu buat ulang pratinjau agar yang dicetak adalah data terbaru
        // (termasuk nilai dari Pengolah Nilai).
        previewRaporSiswa();
        const container = document.getElementById('rapor-preview-container');
        const sheets = container ? [...container.querySelectorAll('.rapor-paper-sheet')] : [];
        if (sheets.length === 0) {
            return Swal.fire('Belum Ada Rapor', 'Rapor belum bisa ditampilkan. Pastikan kelas memiliki siswa, lalu klik "Tampilkan Pratinjau Rapor".', 'info');
        }

        // Cetak lewat halaman khusus (iframe tersembunyi) yang hanya berisi lembar rapor,
        // sehingga banner, menu, dan tombol aplikasi tidak ikut tercetak.
        // Margin kertas dibuat lewat thead/tfoot yang diulang di setiap halaman, dan
        // @page margin 0 agar tanggal/URL bawaan browser tidak ikut tercetak.
        const pages = sheets.map(sh => {
            const clone = sh.cloneNode(true);
            clone.querySelectorAll('.no-print').forEach(n => n.remove());
            return `<table class="pp"><thead><tr><td class="pp-top"></td></tr></thead>` +
                `<tbody><tr><td class="pp-body">${clone.innerHTML}</td></tr></tbody>` +
                `<tfoot><tr><td class="pp-bottom"></td></tr></tfoot></table>`;
        }).join('');

        const studentId = document.getElementById('rapor-siswa')?.value;
        const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
        const studentName = studentId && studentId !== 'ALL' ? (db.students.find(s => String(s.StudentID) === String(studentId))?.StudentName || '') : 'Semua Siswa';
        const docTitle = `Rapor ${studentName} - ${className}`.replace(/[\\/:*?"<>|]+/g, '-');

        const html = `<!doctype html><html lang="id"><head><meta charset="utf-8"><title>${escapeHtml(docTitle)}</title><style>
            @page { size: A4 portrait; margin: 0; }
            html, body { margin: 0; padding: 0; background: #fff; color: #000; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            body { font-family: 'Times New Roman', Times, serif; line-height: 1.35; }
            .pp { width: 100%; border-collapse: collapse; break-after: page; page-break-after: always; }
            .pp:last-child { break-after: auto; page-break-after: auto; }
            .pp > thead > tr > td, .pp > tbody > tr > td, .pp > tfoot > tr > td { border: none; padding: 0; }
            .pp-top { height: 12mm; }
            .pp-bottom { height: 12mm; }
            .pp > tbody > tr > td.pp-body { padding: 0 15mm; vertical-align: top; }
            table { page-break-inside: auto; }
            tr { page-break-inside: avoid; break-inside: avoid; }
            img { max-width: 100%; }
        </style></head><body>${pages}</body></html>`;

        const old = document.getElementById('rapor-print-frame');
        if (old) old.remove();
        const frame = document.createElement('iframe');
        frame.id = 'rapor-print-frame';
        frame.setAttribute('aria-hidden', 'true');
        frame.style.cssText = 'position:fixed; right:0; bottom:0; width:0; height:0; border:0; visibility:hidden;';
        document.body.appendChild(frame);
        const fdoc = frame.contentWindow.document;
        fdoc.open(); fdoc.write(html); fdoc.close();

        // Judul dokumen menjadi nama file default jika guru memilih "Simpan sebagai PDF"
        const prevTitle = document.title;
        document.title = docTitle;
        setTimeout(() => {
            try {
                frame.contentWindow.focus();
                frame.contentWindow.print();
            } catch (e) {
                // Cadangan: cetak halaman utama memakai aturan @media print
                window.print();
            } finally {
                document.title = prevTitle;
            }
        }, 300);
    }

    async function prosesCetakRaporSiswaPDF(lewatiCekPengaturan) {
        if (lewatiCekPengaturan !== true) return cekPengaturanRapor(() => prosesCetakRaporSiswaPDF(true));
        ensureTPBank();
        const cid = document.getElementById('rapor-kelas')?.value;
        const studentId = document.getElementById('rapor-siswa')?.value;
        const semester = document.getElementById('rapor-semester')?.value || 'genap';
        const tapel = document.getElementById('rapor-tapel')?.value.trim() || getActiveTapel();

        const fase = db.config.raporFase || 'D';
        const kota = db.config.schoolCity || RAPOR_KOSONG;
        const tglStr = db.config.raporTanggal || tanggalHariIni();
        const sekolah = db.config.userSchool || RAPOR_KOSONG;
        const alamat = db.config.schoolAddress || RAPOR_KOSONG;
        const waliNama = db.config.userName || RAPOR_KOSONG;
        const waliNip = db.config.userNIP || '';
        const kepsekNama = db.config.headmasterName || RAPOR_KOSONG;
        const kepsekNip = db.config.headmasterNIP || '';

        const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
        const allStds = db.students.filter(s => String(s.ClassID) === String(cid));

        let targetStudents = [];
        if (studentId === 'ALL') {
            targetStudents = allStds;
        } else {
            targetStudents = allStds.filter(s => String(s.StudentID) === String(studentId));
        }

        if (targetStudents.length === 0) {
            return Swal.fire('Error', 'Tidak ada siswa yang dipilih untuk dicetak rapornya.', 'error');
        }

        Swal.fire({
            title: 'Menyusun Dokumen Rapor PDF...',
            text: `Menyiapkan ${targetStudents.length} peserta didik (Standar Kertas A4 Resmi)...`,
            allowOutsideClick: false,
            didOpen: () => Swal.showLoading()
        });

        try {
            await ensurePDF();
            const { jsPDF } = window.jspdf;
            // Standar kertas A4 Resmi: 210 x 297 mm
            const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
            const tglFormatted = formatTanggalIndo(tglStr).replace(/^[A-Za-z]+,\s*/, '');
            const isGenap = semester === 'genap';

            targetStudents.forEach((student, sIdx) => {
                if (sIdx > 0) doc.addPage();

                const nis = student.NIS || '-';
                const nisn = student.NISN || '-';
                const nisText = `${nis} / ${nisn}`;
                const semText = semester === 'ganjil' ? '1 (Ganjil)' : '2 (Genap)';

                // Helper Header Block Standar A4 (Margin: X=15, Width=180)
                function drawRaporHeaderBlock() {
                    doc.setFont('times', 'normal');
                    doc.setFontSize(8.5);
                    doc.setTextColor(0, 0, 0);

                    // Kolom Kiri
                    doc.text('Nama Peserta Didik', 15, 16);
                    doc.text(':', 48, 16);
                    doc.setFont('times', 'bold');
                    doc.text(student.StudentName, 50, 16);
                    doc.setFont('times', 'normal');

                    doc.text('NIS / NISN', 15, 21);
                    doc.text(':', 48, 21);
                    doc.text(nisText, 50, 21);

                    doc.text('Nama Sekolah', 15, 26);
                    doc.text(':', 48, 26);
                    doc.text(sekolah, 50, 26);

                    doc.text('Alamat Sekolah', 15, 31);
                    doc.text(':', 48, 31);
                    doc.text(alamat, 50, 31);

                    // Kolom Kanan
                    doc.text('Kelas / Rombel', 120, 16);
                    doc.text(':', 148, 16);
                    doc.text(className, 150, 16);

                    doc.text('Fase', 120, 21);
                    doc.text(':', 148, 21);
                    doc.text(fase, 150, 21);

                    doc.text('Semester', 120, 26);
                    doc.text(':', 148, 26);
                    doc.text(semText, 150, 26);

                    doc.text('Tahun Pelajaran', 120, 31);
                    doc.text(':', 148, 31);
                    doc.text(tapel, 150, 31);

                    // Garis Pemisah Tipis
                    doc.setDrawColor(0, 0, 0);
                    doc.setLineWidth(0.3);
                    doc.line(15, 34, 195, 34);
                }

                // ================= LEMBAR 1: CAPAIAN HASIL BELAJAR =================
                drawRaporHeaderBlock();

                // Judul
                doc.setFont('times', 'bold');
                doc.setFontSize(10.5);
                doc.text('LAPORAN HASIL BELAJAR', 105, 40, { align: 'center' });

                // Table Nilai
                const allGrades = getStudentAllGrades(cid, student.StudentID, semester, tapel);
                let tableBody = [];
                const catOrder = ['Mata Pelajaran Umum', 'Mata Pelajaran Pilihan', 'Muatan Lokal'];
                const catMap = {
                    'Mata Pelajaran Umum': allGrades.filter(g => g.category === 'Umum'),
                    'Mata Pelajaran Pilihan': allGrades.filter(g => g.category === 'Pilihan'),
                    'Muatan Lokal': allGrades.filter(g => g.category === 'Muatan Lokal')
                };

                let itemNo = 0;
                catOrder.forEach(cat => {
                    const list = catMap[cat];
                    if (list && list.length > 0) {
                        tableBody.push([
                            { content: cat, colSpan: 4, styles: { fontStyle: 'bold', fillColor: [248, 250, 252], textColor: [0, 0, 0], halign: 'left', font: 'times', cellPadding: 1.8 } }
                        ]);
                        list.forEach((g) => {
                            itemNo++;
                            tableBody.push([
                                itemNo.toString(),
                                g.subject,
                                g.hasScore && g.finalScore !== null ? g.finalScore.toString() : '-',
                                g.capaian || '-'
                            ]);
                        });
                    }
                });

                if (tableBody.length === 0) {
                    tableBody.push([{ content: 'Belum ada data mata pelajaran atau nilai.', colSpan: 4, styles: { halign: 'center', fontStyle: 'italic' } }]);
                }

                doc.autoTable({
                    head: [['No', 'Mata Pelajaran', 'Nilai Akhir', 'Capaian Kompetensi']],
                    body: tableBody,
                    startY: 44,
                    theme: 'grid',
                    headStyles: {
                        fillColor: [240, 240, 240],
                        textColor: [0, 0, 0],
                        fontSize: 8,
                        fontStyle: 'bold',
                        font: 'times',
                        halign: 'center',
                        valign: 'middle',
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    styles: {
                        font: 'times',
                        fontSize: 7.5,
                        textColor: [0, 0, 0],
                        cellPadding: 1.8,
                        overflow: 'linebreak',
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    columnStyles: {
                        0: { cellWidth: 9, halign: 'center', valign: 'top' },
                        1: { cellWidth: 46, valign: 'top' },
                        2: { cellWidth: 17, halign: 'center', valign: 'top' },
                        3: { cellWidth: 108, valign: 'top' }
                    },
                    margin: { left: 15, right: 15 }
                });

                // ================= LEMBAR 2: KOKURIKULER, EKSKUL, PRESENSI, PENGESAHAN =================
                doc.addPage();
                drawRaporHeaderBlock();

                // 1. Box Kokurikuler
                const kokuSummary = getStudentKokuSummary(cid, student.StudentID, semester);
                doc.autoTable({
                    head: [['Kokurikuler']],
                    body: [[kokuSummary.text]],
                    startY: 38,
                    theme: 'grid',
                    headStyles: {
                        fillColor: [240, 240, 240],
                        textColor: [0, 0, 0],
                        fontSize: 8,
                        fontStyle: 'bold',
                        font: 'times',
                        halign: 'center',
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    styles: {
                        font: 'times',
                        fontSize: 7.5,
                        textColor: [0, 0, 0],
                        cellPadding: 2.2,
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    margin: { left: 15, right: 15 }
                });

                // 2. Table Ekstrakurikuler
                let nextY = doc.lastAutoTable.finalY + 4;
                const ekskulList = getStudentEkskulSummary(student.StudentID, semester, tapel);
                let ekskulBody = [];
                if (ekskulList.length > 0) {
                    ekskulList.forEach((e, idx) => {
                        ekskulBody.push([
                            (idx + 1).toString(),
                            e.NamaEkskul,
                            e.Predikat || '-',
                            e.Keterangan || '-'
                        ]);
                    });
                } else {
                    ekskulBody.push([{ content: 'Tidak mengikuti kegiatan ekstrakurikuler / Belum ada data', colSpan: 4, styles: { halign: 'center', fontStyle: 'italic' } }]);
                }

                doc.autoTable({
                    head: [['No', 'Kegiatan Ekstrakurikuler', 'Predikat', 'Keterangan']],
                    body: ekskulBody,
                    startY: nextY,
                    theme: 'grid',
                    headStyles: {
                        fillColor: [240, 240, 240],
                        textColor: [0, 0, 0],
                        fontSize: 8,
                        fontStyle: 'bold',
                        font: 'times',
                        halign: 'center',
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    styles: {
                        font: 'times',
                        fontSize: 7.5,
                        textColor: [0, 0, 0],
                        cellPadding: 1.8,
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    columnStyles: {
                        0: { cellWidth: 9, halign: 'center' },
                        1: { cellWidth: 46, halign: 'left' },
                        2: { cellWidth: 22, halign: 'center' },
                        3: { cellWidth: 103, halign: 'left' }
                    },
                    margin: { left: 15, right: 15 }
                });

                // 3. Side-by-side Ketidakhadiran & Catatan Wali Kelas
                nextY = doc.lastAutoTable.finalY + 4;
                const att = getStudentAttendanceSummary(cid, student.StudentID, semester, tapel);
                const notes = getStudentRaporNotes(student.StudentID, semester, tapel, className);

                // Tabel Ketidakhadiran (Kiri: X=15, width=65)
                doc.autoTable({
                    head: [[{ content: 'Ketidakhadiran', colSpan: 2, styles: { halign: 'center', fontStyle: 'bold' } }]],
                    body: [
                        ['Sakit', `: ${att.hasRecords ? att.S + ' hari' : '-'}`],
                        ['Izin', `: ${att.hasRecords ? att.I + ' hari' : '-'}`],
                        ['Tanpa Keterangan', `: ${att.hasRecords ? att.A + ' hari' : '-'}`]
                    ],
                    startY: nextY,
                    theme: 'grid',
                    headStyles: {
                        fillColor: [240, 240, 240],
                        textColor: [0, 0, 0],
                        fontSize: 8,
                        font: 'times',
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    styles: {
                        font: 'times',
                        fontSize: 7.5,
                        textColor: [0, 0, 0],
                        cellPadding: 1.6,
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    columnStyles: {
                        0: { cellWidth: 35 },
                        1: { cellWidth: 30 }
                    },
                    margin: { left: 15, right: 130 }
                });

                const attFinalY = doc.lastAutoTable.finalY;

                // Box Catatan Wali Kelas (Kanan: X=84, width=111)
                doc.autoTable({
                    head: [['Catatan Wali Kelas']],
                    body: [[notes.catatanWali]],
                    startY: nextY,
                    theme: 'grid',
                    headStyles: {
                        fillColor: [240, 240, 240],
                        textColor: [0, 0, 0],
                        fontSize: 8,
                        fontStyle: 'bold',
                        font: 'times',
                        halign: 'center',
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    styles: {
                        font: 'times',
                        fontSize: 7.5,
                        textColor: [0, 0, 0],
                        cellPadding: 2,
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25,
                        minCellHeight: 18
                    },
                    margin: { left: 84, right: 15 }
                });

                const catatanFinalY = doc.lastAutoTable.finalY;
                nextY = Math.max(attFinalY, catatanFinalY) + 4;

                // 4. Box Keputusan Kenaikan (HANYA DITAMPILKAN PADA SEMESTER GENAP!)
                if (isGenap) {
                    const keputusanContent = `Berdasarkan pencapaian seluruh tujuan pembelajaran, Peserta Didik ditetapkan: ${notes.keputusan || 'NAIK KELAS'}`;
                    doc.autoTable({
                        head: [['Keputusan Kenaikan / Kelulusan']],
                        body: [[keputusanContent]],
                        startY: nextY,
                        theme: 'grid',
                        headStyles: {
                            fillColor: [240, 240, 240],
                            textColor: [0, 0, 0],
                            fontSize: 8,
                            fontStyle: 'bold',
                            font: 'times',
                            halign: 'center',
                            lineColor: [0, 0, 0],
                            lineWidth: 0.25
                        },
                        styles: {
                            font: 'times',
                            fontSize: 7.5,
                            textColor: [0, 0, 0],
                            cellPadding: 2,
                            lineColor: [0, 0, 0],
                            lineWidth: 0.25
                        },
                        margin: { left: 15, right: 15 }
                    });
                    nextY = doc.lastAutoTable.finalY + 4;
                }

                // 5. Box Tanggapan Orang Tua/Wali Murid
                doc.autoTable({
                    head: [['Tanggapan Orang Tua / Wali Murid']],
                    body: [['\n\n']],
                    startY: nextY,
                    theme: 'grid',
                    headStyles: {
                        fillColor: [240, 240, 240],
                        textColor: [0, 0, 0],
                        fontSize: 8,
                        fontStyle: 'bold',
                        font: 'times',
                        halign: 'center',
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25
                    },
                    styles: {
                        font: 'times',
                        fontSize: 7.5,
                        lineColor: [0, 0, 0],
                        lineWidth: 0.25,
                        minCellHeight: 12
                    },
                    margin: { left: 15, right: 15 }
                });

                // 6. Signatures (3 Pihak)
                const sigY = doc.lastAutoTable.finalY + 6;
                doc.setFont('times', 'normal');
                doc.setFontSize(8);

                // Baris 1 Kanan: Titimangsa & Wali Kelas (Kolom Kanan x=155)
                doc.text(`${kota}, ${tglFormatted}`, 155, sigY, { align: 'center' });
                doc.text('Wali Kelas,', 155, sigY + 4, { align: 'center' });
                doc.setFont('times', 'bold');
                doc.text(waliNama, 155, sigY + 18, { align: 'center' });
                doc.setFont('times', 'normal');
                doc.setFontSize(7.5);
                if (waliNip && waliNip !== '-') doc.text(`NIP. ${waliNip}`, 155, sigY + 22, { align: 'center' });

                // Baris 1 Kiri: Orang Tua / Wali (Kolom Kiri x=55, Sejajar dengan Wali Kelas)
                doc.setFontSize(8);
                doc.text('Orang Tua / Wali Murid,', 55, sigY + 4, { align: 'center' });
                doc.text('..................................................', 55, sigY + 18, { align: 'center' });

                // Baris 2 Tengah: Kepala Sekolah (Tengah Bawah x=105, Rata Tengah)
                doc.setFontSize(8);
                doc.text('Mengetahui,', 105, sigY + 28, { align: 'center' });
                doc.text(`Kepala ${sekolah}`, 105, sigY + 32, { align: 'center' });
                doc.setFont('times', 'bold');
                doc.text(kepsekNama, 105, sigY + 46, { align: 'center' });
                doc.setFont('times', 'normal');
                doc.setFontSize(7.5);
                if (kepsekNip && kepsekNip !== '-') doc.text(`NIP. ${kepsekNip}`, 105, sigY + 50, { align: 'center' });
            });

            const fileName = targetStudents.length === 1 
                ? `Rapor_A4_${targetStudents[0].StudentName.replace(/[^a-zA-Z0-9_-]/g, '_')}_Semester_${semester}.pdf`
                : `Rapor_A4_Kelas_${className.replace(/[^a-zA-Z0-9_-]/g, '_')}_Semester_${semester}.pdf`;

            doc.save(fileName);
            Swal.fire({
                icon: 'success',
                title: 'Rapor Berhasil Diunduh!',
                text: `Dokumen resmi Laporan Hasil Belajar (${targetStudents.length} siswa) berhasil disimpan dalam ukuran A4.`,
                timer: 2000,
                showConfirmButton: false
            });
        } catch (err) {
            console.error("Gagal membuat PDF Rapor:", err);
            Swal.fire('Error', 'Terjadi kesalahan saat menyusun dokumen PDF: ' + err.message, 'error');
        }
    }

    // --- Nilai Ekstrakurikuler Siswa (Terhubung ke Rapor) ---
    function switchEkskulTab(tab) {
        const subInput = document.getElementById('ekskul-subview-input');
        const subRiwayat = document.getElementById('ekskul-subview-riwayat');
        const btnInput = document.getElementById('btn-tab-ekskul-input');
        const btnRiwayat = document.getElementById('btn-tab-ekskul-riwayat');
        if (!subInput || !subRiwayat) return;

        if (tab === 'riwayat') {
            subInput.style.display = 'none';
            subRiwayat.style.display = 'block';
            if (btnInput) { btnInput.classList.remove('btn-primary'); btnInput.classList.add('btn-outline-primary'); }
            if (btnRiwayat) { btnRiwayat.classList.remove('btn-outline-primary'); btnRiwayat.classList.add('btn-primary'); }
            renderEkskulHistory();
        } else {
            subInput.style.display = 'block';
            subRiwayat.style.display = 'none';
            if (btnInput) { btnInput.classList.remove('btn-outline-primary'); btnInput.classList.add('btn-primary'); }
            if (btnRiwayat) { btnRiwayat.classList.remove('btn-primary'); btnRiwayat.classList.add('btn-outline-primary'); }
        }
    }

    function loadEkskulData() {
        populateClassSelects();
        onEkskulClassChange();
        const curDate = new Date();
        const curYear = curDate.getFullYear();
        const curMonth = curDate.getMonth() + 1;
        const academicYear = curMonth >= 7 ? `${curYear}/${curYear + 1}` : `${curYear - 1}/${curYear}`;
        const defaultSem = curMonth >= 7 ? 'ganjil' : 'genap';
        if (document.getElementById('ekskul-tapel') && !document.getElementById('ekskul-tapel').value) {
            document.getElementById('ekskul-tapel').value = getActiveTapel(); // ikut TP yang dipilih di atas
        }
        const elEkskulSem = document.getElementById('ekskul-semester');
        if (elEkskulSem && !elEkskulSem.dataset.userSet) {
            elEkskulSem.value = defaultSem;
        }
    }

    // Semester yang dipilih sendiri oleh guru tidak ditimpa lagi oleh nilai bawaan
    document.addEventListener('change', e => {
        if (e.target && (e.target.id === 'rapor-semester' || e.target.id === 'ekskul-semester')) e.target.dataset.userSet = '1';
    });

    function onEkskulClassChange() {
        const cid = document.getElementById('ekskul-kelas')?.value;
        const selectSiswa = document.getElementById('ekskul-siswa');
        if (!selectSiswa) return;
        const stds = db.students.filter(s => String(s.ClassID) === String(cid));
        let html = '<option value="">-- Pilih Siswa --</option>';
        stds.forEach(s => {
            html += `<option value="${s.StudentID}">${escapeHtml(s.StudentName)}</option>`;
        });
        selectSiswa.innerHTML = html;
    }

    function setEkskulTemplate(text) {
        const el = document.getElementById('ekskul-keterangan');
        if (el) el.value = text;
    }

    function saveEkskul() {
        const editId = document.getElementById('ekskul-edit-id')?.value;
        const cid = document.getElementById('ekskul-kelas')?.value;
        const sid = document.getElementById('ekskul-siswa')?.value;
        const semester = document.getElementById('ekskul-semester')?.value || 'genap';
        const tapel = document.getElementById('ekskul-tapel')?.value.trim() || getActiveTapel();
        const nama = document.getElementById('ekskul-nama')?.value.trim();
        const predikat = document.getElementById('ekskul-predikat')?.value || 'Baik';
        const keterangan = document.getElementById('ekskul-keterangan')?.value.trim();

        if (!cid) return Swal.fire('Peringatan', 'Silakan pilih kelas terlebih dahulu.', 'warning');
        if (!sid) return Swal.fire('Peringatan', 'Silakan pilih siswa yang dinilai.', 'warning');
        if (!nama) return Swal.fire('Peringatan', 'Silakan isi nama kegiatan ekstrakurikuler (cth: Pramuka).', 'warning');

        const student = db.students.find(s => String(s.StudentID) === String(sid));
        const studentName = student ? student.StudentName : '-';

        if (!db.ekskul) db.ekskul = [];

        if (editId) {
            const idx = db.ekskul.findIndex(e => String(e.EkskulID) === String(editId));
            if (idx > -1) {
                db.ekskul[idx].ClassID = cid;
                db.ekskul[idx].StudentID = sid;
                db.ekskul[idx].StudentName = studentName;
                db.ekskul[idx].Semester = semester;
                db.ekskul[idx].TahunPelajaran = tapel;
                db.ekskul[idx].NamaEkskul = nama;
                db.ekskul[idx].Predikat = predikat;
                db.ekskul[idx].Keterangan = keterangan;
            }
            cancelEditEkskul();
            Swal.fire('Berhasil!', 'Data nilai ekstrakurikuler berhasil diperbarui.', 'success');
        } else {
            const newRecord = {
                EkskulID: 'eks_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
                ClassID: cid,
                StudentID: sid,
                StudentName: studentName,
                Semester: semester,
                TahunPelajaran: tapel,
                NamaEkskul: nama,
                Predikat: predikat,
                Keterangan: keterangan
            };
            db.ekskul.push(newRecord);
            document.getElementById('ekskul-nama').value = '';
            document.getElementById('ekskul-keterangan').value = '';
            Swal.fire('Berhasil!', 'Nilai ekstrakurikuler berhasil disimpan & terhubung ke Rapor.', 'success');
        }

        syncDB();
        renderEkskulHistory();
        if (document.getElementById('view-rapor')?.classList.contains('active')) {
            previewRaporSiswa();
        }
    }

    function renderEkskulHistory() {
        const tbody = document.getElementById('ekskul-history-tbody');
        if (!tbody) return;
        const filterClass = document.getElementById('ekskul-riwayat-kelas')?.value;
        const filterSem = document.getElementById('ekskul-riwayat-semester')?.value;

        if (!db.ekskul) db.ekskul = [];

        let filtered = db.ekskul;
        if (filterClass && filterClass !== 'All') {
            filtered = filtered.filter(e => String(e.ClassID) === String(filterClass));
        }
        if (filterSem) {
            filtered = filtered.filter(e => String(e.Semester) === String(filterSem));
        }

        if (filtered.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" class="text-center py-4 text-muted"><i class="bi bi-info-circle me-1"></i>Belum ada data nilai ekstrakurikuler tercatat.</td></tr>';
            return;
        }

        let html = '';
        filtered.forEach((e, idx) => {
            const className = db.classes.find(c => String(c.ClassID) === String(e.ClassID))?.ClassName || '-';
            let badgeClass = 'bg-primary';
            if (e.Predikat === 'Sangat Baik') badgeClass = 'bg-success';
            else if (e.Predikat === 'Cukup') badgeClass = 'bg-warning text-dark';
            else if (e.Predikat === 'Kurang') badgeClass = 'bg-danger';

            html += `<tr>
                <td class="text-center fw-bold">${idx + 1}</td>
                <td class="fw-bold" style="color: var(--navy-deep);">${escapeHtml(e.StudentName)}</td>
                <td>${escapeHtml(className)}</td>
                <td class="fw-semibold text-primary"><i class="bi bi-trophy me-1 text-warning"></i>${escapeHtml(e.NamaEkskul)}</td>
                <td class="text-center"><span class="badge ${badgeClass} px-2 py-1">${escapeHtml(e.Predikat)}</span></td>
                <td class="small text-muted">${escapeHtml(e.Keterangan || '-')}</td>
                <td class="text-center text-nowrap">
                    <button class="btn btn-sm btn-outline-primary me-1" onclick="editEkskul('${e.EkskulID}')" title="Edit"><i class="bi bi-pencil-square"></i></button>
                    <button class="btn btn-sm btn-outline-danger" onclick="deleteEkskul('${e.EkskulID}')" title="Hapus"><i class="bi bi-trash"></i></button>
                </td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    function editEkskul(id) {
        if (!db.ekskul) return;
        const item = db.ekskul.find(e => String(e.EkskulID) === String(id));
        if (!item) return;

        switchEkskulTab('input');
        document.getElementById('ekskul-edit-id').value = item.EkskulID;
        document.getElementById('ekskul-kelas').value = item.ClassID;
        onEkskulClassChange();
        setTimeout(() => {
            if (document.getElementById('ekskul-siswa')) document.getElementById('ekskul-siswa').value = item.StudentID;
        }, 50);
        document.getElementById('ekskul-semester').value = item.Semester || 'genap';
        document.getElementById('ekskul-tapel').value = item.TahunPelajaran || getActiveTapel();
        document.getElementById('ekskul-nama').value = item.NamaEkskul || '';
        document.getElementById('ekskul-predikat').value = item.Predikat || 'Baik';
        document.getElementById('ekskul-keterangan').value = item.Keterangan || '';

        document.getElementById('ekskul-form-title').innerHTML = '<i class="bi bi-pencil-square text-warning"></i> Edit Nilai Ekstrakurikuler Siswa';
        document.getElementById('btn-cancel-edit-ekskul').style.display = 'inline-block';
    }

    function cancelEditEkskul() {
        document.getElementById('ekskul-edit-id').value = '';
        document.getElementById('ekskul-nama').value = '';
        document.getElementById('ekskul-keterangan').value = '';
        document.getElementById('ekskul-form-title').innerHTML = '<i class="bi bi-pencil-square" style="color: var(--teal-accent);"></i> Form Input Nilai Ekstrakurikuler Siswa';
        document.getElementById('btn-cancel-edit-ekskul').style.display = 'none';
    }

    function deleteEkskul(id) {
        Swal.fire({
            title: 'Hapus Nilai Ekskul?',
            text: 'Data nilai ekstrakurikuler ini akan dihapus permanen.',
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#d33',
            confirmButtonText: 'Ya, Hapus'
        }).then(res => {
            if (res.isConfirmed) {
                db.ekskul = (db.ekskul || []).filter(e => String(e.EkskulID) !== String(id));
                syncDB();
                renderEkskulHistory();
                if (document.getElementById('view-rapor')?.classList.contains('active')) {
                    previewRaporSiswa();
                }
                Swal.fire('Terhapus!', 'Data nilai ekskul telah dihapus.', 'success');
            }
        });
    }

    // --- Kokurikuler (Pemisahan Tab Input Nilai & Riwayat Nilai) ---
    function switchKokuTab(tab) {
        const inputView = document.getElementById('koku-subview-input');
        const riwayatView = document.getElementById('koku-subview-riwayat');
        const btnInput = document.getElementById('btn-tab-koku-input');
        const btnRiwayat = document.getElementById('btn-tab-koku-riwayat');

        if (!inputView || !riwayatView || !btnInput || !btnRiwayat) return;

        if (tab === 'riwayat') {
            inputView.style.display = 'none';
            riwayatView.style.display = 'block';

            btnInput.className = 'btn btn-outline-primary';
            btnRiwayat.className = 'btn btn-primary';

            // Sinkronkan pilihan kelas
            const currentCid = document.getElementById('koku-kelas').value;
            const riwayatSelect = document.getElementById('koku-riwayat-kelas');
            if (riwayatSelect && currentCid) {
                riwayatSelect.value = currentCid;
            }
            loadKokuHistory();
        } else {
            // Default: 'input'
            inputView.style.display = 'block';
            riwayatView.style.display = 'none';

            btnInput.className = 'btn btn-primary';
            btnRiwayat.className = 'btn btn-outline-primary';
        }
    }

    function onKokuClassChange() {
        const val = document.getElementById('koku-kelas').value;
        const riwayatSelect = document.getElementById('koku-riwayat-kelas');
        if (riwayatSelect) riwayatSelect.value = val;
        // Sembunyikan hasil tabel penilaian sebelumnya saat ganti kelas
        const resSec = document.getElementById('koku-result-section');
        if (resSec) resSec.style.display = 'none';
        loadKokuHistory();
    }

    function onKokuRiwayatClassChange() {
        const val = document.getElementById('koku-riwayat-kelas').value;
        const inputSelect = document.getElementById('koku-kelas');
        if (inputSelect) inputSelect.value = val;
        loadKokuHistory();
    }

    function renderKokuDimensi() {
        let html = '';
        ['Keimanan & Ketakwaan', 'Kewargaan', 'Penalaran Kritis', 'Kreativitas', 'Kolaborasi', 'Kemandirian', 'Kesehatan', 'Komunikasi'].forEach((d, i) => {
            html += `<div class="col-md-3 col-6"><div class="form-check"><input class="form-check-input dim-checkbox" type="checkbox" value="${d}" id="dim-chk-${i}"><label class="form-check-label" style="font-size:0.85rem" for="dim-chk-${i}">${d}</label></div></div>`;
        });
        document.getElementById('koku-dimensi-container').innerHTML = html;
        if (!document.getElementById('koku-tanggal').value) {
            document.getElementById('koku-tanggal').value = tanggalHariIni();
        }
    }

    function loadKokuHistory() {
        const cid = document.getElementById('koku-riwayat-kelas')?.value || document.getElementById('koku-kelas')?.value;
        const tbody = document.getElementById('koku-history-tbody');
        const filterEl = document.getElementById('koku-export-filter');
        if (!tbody) return;

        if (!cid) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-center py-4 text-muted"><i class="bi bi-arrow-up-circle fs-3 d-block mb-1 text-primary"></i>Silakan pilih kelas terlebih dahulu untuk melihat riwayat penilaian.</td></tr>';
            if (filterEl) filterEl.innerHTML = '<option value="">Semua Projek (Kelas Ini)</option>';
            return;
        }

        const records = db.koku.filter(k => String(k.ClassID) === String(cid));
        if (records.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-center py-4 text-muted"><i class="bi bi-folder-x fs-3 d-block mb-1"></i>Belum ada data penilaian kokurikuler tersimpan untuk kelas ini.</td></tr>';
            if (filterEl) filterEl.innerHTML = '<option value="">Semua Projek (Kelas Ini)</option>';
            return;
        }

        let html = '';
        let filterOpts = '<option value="">Semua Projek (Kelas Ini)</option>';
        records.forEach(k => {
            let jumlah = 0;
            try { jumlah = Object.keys(JSON.parse(k.KokuData || '{}')).length; } catch (e) { }
            const jenisLabel = k.Jenis === '7KAIH' ? '7KAIH' : 'Lintas Mapel';
            const tglFormatted = formatTanggalIndo(k.Tanggal);
            html += `<tr>
                <td class="fw-bold">${tglFormatted}</td>
                <td><span class="badge bg-light text-primary border">${jenisLabel}</span></td>
                <td class="fw-semibold text-primary">${escapeHtml(k.Projek || '-')}</td>
                <td class="text-center"><span class="badge bg-success-subtle text-success border border-success-subtle">${jumlah} Siswa</span></td>
                <td class="text-nowrap text-center">
                    <button class="btn btn-sm btn-outline-primary me-1 mb-1" onclick="editKokuRecord('${k.RecordID}')" title="Lihat/Edit"><i class="bi bi-pencil-square"></i> Edit</button>
                    <button class="btn btn-sm btn-outline-danger mb-1" onclick="deleteKokuRecord('${k.RecordID}')" title="Hapus"><i class="bi bi-trash"></i></button>
                </td>
            </tr>`;
            filterOpts += `<option value="${k.RecordID}">${escapeHtml(k.Projek || '-')} (${k.Tanggal || '-'})</option>`;
        });
        tbody.innerHTML = html;
        if (filterEl) filterEl.innerHTML = filterOpts;
    }

    function generateKokuTable() {
        const cid = document.getElementById('koku-kelas').value, 
              p = document.getElementById('koku-projek').value, 
              desc = document.getElementById('koku-deskripsi').value;
        let selectedDims = []; 
        document.querySelectorAll('.dim-checkbox:checked').forEach(chk => selectedDims.push(chk.value));

        if (!cid || !p || !desc || selectedDims.length === 0) {
            return Swal.fire('Peringatan', 'Harap lengkapi kelas, nama projek, deskripsi, dan pilih minimal 1 dimensi profil!', 'warning');
        }

        const stds = db.students.filter(s => String(s.ClassID) === String(cid)); 
        if (stds.length === 0) return Swal.fire('Informasi', 'Tidak ada data siswa pada kelas ini.', 'info');

        editingKokuId = null;
        window.activeKokuProject = { p, desc, selectedDims }; 
        let html = '';
        stds.forEach((s, i) => {
            html += `<tr data-sid="${s.StudentID}" class="koku-row">
                <td class="text-center">${i + 1}</td>
                <td class="fw-bold text-primary">${escapeHtml(s.StudentName)}</td>
                <td>
                    <select class="form-select form-select-sm koku-cat-select" onchange="updateKokuNarasi(this)">
                        <option value="">Pilih Predikat</option>
                        <option value="Sangat Baik">Sangat Baik</option>
                        <option value="Baik">Baik</option>
                        <option value="Cukup">Cukup</option>
                        <option value="Kurang">Kurang</option>
                    </select>
                </td>
                <td><small class="koku-narasi-text text-muted">Pilih kategori...</small></td>
            </tr>`;
        });
        document.getElementById('koku-tbody').innerHTML = html; 
        document.getElementById('koku-result-section').style.display = 'block';

        const cancelBtn = document.getElementById('btn-cancel-edit-koku');
        if (cancelBtn) cancelBtn.style.display = 'inline-block';
    }

    function updateKokuNarasi(selectElem) {
        const cat = selectElem.value, row = selectElem.closest('tr'), narasiContainer = row.querySelector('.koku-narasi-text'), sName = row.cells[1].innerText;
        if (!cat) { narasiContainer.innerHTML = 'Pilih kategori...'; return; }
        const proj = window.activeKokuProject;
        let dimStr = proj.selectedDims.length > 1 ? proj.selectedDims.slice(0, -1).join(', ') + ' dan ' + proj.selectedDims.slice(-1) : proj.selectedDims[0];
        let saran = cat === 'Sangat Baik' ? 'Terus pertahankan inisiatif dan kepemimpinan ini.' : cat === 'Baik' ? 'Kembangkan terus potensi dan keaktifan ini.' : cat === 'Cukup' ? 'Perlu ditingkatkan fokus dan konsistensi belajarnya.' : 'Sangat disarankan lebih aktif bertanya dan berkolaborasi.';
        narasiContainer.innerHTML = `Ananda <strong>${escapeHtml(sName)}</strong> menunjukkan profil <strong>${dimStr}</strong> yang <strong>${cat.toLowerCase()}</strong> dalam kegiatan ${escapeHtml(proj.p)}. ${escapeHtml(proj.desc)} ${saran}`; 
        narasiContainer.classList.remove('text-muted');
    }

    function editKokuRecord(id) {
        const k = db.koku.find(x => x.RecordID === id);
        if (!k) return;

        document.getElementById('koku-kelas').value = k.ClassID;
        const riwayatSelect = document.getElementById('koku-riwayat-kelas');
        if (riwayatSelect) riwayatSelect.value = k.ClassID;

        document.getElementById('koku-tanggal').value = k.Tanggal || '';
        document.getElementById('koku-jenis').value = k.Jenis || 'LintasMapel';
        document.getElementById('koku-projek').value = k.Projek || '';
        document.getElementById('koku-deskripsi').value = k.Deskripsi || '';

        let selectedDims = []; try { selectedDims = JSON.parse(k.Dimensi || '[]'); } catch (e) { }
        document.querySelectorAll('.dim-checkbox').forEach(chk => { chk.checked = selectedDims.includes(chk.value); });

        let kokuData = {}; try { kokuData = JSON.parse(k.KokuData || '{}'); } catch (e) { }

        const stds = db.students.filter(s => String(s.ClassID) === String(k.ClassID));
        window.activeKokuProject = { p: k.Projek, desc: k.Deskripsi, selectedDims };
        let html = '';
        stds.forEach((s, i) => {
            const saved = kokuData[s.StudentID];
            const cat = saved ? saved.Kategori : '';
            const narasi = saved ? saved.Narasi : 'Pilih kategori...';
            html += `<tr data-sid="${s.StudentID}" class="koku-row">
                <td class="text-center">${i + 1}</td>
                <td class="fw-bold text-primary">${escapeHtml(s.StudentName)}</td>
                <td>
                    <select class="form-select form-select-sm koku-cat-select" onchange="updateKokuNarasi(this)">
                        <option value="">Pilih</option>
                        <option value="Sangat Baik" ${cat === 'Sangat Baik' ? 'selected' : ''}>Sangat Baik</option>
                        <option value="Baik" ${cat === 'Baik' ? 'selected' : ''}>Baik</option>
                        <option value="Cukup" ${cat === 'Cukup' ? 'selected' : ''}>Cukup</option>
                        <option value="Kurang" ${cat === 'Kurang' ? 'selected' : ''}>Kurang</option>
                    </select>
                </td>
                <td><small class="koku-narasi-text ${cat ? '' : 'text-muted'}">${narasi}</small></td>
            </tr>`;
        });
        document.getElementById('koku-tbody').innerHTML = html;
        document.getElementById('koku-result-section').style.display = 'block';
        editingKokuId = id;

        // Pindah ke tab Input Nilai
        switchKokuTab('input');

        const titleElem = document.getElementById('koku-form-title');
        if (titleElem) {
            titleElem.innerHTML = '<i class="bi bi-pencil-square text-warning"></i> Edit Penilaian Kokurikuler: ' + escapeHtml(k.Projek);
        }
        const cancelBtn = document.getElementById('btn-cancel-edit-koku');
        if (cancelBtn) cancelBtn.style.display = 'inline-block';

        window.scrollTo({ top: 0, behavior: 'smooth' });
        Swal.fire({ 
            icon: 'info', 
            title: 'Mode Edit Kokurikuler', 
            text: `Data projek "${k.Projek}" dimuat ke formulir. Silakan ubah lalu klik "Simpan Data Penilaian".`, 
            timer: 2500, 
            showConfirmButton: false 
        });
    }

    function cancelEditKoku() {
        editingKokuId = null;
        document.getElementById('koku-projek').value = '';
        document.getElementById('koku-deskripsi').value = '';
        document.getElementById('koku-result-section').style.display = 'none';
        document.querySelectorAll('.dim-checkbox').forEach(chk => chk.checked = false);
        const titleElem = document.getElementById('koku-form-title');
        if (titleElem) {
            titleElem.innerHTML = '<i class="bi bi-pencil-square" style="color: var(--teal-accent);"></i> Form Input Penilaian Kokurikuler';
        }
        const cancelBtn = document.getElementById('btn-cancel-edit-koku');
        if (cancelBtn) cancelBtn.style.display = 'none';
    }

    function deleteKokuRecord(id) {
        Swal.fire({
            title: 'Hapus Data Kokurikuler?', text: 'Seluruh nilai kokurikuler pada projek ini akan dihapus.', icon: 'warning',
            showCancelButton: true, confirmButtonColor: '#C44747', confirmButtonText: 'Ya, Hapus!', cancelButtonText: 'Batal'
        }).then((result) => {
            if (result.isConfirmed) {
                db.koku = db.koku.filter(k => k.RecordID !== id);
                if (editingKokuId === id) cancelEditKoku();
                syncDB(); loadKokuHistory();
                Swal.fire('Terhapus!', 'Data kokurikuler berhasil dihapus.', 'success');
            }
        });
    }

    function saveAllKoku() {
        const cid = document.getElementById('koku-kelas').value;
        const tanggal = document.getElementById('koku-tanggal').value;
        const jenis = document.getElementById('koku-jenis').value;
        const proj = window.activeKokuProject;
        if (!proj) return Swal.fire('Error', 'Silakan buat format tabel terlebih dahulu.', 'error');

        let kokuData = {}, count = 0;
        document.querySelectorAll('.koku-row').forEach(row => {
            const sid = row.getAttribute('data-sid'), cat = row.querySelector('.koku-cat-select').value;
            if (cat) { kokuData[sid] = { Kategori: cat, Narasi: row.querySelector('.koku-narasi-text').innerHTML }; count++; }
        });
        if (count === 0) return Swal.fire('Peringatan', 'Belum ada penilaian siswa yang dipilih!', 'warning');

        const wasEditing = !!editingKokuId;
        const record = {
            RecordID: editingKokuId || ("K" + Date.now()),
            ClassID: cid, Tanggal: tanggal, Jenis: jenis,
            Projek: proj.p, Deskripsi: proj.desc, Dimensi: JSON.stringify(proj.selectedDims),
            KokuData: JSON.stringify(kokuData)
        };

        if (editingKokuId) {
            const idx = db.koku.findIndex(k => k.RecordID === editingKokuId);
            if (idx > -1) db.koku[idx] = record; else db.koku.push(record);
        } else {
            db.koku.push(record);
        }

        cancelEditKoku();
        syncDB(); loadKokuHistory();

        Swal.fire({
            title: wasEditing ? 'Penilaian Diperbarui' : 'Penilaian Disimpan',
            text: `${count} Nilai kokurikuler siswa berhasil disimpan ke sistem.`,
            icon: 'success',
            showCancelButton: true,
            confirmButtonText: '<i class="bi bi-clock-history me-1"></i> Lihat Riwayat Nilai',
            cancelButtonText: '<i class="bi bi-plus-lg me-1"></i> Input Nilai Baru'
        }).then((result) => {
            if (result.isConfirmed) {
                switchKokuTab('riwayat');
            } else {
                switchKokuTab('input');
            }
        });
    }

    function getKokuExportData() {
        const cid = document.getElementById('koku-riwayat-kelas')?.value || document.getElementById('koku-kelas')?.value;
        if (!cid) { Swal.fire('Error', 'Pilih kelas terlebih dahulu.', 'error'); return null; }
        const filterId = document.getElementById('koku-export-filter').value;
        let records = db.koku.filter(k => String(k.ClassID) === String(cid));
        if (filterId) records = records.filter(k => k.RecordID === filterId);
        if (records.length === 0) { Swal.fire('Info', 'Belum ada data kokurikuler untuk diunduh pada kelas ini.', 'info'); return null; }

        const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
        const fileSuffix = filterId ? `_${records[0].Projek.replace(/[^a-z0-9]+/gi, '_')}` : '';
        return { records, className, filterId, fileSuffix, cid };
    }

    // --- Unduh PDF Nilai Kokurikuler Berformat Sekolah Profesional ---
    async function downloadKokuPDF() {
        const data = getKokuExportData();
        if (!data) return;
        const { records, className, filterId, fileSuffix } = data;

        Swal.fire({
            title: 'Membuat Dokumen PDF...',
            text: 'Menyusun format rekapitulasi penilaian kokurikuler profesional...',
            allowOutsideClick: false,
            didOpen: () => { Swal.showLoading(); }
        });

        try {
            await ensurePDF();
            const { jsPDF } = window.jspdf;

            // Dokumen A4 Landscape
            const doc = new jsPDF({
                orientation: 'landscape',
                unit: 'mm',
                format: 'a4'
            });

            // 1. Kop Dokumen Resmi Sekolah
            const schoolName = (db.config.userSchool || 'SEKOLAH DASAR / MENENGAH').toUpperCase();
            const teacherName = db.config.userName || RAPOR_KOSONG;

            doc.setFont('helvetica', 'bold');
            doc.setFontSize(14);
            doc.setTextColor(23, 43, 77); // Deep Navy #172B4D
            doc.text(schoolName, 148.5, 15, { align: 'center' });

            doc.setFontSize(12);
            doc.setTextColor(36, 59, 90); // Navy Secondary #243B5A
            doc.text('REKAPITULASI PENILAIAN PROJEK KOKURIKULER', 148.5, 21, { align: 'center' });

            // Garis Pembatas Kop Ganda
            doc.setDrawColor(23, 43, 77);
            doc.setLineWidth(0.6);
            doc.line(15, 24.5, 282, 24.5);
            doc.setDrawColor(148, 163, 184);
            doc.setLineWidth(0.2);
            doc.line(15, 25.5, 282, 25.5);

            // Hitung Tahun Pelajaran & Semester Berjalan
            const curDate = new Date();
            const curYear = curDate.getFullYear();
            const curMonth = curDate.getMonth() + 1;
            const academicYear = curMonth >= 7 ? `${curYear}/${curYear + 1}` : `${curYear - 1}/${curYear}`;
            const semester = curMonth >= 7 ? 'Ganjil' : 'Genap';

            // 2. Informasi Metadata Dokumen Format Sekolah
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            doc.setTextColor(31, 41, 55);

            // Kolom Kiri
            doc.text(`Satuan Pendidikan : ${schoolName}`, 15, 30);
            doc.text(`Kelas / Rombel       : ${className}`, 15, 35);

            // Kolom Kanan
            doc.text(`Tahun Pelajaran   : ${academicYear} (Semester ${semester})`, 185, 30);
            doc.text(`Fasilitator / Guru : ${teacherName}`, 185, 35);

            let startY = 40;

            // Jika filter projek tunggal dipilih, sertakan rincian tema projek
            if (filterId && records.length === 1) {
                const rec = records[0];
                let dims = [];
                try { dims = JSON.parse(rec.Dimensi || '[]'); } catch (e) { }
                const dimStr = dims.length > 0 ? dims.join(', ') : '-';

                doc.setFont('helvetica', 'bold');
                doc.text(`Nama Projek : ${rec.Projek} (${rec.Jenis === '7KAIH' ? '7KAIH' : 'Lintas Mata Pelajaran'})`, 15, 41);
                doc.setFont('helvetica', 'normal');
                doc.text(`Tanggal: ${formatTanggalIndo(rec.Tanggal)}  |  Dimensi Profil: ${dimStr}`, 15, 46);
                startY = 51;
            }

            // 3. Susun Baris Data Tabel
            let body = [];
            let rowIdx = 1;

            records.forEach(k => {
                let kokuData = {};
                try { kokuData = JSON.parse(k.KokuData || '{}'); } catch (e) { }

                Object.keys(kokuData).forEach(sid => {
                    const student = db.students.find(s => String(s.StudentID) === String(sid));
                    const studentName = student ? student.StudentName : sid;
                    const cat = kokuData[sid].Kategori || '-';
                    const rawNarasi = (kokuData[sid].Narasi || '').replace(/<[^>]+>/g, '').trim();

                    body.push([
                        (rowIdx++).toString(),
                        studentName,
                        k.Projek || '-',
                        cat,
                        rawNarasi
                    ]);
                });
            });

            const tableHead = [['NO', 'NAMA SISWA', 'NAMA PROJEK', 'PREDIKAT / CAPAIAN', 'DESKRIPSI CAPAIAN PROFIL PELAJAR']];

            doc.autoTable({
                head: tableHead,
                body: body,
                startY: startY,
                margin: { left: 15, right: 15, bottom: 18 },
                theme: 'grid',
                headStyles: {
                    fillColor: [23, 43, 77], // #172B4D Deep Navy
                    textColor: [255, 255, 255],
                    fontStyle: 'bold',
                    halign: 'center',
                    valign: 'middle',
                    fontSize: 8.5,
                    cellPadding: 3
                },
                bodyStyles: {
                    fontSize: 8,
                    textColor: [31, 41, 55],
                    valign: 'middle',
                    lineColor: [226, 232, 240],
                    lineWidth: 0.2,
                    cellPadding: 2.5
                },
                alternateRowStyles: {
                    fillColor: [248, 250, 252]
                },
                columnStyles: {
                    0: { cellWidth: 12, halign: 'center' }, // No
                    1: { cellWidth: 50, halign: 'left' },   // Nama Siswa
                    2: { cellWidth: 44, halign: 'left' },   // Projek
                    3: { cellWidth: 35, halign: 'center', fontStyle: 'bold' }, // Predikat
                    4: { cellWidth: 126, halign: 'left' }   // Deskripsi
                }
            });

            // 4. Kolom Tanda Tangan / Pengesahan Dokumen
            let finalY = doc.lastAutoTable.finalY + 10;
            if (finalY + 38 > 210 - 15) {
                doc.addPage();
                finalY = 22;
            }

            const colLeft = 25;
            const colRight = 205;

            doc.setFont('helvetica', 'normal');
            doc.setFontSize(8.5);
            doc.setTextColor(31, 41, 55);

            // Kolom Kiri: Kepala Sekolah
            doc.text('Mengetahui,', colLeft, finalY);
            doc.text('Kepala Sekolah', colLeft, finalY + 4.5);

            // Kolom Kanan: Fasilitator / Guru
            const signDateStr = formatTanggalIndo(tanggalHariIni()).replace(/^[A-Za-z]+,\s*/, '');
            doc.text(`...................., ${signDateStr}`, colRight, finalY);
            doc.text('Fasilitator Projek / Guru Kelas,', colRight, finalY + 4.5);

            // Ruang Tanda Tangan
            finalY += 24;
            doc.setFont('helvetica', 'bold');
            doc.text('( .................................................... )', colLeft, finalY);
            doc.setFont('helvetica', 'normal');
            doc.text('NIP. ............................................', colLeft, finalY + 4.5);

            doc.setFont('helvetica', 'bold');
            doc.text(`( ${teacherName} )`, colRight, finalY);
            doc.setFont('helvetica', 'normal');
            doc.text('NIP. ............................................', colRight, finalY + 4.5);

            // 5. Nomor Halaman & Footer
            const totalPages = doc.internal.getNumberOfPages();
            for (let i = 1; i <= totalPages; i++) {
                doc.setPage(i);
                doc.setFontSize(8);
                doc.setFont('helvetica', 'italic');
                doc.setTextColor(148, 163, 184);
                doc.text('AgendaGuru – Aplikasi Manajemen Kelas Digital', 15, 203);
                doc.text(`Halaman ${i} dari ${totalPages}`, 282, 203, { align: 'right' });
            }

            Swal.close();
            const safeClassName = className.replace(/[^a-zA-Z0-9_-]/g, '_');
            doc.save(`Rekap_Kokurikuler_${safeClassName}${fileSuffix}.pdf`);
            Swal.fire('Berhasil!', 'Dokumen PDF Penilaian Kokurikuler profesional berhasil diunduh.', 'success');
        } catch (err) {
            console.error('Error saat membuat PDF kokurikuler:', err);
            Swal.fire('Error', 'Gagal memproses dokumen PDF: ' + (err.message || err), 'error');
        }
    }

    // Export ke Word (.doc) memakai trik HTML-sebagai-dokumen-Word
    function downloadKokuWord() {
        const data = getKokuExportData();
        if (!data) return;
        const { records, className, fileSuffix } = data;

        let bodyHtml = `<h2 style="font-family:Calibri,Arial,sans-serif; text-align:center;">REKAPITULASI PENILAIAN PROJEK KOKURIKULER</h2>`;
        bodyHtml += `<p style="text-align:center; font-family:Calibri,Arial,sans-serif;">Kelas: ${escapeHtml(className)} | Satuan Pendidikan: ${escapeHtml(db.config.userSchool || 'Sekolah')}</p><hr/>`;

        records.forEach(k => {
            let kokuData = {}; try { kokuData = JSON.parse(k.KokuData || '{}'); } catch (e) { }
            bodyHtml += `<h3 style="font-family:Calibri,Arial,sans-serif;">${escapeHtml(k.Projek || '-')}${k.Tanggal ? ' (' + formatTanggalIndo(k.Tanggal) + ')' : ''}</h3>`;
            if (k.Deskripsi) bodyHtml += `<p style="font-family:Calibri,Arial,sans-serif;"><i>${escapeHtml(k.Deskripsi)}</i></p>`;
            bodyHtml += `<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; width:100%; font-family:Calibri,Arial,sans-serif; font-size:10pt;">
                    <tr style="background:#172B4D; color:#ffffff;"><th style="width:25%;">Nama Siswa</th><th style="width:20%;">Predikat</th><th>Deskripsi Capaian Profil</th></tr>`;
            Object.keys(kokuData).forEach(sid => {
                const student = db.students.find(s => String(s.StudentID) === String(sid));
                const name = student ? student.StudentName : sid;
                bodyHtml += `<tr><td>${escapeHtml(name)}</td><td style="text-align:center; font-weight:bold;">${escapeHtml(kokuData[sid].Kategori)}</td><td>${kokuData[sid].Narasi.replace(/<[^>]+>/g, '')}</td></tr>`;
            });
            bodyHtml += `</table><br>`;
        });

        const htmlDoc = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
                <head><meta charset="utf-8"><title>Rekap Kokurikuler</title></head>
                <body>${bodyHtml}</body></html>`;

        const blob = new Blob(['\ufeff', htmlDoc], { type: 'application/msword' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `Rekap_Kokurikuler_${className}${fileSuffix}.doc`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
    }

    // --- Jurnal Harian (Pemisahan Tampilan Input Jurnal & Riwayat Jurnal) ---
    function switchJurnalTab(tab) {
        const inputView = document.getElementById('jurnal-subview-input');
        const riwayatView = document.getElementById('jurnal-subview-riwayat');
        const btnInput = document.getElementById('btn-tab-jurnal-input');
        const btnRiwayat = document.getElementById('btn-tab-jurnal-riwayat');

        if (!inputView || !riwayatView || !btnInput || !btnRiwayat) return;

        if (tab === 'riwayat') {
            inputView.style.display = 'none';
            riwayatView.style.display = 'block';

            btnInput.className = 'btn btn-outline-primary';
            btnRiwayat.className = 'btn btn-primary';

            renderJurnal();
        } else {
            // Default: 'input'
            inputView.style.display = 'block';
            riwayatView.style.display = 'none';

            btnInput.className = 'btn btn-primary';
            btnRiwayat.className = 'btn btn-outline-primary';

            const dateInput = document.getElementById('jurnal-date');
            if (dateInput && !dateInput.value) {
                dateInput.value = tanggalHariIni();
            }
        }
    }

    // Helper format tanggal Indonesia
    function formatTanggalIndo(dateStr) {
        if (!dateStr) return '-';
        try {
            const parts = dateStr.split('-');
            if (parts.length === 3) {
                const bln = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
                const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
                const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
                if (!isNaN(d.getTime())) {
                    return `${days[d.getDay()]}, ${d.getDate()} ${bln[d.getMonth()]} ${d.getFullYear()}`;
                }
            }
        } catch (e) { }
        return dateStr;
    }

    // Helper kompresi gambar foto dokumentasi untuk efisiensi penyimpanan & PDF
    function compressJournalImage(file) {
        return new Promise((resolve) => {
            if (!file) return resolve(null);
            if (!file.type || !file.type.startsWith('image/')) {
                const r = new FileReader();
                r.onload = (e) => resolve({ isImage: false, dataUrl: e.target.result, raw: e.target.result });
                r.onerror = () => resolve(null);
                r.readAsDataURL(file);
                return;
            }
            const reader = new FileReader();
            reader.onload = (e) => {
                const img = new Image();
                img.onload = () => {
                    const maxWidth = 900;
                    const maxHeight = 900;
                    let w = img.width;
                    let h = img.height;
                    if (w > maxWidth || h > maxHeight) {
                        if (w / h > maxWidth / maxHeight) {
                            h = Math.round((h * maxWidth) / w);
                            w = maxWidth;
                        } else {
                            w = Math.round((w * maxHeight) / h);
                            h = maxHeight;
                        }
                    }
                    const canvas = document.createElement('canvas');
                    canvas.width = w;
                    canvas.height = h;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0, w, h);
                    const compressed = canvas.toDataURL('image/jpeg', 0.75);
                    resolve({ isImage: true, dataUrl: compressed, raw: e.target.result, width: w, height: h });
                };
                img.onerror = () => resolve({ isImage: true, dataUrl: e.target.result, raw: e.target.result });
                img.src = e.target.result;
            };
            reader.onerror = () => resolve(null);
            reader.readAsDataURL(file);
        });
    }

    // Helper konversi URL gambar ke Base64 (jika ada gambar eksternal)
    function loadImageToDataUrl(url) {
        return new Promise((resolve) => {
            if (!url) return resolve(null);
            if (url.startsWith('data:image')) return resolve(url);
            const img = new Image();
            img.crossOrigin = 'Anonymous';
            img.onload = () => {
                try {
                    const canvas = document.createElement('canvas');
                    canvas.width = img.naturalWidth || img.width;
                    canvas.height = img.naturalHeight || img.height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0);
                    resolve(canvas.toDataURL('image/jpeg', 0.8));
                } catch (e) {
                    resolve(null);
                }
            };
            img.onerror = () => resolve(null);
            img.src = url;
        });
    }

    let currentJurnalPhotoBase64 = null; // null: pertahankan lama, '': hapus foto, string: data foto baru
    let currentJurnalFileName = '';

    async function previewJurnalFile(event) {
        const file = event.target.files[0];
        const previewBox = document.getElementById('jurnal-file-preview');
        const previewImg = document.getElementById('jurnal-file-preview-img');
        const fileNameEl = document.getElementById('jurnal-file-name');
        const fileDescEl = document.getElementById('jurnal-file-desc');

        if (!file) {
            clearJurnalFile();
            return;
        }

        currentJurnalFileName = file.name;
        if (file.type && file.type.startsWith('image/')) {
            const res = await compressJournalImage(file);
            if (res && res.dataUrl) {
                currentJurnalPhotoBase64 = res.dataUrl;
                if (previewImg) {
                    previewImg.src = res.dataUrl;
                    previewImg.style.display = 'block';
                }
                if (fileNameEl) fileNameEl.textContent = file.name;
                if (fileDescEl) fileDescEl.textContent = 'Foto Dokumentasi (' + Math.round(file.size / 1024) + ' KB)';
                if (previewBox) previewBox.style.display = 'block';
            }
        } else {
            currentJurnalPhotoBase64 = null;
            if (previewImg) previewImg.style.display = 'none';
            if (fileNameEl) fileNameEl.textContent = file.name;
            if (fileDescEl) fileDescEl.textContent = 'Dokumen Lampiran (' + Math.round(file.size / 1024) + ' KB)';
            if (previewBox) previewBox.style.display = 'block';
        }
    }

    function clearJurnalFile() {
        const input = document.getElementById('jurnal-file');
        if (input) input.value = '';
        currentJurnalPhotoBase64 = ''; // tanda dihapus oleh user
        currentJurnalFileName = '';
        const previewBox = document.getElementById('jurnal-file-preview');
        if (previewBox) previewBox.style.display = 'none';
        const previewImg = document.getElementById('jurnal-file-preview-img');
        if (previewImg) previewImg.src = '';
    }

    function previewImageModal(src) {
        if (!src) return;
        Swal.fire({
            title: 'Dokumentasi Kegiatan Pembelajaran',
            imageUrl: src,
            imageAlt: 'Foto Dokumentasi Pembelajaran',
            showConfirmButton: false,
            showCloseButton: true,
            width: 'auto'
        });
    }

    function renderJurnal() {
        const classFilter = document.getElementById('filter-jurnal-kelas').value;
        const tbody = document.getElementById('jurnal-tbody');

        if (!classFilter) {
            tbody.innerHTML = '<tr><td colspan="6" class="text-center py-4 text-muted"><i class="bi bi-arrow-up-circle fs-3 d-block mb-1 text-primary"></i>Silakan pilih kelas terlebih dahulu untuk melihat riwayat jurnal.</td></tr>';
            return;
        }

        let filtered = db.journal.filter(j => String(j.ClassID) === String(classFilter) || j.ClassID === 'All');

        // Pencarian jika kata kunci diinput
        const searchInput = document.getElementById('search-jurnal-text');
        const query = searchInput ? searchInput.value.toLowerCase().trim() : '';
        if (query) {
            filtered = filtered.filter(j => 
                (j.Subject && j.Subject.toLowerCase().includes(query)) ||
                (j.Content && j.Content.toLowerCase().includes(query)) ||
                (j.Date && j.Date.toLowerCase().includes(query))
            );
        }

        if (filtered.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" class="text-center py-4 text-muted"><i class="bi bi-journal-x fs-3 d-block mb-1"></i>Belum ada catatan jurnal pada kelas ini' + (query ? ' yang cocok dengan kata kunci "' + query + '"' : '') + '.</td></tr>';
            return;
        }

        let html = '';
        filtered.forEach(j => {
            const imgSrc = j.AttachmentData || (j.AttachmentUrl && j.AttachmentUrl.startsWith('data:image') ? j.AttachmentUrl : '');
            let attachHtml = '<span class="text-muted small">-</span>';
            if (imgSrc) {
                attachHtml = `<div class="d-flex align-items-center gap-2">
                    <img src="${imgSrc}" alt="Bukti Foto" class="rounded border shadow-sm" style="width: 44px; height: 44px; object-fit: cover; cursor: pointer; transition: transform 0.15s;" onmouseover="this.style.transform='scale(1.08)'" onmouseout="this.style.transform='scale(1)'" onclick="previewImageModal('${imgSrc}')" title="Klik untuk memperbesar foto">
                    <span class="badge bg-light text-primary border" style="font-size:0.75rem;"><i class="bi bi-image"></i> Foto</span>
                </div>`;
            } else if (j.AttachmentUrl && j.AttachmentUrl !== 'https://simulasi-url-drive.com/file') {
                attachHtml = `<a href="${j.AttachmentUrl}" target="_blank" class="btn btn-sm btn-outline-primary py-0 px-2" style="font-size:0.75rem;"><i class="bi bi-paperclip"></i> Lampiran</a>`;
            }

            let classBadge = j.ClassID === 'All' ? 'Semua Kelas' : (db.classes.find(c => String(c.ClassID) === String(j.ClassID))?.ClassName || j.ClassID);
            html += `<tr>
                <td class="fw-bold">${j.Date}</td>
                <td><span class="badge bg-light text-primary border">${classBadge}</span></td>
                <td class="fw-semibold text-primary">${escapeHtml(j.Subject)}</td>
                <td><small class="text-secondary">${escapeHtml(j.Content)}</small></td>
                <td>${attachHtml}</td>
                <td class="text-nowrap text-center">
                    <button class="btn btn-sm btn-outline-primary me-1 mb-1" onclick="editJournal('${j.RecordID}')" title="Edit Jurnal"><i class="bi bi-pencil-square"></i> Edit</button>
                    <button class="btn btn-sm btn-outline-danger mb-1" onclick="deleteJournal('${j.RecordID}')" title="Hapus Jurnal"><i class="bi bi-trash"></i></button>
                </td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    async function saveJurnal() {
        const d = document.getElementById('jurnal-date').value, 
              k = document.getElementById('jurnal-kelas').value, 
              m = document.getElementById('jurnal-mapel').value, 
              c = document.getElementById('jurnal-content').value;
        const fileInput = document.getElementById('jurnal-file');
        const syncCal = document.getElementById('jurnal-sync-cal').checked;
        if (!d || !k || !m || !c) return Swal.fire('Error', 'Harap isi semua field utama!', 'error');

        Swal.fire({ title: 'Menyimpan...', allowOutsideClick: false, didOpen: () => { Swal.showLoading(); } });

        if (fileInput.files.length > 0) {
            const file = fileInput.files[0];
            const processed = await compressJournalImage(file);
            const photoDataUrl = (processed && processed.isImage) ? processed.dataUrl : '';
            const rawDataUrl = processed ? processed.raw : '';

            if (typeof google !== 'undefined') {
                google.script.run
                    .withSuccessHandler(function (url) {
                        finishSaveJurnal(d, k, m, c, url, photoDataUrl, syncCal);
                    })
                    .withFailureHandler(function () {
                        // Jika upload ke drive gagal, tetap simpan base64 foto secara lokal
                        finishSaveJurnal(d, k, m, c, photoDataUrl || "", photoDataUrl, syncCal);
                    })
                    .uploadFileToDrive(rawDataUrl, file.name, file.type, getToken());
            } else {
                // Simulasi offline/lokal: gunakan base64 langsung sebagai URL dan data foto
                finishSaveJurnal(d, k, m, c, photoDataUrl || "https://simulasi-url-drive.com/file", photoDataUrl, syncCal);
            }
        } else {
            // Tidak ada file baru yang dipilih pada input
            if (currentJurnalPhotoBase64 === '') {
                // User klik tombol hapus foto saat edit
                finishSaveJurnal(d, k, m, c, "", "", syncCal);
            } else {
                // Pertahankan file lama jika dalam mode edit
                finishSaveJurnal(d, k, m, c, null, null, syncCal);
            }
        }
    }

    function finishSaveJurnal(d, k, m, c, fileUrl, photoBase64, syncCal) {
        let className = db.classes.find(cls => String(cls.ClassID) === String(k))?.ClassName || k;
        if (syncCal && typeof google !== 'undefined') {
            google.script.run.syncToCalendar('jurnal', { date: d, subject: m, className: className, content: c }, getToken());
        }

        const wasEditing = !!editingJournalId;
        if (editingJournalId) {
            const idx = db.journal.findIndex(j => j.RecordID === editingJournalId);
            if (idx > -1) {
                db.journal[idx].Date = d; 
                db.journal[idx].ClassID = k; 
                db.journal[idx].Subject = m; 
                db.journal[idx].Content = c;
                if (fileUrl !== null) db.journal[idx].AttachmentUrl = fileUrl;
                if (photoBase64 !== null) db.journal[idx].AttachmentData = photoBase64;
            }
            cancelEditJurnal();
        } else {
            db.journal.push({ 
                RecordID: "J" + Date.now(), 
                Date: d, 
                ClassID: k, 
                Subject: m, 
                Content: c, 
                AttachmentUrl: fileUrl || "",
                AttachmentData: photoBase64 || ""
            });
        }

        document.getElementById('filter-jurnal-kelas').value = k; 
        renderJurnal(); 
        syncDB();
        document.getElementById('jurnal-content').value = ''; 
        clearJurnalFile();

        Swal.fire({
            title: wasEditing ? 'Jurnal Diperbarui' : 'Jurnal Disimpan',
            text: wasEditing ? 'Perubahan catatan jurnal berhasil disimpan.' : 'Catatan jurnal pembelajaran baru berhasil ditambahkan.',
            icon: 'success',
            showCancelButton: true,
            confirmButtonText: '<i class="bi bi-clock-history me-1"></i> Lihat Riwayat Jurnal',
            cancelButtonText: '<i class="bi bi-plus-lg me-1"></i> Tambah Jurnal Lagi'
        }).then((result) => {
            if (result.isConfirmed) {
                switchJurnalTab('riwayat');
            } else {
                switchJurnalTab('input');
            }
        });
    }

    function editJournal(id) {
        const j = db.journal.find(x => x.RecordID === id);
        if (!j) return;
        editingJournalId = id;
        document.getElementById('jurnal-date').value = j.Date;
        document.getElementById('jurnal-kelas').value = j.ClassID;
        document.getElementById('jurnal-mapel').value = j.Subject;
        document.getElementById('jurnal-content').value = j.Content;
        
        // Reset file input dan periksa apakah ada lampiran/foto tersimpan
        const input = document.getElementById('jurnal-file');
        if (input) input.value = '';
        currentJurnalPhotoBase64 = null; // null = jika tidak diubah, pertahankan
        
        const previewBox = document.getElementById('jurnal-file-preview');
        const previewImg = document.getElementById('jurnal-file-preview-img');
        const fileNameEl = document.getElementById('jurnal-file-name');
        const fileDescEl = document.getElementById('jurnal-file-desc');
        
        const existingImg = j.AttachmentData || (j.AttachmentUrl && j.AttachmentUrl.startsWith('data:image') ? j.AttachmentUrl : '');
        if (existingImg) {
            if (previewImg) { previewImg.src = existingImg; previewImg.style.display = 'block'; }
            if (fileNameEl) fileNameEl.textContent = 'Foto Dokumentasi Tersimpan';
            if (fileDescEl) fileDescEl.textContent = 'Klik ikon "x" jika ingin menghapus atau pilih file baru untuk mengganti';
            if (previewBox) previewBox.style.display = 'block';
        } else if (j.AttachmentUrl && j.AttachmentUrl !== 'https://simulasi-url-drive.com/file') {
            if (previewImg) previewImg.style.display = 'none';
            if (fileNameEl) fileNameEl.textContent = 'Lampiran Tersimpan';
            if (fileDescEl) fileDescEl.textContent = j.AttachmentUrl;
            if (previewBox) previewBox.style.display = 'block';
        } else {
            if (previewBox) previewBox.style.display = 'none';
        }

        document.getElementById('btn-save-jurnal').innerHTML = '<i class="bi bi-save me-1"></i> Update Jurnal';
        document.getElementById('btn-cancel-edit-jurnal').style.display = 'inline-block';

        // Pindah otomatis ke tampilan formulir Input Jurnal
        switchJurnalTab('input');

        const titleElem = document.getElementById('jurnal-form-title');
        if (titleElem) {
            titleElem.innerHTML = '<i class="bi bi-pencil-square text-warning"></i> Edit Jurnal Harian';
        }
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function cancelEditJurnal() {
        editingJournalId = null;
        document.getElementById('btn-save-jurnal').innerHTML = '<i class="bi bi-save me-1"></i> Simpan';
        document.getElementById('btn-cancel-edit-jurnal').style.display = 'none';
        document.getElementById('jurnal-content').value = '';
        clearJurnalFile();
        currentJurnalPhotoBase64 = null;
        const titleElem = document.getElementById('jurnal-form-title');
        if (titleElem) {
            titleElem.innerHTML = '<i class="bi bi-pencil-square" style="color: var(--teal-accent);"></i> Form Input Jurnal Harian';
        }
    }

    function deleteJournal(id) {
        Swal.fire({
            title: 'Hapus Jurnal?', text: "Apakah Anda yakin ingin menghapus catatan jurnal ini?", icon: 'warning',
            showCancelButton: true, confirmButtonColor: '#C44747', confirmButtonText: 'Ya, Hapus!'
        }).then((result) => {
            if (result.isConfirmed) {
                db.journal = db.journal.filter(j => j.RecordID !== id);
                if (editingJournalId === id) cancelEditJurnal();
                syncDB(); renderJurnal();
                Swal.fire('Terhapus!', 'Jurnal berhasil dihapus.', 'success');
            }
        });
    }

    // --- Ekspor Dokumen PDF Jurnal Harian Profesional dengan Foto Dokumentasi ---
    async function downloadJurnalPDF() {
        const cFilter = document.getElementById('filter-jurnal-kelas').value;
        if (!cFilter) return Swal.fire('Peringatan', 'Silakan pilih kelas terlebih dahulu sebelum mengunduh PDF!', 'warning');

        let filtered = db.journal.filter(j => String(j.ClassID) === String(cFilter) || j.ClassID === 'All');
        if (filtered.length === 0) return Swal.fire('Informasi', 'Belum ada catatan jurnal pada kelas yang dipilih.', 'info');

        // Urutkan data berdasarkan tanggal
        filtered.sort((a, b) => (a.Date > b.Date ? 1 : -1));

        const className = cFilter === 'All' ? 'Semua Kelas' : (db.classes.find(c => String(c.ClassID) === String(cFilter))?.ClassName || cFilter);

        Swal.fire({
            title: 'Membuat Dokumen PDF...',
            text: 'Menyusun format jurnal profesional & memuat foto dokumentasi...',
            allowOutsideClick: false,
            didOpen: () => { Swal.showLoading(); }
        });

        try {
            await ensurePDF();
            const { jsPDF } = window.jspdf;

            // Dokumen A4 Landscape untuk keterbacaan tabel dan kolom foto yang proporsional
            const doc = new jsPDF({
                orientation: 'landscape',
                unit: 'mm',
                format: 'a4'
            });

            // 1. Kop Dokumen Resmi
            const schoolName = (db.config.userSchool || 'SEKOLAH DASAR / MENENGAH').toUpperCase();
            const teacherName = db.config.userName || RAPOR_KOSONG;

            doc.setFont('helvetica', 'bold');
            doc.setFontSize(14);
            doc.setTextColor(23, 43, 77); // Deep Navy #172B4D
            doc.text(schoolName, 148.5, 15, { align: 'center' });

            doc.setFontSize(12);
            doc.setTextColor(36, 59, 90); // Navy Secondary #243B5A
            doc.text('AGENDA & JURNAL KEGIATAN PEMBELAJARAN GURU', 148.5, 21, { align: 'center' });

            // Garis Pembatas Kop Ganda
            doc.setDrawColor(23, 43, 77);
            doc.setLineWidth(0.6);
            doc.line(15, 24.5, 282, 24.5);
            doc.setDrawColor(148, 163, 184);
            doc.setLineWidth(0.2);
            doc.line(15, 25.5, 282, 25.5);

            // Hitung Tahun Pelajaran & Semester Berjalan
            const curDate = new Date();
            const curYear = curDate.getFullYear();
            const curMonth = curDate.getMonth() + 1;
            const academicYear = curMonth >= 7 ? `${curYear}/${curYear + 1}` : `${curYear - 1}/${curYear}`;
            const semester = curMonth >= 7 ? 'Ganjil' : 'Genap';

            // Informasi Metadata Dokumen Format Sekolah Resmi
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            doc.setTextColor(31, 41, 55);

            // Kolom Kiri
            doc.text(`Satuan Pendidikan : ${schoolName}`, 15, 30);
            doc.text(`Kelas / Rombel       : ${className}`, 15, 35);

            // Kolom Kanan (Informasi Administrasi Akademik Sekolah)
            doc.text(`Tahun Pelajaran   : ${academicYear} (Semester ${semester})`, 185, 30);
            doc.text(`Guru Pengampu     : ${teacherName}`, 185, 35);

            // 2. Preload & Siapkan Data Baris beserta Gambar Foto
            const rowsData = [];
            for (let i = 0; i < filtered.length; i++) {
                const j = filtered[i];
                let imgSrc = j.AttachmentData || '';
                if (!imgSrc && j.AttachmentUrl && j.AttachmentUrl.startsWith('data:image')) {
                    imgSrc = j.AttachmentUrl;
                }
                if (!imgSrc && j.AttachmentUrl && (j.AttachmentUrl.startsWith('http') || j.AttachmentUrl.startsWith('blob'))) {
                    try {
                        imgSrc = await loadImageToDataUrl(j.AttachmentUrl);
                    } catch (e) {
                        imgSrc = null;
                    }
                }

                const clsName = j.ClassID === 'All' ? 'Semua Kelas' : (db.classes.find(c => String(c.ClassID) === String(j.ClassID))?.ClassName || j.ClassID);
                const tglFormatted = formatTanggalIndo(j.Date);

                rowsData.push({
                    no: (i + 1).toString(),
                    tanggal: tglFormatted,
                    kelas: clsName,
                    mapel: j.Subject || '-',
                    isi: j.Content || '-',
                    imgData: imgSrc,
                    hasAttachNonImg: !imgSrc && j.AttachmentUrl && j.AttachmentUrl !== 'https://simulasi-url-drive.com/file' ? 'Ada Dokumen' : ''
                });
            }

            const tableHead = [['NO', 'HARI / TANGGAL', 'KELAS', 'MATA PELAJARAN', 'URAIAN KEGIATAN PEMBELAJARAN', 'DOKUMENTASI FOTO']];
            const tableBody = rowsData.map(r => [
                r.no,
                r.tanggal,
                r.kelas,
                r.mapel,
                r.isi,
                r.imgData ? '' : (r.hasAttachNonImg ? 'Lampiran File' : 'Tidak Ada Foto')
            ]);

            doc.autoTable({
                head: tableHead,
                body: tableBody,
                startY: 40,
                margin: { left: 15, right: 15, bottom: 18 },
                theme: 'grid',
                headStyles: {
                    fillColor: [23, 43, 77], // #172B4D Deep Navy
                    textColor: [255, 255, 255],
                    fontStyle: 'bold',
                    halign: 'center',
                    valign: 'middle',
                    fontSize: 8.5,
                    cellPadding: 3
                },
                bodyStyles: {
                    fontSize: 8,
                    textColor: [31, 41, 55],
                    valign: 'middle',
                    lineColor: [226, 232, 240],
                    lineWidth: 0.2,
                    cellPadding: 2.5
                },
                alternateRowStyles: {
                    fillColor: [248, 250, 252]
                },
                columnStyles: {
                    0: { cellWidth: 12, halign: 'center' }, // No
                    1: { cellWidth: 34, halign: 'center' }, // Hari/Tanggal
                    2: { cellWidth: 24, halign: 'center' }, // Kelas
                    3: { cellWidth: 36, halign: 'left' },   // Mapel
                    4: { cellWidth: 111, halign: 'left' },  // Uraian Kegiatan
                    5: { cellWidth: 50, halign: 'center' }  // Foto Dokumentasi
                },
                didParseCell: function (data) {
                    if (data.section === 'body' && data.column.index === 5) {
                        const row = rowsData[data.row.index];
                        if (row && row.imgData) {
                            // Berikan tinggi minimal yang lega untuk sel foto
                            data.cell.styles.minCellHeight = 30;
                        }
                    }
                },
                didDrawCell: function (data) {
                    if (data.section === 'body' && data.column.index === 5) {
                        const row = rowsData[data.row.index];
                        if (row && row.imgData) {
                            const padding = 2;
                            const maxW = data.cell.width - (padding * 2);
                            const maxH = data.cell.height - (padding * 2);
                            const targetW = Math.min(maxW, 44);
                            const targetH = Math.min(maxH, 26);
                            const posX = data.cell.x + (data.cell.width - targetW) / 2;
                            const posY = data.cell.y + (data.cell.height - targetH) / 2;
                            try {
                                doc.addImage(row.imgData, 'JPEG', posX, posY, targetW, targetH);
                            } catch (e) {
                                console.warn('Gagal me-render foto pada baris PDF:', e);
                            }
                        }
                    }
                }
            });

            // 3. Kolom Tanda Tangan / Pengesahan Dokumen
            let finalY = doc.lastAutoTable.finalY + 10;
            // Cek apakah halaman masih muat untuk blok tanda tangan (~38mm)
            if (finalY + 38 > 210 - 15) {
                doc.addPage();
                finalY = 22;
            }

            const colLeft = 25;
            const colRight = 205;

            doc.setFont('helvetica', 'normal');
            doc.setFontSize(8.5);
            doc.setTextColor(31, 41, 55);

            // Kolom Kiri: Kepala Sekolah
            doc.text('Mengetahui,', colLeft, finalY);
            doc.text('Kepala Sekolah', colLeft, finalY + 4.5);

            // Kolom Kanan: Guru Pengampu
            const signDateStr = formatTanggalIndo(tanggalHariIni()).replace(/^[A-Za-z]+,\s*/, '');
            doc.text(`...................., ${signDateStr}`, colRight, finalY);
            doc.text('Guru Kelas / Mata Pelajaran,', colRight, finalY + 4.5);

            // Ruang Tanda Tangan
            finalY += 24;
            doc.setFont('helvetica', 'bold');
            doc.text('( .................................................... )', colLeft, finalY);
            doc.setFont('helvetica', 'normal');
            doc.text('NIP. ............................................', colLeft, finalY + 4.5);

            doc.setFont('helvetica', 'bold');
            doc.text(`( ${teacherName} )`, colRight, finalY);
            doc.setFont('helvetica', 'normal');
            doc.text('NIP. ............................................', colRight, finalY + 4.5);

            // 4. Nomor Halaman & Footer
            const totalPages = doc.internal.getNumberOfPages();
            for (let i = 1; i <= totalPages; i++) {
                doc.setPage(i);
                doc.setFontSize(8);
                doc.setFont('helvetica', 'italic');
                doc.setTextColor(148, 163, 184);
                doc.text('AgendaGuru – Aplikasi Manajemen Kelas Digital', 15, 203);
                doc.text(`Halaman ${i} dari ${totalPages}`, 282, 203, { align: 'right' });
            }

            Swal.close();
            const safeClassName = className.replace(/[^a-zA-Z0-9_-]/g, '_');
            doc.save(`Jurnal_Harian_${safeClassName}.pdf`);
            Swal.fire('Berhasil!', 'Dokumen PDF Jurnal Harian profesional beserta foto dokumentasi berhasil diunduh.', 'success');
        } catch (err) {
            console.error('Error saat men-download PDF jurnal:', err);
            Swal.fire('Error', 'Gagal memproses dokumen PDF: ' + (err.message || err), 'error');
        }
    }

    // --- Profile ---
    let tempPhotoBase64 = "";
    function updateProfileUI() {
        const profName = db.config.userName || 'Guru', profSchool = db.config.userSchool || 'Guru Kelas';
        document.getElementById('profile-name').innerText = profName; document.getElementById('profile-school').innerText = profSchool;
        document.getElementById('profile-img').src = db.config.userPhoto || ('https://ui-avatars.com/api/?name=' + encodeURIComponent(profName) + '&background=172B4D&color=fff');
    }

    function openProfileModal() {
        document.getElementById('prof-name').value = db.config.userName || ''; document.getElementById('prof-school').value = db.config.userSchool || '';
        document.getElementById('modal-prof-img').src = db.config.userPhoto || ('https://ui-avatars.com/api/?name=' + encodeURIComponent(db.config.userName) + '&background=172B4D&color=fff');
        tempPhotoBase64 = db.config.userPhoto || ""; document.getElementById('prof-photo-file').value = "";
        let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalProfile'));
        if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalProfile')); modalInstance.show();
    }

    function closeModal(id) { let modalInstance = bootstrap.Modal.getInstance(document.getElementById(id)); if (modalInstance) modalInstance.hide(); }

    function previewProfilePhoto(event) {
        const file = event.target.files[0];
        if (file) {
            const r = new FileReader(); r.onload = function (e) {
                const img = new Image(); img.onload = function () {
                    const canvas = document.createElement('canvas'); const ctx = canvas.getContext('2d'); canvas.width = 300; canvas.height = 300; ctx.drawImage(img, 0, 0, 300, 300);
                    tempPhotoBase64 = canvas.toDataURL('image/jpeg', 0.7); document.getElementById('modal-prof-img').src = tempPhotoBase64;
                }; img.src = e.target.result;
            }; r.readAsDataURL(file);
        }
    }

    function saveProfile() {
        db.config.userName = document.getElementById('prof-name').value || 'Guru'; db.config.userSchool = document.getElementById('prof-school').value || 'Sekolah';
        if (tempPhotoBase64) db.config.userPhoto = tempPhotoBase64;
        syncDB(); updateProfileUI(); closeModal('modalProfile'); Swal.fire('Tersimpan', 'Profil diperbarui!', 'success');
    }

    // --- Schedule / Jadwal ---
    function renderSchedule() {
        const cFilter = document.getElementById('filter-sched-class').value, container = document.getElementById('schedule-container');
        if (!cFilter) { container.innerHTML = `<div class="col-12 text-center text-muted py-5"><i class="bi bi-calendar-x" style="font-size: 3rem;"></i><p>Pilih kelas terlebih dahulu.</p></div>`; return; }
        let html = '';
        ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'].forEach(day => {
            let daySch = db.schedule.filter(s => s.Day === day && (String(s.ClassID) === String(cFilter) || s.ClassID === 'All')).sort((a, b) => a.TimeStart.localeCompare(b.TimeStart));
            html += `<div class="col-md-4 mb-4"><div class="card shadow-sm border-0 h-100"><div class="card-header bg-primary text-white fw-bold text-center py-2">${day}</div><ul class="list-group list-group-flush">`;
            if (daySch.length === 0) {
                html += `<li class="list-group-item text-center text-muted small py-3">Libur / Kosong</li>`;
            } else {
                daySch.forEach(s => {
                    let className = s.ClassID === 'All' ? 'Semua Kelas' : (db.classes.find(c => String(c.ClassID) === String(s.ClassID))?.ClassName || s.ClassID);
                    html += `
                        <li class="list-group-item px-3 py-2">
                            <div class="d-flex justify-content-between align-items-center mb-1">
                                <span class="badge bg-warning text-dark"><i class="bi bi-clock"></i> ${s.TimeStart} - ${s.TimeEnd}</span>
                                <div>
                                    <button class="btn btn-sm btn-link text-primary p-0 me-2" onclick="editSchedule('${s.RecordID}')" title="Edit Jadwal"><i class="bi bi-pencil-square"></i></button>
                                    <button class="btn btn-sm btn-link text-danger p-0" onclick="deleteSchedule('${s.RecordID}')" title="Hapus Jadwal"><i class="bi bi-trash"></i></button>
                                </div>
                            </div>
                            <div class="fw-semibold text-primary">${escapeHtml(s.Subject)}</div>
                            <small class="text-muted fw-bold" style="font-size:0.75rem">${className}</small>
                        </li>`;
                });
            }
            html += `</ul></div></div>`;
        }); container.innerHTML = html;
    }

    function showAddScheduleModal() {
        editingScheduleId = null;
        document.getElementById('modal-schedule-title').innerText = "Tambah Jadwal Pelajaran";
        document.getElementById('sched-day').value = 'Senin';
        document.getElementById('sched-start').value = '';
        document.getElementById('sched-end').value = '';
        document.getElementById('sched-subject').value = '';
        document.getElementById('sched-class').value = '';

        let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalAddSchedule'));
        if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalAddSchedule'));
        modalInstance.show();
    }

    function editSchedule(id) {
        const s = db.schedule.find(x => x.RecordID === id);
        if (!s) return;

        editingScheduleId = id;
        document.getElementById('modal-schedule-title').innerText = "Edit Jadwal Pelajaran";
        document.getElementById('sched-day').value = s.Day;
        document.getElementById('sched-start').value = s.TimeStart;
        document.getElementById('sched-end').value = s.TimeEnd;
        document.getElementById('sched-subject').value = s.Subject;
        document.getElementById('sched-class').value = s.ClassID;

        let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalAddSchedule'));
        if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalAddSchedule'));
        modalInstance.show();
    }

    function deleteSchedule(id) {
        Swal.fire({
            title: 'Hapus Jadwal?', text: "Apakah Anda yakin ingin menghapus jadwal ini?", icon: 'warning',
            showCancelButton: true, confirmButtonColor: '#C44747', confirmButtonText: 'Ya, Hapus!'
        }).then((result) => {
            if (result.isConfirmed) {
                db.schedule = db.schedule.filter(s => s.RecordID !== id);
                syncDB(); renderSchedule(); Swal.fire('Terhapus!', 'Jadwal berhasil dihapus.', 'success');
            }
        });
    }

    function saveSchedule() {
        const day = document.getElementById('sched-day').value,
            start = document.getElementById('sched-start').value,
            end = document.getElementById('sched-end').value,
            subj = document.getElementById('sched-subject').value,
            cid = document.getElementById('sched-class').value;
        const syncCal = document.getElementById('sched-sync-cal').checked;

        if (!start || !end || !subj || !cid) return Swal.fire('Error', 'Harap lengkapi semua form!', 'error');

        let className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;

        if (editingScheduleId) {
            let idx = db.schedule.findIndex(s => s.RecordID === editingScheduleId);
            if (idx >= 0) {
                db.schedule[idx].Day = day; db.schedule[idx].TimeStart = start; db.schedule[idx].TimeEnd = end; db.schedule[idx].Subject = subj; db.schedule[idx].ClassID = cid;
            }
        } else {
            db.schedule.push({ RecordID: "SCH" + Date.now(), Day: day, TimeStart: start, TimeEnd: end, Subject: subj, ClassID: cid });

            // Tambah ke Calendar otomatis khusus saat Jadwal Baru dibuat (bukan edit)
            if (syncCal && typeof google !== 'undefined') {
                google.script.run.syncToCalendar('schedule', { day: day, start: start, end: end, subject: subj, className: className }, getToken());
            }
        }

        closeModal('modalAddSchedule'); renderSchedule(); syncDB();
        let msg = editingScheduleId ? 'Jadwal diperbarui.' : 'Jadwal ditambahkan.';
        Swal.fire('Tersimpan', msg, 'success');
        editingScheduleId = null;
    }

    // --- Manajemen Kelas (Pemisahan Tab Siswa/Kelas & Mata Pelajaran) ---
    let editingManajemenMapelName = null;

    function switchKelasTab(tab) {
        const siswaView = document.getElementById('kelas-subview-siswa');
        const mapelView = document.getElementById('kelas-subview-mapel');
        const btnSiswa = document.getElementById('btn-tab-kelas-siswa');
        const btnMapel = document.getElementById('btn-tab-kelas-mapel');

        if (!siswaView || !btnSiswa) return;

        if (tab === 'mapel') {
            siswaView.style.display = 'none';
            if (mapelView) mapelView.style.display = 'block';

            btnSiswa.className = 'btn btn-outline-primary';
            if (btnMapel) btnMapel.className = 'btn btn-primary';

            renderManajemenMapelTable();
        } else {
            // Default: 'siswa'
            siswaView.style.display = 'block';
            if (mapelView) mapelView.style.display = 'none';

            btnSiswa.className = 'btn btn-primary';
            if (btnMapel) btnMapel.className = 'btn btn-outline-primary';

            renderStudentsTable();
        }
    }

    function renderManajemenMapelTable() {
        if (!db.subjects) db.subjects = [];
        const tbody = document.getElementById('manajemen-mapel-tbody');
        const badge = document.getElementById('manajemen-mapel-count-badge');
        if (badge) badge.innerText = `${db.subjects.length} Mata Pelajaran`;
        if (!tbody) return;

        if (db.subjects.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted py-4"><i class="bi bi-journal-x fs-3 d-block mb-1"></i>Belum ada mata pelajaran. Silakan tambahkan melalui form di atas.</td></tr>';
            return;
        }

        let html = '';
        db.subjects.forEach((m, idx) => {
            const countNilai = db.grades.filter(g => g.Subject === m).length;
            const countJurnal = db.journal.filter(j => j.Subject === m).length;
            const countJadwal = db.schedule.filter(s => s.Subject === m).length;

            html += `<tr>
                <td class="text-center fw-bold align-middle">${idx + 1}</td>
                <td class="align-middle">
                    <span class="fw-bold" style="color: var(--navy-deep);">${escapeHtml(m)}</span>
                </td>
                <td class="text-center align-middle">
                    <div class="d-inline-flex flex-wrap gap-1 justify-content-center">
                        <span class="badge bg-light text-primary border" title="Terhubung ke Data Nilai (${countNilai} data)"><i class="bi bi-award"></i> Nilai (${countNilai})</span>
                        <span class="badge bg-light text-success border" title="Terhubung ke Jurnal Harian (${countJurnal} data)"><i class="bi bi-journal-text"></i> Jurnal (${countJurnal})</span>
                        <span class="badge bg-light text-warning text-dark border" title="Terhubung ke Jadwal Pelajaran (${countJadwal} data)"><i class="bi bi-clock"></i> Jadwal (${countJadwal})</span>
                    </div>
                </td>
                <td class="text-center align-middle text-nowrap">
                    <button class="btn btn-sm btn-outline-primary me-1" onclick="editManajemenMapel('${encodeURIComponent(m)}')" title="Edit Nama Mapel"><i class="bi bi-pencil-square"></i> Edit</button>
                    <button class="btn btn-sm btn-outline-danger" onclick="deleteMapel('${encodeURIComponent(m)}')" title="Hapus Mapel"><i class="bi bi-trash"></i> Hapus</button>
                </td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    function handleSaveManajemenMapel(e) {
        e.preventDefault();
        const input = document.getElementById('manajemen-mapel-input');
        if (!input) return;
        const mapelName = input.value.trim();
        if (!mapelName) return;

        if (!db.subjects) db.subjects = [];

        if (editingManajemenMapelName) {
            // Mode Edit
            const oldName = editingManajemenMapelName;
            if (mapelName !== oldName && db.subjects.includes(mapelName)) {
                return Swal.fire('Gagal', 'Nama mata pelajaran tersebut sudah ada di daftar!', 'error');
            }

            const idx = db.subjects.indexOf(oldName);
            if (idx > -1) db.subjects[idx] = mapelName;

            // Perbarui data yang terhubung ke mapel ini
            db.grades.forEach(g => { if (g.Subject === oldName) g.Subject = mapelName; });
            db.journal.forEach(j => { if (j.Subject === oldName) j.Subject = mapelName; });
            db.schedule.forEach(s => { if (s.Subject === oldName) s.Subject = mapelName; });

            cancelEditManajemenMapel();
            populateMapelSelects();
            syncDB();
            renderManajemenMapelTable();

            // Refresh tampilan lain yang mungkin aktif
            if (document.getElementById('view-nilai').classList.contains('active')) loadNilaiData();
            if (document.getElementById('view-jurnal').classList.contains('active')) renderJurnal();
            if (document.getElementById('view-jadwal').classList.contains('active')) renderSchedule();
            renderManajemenSchedule();

            Swal.fire('Tersimpan', `Mata Pelajaran berhasil diubah menjadi "${mapelName}".`, 'success');
        } else {
            // Mode Tambah Baru
            if (db.subjects.includes(mapelName)) {
                return Swal.fire('Info', 'Mata Pelajaran tersebut sudah ada di daftar.', 'info');
            }

            db.subjects.push(mapelName);
            input.value = '';
            populateMapelSelects();
            syncDB();
            renderManajemenMapelTable();

            Swal.fire('Berhasil!', `Mata Pelajaran "${mapelName}" berhasil ditambahkan dan langsung terhubung ke Data Nilai, Jurnal Harian, dan Jadwal Pelajaran.`, 'success');
        }
    }

    function editManajemenMapel(encodedName) {
        const name = decodeURIComponent(encodedName);
        editingManajemenMapelName = name;

        const input = document.getElementById('manajemen-mapel-input');
        const title = document.getElementById('manajemen-mapel-form-title');
        const btnSave = document.getElementById('btn-save-manajemen-mapel');
        const btnCancel = document.getElementById('btn-cancel-manajemen-mapel');

        if (input) input.value = name;
        if (title) title.innerHTML = '<i class="bi bi-pencil-square text-warning"></i> Form Edit Mata Pelajaran';
        if (btnSave) btnSave.innerHTML = '<i class="bi bi-save me-1"></i> Perbarui Mapel';
        if (btnCancel) btnCancel.style.display = 'inline-block';

        input?.focus();
        title?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    function cancelEditManajemenMapel() {
        editingManajemenMapelName = null;
        const input = document.getElementById('manajemen-mapel-input');
        const title = document.getElementById('manajemen-mapel-form-title');
        const btnSave = document.getElementById('btn-save-manajemen-mapel');
        const btnCancel = document.getElementById('btn-cancel-manajemen-mapel');

        if (input) input.value = '';
        if (title) title.innerHTML = '<i class="bi bi-plus-circle" style="color: var(--teal-accent);"></i> Form Tambah Mata Pelajaran Baru';
        if (btnSave) btnSave.innerHTML = '<i class="bi bi-plus-lg me-1"></i> Simpan Mapel';
        if (btnCancel) btnCancel.style.display = 'none';
    }

    // --- Poin Siswa (Pemisahan Tab Input Poin & Riwayat Poin) ---
    function switchPoinTab(tab) {
        const inputView = document.getElementById('poin-subview-input');
        const riwayatView = document.getElementById('poin-subview-riwayat');
        const btnInput = document.getElementById('btn-tab-poin-input');
        const btnRiwayat = document.getElementById('btn-tab-poin-riwayat');

        if (!inputView || !riwayatView || !btnInput || !btnRiwayat) return;

        if (tab === 'riwayat') {
            inputView.style.display = 'none';
            riwayatView.style.display = 'block';

            btnInput.className = 'btn btn-outline-primary';
            btnRiwayat.className = 'btn btn-primary';

            // Sinkronkan pilihan kelas riwayat dengan kelas input
            const inputClass = document.getElementById('poin-kelas')?.value;
            const filterClass = document.getElementById('filter-poin-kelas');
            if (filterClass && inputClass && !filterClass.value) {
                filterClass.value = inputClass;
            }
            loadPoinData();
        } else {
            // Default: 'input'
            inputView.style.display = 'block';
            riwayatView.style.display = 'none';

            btnInput.className = 'btn btn-primary';
            btnRiwayat.className = 'btn btn-outline-primary';

            updateQuickPoinSummary();
        }
    }

    function populatePoinStudentSelect() {
        const cid = document.getElementById('poin-kelas').value;
        const select = document.getElementById('poin-siswa');
        if (!cid) {
            select.innerHTML = '<option value="">-- Pilih Kelas Dahulu --</option>';
            updateQuickPoinSummary();
            return;
        }

        const students = db.students.filter(s => String(s.ClassID) === String(cid));
        if (students.length === 0) {
            select.innerHTML = '<option value="">-- Tidak ada siswa di kelas ini --</option>';
            updateQuickPoinSummary();
            return;
        }

        let opts = '<option value="">-- Pilih Siswa --</option>';
        students.forEach(s => opts += `<option value="${escapeHtml(s.StudentID)}">${escapeHtml(s.StudentName)}</option>`);
        select.innerHTML = opts;

        // Sinkronkan juga filter kelas riwayat
        const filterClass = document.getElementById('filter-poin-kelas');
        if (filterClass && !filterClass.value) {
            filterClass.value = cid;
        }

        updateQuickPoinSummary();
    }

    function updateQuickPoinSummary() {
        const cid = document.getElementById('poin-kelas')?.value;
        const tbody = document.getElementById('poin-quick-summary-tbody');
        if (!tbody) return;

        if (!cid) {
            tbody.innerHTML = '<tr><td colspan="4" class="text-center py-3 text-muted">Silakan pilih kelas pada formulir di atas untuk melihat ringkasan poin.</td></tr>';
            return;
        }

        const students = db.students.filter(s => String(s.ClassID) === String(cid));
        if (students.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4" class="text-center py-3 text-muted">Belum ada siswa pada kelas ini.</td></tr>';
            return;
        }

        const records = db.poin.filter(p => String(p.ClassID) === String(cid));
        let rekap = {};
        students.forEach(s => {
            rekap[s.StudentID] = { name: s.StudentName, reward: 0, pelanggaran: 0 };
        });

        records.forEach(p => {
            if (rekap[p.StudentID]) {
                rekap[p.StudentID].reward += Number(p.PoinReward) || 0;
                rekap[p.StudentID].pelanggaran += Number(p.PoinPelanggaran) || 0;
            }
        });

        let html = '';
        Object.keys(rekap).sort((a, b) => rekap[a].name.localeCompare(rekap[b].name)).forEach(sid => {
            const r = rekap[sid];
            const total = r.reward - r.pelanggaran;
            html += `<tr>
                <td class="fw-bold">${escapeHtml(r.name)}</td>
                <td class="text-center text-success fw-bold">+${r.reward}</td>
                <td class="text-center text-danger fw-bold">-${r.pelanggaran}</td>
                <td class="text-center fw-bold ${total >= 0 ? 'text-success' : 'text-danger'}">${total >= 0 ? '+' : ''}${total}</td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    function savePoin() {
        const k = document.getElementById('poin-kelas').value;
        const s = document.getElementById('poin-siswa').value;
        const d = document.getElementById('poin-tanggal').value;
        const r = Number(document.getElementById('poin-reward').value) || 0;
        const p = Number(document.getElementById('poin-pelanggaran').value) || 0;
        const ket = document.getElementById('poin-keterangan').value;

        if (!k || !s || !d) return Swal.fire('Error', 'Harap lengkapi kelas, siswa, dan tanggal!', 'error');
        if (r === 0 && p === 0) return Swal.fire('Peringatan', 'Harap isi poin reward atau poin pelanggaran (minimal salah satu lebih dari 0)!', 'warning');

        const wasEditing = !!editingPoinId;
        if (editingPoinId) {
            const idx = db.poin.findIndex(x => x.RecordID === editingPoinId);
            if (idx > -1) {
                db.poin[idx].ClassID = k;
                db.poin[idx].StudentID = s;
                db.poin[idx].Tanggal = d;
                db.poin[idx].PoinReward = r;
                db.poin[idx].PoinPelanggaran = p;
                db.poin[idx].Keterangan = ket;
            }
            cancelEditPoin();
        } else {
            db.poin.push({
                RecordID: "P" + Date.now(),
                ClassID: k,
                StudentID: s,
                Tanggal: d,
                PoinReward: r,
                PoinPelanggaran: p,
                Keterangan: ket
            });
        }

        syncDB();
        loadPoinData();
        updateQuickPoinSummary();

        document.getElementById('poin-reward').value = 0;
        document.getElementById('poin-pelanggaran').value = 0;
        document.getElementById('poin-keterangan').value = '';

        Swal.fire({
            title: wasEditing ? 'Poin Diperbarui' : 'Poin Disimpan',
            text: wasEditing ? 'Perubahan catatan poin berhasil disimpan.' : 'Catatan poin siswa berhasil disimpan ke sistem.',
            icon: 'success',
            showCancelButton: true,
            confirmButtonText: '<i class="bi bi-clock-history me-1"></i> Lihat Riwayat Poin',
            cancelButtonText: '<i class="bi bi-plus-lg me-1"></i> Input Poin Lagi'
        }).then((result) => {
            if (result.isConfirmed) {
                switchPoinTab('riwayat');
            } else {
                switchPoinTab('input');
            }
        });
    }

    function editPoin(id) {
        const p = db.poin.find(x => x.RecordID === id);
        if (!p) return;
        editingPoinId = id;

        document.getElementById('poin-kelas').value = p.ClassID;
        populatePoinStudentSelect();
        document.getElementById('poin-siswa').value = p.StudentID;
        document.getElementById('poin-tanggal').value = p.Tanggal || '';
        document.getElementById('poin-reward').value = p.PoinReward;
        document.getElementById('poin-pelanggaran').value = p.PoinPelanggaran;
        document.getElementById('poin-keterangan').value = p.Keterangan || '';
        
        document.getElementById('btn-save-poin').innerHTML = '<i class="bi bi-save me-1"></i> Update Poin';
        document.getElementById('btn-cancel-edit-poin').style.display = 'block';

        // Pindah otomatis ke tampilan form input
        switchPoinTab('input');

        const titleElem = document.getElementById('poin-form-title');
        if (titleElem) {
            titleElem.innerHTML = '<i class="bi bi-pencil-square text-warning"></i> Edit Catatan Poin Siswa';
        }

        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function cancelEditPoin() {
        editingPoinId = null;
        document.getElementById('btn-save-poin').innerHTML = '<i class="bi bi-save me-1"></i> Simpan Data Poin';
        document.getElementById('btn-cancel-edit-poin').style.display = 'none';
        document.getElementById('poin-reward').value = 0;
        document.getElementById('poin-pelanggaran').value = 0;
        document.getElementById('poin-keterangan').value = '';

        const titleElem = document.getElementById('poin-form-title');
        if (titleElem) {
            titleElem.innerHTML = '<i class="bi bi-pencil-square" style="color: var(--teal-accent);"></i> Form Input Poin Siswa';
        }
    }

    function deletePoin(id) {
        Swal.fire({
            title: 'Hapus Data Poin?', 
            text: 'Catatan poin siswa ini akan dihapus secara permanen.', 
            icon: 'warning',
            showCancelButton: true, 
            confirmButtonColor: '#C44747', 
            confirmButtonText: 'Ya, Hapus!', 
            cancelButtonText: 'Batal'
        }).then((result) => {
            if (result.isConfirmed) {
                db.poin = db.poin.filter(p => p.RecordID !== id);
                if (editingPoinId === id) cancelEditPoin();
                syncDB(); 
                loadPoinData();
                updateQuickPoinSummary();
                Swal.fire('Terhapus!', 'Data poin berhasil dihapus.', 'success');
            }
        });
    }

    function loadPoinData() {
        const cid = document.getElementById('filter-poin-kelas')?.value;
        const jenisFilter = document.getElementById('filter-poin-jenis')?.value || 'all';
        const searchQuery = (document.getElementById('search-poin-text')?.value || '').toLowerCase().trim();
        const tbody = document.getElementById('poin-tbody'), rekapTbody = document.getElementById('poin-rekap-tbody');

        if (!tbody || !rekapTbody) return;

        if (!cid) {
            tbody.innerHTML = '<tr><td colspan="8" class="text-center py-4 text-muted"><i class="bi bi-arrow-up-circle fs-3 d-block mb-1 text-primary"></i>Silakan pilih kelas terlebih dahulu untuk melihat data poin.</td></tr>';
            rekapTbody.innerHTML = '<tr><td colspan="4" class="text-center py-4 text-muted">Pilih kelas terlebih dahulu.</td></tr>';
            return;
        }

        let records = db.poin.filter(p => cid === 'All' || String(p.ClassID) === String(cid))
            .slice().sort((a, b) => String(b.Tanggal).localeCompare(String(a.Tanggal)));

        // Filter berdasarkan jenis / kategori
        if (jenisFilter === 'reward') {
            records = records.filter(p => (Number(p.PoinReward) || 0) > 0);
        } else if (jenisFilter === 'pelanggaran') {
            records = records.filter(p => (Number(p.PoinPelanggaran) || 0) > 0);
        }

        // Filter pencarian
        if (searchQuery) {
            records = records.filter(p => {
                const sName = (db.students.find(s => s.StudentID === p.StudentID)?.StudentName || '').toLowerCase();
                const ket = (p.Keterangan || '').toLowerCase();
                return sName.includes(searchQuery) || ket.includes(searchQuery);
            });
        }

        if (records.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8" class="text-center py-4 text-muted"><i class="bi bi-journal-x fs-3 d-block mb-1"></i>Belum ada catatan poin siswa yang cocok dengan filter.</td></tr>';
            rekapTbody.innerHTML = '<tr><td colspan="4" class="text-center py-3 text-muted">Belum ada data poin siswa.</td></tr>';
            return;
        }

        let html = '';
        records.forEach(p => {
            const student = db.students.find(s => s.StudentID === p.StudentID);
            const name = student ? student.StudentName : '(Siswa tidak ditemukan)';
            const className = db.classes.find(c => String(c.ClassID) === String(p.ClassID))?.ClassName || p.ClassID;
            const reward = Number(p.PoinReward) || 0, pelanggaran = Number(p.PoinPelanggaran) || 0;
            const total = reward - pelanggaran;
            const tglFormatted = formatTanggalIndo(p.Tanggal);

            html += `<tr>
                <td class="fw-bold">${tglFormatted}</td>
                <td class="fw-semibold text-primary">${escapeHtml(name)}</td>
                <td><span class="badge bg-light text-secondary border">${escapeHtml(className)}</span></td>
                <td class="text-center text-success fw-bold">${reward > 0 ? '+' + reward : '-'}</td>
                <td class="text-center text-danger fw-bold">${pelanggaran > 0 ? '-' + pelanggaran : '-'}</td>
                <td class="text-center fw-bold ${total >= 0 ? 'text-success' : 'text-danger'}">${total >= 0 ? '+' : ''}${total}</td>
                <td><small class="text-secondary">${escapeHtml(p.Keterangan || '-')}</small></td>
                <td class="text-nowrap text-center">
                    <button class="btn btn-sm btn-outline-primary me-1 mb-1" onclick="editPoin('${p.RecordID}')" title="Edit"><i class="bi bi-pencil-square"></i> Edit</button>
                    <button class="btn btn-sm btn-outline-danger mb-1" onclick="deletePoin('${p.RecordID}')" title="Hapus"><i class="bi bi-trash"></i></button>
                </td>
            </tr>`;
        });
        tbody.innerHTML = html;

        // Rekap total akumulasi poin per siswa
        let rekap = {};
        records.forEach(p => {
            if (!rekap[p.StudentID]) rekap[p.StudentID] = { reward: 0, pelanggaran: 0 };
            rekap[p.StudentID].reward += Number(p.PoinReward) || 0;
            rekap[p.StudentID].pelanggaran += Number(p.PoinPelanggaran) || 0;
        });

        let rekapHtml = '';
        Object.keys(rekap).sort((a, b) => {
            const nameA = db.students.find(s => s.StudentID === a)?.StudentName || '';
            const nameB = db.students.find(s => s.StudentID === b)?.StudentName || '';
            return nameA.localeCompare(nameB);
        }).forEach(sid => {
            const student = db.students.find(s => s.StudentID === sid);
            const name = student ? student.StudentName : '(Siswa tidak ditemukan)';
            const r = rekap[sid], total = r.reward - r.pelanggaran;
            rekapHtml += `<tr>
                <td class="fw-bold text-primary">${escapeHtml(name)}</td>
                <td class="text-center text-success fw-bold">+${r.reward}</td>
                <td class="text-center text-danger fw-bold">-${r.pelanggaran}</td>
                <td class="text-center fw-bold ${total >= 0 ? 'text-success' : 'text-danger'}">${total >= 0 ? '+' : ''}${total}</td>
            </tr>`;
        });
        rekapTbody.innerHTML = rekapHtml;
    }

    // --- Unduh Excel Poin Siswa Berformat Sekolah Profesional ---
    async function downloadPoinExcel() {
        const cid = document.getElementById('filter-poin-kelas')?.value || document.getElementById('poin-kelas')?.value;
        if (!cid) return Swal.fire('Error', 'Silakan pilih kelas terlebih dahulu sebelum mengunduh Excel!', 'error');

        const jenisFilter = document.getElementById('filter-poin-jenis')?.value || 'all';
        const searchQuery = (document.getElementById('search-poin-text')?.value || '').toLowerCase().trim();

        let records = db.poin.filter(p => cid === 'All' || String(p.ClassID) === String(cid))
            .slice().sort((a, b) => String(b.Tanggal).localeCompare(String(a.Tanggal)));

        if (jenisFilter === 'reward') records = records.filter(p => (Number(p.PoinReward) || 0) > 0);
        else if (jenisFilter === 'pelanggaran') records = records.filter(p => (Number(p.PoinPelanggaran) || 0) > 0);

        if (searchQuery) {
            records = records.filter(p => {
                const sName = (db.students.find(s => s.StudentID === p.StudentID)?.StudentName || '').toLowerCase();
                const ket = (p.Keterangan || '').toLowerCase();
                return sName.includes(searchQuery) || ket.includes(searchQuery);
            });
        }

        const className = cid === 'All' ? 'Semua Kelas' : (db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid);
        const schoolName = (db.config.userSchool || 'SEKOLAH DASAR / MENENGAH').toUpperCase();
        const teacherName = db.config.userName || RAPOR_KOSONG;

        const curDate = new Date();
        const curYear = curDate.getFullYear();
        const curMonth = curDate.getMonth() + 1;
        const academicYear = curMonth >= 7 ? `${curYear}/${curYear + 1}` : `${curYear - 1}/${curYear}`;
        const semester = curMonth >= 7 ? 'Ganjil' : 'Genap';
        const signDateStr = formatTanggalIndo(tanggalHariIni()).replace(/^[A-Za-z]+,\s*/, '');

        let filterLabel = 'Semua Kategori';
        if (jenisFilter === 'reward') filterLabel = 'Hanya Reward / Prestasi (+)';
        else if (jenisFilter === 'pelanggaran') filterLabel = 'Hanya Pelanggaran (-)';

        Swal.fire({
            title: 'Menyiapkan Excel...',
            text: 'Menyusun tabel Excel poin siswa format sekolah profesional...',
            allowOutsideClick: false,
            didOpen: () => { Swal.showLoading(); }
        });

        try {
            await ensureXLSX();
            const wb = XLSX.utils.book_new();

            // ==========================================
            // SHEET 1: REKAPITULASI POIN SISWA (TABEL UTAMA KELAS)
            // ==========================================
            let allStudentsInClass = db.students.filter(s => cid === 'All' || String(s.ClassID) === String(cid))
                .slice().sort((a, b) => a.StudentName.localeCompare(b.StudentName));

            // Jika tidak ada data siswa khusus, ambil dari record poin
            if (allStudentsInClass.length === 0 && records.length > 0) {
                let sMap = {};
                records.forEach(r => {
                    if (!sMap[r.StudentID]) {
                        sMap[r.StudentID] = true;
                        allStudentsInClass.push({ StudentID: r.StudentID, StudentName: r.StudentID, ClassID: r.ClassID });
                    }
                });
            }

            let rekapMap = {};
            allStudentsInClass.forEach(s => {
                rekapMap[s.StudentID] = {
                    id: s.StudentID,
                    name: s.StudentName,
                    classId: s.ClassID,
                    reward: 0,
                    pelanggaran: 0
                };
            });

            // Akumulasi seluruh catatan poin untuk siswa di kelas ini
            db.poin.filter(p => cid === 'All' || String(p.ClassID) === String(cid)).forEach(p => {
                if (rekapMap[p.StudentID]) {
                    rekapMap[p.StudentID].reward += Number(p.PoinReward) || 0;
                    rekapMap[p.StudentID].pelanggaran += Number(p.PoinPelanggaran) || 0;
                }
            });

            let rekapAOA = [
                [schoolName],
                ["REKAPITULASI AKUMULASI POIN & PRESTASI KARAKTER SISWA"],
                [`TAHUN PELAJARAN ${academicYear} - SEMESTER ${semester.toUpperCase()}`],
                [],
                ["", "Satuan Pendidikan", `: ${schoolName}`, "", "", "Kelas / Rombel", `: ${className}`, ""],
                ["", "Guru Pengampu / BK", `: ${teacherName}`, "", "", "Tanggal Unduh", `: ${signDateStr}`, ""],
                [],
                ["NO", "NISN / ID SISWA", "NAMA LENGKAP SISWA", "KELAS", "TOTAL REWARD (+)", "TOTAL PELANGGARAN (-)", "TOTAL AKUMULASI POIN", "PREDIKAT / STATUS PEMBINAAN"]
            ];

            let sumAllReward = 0;
            let sumAllPelanggaran = 0;
            let sumAllAkumulasi = 0;
            let studentRowIdx = 1;

            allStudentsInClass.forEach(s => {
                const r = rekapMap[s.StudentID] || { reward: 0, pelanggaran: 0 };
                const net = r.reward - r.pelanggaran;
                const sClass = db.classes.find(c => String(c.ClassID) === String(s.ClassID))?.ClassName || className;

                sumAllReward += r.reward;
                sumAllPelanggaran += r.pelanggaran;
                sumAllAkumulasi += net;

                let status = 'Baik / Tertib';
                if (net >= 20) status = 'Sangat Teladan / Berprestasi Tinggi';
                else if (net >= 10) status = 'Teladan / Berprestasi';
                else if (net < 0 && net >= -10) status = 'Perlu Pembinaan Ringan';
                else if (net < -10 && net >= -25) status = 'Perlu Bimbingan Konseling (BK)';
                else if (net < -25) status = 'Panggilan Orang Tua / Khusus';

                rekapAOA.push([
                    studentRowIdx++,
                    s.StudentID || '-',
                    s.StudentName,
                    sClass,
                    r.reward,
                    r.pelanggaran,
                    net,
                    status
                ]);
            });

            const totalRekapRowIndex = rekapAOA.length;
            // Baris Total Akumulasi
            rekapAOA.push([
                "",
                "TOTAL KESELURUHAN KELAS",
                "",
                "",
                sumAllReward,
                sumAllPelanggaran,
                sumAllAkumulasi,
                `Rata-rata: ${allStudentsInClass.length > 0 ? (sumAllAkumulasi / allStudentsInClass.length).toFixed(1) : 0} Poin/Siswa`
            ]);

            // Ruang Tanda Tangan
            rekapAOA.push([]);
            rekapAOA.push(["", "", "Mengetahui,", "", "", "", `...................., ${signDateStr}`, ""]);
            rekapAOA.push(["", "", "Kepala Sekolah,", "", "", "", "Guru Kelas / Guru Bimbingan Konseling,", ""]);
            rekapAOA.push([]);
            rekapAOA.push([]);
            rekapAOA.push(["", "", "( ................................................ )", "", "", "", `( ${teacherName} )`, ""]);
            rekapAOA.push(["", "", "NIP. ........................................", "", "", "", "NIP. ........................................", ""]);

            const wsRekap = XLSX.utils.aoa_to_sheet(rekapAOA);

            // Merges Sheet 1
            wsRekap['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }, // Kop 1
                { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } }, // Kop 2
                { s: { r: 2, c: 0 }, e: { r: 2, c: 7 } }, // Kop 3
                { s: { r: 4, c: 2 }, e: { r: 4, c: 4 } }, // Info Satuan Pendidikan
                { s: { r: 5, c: 2 }, e: { r: 5, c: 4 } }, // Info Guru
                { s: { r: 4, c: 6 }, e: { r: 4, c: 7 } }, // Info Kelas
                { s: { r: 5, c: 6 }, e: { r: 5, c: 7 } }, // Info Tanggal
                { s: { r: totalRekapRowIndex, c: 1 }, e: { r: totalRekapRowIndex, c: 3 } } // Total label B:D
            ];

            // Lebar Kolom Sheet 1
            wsRekap['!cols'] = [
                { wch: 6 },  // NO
                { wch: 18 }, // NISN
                { wch: 32 }, // NAMA SISWA
                { wch: 16 }, // KELAS
                { wch: 18 }, // TOTAL REWARD
                { wch: 22 }, // TOTAL PELANGGARAN
                { wch: 22 }, // TOTAL AKUMULASI
                { wch: 34 }  // STATUS PEMBINAAN
            ];

            // Pasang Autofilter pada baris Header Tabel (Baris 8 -> index 7)
            if (allStudentsInClass.length > 0) {
                wsRekap['!autofilter'] = { ref: `A8:H${7 + allStudentsInClass.length}` };
            }

            XLSX.utils.book_append_sheet(wb, wsRekap, "Rekapitulasi Poin");

            // ==========================================
            // SHEET 2: BUKU CATATAN & RIWAYAT TRANSAKSI POIN
            // ==========================================
            let riwayatAOA = [
                [schoolName],
                ["BUKU CATATAN HARIAN POIN PELANGGARAN & REWARD SISWA"],
                [`TAHUN PELAJARAN ${academicYear} - SEMESTER ${semester.toUpperCase()}`],
                [],
                ["", "Satuan Pendidikan", `: ${schoolName}`, "", "", "Kelas / Rombel", `: ${className}`, ""],
                ["", "Guru Pengampu / BK", `: ${teacherName}`, "", "", "Kategori Filter", `: ${filterLabel}`, ""],
                [],
                ["NO", "HARI / TANGGAL", "NAMA SISWA", "KELAS", "REWARD (+)", "PELANGGARAN (-)", "TOTAL POIN", "KETERANGAN / TINDAK LANJUT"]
            ];

            let sumLogReward = 0;
            let sumLogPelanggaran = 0;
            let sumLogNet = 0;

            records.forEach((p, idx) => {
                const student = db.students.find(s => s.StudentID === p.StudentID);
                const name = student ? student.StudentName : '(Siswa tidak ditemukan)';
                const pClass = db.classes.find(c => String(c.ClassID) === String(p.ClassID))?.ClassName || p.ClassID;
                const reward = Number(p.PoinReward) || 0;
                const pelanggaran = Number(p.PoinPelanggaran) || 0;
                const total = reward - pelanggaran;

                sumLogReward += reward;
                sumLogPelanggaran += pelanggaran;
                sumLogNet += total;

                riwayatAOA.push([
                    idx + 1,
                    formatTanggalIndo(p.Tanggal),
                    name,
                    pClass,
                    reward,
                    pelanggaran,
                    total,
                    p.Keterangan || '-'
                ]);
            });

            const totalRiwayatRowIndex = riwayatAOA.length;
            // Baris Total Log
            riwayatAOA.push([
                "",
                "TOTAL SELURUH CATATAN",
                "",
                "",
                sumLogReward,
                sumLogPelanggaran,
                sumLogNet,
                `${records.length} Catatan Kejadian Poin Terdaftar`
            ]);

            // Ruang Tanda Tangan Sheet 2
            riwayatAOA.push([]);
            riwayatAOA.push(["", "", "Mengetahui,", "", "", "", `...................., ${signDateStr}`, ""]);
            riwayatAOA.push(["", "", "Kepala Sekolah,", "", "", "", "Guru Kelas / Guru Bimbingan Konseling,", ""]);
            riwayatAOA.push([]);
            riwayatAOA.push([]);
            riwayatAOA.push(["", "", "( ................................................ )", "", "", "", `( ${teacherName} )`, ""]);
            riwayatAOA.push(["", "", "NIP. ........................................", "", "", "", "NIP. ........................................", ""]);

            const wsRiwayat = XLSX.utils.aoa_to_sheet(riwayatAOA);

            // Merges Sheet 2
            wsRiwayat['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }, // Kop 1
                { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } }, // Kop 2
                { s: { r: 2, c: 0 }, e: { r: 2, c: 7 } }, // Kop 3
                { s: { r: 4, c: 2 }, e: { r: 4, c: 4 } }, // Satuan Pendidikan
                { s: { r: 5, c: 2 }, e: { r: 5, c: 4 } }, // Guru
                { s: { r: 4, c: 6 }, e: { r: 4, c: 7 } }, // Kelas
                { s: { r: 5, c: 6 }, e: { r: 5, c: 7 } }, // Filter
                { s: { r: totalRiwayatRowIndex, c: 1 }, e: { r: totalRiwayatRowIndex, c: 3 } } // Total label B:D
            ];

            // Lebar Kolom Sheet 2
            wsRiwayat['!cols'] = [
                { wch: 6 },  // NO
                { wch: 24 }, // TANGGAL
                { wch: 30 }, // NAMA SISWA
                { wch: 15 }, // KELAS
                { wch: 14 }, // REWARD
                { wch: 16 }, // PELANGGARAN
                { wch: 14 }, // TOTAL
                { wch: 45 }  // KETERANGAN
            ];

            // Pasang Autofilter pada baris Header Tabel (Baris 8 -> index 7)
            if (records.length > 0) {
                wsRiwayat['!autofilter'] = { ref: `A8:H${7 + records.length}` };
            }

            XLSX.utils.book_append_sheet(wb, wsRiwayat, "Log Riwayat Poin");

            Swal.close();
            const safeClassName = className.replace(/[^a-zA-Z0-9_-]/g, '_');
            XLSX.writeFile(wb, `Rekapitulasi_Poin_Siswa_${safeClassName}.xlsx`);
            Swal.fire('Berhasil!', 'Dokumen Excel Poin Siswa tabel format sekolah profesional berhasil diunduh.', 'success');
        } catch (err) {
            console.error('Error saat download Excel poin:', err);
            Swal.fire('Error', 'Gagal memproses dokumen Excel: ' + (err.message || err), 'error');
        }
    }
    // --- Unduh PDF Poin Siswa Berformat Sekolah Profesional ---
    async function downloadPoinPDF() {
        const cid = document.getElementById('filter-poin-kelas')?.value || document.getElementById('poin-kelas')?.value;
        if (!cid) return Swal.fire('Error', 'Silakan pilih kelas terlebih dahulu sebelum mengunduh PDF!', 'error');

        const jenisFilter = document.getElementById('filter-poin-jenis')?.value || 'all';
        const searchQuery = (document.getElementById('search-poin-text')?.value || '').toLowerCase().trim();

        let records = db.poin.filter(p => cid === 'All' || String(p.ClassID) === String(cid))
            .slice().sort((a, b) => String(b.Tanggal).localeCompare(String(a.Tanggal)));

        if (jenisFilter === 'reward') records = records.filter(p => (Number(p.PoinReward) || 0) > 0);
        else if (jenisFilter === 'pelanggaran') records = records.filter(p => (Number(p.PoinPelanggaran) || 0) > 0);

        if (searchQuery) {
            records = records.filter(p => {
                const sName = (db.students.find(s => s.StudentID === p.StudentID)?.StudentName || '').toLowerCase();
                const ket = (p.Keterangan || '').toLowerCase();
                return sName.includes(searchQuery) || ket.includes(searchQuery);
            });
        }

        if (records.length === 0) return Swal.fire('Informasi', 'Belum ada data poin siswa yang cocok untuk diunduh.', 'info');

        const className = cid === 'All' ? 'Semua Kelas' : (db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid);
        const schoolName = (db.config.userSchool || 'SEKOLAH DASAR / MENENGAH').toUpperCase();
        const teacherName = db.config.userName || RAPOR_KOSONG;

        const curDate = new Date();
        const curYear = curDate.getFullYear();
        const curMonth = curDate.getMonth() + 1;
        const academicYear = curMonth >= 7 ? `${curYear}/${curYear + 1}` : `${curYear - 1}/${curYear}`;
        const semester = curMonth >= 7 ? 'Ganjil' : 'Genap';

        Swal.fire({
            title: 'Membuat Dokumen PDF...',
            text: 'Menyusun dokumen PDF buku catatan poin siswa profesional...',
            allowOutsideClick: false,
            didOpen: () => { Swal.showLoading(); }
        });

        try {
            await ensurePDF();
            const { jsPDF } = window.jspdf;

            // Format A4 Landscape
            const doc = new jsPDF({
                orientation: 'landscape',
                unit: 'mm',
                format: 'a4'
            });

            // 1. Kop Dokumen Resmi Sekolah
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(14);
            doc.setTextColor(23, 43, 77); // Deep Navy #172B4D
            doc.text(schoolName, 148.5, 15, { align: 'center' });

            doc.setFontSize(12);
            doc.setTextColor(36, 59, 90); // Navy Secondary #243B5A
            doc.text('BUKU CATATAN DAN REKAPITULASI POIN SISWA', 148.5, 21, { align: 'center' });

            // Garis Pembatas Kop Ganda
            doc.setDrawColor(23, 43, 77);
            doc.setLineWidth(0.6);
            doc.line(15, 24.5, 282, 24.5);
            doc.setDrawColor(148, 163, 184);
            doc.setLineWidth(0.2);
            doc.line(15, 25.5, 282, 25.5);

            // 2. Informasi Metadata Dokumen
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            doc.setTextColor(31, 41, 55);

            // Kolom Kiri
            doc.text(`Satuan Pendidikan : ${schoolName}`, 15, 30);
            doc.text(`Kelas / Rombel       : ${className}`, 15, 35);

            // Kolom Kanan
            doc.text(`Tahun Pelajaran   : ${academicYear} (Semester ${semester})`, 185, 30);
            doc.text(`Guru Pengampu / BK: ${teacherName}`, 185, 35);

            // 3. Susun Baris Data Tabel
            const tableHead = [['NO', 'HARI / TANGGAL', 'NAMA SISWA', 'KELAS', 'REWARD (+)', 'PELANGGARAN (-)', 'TOTAL', 'KETERANGAN / CATATAN PEMBINAAN']];
            const tableBody = records.map((p, idx) => {
                const student = db.students.find(s => s.StudentID === p.StudentID);
                const name = student ? student.StudentName : '(Siswa tidak ditemukan)';
                const pClass = db.classes.find(c => String(c.ClassID) === String(p.ClassID))?.ClassName || p.ClassID;
                const reward = Number(p.PoinReward) || 0;
                const pelanggaran = Number(p.PoinPelanggaran) || 0;
                const total = reward - pelanggaran;

                return [
                    (idx + 1).toString(),
                    formatTanggalIndo(p.Tanggal),
                    name,
                    pClass,
                    reward > 0 ? '+' + reward : '-',
                    pelanggaran > 0 ? '-' + pelanggaran : '-',
                    (total >= 0 ? '+' : '') + total,
                    p.Keterangan || '-'
                ];
            });

            doc.autoTable({
                head: tableHead,
                body: tableBody,
                startY: 40,
                margin: { left: 15, right: 15, bottom: 18 },
                theme: 'grid',
                headStyles: {
                    fillColor: [23, 43, 77], // #172B4D Deep Navy
                    textColor: [255, 255, 255],
                    fontStyle: 'bold',
                    halign: 'center',
                    valign: 'middle',
                    fontSize: 8.5,
                    cellPadding: 3
                },
                bodyStyles: {
                    fontSize: 8,
                    textColor: [31, 41, 55],
                    valign: 'middle',
                    lineColor: [226, 232, 240],
                    lineWidth: 0.2,
                    cellPadding: 2.5
                },
                alternateRowStyles: {
                    fillColor: [248, 250, 252]
                },
                columnStyles: {
                    0: { cellWidth: 12, halign: 'center' }, // No
                    1: { cellWidth: 32, halign: 'center' }, // Tanggal
                    2: { cellWidth: 48, halign: 'left' },   // Nama Siswa
                    3: { cellWidth: 24, halign: 'center' }, // Kelas
                    4: { cellWidth: 24, halign: 'center', textColor: [33, 134, 91], fontStyle: 'bold' }, // Reward
                    5: { cellWidth: 26, halign: 'center', textColor: [196, 71, 71], fontStyle: 'bold' }, // Pelanggaran
                    6: { cellWidth: 22, halign: 'center', fontStyle: 'bold' }, // Total
                    7: { cellWidth: 79, halign: 'left' }    // Keterangan
                }
            });

            // 4. Kolom Tanda Tangan / Pengesahan Dokumen
            let finalY = doc.lastAutoTable.finalY + 10;
            if (finalY + 38 > 210 - 15) {
                doc.addPage();
                finalY = 22;
            }

            const colLeft = 25;
            const colRight = 205;

            doc.setFont('helvetica', 'normal');
            doc.setFontSize(8.5);
            doc.setTextColor(31, 41, 55);

            // Kolom Kiri: Kepala Sekolah
            doc.text('Mengetahui,', colLeft, finalY);
            doc.text('Kepala Sekolah', colLeft, finalY + 4.5);

            // Kolom Kanan: Guru Pengampu / BK
            const signDateStr = formatTanggalIndo(tanggalHariIni()).replace(/^[A-Za-z]+,\s*/, '');
            doc.text(`...................., ${signDateStr}`, colRight, finalY);
            doc.text('Guru Kelas / Guru Bimbingan Konseling,', colRight, finalY + 4.5);

            // Ruang Tanda Tangan
            finalY += 24;
            doc.setFont('helvetica', 'bold');
            doc.text('( .................................................... )', colLeft, finalY);
            doc.setFont('helvetica', 'normal');
            doc.text('NIP. ............................................', colLeft, finalY + 4.5);

            doc.setFont('helvetica', 'bold');
            doc.text(`( ${teacherName} )`, colRight, finalY);
            doc.setFont('helvetica', 'normal');
            doc.text('NIP. ............................................', colRight, finalY + 4.5);

            // 5. Nomor Halaman & Footer
            const totalPages = doc.internal.getNumberOfPages();
            for (let i = 1; i <= totalPages; i++) {
                doc.setPage(i);
                doc.setFontSize(8);
                doc.setFont('helvetica', 'italic');
                doc.setTextColor(148, 163, 184);
                doc.text('AgendaGuru – Aplikasi Manajemen Kelas Digital', 15, 203);
                doc.text(`Halaman ${i} dari ${totalPages}`, 282, 203, { align: 'right' });
            }

            Swal.close();
            const safeClassName = className.replace(/[^a-zA-Z0-9_-]/g, '_');
            doc.save(`Catatan_Poin_Siswa_${safeClassName}.pdf`);
            Swal.fire('Berhasil!', 'Dokumen PDF Catatan Poin Siswa profesional berhasil diunduh.', 'success');
        } catch (err) {
            console.error('Error saat download PDF poin:', err);
            Swal.fire('Error', 'Gagal memproses dokumen PDF: ' + (err.message || err), 'error');
        }
    }

    // =====================================
    // INTEGRASI WHATSAPP HELPER
    // =====================================
    function sendWAAbsen(sid) {
        let student = db.students.find(s => s.StudentID === sid);
        if (!student || !student.ParentPhone) return Swal.fire('Gagal', 'Nomor HP Orang Tua belum diatur untuk siswa ini.', 'error');

        // Perbaikan Bug: Membersihkan karakter selain angka (seperti spasi atau strip)
        let phone = student.ParentPhone.replace(/\D/g, '');
        if (phone.startsWith('0')) phone = phone.replace(/^0/, '62');

        let selectElem = document.getElementById('absen-bulan');
        let month = selectElem.options[selectElem.selectedIndex].text;

        // Hitung rekap
        let tH = 0, tS = 0, tI = 0, tA = 0;
        document.querySelectorAll(`.select-absen-val[data-sid="${sid}"]`).forEach(sel => {
            let v = sel.value;
            if (v === 'H') tH++; else if (v === 'S') tS++; else if (v === 'I') tI++; else if (v === 'A') tA++;
        });

        let guruName = db.config.userName || 'Guru Kelas';
        let text = `Halo Bapak/Ibu Wali Murid dari *${student.StudentName}*,\n\nBerikut adalah rekap sementara kehadiran ananda pada bulan *${month}*:\n\n✔️ Hadir (H): ${tH}\n🤒 Sakit (S): ${tS}\n📩 Izin (I): ${tI}\n❌ Alpha (A): ${tA}\n\nMohon bantuannya untuk memantau kehadiran ananda. Terima kasih,\n*${guruName}*`;

        // Perbaikan Bug: Menggunakan metode pembuatan elemen anchor (<a>) dinamis untuk menghindari pemblokiran Pop-up dari Iframe
        let waUrl = 'https://api.whatsapp.com/send?phone=' + phone + '&text=' + encodeURIComponent(text);
        let a = document.createElement('a');
        a.target = '_blank';
        a.href = waUrl;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    function sendWANilai(sid) {
        let student = db.students.find(s => s.StudentID === sid);
        if (!student || !student.ParentPhone) return Swal.fire('Gagal', 'Nomor HP Orang Tua belum diatur untuk siswa ini.', 'error');

        // Perbaikan Bug
        let phone = student.ParentPhone.replace(/\D/g, '');
        if (phone.startsWith('0')) phone = phone.replace(/^0/, '62');

        let mapel = document.getElementById('nilai-mapel').value;
        let selJenis = document.getElementById('nilai-jenis');
        let jenis = selJenis.options[selJenis.selectedIndex].text;

        let vals = [];
        document.querySelectorAll(`.input-nilai-val[data-sid="${sid}"]`).forEach(inp => vals.push(inp.value || '-'));

        let guruName = db.config.userName || 'Guru Kelas';
        let text = `Halo Bapak/Ibu Wali Murid dari *${student.StudentName}*,\n\nBerikut adalah hasil penilaian *${jenis}* untuk mata pelajaran *${mapel}*:\n\n*Nilai: ${vals.join(', ')}*\n\nTerima kasih atas bimbingan Bapak/Ibu di rumah,\n*${guruName}*`;

        // Perbaikan Bug
        let waUrl = 'https://api.whatsapp.com/send?phone=' + phone + '&text=' + encodeURIComponent(text);
        let a = document.createElement('a');
        a.target = '_blank';
        a.href = waUrl;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    // =====================================
    // FITUR BROADCAST WHATSAPP (UMUM - MANAJEMEN KELAS SAJA)
    // =====================================
    let broadcastQueue = [];
    let currentBroadcastIndex = 0;

    function toggleAllStudents(source) {
        const checkboxes = document.querySelectorAll('.chk-std:not(:disabled)');
        checkboxes.forEach(chk => chk.checked = source.checked);
    }

    function openBroadcastModal() {
        let selectedChks = document.querySelectorAll('.chk-std:checked');

        if (!selectedChks || selectedChks.length === 0) {
            return Swal.fire('Info', 'Pilih minimal 1 siswa dengan mencentang kotak di tabel. (Pastikan siswa tersebut memiliki No HP)', 'info');
        }

        broadcastQueue = [];
        let htmlList = '';

        selectedChks.forEach(chk => {
            let student = db.students.find(s => s.StudentID === chk.value);
            if (student && student.ParentPhone) {
                broadcastQueue.push(student);
                htmlList += `<li class="list-group-item px-3 py-2 small d-flex align-items-center" id="br-item-${student.StudentID}"><i class="bi bi-clock-history text-secondary me-2 fs-5"></i> <div><span class="fw-bold">${escapeHtml(student.StudentName)}</span><br><small class="text-muted">${escapeHtml(student.ParentPhone)}</small></div></li>`;
            }
        });

        if (broadcastQueue.length === 0) return Swal.fire('Info', 'Siswa yang dipilih tidak memiliki No HP.', 'info');

        document.getElementById('broadcast-count').innerText = broadcastQueue.length;
        document.getElementById('broadcast-list').innerHTML = htmlList;

        document.getElementById('broadcast-msg-wrapper').style.display = 'block';
        document.getElementById('broadcast-msg').value = '';

        currentBroadcastIndex = 0;
        const btnSend = document.getElementById('btn-send-broadcast');
        btnSend.innerText = `Kirim Pesan (1 dari ${broadcastQueue.length})`;
        btnSend.classList.replace('btn-secondary', 'btn-success');
        btnSend.disabled = false;

        let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalBroadcast'));
        if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalBroadcast'));
        modalInstance.show();
    }

    function sendNextBroadcast() {
        const msgTemplate = document.getElementById('broadcast-msg').value.trim();
        if (!msgTemplate) return Swal.fire('Peringatan', 'Pesan pengumuman tidak boleh kosong!', 'warning');

        if (currentBroadcastIndex >= broadcastQueue.length) {
            Swal.fire('Selesai', 'Semua pesan broadcast telah diproses!', 'success');
            return;
        }

        let student = broadcastQueue[currentBroadcastIndex];

        // Ganti placeholder [Nama] dengan nama asli siswa
        let text = msgTemplate.replace(/\[Nama\]/gi, student.StudentName);

        let phone = student.ParentPhone.replace(/\D/g, '');
        if (phone.startsWith('0')) phone = phone.replace(/^0/, '62');

        // Coret siswa dari antrean karena link WA sudah dibuka
        let liElem = document.getElementById(`br-item-${student.StudentID}`);
        if (liElem) {
            liElem.innerHTML = `<i class="bi bi-check-circle-fill text-success me-2 fs-5"></i> <div><s class="fw-bold">${escapeHtml(student.StudentName)}</s><br><small class="text-muted">${escapeHtml(student.ParentPhone)}</small></div>`;
            liElem.classList.add('bg-light');
        }

        let waUrl = 'https://api.whatsapp.com/send?phone=' + phone + '&text=' + encodeURIComponent(text);
        let a = document.createElement('a');
        a.target = '_blank';
        a.href = waUrl;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);

        currentBroadcastIndex++;

        const btnSend = document.getElementById('btn-send-broadcast');
        if (currentBroadcastIndex < broadcastQueue.length) {
            btnSend.innerHTML = `Kirim Selanjutnya (${currentBroadcastIndex + 1} dari ${broadcastQueue.length})`;
        } else {
            btnSend.innerHTML = `<i class="bi bi-check2-all"></i> Selesai`;
            btnSend.classList.replace('btn-success', 'btn-secondary');
        }
    }

    // =====================================================================
    // PENGOLAH NILAI
    // Mengolah nilai harian (formatif), PTS (SumatifLM), PAS (SumatifAS), dan
    // kehadiran menjadi nilai akhir rapor dengan bobot, KKM, dan penyesuaian.
    // Hasil disimpan di db.nilaiAkhir (sheet "X_nilaiAkhir") dan otomatis
    // dipakai sebagai nilai akhir di Cetak Rapor & Rekap Laporan.
    // =====================================================================
    const PN_DEFAULT = { bobot: { harian: 40, pts: 30, pas: 30, hadir: 0 }, kkm: 75, bulat: true, metode: 'none', rentangMin: 75, rentangMax: 80 };
    const PN_SETTINGS_KEY = 'agendaGuru_pnSettings';
    let pnRows = [];          // hasil hitung terakhir per siswa
    let pnSemesterTouched = false;

    function pnNum(v) { if (v === '' || v === null || v === undefined) return null; const n = parseFloat(v); return isNaN(n) ? null : n; }
    function pnFmt(v) { return v === null || v === undefined ? '-' : (Math.round(v * 10) / 10).toString().replace('.', ','); }
    function pnSameTapel(a, b) { return String(a || '').replace(/\s+/g, '') === String(b || '').replace(/\s+/g, ''); }

    // Komponen nilai mentah satu siswa (cara mengambil datanya sama dengan perhitungan rapor)
    function getNilaiKomponen(cid, sid, subj, semester, tapel) {
        const months = semester === 'ganjil' ? ['07', '08', '09', '10', '11', '12'] : ['01', '02', '03', '04', '05', '06'];
        let sum = 0, count = 0;
        months.forEach(m => {
            const rec = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(subj) && String(g.Type) === 'Formatif' && String(g.MonthYear) === String(m) && matchTapel(g, tapel));
            if (!rec || !rec.GradesData) return;
            try { (JSON.parse(rec.GradesData)[sid] || []).forEach(v => { const n = pnNum(v); if (n !== null) { sum += n; count++; } }); } catch (e) { }
        });
        const single = type => {
            const rec = db.grades.slice().reverse().find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(subj) && String(g.Type) === type && isGradeMatchingSemester(g, semester) && matchTapel(g, tapel));
            if (!rec) return null;
            try { return pnNum((JSON.parse(rec.GradesData)[sid] || [])[0]); } catch (e) { return null; }
        };
        const att = getStudentAttendanceSummary(cid, sid, semester, tapel);
        const totalAtt = att.H + att.S + att.I + att.A;
        return {
            harian: count > 0 ? sum / count : null,
            pts: single('SumatifLM'),
            pas: single('SumatifAS'),
            hadir: totalAtt > 0 ? (att.H / totalAtt) * 100 : null
        };
    }

    function findNilaiOlahan(cid, subj, semester, tapel) {
        return (db.nilaiAkhir || []).find(r => String(r.ClassID) === String(cid) && String(r.Subject) === String(subj) &&
            String(r.Semester || '').toLowerCase() === String(semester || '').toLowerCase() && pnSameTapel(r.TahunPelajaran, tapel));
    }
    function parseNilaiData(rec) {
        try { return typeof rec.NilaiData === 'string' ? JSON.parse(rec.NilaiData || '{}') : (rec.NilaiData || {}); } catch (e) { return {}; }
    }
    // Nilai akhir hasil Pengolah Nilai untuk satu siswa (null jika belum diolah)
    function getNilaiOlahanSiswa(cid, subj, semester, tapel, sid) {
        if (!db.nilaiAkhir || db.nilaiAkhir.length === 0) return null;
        const rec = findNilaiOlahan(cid, subj, semester, tapel || getActiveTapel());
        if (!rec) return null;
        const d = parseNilaiData(rec)[sid];
        return d ? pnNum(d.akhir) : null;
    }

    function pnReadSettings() {
        const n = (id, def) => { const v = pnNum(document.getElementById(id)?.value); return v === null ? def : v; };
        return {
            bobot: { harian: n('pn-bobot-harian', 0), pts: n('pn-bobot-pts', 0), pas: n('pn-bobot-pas', 0), hadir: n('pn-bobot-hadir', 0) },
            kkm: n('pn-kkm', 75),
            bulat: !!document.getElementById('pn-bulat')?.checked,
            metode: document.getElementById('pn-metode')?.value || 'none',
            rentangMin: n('pn-rentang-min', 75),
            rentangMax: n('pn-rentang-max', 80)
        };
    }
    function pnApplySettings(s) {
        s = Object.assign({}, PN_DEFAULT, s || {});
        s.bobot = Object.assign({}, PN_DEFAULT.bobot, s.bobot || {});
        const set = (id, v) => { const el = document.getElementById(id); if (el) el.value = v; };
        set('pn-bobot-harian', s.bobot.harian); set('pn-bobot-pts', s.bobot.pts); set('pn-bobot-pas', s.bobot.pas); set('pn-bobot-hadir', s.bobot.hadir);
        set('pn-kkm', s.kkm); set('pn-metode', s.metode); set('pn-rentang-min', s.rentangMin); set('pn-rentang-max', s.rentangMax);
        const b = document.getElementById('pn-bulat'); if (b) b.checked = !!s.bulat;
        pnUpdateSettingUI();
    }
    function pnLastSettings() {
        try { return JSON.parse(localStorage.getItem(PN_SETTINGS_KEY) || 'null') || PN_DEFAULT; } catch (e) { return PN_DEFAULT; }
    }
    function pnUpdateSettingUI() {
        const s = pnReadSettings();
        const total = s.bobot.harian + s.bobot.pts + s.bobot.pas + s.bobot.hadir;
        const totEl = document.getElementById('pn-bobot-total');
        if (totEl) totEl.innerHTML = total === 100
            ? '<span class="text-success fw-semibold"><i class="bi bi-check-circle me-1"></i>Total bobot 100%</span>'
            : `<span class="text-danger fw-semibold"><i class="bi bi-exclamation-circle me-1"></i>Total bobot ${total}% (harus 100%)</span>`;
        const showRange = s.metode !== 'none';
        ['pn-rentang-wrap-min', 'pn-rentang-wrap-max'].forEach(id => { const el = document.getElementById(id); if (el) el.style.display = showRange ? '' : 'none'; });
        const help = document.getElementById('pn-metode-help');
        if (help) help.textContent = s.metode === 'kkm'
            ? `Nilai di bawah KKM (${s.kkm}) dipetakan berurutan ke rentang ${s.rentangMin}–${s.rentangMax}. Nilai yang sudah mencapai KKM tidak diubah.`
            : s.metode === 'rentang'
                ? `Nilai terendah di kelas menjadi ${s.rentangMin} dan tertinggi menjadi ${s.rentangMax}; nilai lain menyesuaikan secara proporsional.`
                : 'Nilai akhir = nilai mentah hasil pembobotan.';
    }

    function pnSelection() {
        return {
            cid: document.getElementById('pn-kelas')?.value || '',
            subj: document.getElementById('pn-mapel')?.value || '',
            sem: document.getElementById('pn-semester')?.value || 'ganjil',
            tapel: getActiveTapel()
        };
    }

    function loadPengolahView() {
        const elK = document.getElementById('pn-kelas'), elM = document.getElementById('pn-mapel');
        if (!elK || !elM) return;
        const curK = elK.value, curM = elM.value;
        elK.innerHTML = '<option value="">-- Pilih Kelas --</option>' + db.classes.map(c => `<option value="${escapeHtml(c.ClassID)}">${escapeHtml(c.ClassName)}</option>`).join('');
        if (curK && db.classes.some(c => String(c.ClassID) === String(curK))) elK.value = curK;
        else if (db.classes.length === 1) elK.value = db.classes[0].ClassID;
        const subjects = [...(db.subjects || [])];
        db.grades.forEach(g => { if (g.Subject && !subjects.includes(g.Subject)) subjects.push(g.Subject); });
        elM.innerHTML = '<option value="">-- Pilih Mapel --</option>' + subjects.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
        if (curM && subjects.includes(curM)) elM.value = curM;
        if (!pnSemesterTouched) {
            const t = getActiveTapel(), now = new Date();
            document.getElementById('pn-semester').value = (t === getDefaultTapel() && now.getMonth() + 1 < 7) ? 'genap' : 'ganjil';
        }
        const lbl = document.getElementById('pn-tapel-label'); if (lbl) lbl.textContent = 'TP ' + getActiveTapel();
        loadPengolahNilai();
    }

    // Dipanggil saat kelas/mapel/semester berubah: pakai pengaturan tersimpan jika ada
    function loadPengolahNilai() {
        const { cid, subj, sem, tapel } = pnSelection();
        const rec = cid && subj ? findNilaiOlahan(cid, subj, sem, tapel) : null;
        let saved = null;
        if (rec) { try { saved = JSON.parse(rec.Pengaturan || 'null'); } catch (e) { } }
        pnApplySettings(saved || pnLastSettings());
        hitungPengolahNilai(true);
        if (rec) {
            // Tampilkan nilai akhir yang tersimpan apa adanya
            const data = parseNilaiData(rec);
            let changed = 0;
            pnRows.forEach(r => {
                const d = data[r.sid];
                if (!d) return;
                const savedAkhir = pnNum(d.akhir);
                if (savedAkhir !== r.akhir) { r.akhir = savedAkhir; r.manual = !!d.manual; if (!d.manual) changed++; }
                else if (d.manual) r.manual = true;
            });
            renderPengolahTable();
            pnSetStatus(rec, changed);
        } else {
            pnSetStatus(null, 0);
        }
    }

    function pnSetStatus(rec, staleCount) {
        const el = document.getElementById('pn-status-badge');
        const btnDel = document.getElementById('pn-btn-hapus');
        if (btnDel) btnDel.disabled = !rec;
        if (!el) return;
        const { cid, subj } = pnSelection();
        if (!cid || !subj) { el.className = 'badge rounded-pill bg-secondary-subtle text-secondary-emphasis px-3 py-2'; el.textContent = 'Pilih kelas & mapel'; return; }
        if (!rec) { el.className = 'badge rounded-pill bg-warning-subtle text-warning-emphasis px-3 py-2'; el.textContent = 'Belum disimpan – rapor memakai rata-rata biasa'; return; }
        const t = rec.UpdatedAt ? new Date(rec.UpdatedAt) : null;
        const when = t && !isNaN(t) ? t.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' }) : '';
        el.className = 'badge rounded-pill bg-success-subtle text-success-emphasis px-3 py-2';
        el.textContent = 'Tersimpan di rapor' + (when ? ' · ' + when : '');
        if (staleCount > 0) {
            el.className = 'badge rounded-pill bg-info-subtle text-info-emphasis px-3 py-2';
            el.textContent = `Tersimpan · ${staleCount} nilai berubah sejak disimpan, klik Hitung Ulang lalu Simpan`;
        }
    }

    // Hitung nilai mentah & nilai akhir. resetManual=false mempertahankan ubahan manual guru.
    function hitungPengolahNilai(resetManual) {
        const { cid, subj, sem, tapel } = pnSelection();
        const tbody = document.getElementById('pn-tbody');
        if (!cid || !subj) {
            pnRows = [];
            if (tbody) tbody.innerHTML = '<tr><td colspan="9" class="text-center text-muted py-4">Pilih kelas dan mata pelajaran terlebih dahulu.</td></tr>';
            const st = document.getElementById('pn-stats'); if (st) st.innerHTML = '';
            pnSetStatus(null, 0);
            return;
        }
        const s = pnReadSettings();
        const keepManual = {};
        if (!resetManual) pnRows.forEach(r => { if (r.manual) keepManual[r.sid] = r.akhir; });

        const students = db.students.filter(st => String(st.ClassID) === String(cid));
        pnRows = students.map(st => {
            const k = getNilaiKomponen(cid, st.StudentID, subj, sem, tapel);
            const parts = ['harian', 'pts', 'pas', 'hadir'].filter(key => k[key] !== null && s.bobot[key] > 0);
            const wsum = parts.reduce((a, key) => a + s.bobot[key], 0);
            const mentah = wsum > 0 ? parts.reduce((a, key) => a + k[key] * s.bobot[key], 0) / wsum : null;
            return { sid: st.StudentID, name: st.StudentName, nis: st.NIS || '', harian: k.harian, pts: k.pts, pas: k.pas, hadir: k.hadir, mentah, akhir: null, manual: false };
        });

        const valid = pnRows.filter(r => r.mentah !== null);
        const lo = arr => Math.min(...arr), hi = arr => Math.max(...arr);
        if (s.metode === 'kkm') {
            const below = valid.filter(r => r.mentah < s.kkm).map(r => r.mentah);
            const bLo = below.length ? lo(below) : 0, bHi = below.length ? hi(below) : 0;
            pnRows.forEach(r => {
                if (r.mentah === null) return;
                if (r.mentah >= s.kkm) { r.akhir = r.mentah; return; }
                r.akhir = bHi > bLo ? s.rentangMin + (r.mentah - bLo) / (bHi - bLo) * (s.rentangMax - s.rentangMin) : s.rentangMin;
            });
        } else if (s.metode === 'rentang') {
            const vals = valid.map(r => r.mentah);
            const aLo = vals.length ? lo(vals) : 0, aHi = vals.length ? hi(vals) : 0;
            pnRows.forEach(r => {
                if (r.mentah === null) return;
                r.akhir = aHi > aLo ? s.rentangMin + (r.mentah - aLo) / (aHi - aLo) * (s.rentangMax - s.rentangMin) : Math.min(s.rentangMax, Math.max(s.rentangMin, r.mentah));
            });
        } else {
            pnRows.forEach(r => { r.akhir = r.mentah; });
        }
        pnRows.forEach(r => {
            if (r.akhir !== null) {
                r.akhir = Math.min(100, Math.max(0, r.akhir));
                r.akhir = s.bulat ? Math.round(r.akhir) : Math.round(r.akhir * 10) / 10;
            }
            if (Object.prototype.hasOwnProperty.call(keepManual, r.sid)) { r.akhir = keepManual[r.sid]; r.manual = true; }
        });
        renderPengolahTable();
    }

    function renderPengolahTable() {
        const tbody = document.getElementById('pn-tbody');
        if (!tbody) return;
        const s = pnReadSettings();
        if (pnRows.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" class="text-center text-muted py-4">Belum ada siswa di kelas ini.</td></tr>';
        } else {
            tbody.innerHTML = pnRows.map((r, i) => {
                const rounded = r.mentah === null ? null : (s.bulat ? Math.round(r.mentah) : Math.round(r.mentah * 10) / 10);
                const adjusted = r.akhir !== null && rounded !== null && r.akhir !== rounded;
                let ket = '<span class="text-muted small">Belum ada nilai</span>';
                if (r.akhir !== null) {
                    ket = r.akhir >= s.kkm ? '<span class="badge bg-success-subtle text-success-emphasis">Tuntas</span>' : '<span class="badge bg-danger-subtle text-danger-emphasis">Belum tuntas</span>';
                    if (r.manual) ket += ' <span class="badge bg-secondary-subtle text-secondary-emphasis" title="Diubah manual">manual</span>';
                    else if (adjusted) ket += ' <span class="badge bg-info-subtle text-info-emphasis" title="Disesuaikan dari nilai mentah">disesuaikan</span>';
                }
                return `<tr class="text-center">
                    <td>${i + 1}</td>
                    <td class="text-start fw-semibold">${escapeHtml(r.name)}</td>
                    <td>${pnFmt(r.harian)}</td><td>${pnFmt(r.pts)}</td><td>${pnFmt(r.pas)}</td><td>${pnFmt(r.hadir)}</td>
                    <td>${pnFmt(r.mentah)}</td>
                    <td><input type="number" min="0" max="100" step="0.1" class="form-control form-control-sm text-center fw-bold pn-akhir" data-sid="${escapeHtml(r.sid)}" value="${r.akhir === null ? '' : r.akhir}" aria-label="Nilai akhir ${escapeHtml(r.name)}"></td>
                    <td class="text-nowrap">${ket}</td></tr>`;
            }).join('');
        }
        renderPengolahStats();
    }

    function renderPengolahStats() {
        const el = document.getElementById('pn-stats'); if (!el) return;
        const s = pnReadSettings();
        const vals = pnRows.map(r => r.akhir).filter(v => v !== null);
        const card = (label, value, cls) => `<div class="col-6 col-md-3"><div class="p-2 rounded-3 border text-center ${cls || ''}"><div class="small text-muted">${label}</div><div class="fs-5 fw-bold">${value}</div></div></div>`;
        if (vals.length === 0) { el.innerHTML = card('Siswa dinilai', `0 / ${pnRows.length}`); return; }
        const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
        const tuntas = vals.filter(v => v >= s.kkm).length;
        el.innerHTML = card('Rata-rata kelas', pnFmt(avg)) + card('Tertinggi', pnFmt(Math.max(...vals))) + card('Terendah', pnFmt(Math.min(...vals))) +
            card('Tuntas (≥ KKM)', `${tuntas} / ${vals.length}`, tuntas === vals.length ? 'bg-success-subtle' : '');
    }

    // Ubahan pengaturan -> hitung ulang (ubahan manual guru dipertahankan)
    document.addEventListener('input', e => {
        const el = e.target;
        if (!el || !el.classList) return;
        if (el.classList.contains('pn-akhir')) {
            const row = pnRows.find(r => String(r.sid) === String(el.dataset.sid));
            if (row) { row.akhir = el.value === '' ? null : Math.min(100, Math.max(0, pnNum(el.value) ?? 0)); row.manual = true; renderPengolahStats(); }
            return;
        }
        if (el.classList.contains('pn-setting') && el.id !== 'pn-metode' && el.type !== 'checkbox') { pnUpdateSettingUI(); hitungPengolahNilai(false); }
    });
    document.addEventListener('change', e => {
        const el = e.target;
        if (!el) return;
        if (el.classList && el.classList.contains('pn-akhir')) { renderPengolahTable(); return; }
        if (el.id === 'pn-semester') pnSemesterTouched = true;
        if (el.id === 'pn-metode') {
            const kkm = pnNum(document.getElementById('pn-kkm').value) ?? 75;
            if (el.value === 'kkm') { document.getElementById('pn-rentang-min').value = kkm; document.getElementById('pn-rentang-max').value = Math.min(100, kkm + 5); }
            if (el.value === 'rentang') { document.getElementById('pn-rentang-min').value = kkm; document.getElementById('pn-rentang-max').value = 95; }
        }
        if (el.classList && el.classList.contains('pn-setting') && (el.id === 'pn-metode' || el.type === 'checkbox')) { pnUpdateSettingUI(); hitungPengolahNilai(false); }
    });

    function pnServerTooOld() {
        return IS_GAS && window.AGENDA_SERVER_INFO && Number(window.AGENDA_SERVER_INFO.apiVersion || 0) < 2;
    }

    function simpanPengolahNilai() {
        const { cid, subj, sem, tapel } = pnSelection();
        if (!cid || !subj) return Swal.fire('Pilih Data', 'Pilih kelas dan mata pelajaran terlebih dahulu.', 'warning');
        const s = pnReadSettings();
        const total = s.bobot.harian + s.bobot.pts + s.bobot.pas + s.bobot.hadir;
        if (total !== 100) return Swal.fire('Bobot Belum 100%', `Total bobot saat ini ${total}%. Sesuaikan agar totalnya tepat 100%.`, 'warning');
        if (s.metode !== 'none' && s.rentangMin > s.rentangMax) return Swal.fire('Rentang Tidak Valid', 'Batas bawah tidak boleh lebih besar dari batas atas.', 'warning');
        if (!pnRows.some(r => r.akhir !== null)) return Swal.fire('Belum Ada Nilai', 'Belum ada siswa yang memiliki nilai pada mapel dan semester ini.', 'info');
        if (pnServerTooOld()) {
            return Swal.fire({ icon: 'warning', title: 'Perbarui Code.gs', html: 'Fitur Pengolah Nilai butuh <b>Code.gs versi terbaru</b> agar hasilnya tersimpan di Spreadsheet.<br><br>Salin Code.gs terbaru ke Apps Script, lalu <b>Terapkan → Kelola deployment → Edit → Versi baru</b>.' });
        }
        const data = {};
        pnRows.forEach(r => {
            data[r.sid] = { harian: r.harian === null ? '' : Math.round(r.harian * 10) / 10, pts: r.pts ?? '', pas: r.pas ?? '', hadir: r.hadir === null ? '' : Math.round(r.hadir * 10) / 10,
                mentah: r.mentah === null ? '' : Math.round(r.mentah * 10) / 10, akhir: r.akhir ?? '', manual: !!r.manual };
        });
        if (!db.nilaiAkhir) db.nilaiAkhir = [];
        let rec = findNilaiOlahan(cid, subj, sem, tapel);
        if (!rec) { rec = { RecordID: 'NA' + Date.now(), ClassID: cid, Subject: subj, Semester: sem, TahunPelajaran: tapel }; db.nilaiAkhir.push(rec); }
        rec.Pengaturan = JSON.stringify(s);
        rec.NilaiData = JSON.stringify(data);
        rec.UpdatedAt = new Date().toISOString();
        try { localStorage.setItem(PN_SETTINGS_KEY, JSON.stringify(s)); } catch (e) { }
        syncDB();
        pnSetStatus(rec, 0);
        Swal.fire({ icon: 'success', title: 'Nilai Akhir Tersimpan', text: `Nilai akhir ${subj} otomatis dipakai di Cetak Rapor dan Rekap Laporan.`, timer: 2200, showConfirmButton: false });
    }

    function hapusPengolahNilai() {
        const { cid, subj, sem, tapel } = pnSelection();
        const rec = cid && subj ? findNilaiOlahan(cid, subj, sem, tapel) : null;
        if (!rec) return;
        Swal.fire({
            icon: 'warning', title: 'Hapus hasil pengolahan?', text: `Rapor ${subj} akan kembali memakai rata-rata biasa (harian, PTS, PAS). Data nilai asli tidak terhapus.`,
            showCancelButton: true, confirmButtonText: 'Ya, hapus', cancelButtonText: 'Batal', confirmButtonColor: '#C44747'
        }).then(r => {
            if (!r.isConfirmed) return;
            db.nilaiAkhir = db.nilaiAkhir.filter(x => x !== rec);
            syncDB();
            loadPengolahNilai();
            Swal.fire({ icon: 'success', title: 'Dihapus', timer: 1500, showConfirmButton: false });
        });
    }

    function pnSafeName(s) { return String(s || '').replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, '_'); }
    function pnKelasName(cid) { return db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid; }
    function pnMetodeText(s) {
        return s.metode === 'kkm' ? `Nilai < KKM disesuaikan ke ${s.rentangMin}-${s.rentangMax}` : s.metode === 'rentang' ? `Konversi ke rentang ${s.rentangMin}-${s.rentangMax}` : 'Tanpa penyesuaian';
    }

    async function downloadPengolahExcel() {
        const { cid, subj, sem, tapel } = pnSelection();
        if (!cid || !subj || pnRows.length === 0) return Swal.fire('Pilih Data', 'Pilih kelas dan mata pelajaran yang memiliki siswa.', 'warning');
        await ensureXLSX();
        const s = pnReadSettings();
        const v = x => x === null || x === undefined ? '' : Math.round(x * 10) / 10;
        const aoa = [
            ['PENGOLAHAN NILAI AKHIR'],
            ['Sekolah', db.config.userSchool || ''], ['Kelas', pnKelasName(cid)], ['Mata Pelajaran', subj],
            ['Semester', sem === 'ganjil' ? '1 (Ganjil)' : '2 (Genap)'], ['Tahun Pelajaran', tapel],
            ['Bobot', `Harian ${s.bobot.harian}% | PTS ${s.bobot.pts}% | PAS ${s.bobot.pas}% | Kehadiran ${s.bobot.hadir}%`],
            ['KKM', s.kkm], ['Penyesuaian', pnMetodeText(s)], [],
            ['No', 'NIS', 'Nama Siswa', 'Harian', 'PTS', 'PAS', 'Hadir (%)', 'Nilai Mentah', 'Nilai Akhir', 'Keterangan']
        ];
        pnRows.forEach((r, i) => aoa.push([i + 1, r.nis, r.name, v(r.harian), v(r.pts), v(r.pas), v(r.hadir), v(r.mentah), r.akhir ?? '', r.akhir === null ? '' : (r.akhir >= s.kkm ? 'Tuntas' : 'Belum tuntas')]));
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!cols'] = [{ wch: 5 }, { wch: 12 }, { wch: 30 }, { wch: 9 }, { wch: 9 }, { wch: 9 }, { wch: 10 }, { wch: 13 }, { wch: 11 }, { wch: 14 }];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Pengolahan Nilai');
        XLSX.writeFile(wb, `PengolahNilai_${pnSafeName(pnKelasName(cid))}_${pnSafeName(subj)}_${sem}_${pnSafeName(tapel)}.xlsx`);
    }

    // =====================================================================
    // REKAP LAPORAN (Excel): leger nilai akhir, kehadiran, dan detail pengolahan
    // =====================================================================
    function loadRekapView() {
        const el = document.getElementById('rekap-kelas'); if (!el) return;
        const cur = el.value;
        el.innerHTML = '<option value="">-- Pilih Kelas --</option>' + db.classes.map(c => `<option value="${escapeHtml(c.ClassID)}">${escapeHtml(c.ClassName)}</option>`).join('');
        if (cur && db.classes.some(c => String(c.ClassID) === String(cur))) el.value = cur;
        else if (db.classes.length === 1) el.value = db.classes[0].ClassID;
        const semEl = document.getElementById('rekap-semester');
        if (semEl && !semEl.dataset.touched) {
            semEl.value = (getActiveTapel() === getDefaultTapel() && new Date().getMonth() + 1 < 7) ? 'genap' : 'ganjil';
            semEl.addEventListener('change', () => { semEl.dataset.touched = '1'; }, { once: true });
        }
        const lbl = document.getElementById('rekap-tapel-label'); if (lbl) lbl.textContent = getActiveTapel();
        renderRekapLaporan();
    }

    function buildRekapData(cid, sem, tapel) {
        const students = db.students.filter(s => String(s.ClassID) === String(cid));
        const perStudent = students.map(st => {
            const grades = getStudentAllGrades(cid, st.StudentID, sem, tapel);
            const att = getStudentAttendanceSummary(cid, st.StudentID, sem, tapel);
            return { st, grades, att };
        });
        const subjects = [];
        perStudent.forEach(p => p.grades.forEach(g => { if (g.hasScore && g.finalScore !== null && !subjects.includes(g.subject)) subjects.push(g.subject); }));
        const rows = perStudent.map(p => {
            const scores = {};
            p.grades.forEach(g => { if (subjects.includes(g.subject)) scores[g.subject] = g.finalScore; });
            const vals = subjects.map(s => scores[s]).filter(v => v !== null && v !== undefined);
            const jumlah = vals.reduce((a, b) => a + b, 0);
            const total = p.att.H + p.att.S + p.att.I + p.att.A;
            return { st: p.st, scores, jumlah: vals.length ? Math.round(jumlah * 10) / 10 : null, rata: vals.length ? Math.round(jumlah / vals.length * 100) / 100 : null, att: p.att, persenHadir: total ? Math.round(p.att.H / total * 1000) / 10 : null };
        });
        const ranked = rows.filter(r => r.rata !== null).sort((a, b) => b.rata - a.rata);
        ranked.forEach((r, i) => { r.rank = i > 0 && ranked[i - 1].rata === r.rata ? ranked[i - 1].rank : i + 1; });
        const olahanSubjects = subjects.filter(s => findNilaiOlahan(cid, s, sem, tapel));
        return { subjects, rows, olahanSubjects };
    }

    function renderRekapLaporan() {
        const cid = document.getElementById('rekap-kelas')?.value, sem = document.getElementById('rekap-semester')?.value || 'ganjil', tapel = getActiveTapel();
        const box = document.getElementById('rekap-preview'), info = document.getElementById('rekap-info');
        if (!box) return;
        if (!cid) { box.innerHTML = '<div class="text-center text-muted py-4">Pilih kelas terlebih dahulu.</div>'; if (info) info.textContent = ''; return; }
        const { subjects, rows, olahanSubjects } = buildRekapData(cid, sem, tapel);
        if (info) info.textContent = `${rows.length} siswa · ${subjects.length} mapel bernilai · ${olahanSubjects.length} mapel memakai nilai dari Pengolah Nilai` +
            (subjects.length > olahanSubjects.length ? ' (mapel lain memakai rata-rata biasa)' : '');
        if (rows.length === 0) { box.innerHTML = '<div class="text-center text-muted py-4">Belum ada siswa di kelas ini.</div>'; return; }
        const head = `<tr class="text-center small"><th>No</th><th class="text-start">Nama Siswa</th>${subjects.map(s => `<th title="${escapeHtml(s)}">${escapeHtml(s.length > 14 ? s.slice(0, 13) + '…' : s)}${olahanSubjects.includes(s) ? ' <i class="bi bi-calculator text-primary" title="Dari Pengolah Nilai"></i>' : ''}</th>`).join('')}<th>Jumlah</th><th>Rata-rata</th><th>Rank</th><th>S</th><th>I</th><th>A</th></tr>`;
        const body = rows.map((r, i) => `<tr class="text-center"><td>${i + 1}</td><td class="text-start text-nowrap fw-semibold">${escapeHtml(r.st.StudentName)}</td>${subjects.map(s => `<td>${pnFmt(r.scores[s] ?? null)}</td>`).join('')}<td>${pnFmt(r.jumlah)}</td><td class="fw-bold">${pnFmt(r.rata)}</td><td>${r.rank || '-'}</td><td>${r.att.S}</td><td>${r.att.I}</td><td>${r.att.A}</td></tr>`).join('');
        box.innerHTML = `<table class="table table-bordered table-sm align-middle mb-0"><thead class="table-light">${head}</thead><tbody>${body}</tbody></table>`;
    }

    async function downloadRekapExcel() {
        const cid = document.getElementById('rekap-kelas')?.value, sem = document.getElementById('rekap-semester')?.value || 'ganjil', tapel = getActiveTapel();
        if (!cid) return Swal.fire('Pilih Kelas', 'Pilih kelas yang akan direkap.', 'warning');
        await ensureXLSX();
        const { subjects, rows } = buildRekapData(cid, sem, tapel);
        const kelas = pnKelasName(cid), semText = sem === 'ganjil' ? '1 (Ganjil)' : '2 (Genap)';
        const header = title => [[title], ['Sekolah', db.config.userSchool || ''], ['Kelas', kelas], ['Semester', semText], ['Tahun Pelajaran', tapel], ['Wali Kelas', db.config.userName || ''], []];
        const num = x => x === null || x === undefined ? '' : x;

        const leger = header('LEGER NILAI AKHIR RAPOR');
        leger.push(['No', 'NIS', 'NISN', 'Nama Siswa', ...subjects, 'Jumlah', 'Rata-rata', 'Peringkat']);
        rows.forEach((r, i) => leger.push([i + 1, r.st.NIS || '', r.st.NISN || '', r.st.StudentName, ...subjects.map(s => num(r.scores[s])), num(r.jumlah), num(r.rata), r.rank || '']));
        const wsLeger = XLSX.utils.aoa_to_sheet(leger);
        wsLeger['!cols'] = [{ wch: 5 }, { wch: 12 }, { wch: 13 }, { wch: 30 }, ...subjects.map(() => ({ wch: 12 })), { wch: 9 }, { wch: 10 }, { wch: 10 }];

        const hadir = header('REKAP KEHADIRAN SISWA');
        hadir.push(['No', 'NIS', 'Nama Siswa', 'Hadir', 'Sakit', 'Izin', 'Alpa', '% Kehadiran']);
        rows.forEach((r, i) => hadir.push([i + 1, r.st.NIS || '', r.st.StudentName, r.att.H, r.att.S, r.att.I, r.att.A, num(r.persenHadir)]));
        const wsHadir = XLSX.utils.aoa_to_sheet(hadir);
        wsHadir['!cols'] = [{ wch: 5 }, { wch: 12 }, { wch: 30 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 12 }];

        const olah = header('DETAIL PENGOLAHAN NILAI');
        olah.push(['Mata Pelajaran', 'No', 'Nama Siswa', 'Harian', 'PTS', 'PAS', 'Hadir (%)', 'Nilai Mentah', 'Nilai Akhir', 'Bobot (H/PTS/PAS/Hadir)', 'KKM', 'Penyesuaian']);
        const recs = (db.nilaiAkhir || []).filter(r => String(r.ClassID) === String(cid) && String(r.Semester || '').toLowerCase() === sem && pnSameTapel(r.TahunPelajaran, tapel));
        if (recs.length === 0) olah.push(['Belum ada mapel yang disimpan dari menu Pengolah Nilai. Nilai di leger memakai rata-rata biasa.']);
        recs.forEach(rec => {
            let s = PN_DEFAULT; try { s = Object.assign({}, PN_DEFAULT, JSON.parse(rec.Pengaturan || '{}')); } catch (e) { }
            const data = parseNilaiData(rec);
            rows.forEach((r, i) => {
                const d = data[r.st.StudentID] || {};
                olah.push([rec.Subject, i + 1, r.st.StudentName, num(pnNum(d.harian)), num(pnNum(d.pts)), num(pnNum(d.pas)), num(pnNum(d.hadir)), num(pnNum(d.mentah)), num(pnNum(d.akhir)),
                    `${s.bobot.harian}/${s.bobot.pts}/${s.bobot.pas}/${s.bobot.hadir}`, s.kkm, pnMetodeText(s)]);
            });
        });
        const wsOlah = XLSX.utils.aoa_to_sheet(olah);
        wsOlah['!cols'] = [{ wch: 24 }, { wch: 5 }, { wch: 30 }, { wch: 9 }, { wch: 9 }, { wch: 9 }, { wch: 10 }, { wch: 13 }, { wch: 11 }, { wch: 22 }, { wch: 6 }, { wch: 34 }];

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, wsLeger, 'Leger Nilai');
        XLSX.utils.book_append_sheet(wb, wsHadir, 'Rekap Kehadiran');
        XLSX.utils.book_append_sheet(wb, wsOlah, 'Pengolahan Nilai');
        XLSX.writeFile(wb, `Rekap_${pnSafeName(kelas)}_Semester-${sem}_${pnSafeName(tapel)}.xlsx`);
    }

    // =====================================================================
    // DATA SEKOLAH & PENANDATANGAN RAPOR
    // Tidak ada lagi data contoh (nama/NIP penulis, kepala sekolah contoh, dsb.).
    // Isian yang belum diisi guru tampil sebagai titik-titik di rapor, dan guru
    // diingatkan untuk melengkapinya sebelum mencetak.
    // =====================================================================
    const RAPOR_KOSONG = '..............................';

    // Versi lama menyimpan data contoh ke Spreadsheet guru secara otomatis. Bersihkan
    // hanya nilai yang persis sama dengan data contoh tersebut.
    function bersihkanDataContohRapor() {
        const c = db.config || {};
        let changed = false;
        const clear = (key, contoh) => { if (c[key] === contoh) { c[key] = ''; changed = true; } };
        const dataContohOtomatis = c.userNIP === '198501012010012001' || c.headmasterNIP === '197005121995121001';
        clear('userNIP', '198501012010012001');
        clear('headmasterNIP', '197005121995121001');
        clear('headmasterName', 'Erik Santoso, S.Pd');
        clear('userSchool', 'SMPN 1 INDONESIA');
        clear('schoolAddress', 'Jl. Raya Indonesia, No. 17');
        if (dataContohOtomatis) { clear('userName', 'Resti Novika, M.Pd.'); clear('schoolCity', 'Bandung'); }
        return changed;
    }

    function getPengaturanRaporKurang() {
        const c = db.config || {}, kurang = [];
        if (!c.userSchool || c.userSchool === 'Sekolah Saya') kurang.push('Nama sekolah');
        if (!c.schoolAddress) kurang.push('Alamat sekolah');
        if (!c.schoolCity) kurang.push('Kota (titimangsa)');
        if (!c.userName) kurang.push('Nama wali kelas');
        if (!c.headmasterName) kurang.push('Nama kepala sekolah');
        return kurang;
    }

    // Sebelum mencetak/unduh PDF: ingatkan jika data sekolah belum lengkap
    function cekPengaturanRapor(lanjut) {
        const kurang = getPengaturanRaporKurang();
        if (kurang.length === 0) return lanjut();
        Swal.fire({
            icon: 'info',
            title: 'Lengkapi Data Sekolah',
            html: `Data berikut belum diisi, sehingga di rapor tampil sebagai titik-titik:<br><b>${kurang.map(escapeHtml).join(', ')}</b>`,
            showDenyButton: true, showCancelButton: true,
            confirmButtonText: 'Isi Sekarang', denyButtonText: 'Tetap Lanjutkan', cancelButtonText: 'Batal'
        }).then(r => {
            if (r.isConfirmed) openModalPengaturanRapor();
            else if (r.isDenied) lanjut();
        });
    }

    // Pemberitahuan di atas pratinjau rapor (tidak ikut tercetak)
    function tampilkanPeringatanPengaturanRapor() {
        const box = document.getElementById('rapor-preview-container');
        if (!box || !box.querySelector('.rapor-paper-sheet')) return;
        const old = document.getElementById('rapor-setting-warning'); if (old) old.remove();
        const kurang = getPengaturanRaporKurang();
        if (kurang.length === 0) return;
        const div = document.createElement('div');
        div.id = 'rapor-setting-warning';
        div.className = 'alert alert-warning d-flex flex-wrap align-items-center justify-content-between gap-2 no-print';
        div.innerHTML = `<div><i class="bi bi-exclamation-triangle me-1"></i> Data sekolah belum lengkap: <b>${kurang.map(escapeHtml).join(', ')}</b>. Bagian ini tampil sebagai titik-titik di rapor.</div>` +
            '<button type="button" class="btn btn-sm btn-warning fw-semibold" onclick="openModalPengaturanRapor()"><i class="bi bi-gear me-1"></i> Isi Pengaturan Sekolah</button>';
        box.prepend(div);
    }
    const _previewRaporSiswaAsli = previewRaporSiswa;
    previewRaporSiswa = function () {
        const hasil = _previewRaporSiswaAsli.apply(this, arguments);
        tampilkanPeringatanPengaturanRapor();
        return hasil;
    };

    // =====================================================================
    // LISENSI
    // Code.gs pembeli mengirim "izin" bertanda tangan digital dari server lisensi
    // penjual. Tanda tangan diperiksa di sini dengan kunci publik: hanya server
    // penjual (pemegang kunci privat) yang bisa membuat izin yang sah.
    // Mode demo (tanpa server) tidak memerlukan lisensi.
    // =====================================================================
    const LICENSE_PUBLIC_KEY = `-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAki4N+toJuvAs+v6EkBPs
bP5gsJd1ZV4AJiYaflloWJiM+X2JsMfh873oJwHQRj8BiBU/XcKRqIHC4mBZL7Ao
xPO+I2MQLSFyTevu8+N40G3jxIjfL1yIxepEyRIDoF4BUZLX+bMNN7rQ1FzgYHy7
inDRAjXdnNWxunyFG/4X0xz4dQbbsy15UEJJHoGHmGM94uEqJhA4bOU6cFlThMmq
l2BCG1tLkz7XfLKRdqq4Nqd0a3xBJ1yobc5abZO/6ulzxoZfQ9ZXKoCny//+27r/
AEITYTlu03HJBSUGs7Xky3o9tMexOx6eNZHN2VGLtazX6FPyqoAr+qT88HOgTlM2
gQIDAQAB
-----END PUBLIC KEY-----`;
    const LICENSE_CACHE_KEY = 'agendaGuru_license';
    const LICENSE_SUPPORT_WA = '6285788408882';
    const LICENSE_KEY_RE = /^AG-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}$/;

    function normalizeLicenseKey(v) {
        let s = String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (s.startsWith('AG')) s = s.slice(2);
        return s.length === 12 ? `AG-${s.slice(0, 4)}-${s.slice(4, 8)}-${s.slice(8)}` : String(v || '').trim().toUpperCase();
    }

    async function verifySignedLicense(lic) {
        if (!lic || !lic.payload || !lic.sig || !window.crypto || !crypto.subtle) return null;
        try {
            const der = Uint8Array.from(atob(LICENSE_PUBLIC_KEY.replace(/-----[^-]+-----|\s/g, '')), c => c.charCodeAt(0));
            const key = await crypto.subtle.importKey('spki', der, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
            const sig = Uint8Array.from(atob(lic.sig), c => c.charCodeAt(0));
            const ok = await crypto.subtle.verify('RSASSA-PKCS1-v1_5', key, sig, new TextEncoder().encode(lic.payload));
            if (!ok) return null;
            const p = JSON.parse(lic.payload);
            if (p.app !== 'AgendaGuru') return null;
            if (lic.scriptId && p.sid !== lic.scriptId) return null;
            return p;
        } catch (e) {
            return null;
        }
    }

    // true = boleh lanjut membuka aplikasi. false = layar aktivasi lisensi ditampilkan.
    async function licenseGate(lic, offline) {
        if (!IS_GAS) return true;
        if (offline) { try { lic = JSON.parse(localStorage.getItem(LICENSE_CACHE_KEY) || 'null'); } catch (e) { lic = null; } }
        if (!lic) { showLicenseScreen(offline ? 'offline' : 'upgrade'); return false; }
        const p = await verifySignedLicense(lic);
        if (p && p.exp > Date.now()) {
            if (!offline) { try { localStorage.setItem(LICENSE_CACHE_KEY, JSON.stringify({ payload: lic.payload, sig: lic.sig, scriptId: lic.scriptId })); } catch (e) { } }
            window.AGENDA_LICENSE = p;
            return true;
        }
        const reason = lic.reason || (p ? (lic.hasKey === false ? 'missing' : 'renew') : (lic.payload ? 'invalid' : (lic.hasKey ? 'renew' : 'missing')));
        showLicenseScreen(reason, lic.message);
        return false;
    }

    function showLicenseScreen(reason, serverMessage) {
        const ov = document.getElementById('loading-overlay'); if (ov) ov.style.display = 'none';
        ['app-view', 'login-view'].forEach(id => { const el = document.getElementById(id); if (el) el.style.display = 'none'; });
        const texts = {
            missing: 'Akun ini belum memiliki lisensi. Masukkan kode lisensi yang Anda terima dari penjual.',
            revoked: 'Lisensi untuk aplikasi ini sudah dinonaktifkan. Hubungi penjual, atau masukkan kode lisensi lain.',
            expired: 'Masa berlaku lisensi sudah habis. Hubungi penjual untuk memperpanjang.',
            limit: 'Kode lisensi ini sudah dipakai di Spreadsheet lain.',
            invalid: 'Data lisensi tidak valid. Masukkan ulang kode lisensi Anda.',
            renew: 'Izin lisensi perlu diperbarui, tetapi server lisensi belum dapat dihubungi. Coba lagi beberapa saat, atau masukkan ulang kode lisensi.',
            offline: 'Perangkat sedang offline dan izin lisensi di perangkat ini belum ada atau sudah habis. Sambungkan internet lalu muat ulang.',
            upgrade: 'Code.gs di Spreadsheet Anda masih versi lama dan belum mendukung lisensi. Salin Code.gs terbaru ke Apps Script, lalu buat versi deployment baru.'
        };
        const canActivate = reason !== 'upgrade' && reason !== 'offline';
        const old = document.getElementById('ag-license'); if (old) old.remove();
        const wrap = document.createElement('div');
        wrap.id = 'ag-license';
        wrap.style.cssText = 'position:fixed;inset:0;z-index:4500;background:#172B4D;display:flex;align-items:center;justify-content:center;padding:16px;overflow:auto';
        wrap.innerHTML = `
            <div style="background:#fff;border-radius:16px;max-width:460px;width:100%;padding:28px;font-family:var(--font-family, system-ui, sans-serif);color:#172B4D">
                <div style="font-size:2rem;line-height:1">🔑</div>
                <h4 style="font-weight:800;margin:10px 0 6px">Aktivasi Lisensi AgendaGuru</h4>
                <p style="color:#5E6C84;font-size:14px;margin:0 0 4px">${escapeHtml(texts[reason] || texts.invalid)}</p>
                ${serverMessage && reason === 'limit' ? `<p style="color:#5E6C84;font-size:13px;margin:0 0 4px">${escapeHtml(serverMessage)}</p>` : ''}
                ${canActivate ? `
                <label for="ag-license-input" style="font-weight:700;font-size:13px;margin-top:12px;display:block">Kode Lisensi</label>
                <input id="ag-license-input" class="form-control form-control-lg text-uppercase" placeholder="AG-XXXX-XXXX-XXXX" autocomplete="off" spellcheck="false" style="margin:6px 0">
                <div id="ag-license-msg" role="status" aria-live="polite" style="font-size:13px;min-height:20px;color:#C44747"></div>
                <button id="ag-license-btn" type="button" class="btn btn-primary w-100 fw-bold">Aktifkan</button>` :
                (reason === 'upgrade' ? `<a href="panduan.html" target="_blank" rel="noopener" class="btn btn-primary w-100 fw-bold mt-3">Buka Panduan Pembaruan</a>` :
                    `<button type="button" class="btn btn-primary w-100 fw-bold mt-3" onclick="location.reload()">Muat Ulang</button>`)}
                <a href="https://wa.me/${LICENSE_SUPPORT_WA}?text=${encodeURIComponent('Halo, saya butuh bantuan lisensi AgendaGuru')}" target="_blank" rel="noopener" class="btn btn-outline-success w-100 mt-2">💬 Hubungi Penjual (WhatsApp)</a>
                <button id="ag-license-logout" type="button" class="btn btn-link w-100 text-secondary mt-1">Keluar Akun</button>
                <p style="color:#5E6C84;font-size:12px;margin:10px 0 0">Data Anda tetap aman di Google Spreadsheet milik Anda.</p>
            </div>`;
        document.body.appendChild(wrap);
        document.getElementById('ag-license-logout').onclick = () => {
            setToken('');
            try { localStorage.removeItem(LICENSE_CACHE_KEY); } catch (e) { }
            location.reload();
        };
        if (!canActivate) return;
        const input = document.getElementById('ag-license-input'), msg = document.getElementById('ag-license-msg'), btn = document.getElementById('ag-license-btn');
        const go = () => {
            const key = normalizeLicenseKey(input.value);
            input.value = key;
            if (!LICENSE_KEY_RE.test(key)) { msg.textContent = 'Format kode: AG-XXXX-XXXX-XXXX'; return; }
            btn.disabled = true; msg.style.color = '#5E6C84'; msg.textContent = 'Memeriksa kode lisensi...';
            google.script.run
                .withSuccessHandler(async res => {
                    if (res && res.authError) { setToken(''); location.reload(); return; }
                    if (res && res.ok && await verifySignedLicense(res.license)) {
                        try { localStorage.setItem(LICENSE_CACHE_KEY, JSON.stringify(res.license)); } catch (e) { }
                        msg.style.color = '#21865B'; msg.textContent = 'Lisensi aktif! Memuat aplikasi...';
                        setTimeout(() => location.reload(), 600);
                        return;
                    }
                    btn.disabled = false; msg.style.color = '#C44747';
                    msg.textContent = (res && res.message) || 'Kode lisensi tidak valid.';
                })
                .withFailureHandler(err => { btn.disabled = false; msg.style.color = '#C44747'; msg.textContent = 'Tidak dapat terhubung: ' + ((err && err.message) || err); })
                .activateLicense(getToken(), key);
        };
        btn.onclick = go;
        input.addEventListener('keydown', e => { if (e.key === 'Enter') go(); });
        input.focus();
    }