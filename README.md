# Panduan Deploy AgendaGuru (Frontend GitHub + Backend Apps Script)

```
 Guru membuka ─▶ https://restinovika123-stack.github.io/agendaguru_/   (frontend, selalu versi terbaru)
                         │  fetch (JSON)
                         ▼
               Apps Script milik guru  (Code.gs = API)  ─▶  Google Sheet milik guru (database)
```

- **Anda (pemilik aplikasi)** cukup memperbarui frontend di GitHub. Semua guru otomatis
  mendapat versi baru saat membuka/memuat ulang aplikasi (paling lama ±10 menit karena cache GitHub).
- **Guru** memasang Code.gs sekali saja. Data tiap guru terpisah di Spreadsheet masing-masing.

---

## A. Untuk Anda: pasang frontend di GitHub (sekali)

1. Buat repository baru di GitHub, misalnya `agendaguru_` (Public).
2. Pastikan baris paling atas `Code.gs` sudah berisi alamat GitHub Pages Anda (sudah diisi):
   ```js
   const FRONTEND_URL = 'https://restinovika123-stack.github.io/agendaguru_/';
   ```
   Jika nama repo diganti, ubah juga bagian `agendaguru_` di alamat ini.
3. Jalankan di folder proyek:
   ```
   node build_frontend.js
   ```
   Hasilnya ada di folder `frontend-github/`.
4. Unggah **isi** folder `frontend-github/` ke repository (tombol *Add file → Upload files*,
   atau pakai git — lihat bagian D).
5. Di GitHub: **Settings → Pages → Source: Deploy from a branch → Branch: main / (root) → Save**.
6. Tunggu 1–2 menit, buka `https://restinovika123-stack.github.io/agendaguru_/`.

## B. Untuk Anda: siapkan TEMPLATE Spreadsheet (sekali, ulangi saat Code.gs diperbarui)

Pembeli cukup membuat salinan template. Kode aplikasi (Code.gs) ikut tersalin, jadi pembeli tidak perlu menyalin kode.

1. Buat Google Spreadsheet baru bernama **TEMPLATE AgendaGuru** (pakai akun penjual).
2. **Ekstensi → Apps Script**. Hapus isinya, tempel seluruh isi `Code.gs` terbaru, lalu **Simpan**.
3. Pilih fungsi **`bersihkanTemplate`** di bagian atas editor, klik **Jalankan**, lalu izinkan akses.
   Semua akun, sesi, lisensi, dan isi data di template dihapus; tersisa sheet **Petunjuk** untuk pembeli.
4. **Jangan** melakukan Deploy atau mendaftar akun di template ini.
5. Klik **Bagikan** → *Akses umum*: **Siapa saja yang memiliki link** → **Pelihat**.
6. Salin link template, lalu ganti bagian akhir `/edit...` menjadi `/copy`. Contoh:
   `https://docs.google.com/spreadsheets/d/ID_TEMPLATE/copy`
   Link ini dipasang di tombol **Buat Salinan Template** pada landing page panduan (`landing/index.html`, https://agendaguru-id.vercel.app).

Setiap kali Code.gs diperbarui, ulangi langkah 2–3 di template agar pembeli baru langsung mendapat versi terbaru.

### Yang dilakukan pembeli (lihat https://agendaguru-id.vercel.app)
1. Klik **Buat Salinan Template** → **Buat salinan**.
2. Di salinan: **Ekstensi → Apps Script → Terapkan → Deployment baru → Aplikasi web**
   (Jalankan sebagai **Saya**, akses **Siapa saja**), lalu izinkan akses.
3. Di Spreadsheet: menu **AgendaGuru → Buka Aplikasi**, lalu daftar akun dengan **Kode Lisensi**.

> Pengaman: jika template terlanjur pernah dipakai, akun/sesi/lisensi yang ikut tersalin otomatis
> dihapus saat salinan pertama kali dibuka, sehingga setiap pembeli memulai dengan akun sendiri.

## C. Memperbarui aplikasi

| Yang diubah | Yang perlu dilakukan |
|---|---|
| Tampilan / fitur (`Index.html`, `JavaScript.html`, `api.js`) | `node build_frontend.js` lalu unggah ulang isi `frontend-github/`. Guru **tidak perlu** melakukan apa-apa. |
| `Code.gs` (API/database) | Jarang. Naikkan `API_VERSION` di Code.gs dan `MIN_API_VERSION` di api.js. Guru menempel Code.gs baru lalu **Terapkan → Kelola deployment → Edit → Versi baru**. Aplikasi menampilkan peringatan ke guru yang Code.gs-nya masih lama. |

Tips: usahakan perubahan fitur hanya di frontend, supaya guru tidak perlu memperbarui Code.gs.

**Code.gs versi 2 (wajib untuk menu Pengolah Nilai):** sejak versi ini, setiap jenis data baru dari
frontend otomatis disimpan ke sheet berawalan `X_` (misalnya `X_nilaiAkhir`). Jadi fitur baru
berikutnya umumnya tidak perlu memperbarui Code.gs lagi.

## D. (Opsional) Unggah dengan git

```
cd frontend-github
git init -b main
git add .
git commit -m "AgendaGuru frontend"
git remote add origin https://github.com/restinovika123-stack/agendaguru_.git
git push -u origin main
```
Pembaruan berikutnya: `node build_frontend.js`, lalu di folder `frontend-github`:
`git add . && git commit -m "update" && git push`.

## E. Lupa password / masalah

- **Lupa password:** di editor Apps Script, pilih fungsi `resetAkunLogin` → **Jalankan**.
  Data tetap ada; buka aplikasi lalu daftar ulang.
- **Foto bukti lama masih publik:** jalankan fungsi `jadikanFolderBuktiPrivat` sekali.
- **"Tidak dapat terhubung ke server":** pastikan deployment aksesnya *Siapa saja* dan URL
  berakhiran `/exec`.

## F. Landing page panduan (Vercel)

- Alamat: **https://agendaguru-id.vercel.app** (panduan) dan **/privasi** (kebijakan privasi).
- Sumber halaman ada di folder `landing/` (`index.html`, `privasi.html`, `vercel.json`).
- Tombol "Salin Code.gs" di landing page mengambil Code.gs terbaru dari GitHub, jadi tidak perlu
  diperbarui setiap kali Code.gs berubah.
- Alamat lama `…github.io/agendaguru_/panduan.html` otomatis dialihkan ke Vercel.
