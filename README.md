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

## B. Untuk guru: pasang backend (sekali per guru)

1. Buat Google Spreadsheet baru (kosong).
2. Menu **Ekstensi → Apps Script**. Hapus isi `Code.gs`, tempel seluruh isi `Code.gs` dari
   `https://restinovika123-stack.github.io/agendaguru_/Code.gs`, lalu **Simpan**.
3. **Terapkan → Deployment baru** → ikon roda gigi → **Aplikasi web**.
   - Jalankan sebagai: **Saya**
   - Yang memiliki akses: **Siapa saja**
   - Klik **Terapkan**, lalu izinkan akses (Tinjau izin → pilih akun → Lanjutan → Buka).
4. Buka **URL Aplikasi web** yang muncul → klik **Buka AgendaGuru** → **Daftar akun**.
5. Simpan (bookmark) alamat GitHub aplikasi. Di HP/laptop lain cukup buka alamat yang sama,
   lalu tempel URL Aplikasi web saat diminta (atau buka lewat URL Aplikasi web lagi).

Sheet database dibuat otomatis saat pertama dipakai.

> Akun Google Workspace sekolah kadang melarang akses "Siapa saja". Jika gagal terhubung,
> gunakan akun Gmail pribadi untuk Spreadsheet-nya.

## C. Memperbarui aplikasi

| Yang diubah | Yang perlu dilakukan |
|---|---|
| Tampilan / fitur (`Index.html`, `JavaScript.html`, `api.js`) | `node build_frontend.js` lalu unggah ulang isi `frontend-github/`. Guru **tidak perlu** melakukan apa-apa. |
| `Code.gs` (API/database) | Jarang. Naikkan `API_VERSION` di Code.gs dan `MIN_API_VERSION` di api.js. Guru menempel Code.gs baru lalu **Terapkan → Kelola deployment → Edit → Versi baru**. Aplikasi menampilkan peringatan ke guru yang Code.gs-nya masih lama. |

Tips: usahakan perubahan fitur hanya di frontend, supaya guru tidak perlu memperbarui Code.gs.

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
