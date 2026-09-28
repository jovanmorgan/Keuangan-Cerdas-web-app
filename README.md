# Keuangan Cerdas — Web App

Frontend (`index.html`) dihosting terpisah di **Vercel** (lewat GitHub), sedangkan
data & logika penyimpanan tetap di **Google Sheets**, dijembatani oleh
**Google Apps Script** yang di-deploy sebagai Web App (REST API sederhana).

```
Browser (Vercel) ──fetch()──▶ Apps Script Web App (/exec) ──▶ Google Sheets
```

## Struktur folder

```
index.html            ← UI, tinggal di-deploy ke Vercel apa adanya (statis, tanpa build)
api/config.js          ← serverless function: expose Environment Variables ke frontend
vercel.json            ← konfigurasi Vercel (opsional, sudah cukup minimal)
apps-script/Code.gs    ← backend, DITEMPEL ke Apps Script project kamu (bukan yang di-deploy Vercel)
```

`apps-script/Code.gs` hanya disimpan di repo ini sebagai salinan/versi cadangan.
File ini **tidak dideploy oleh Vercel** — dia harus ditempel manual (atau lewat
`clasp`) ke project Apps Script yang terhubung ke Google Sheets kamu.

---

## 1) Deploy backend (Apps Script)

1. Buka Google Sheet kamu → **Extensions > Apps Script**.
2. Tempel isi `apps-script/Code.gs` ke `Code.gs` di project itu (timpa yang lama,
   atau merge kalau kamu sudah punya fungsi lain seperti `setupDatabase()`).
3. **(Sangat disarankan)** Atur token rahasia supaya orang lain yang tahu URL
   Web App tidak bisa ikut mengubah data:
   - Klik ⚙️ **Project Settings** → scroll ke **Script Properties**.
   - Tambah properti: key `APP_TOKEN`, value = teks rahasia bebas (mis. string acak panjang).
4. **Deploy > New deployment**:
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Klik **Deploy**, izinkan akses saat diminta.
5. Salin **Web app URL** yang muncul (diakhiri `/exec`). Ini yang nanti dimasukkan
   ke UI di langkah 3.

> Setiap kali kamu mengubah `Code.gs`, buat **New deployment** baru (atau
> "Manage deployments" → edit versi) supaya perubahan ikut terpakai di URL yang
> sama.

---

## 2) Push ke GitHub & hosting di Vercel

```bash
git init
git add .
git commit -m "Keuangan Cerdas web app"
git branch -M main
git remote add origin https://github.com/USERNAME/NAMA-REPO.git
git push -u origin main
```

Lalu di [vercel.com](https://vercel.com):

1. **Add New… > Project** → import repo GitHub tadi.
2. Framework preset: biarkan **Other** (situs ini statis, tidak perlu build step).
3. Root Directory: biarkan default (folder yang berisi `index.html`).
4. Klik **Deploy**. Selesai — Vercel akan memberi URL publik (mis.
   `nama-app.vercel.app`).

---

## 3) Sambungkan frontend ke backend

Ada 2 cara — pilih salah satu (atau dua-duanya, env var jadi default, tombol ⚙️ tetap bisa dipakai untuk override per-browser):

### Cara A — otomatis lewat Environment Variables (disarankan)

Supaya semua pengunjung situs otomatis tersambung tanpa isi URL manual:

1. Di dashboard Vercel → project kamu → **Settings > Environment Variables**.
2. Tambahkan:
   - `API_URL` = URL Web App `/exec` dari langkah 1.
   - `API_TOKEN` = (opsional) sama persis dengan `APP_TOKEN` di Script Properties.
3. **Redeploy** (Deployments → titik tiga → Redeploy) supaya env var terpakai.

Frontend akan memanggil `/api/config` (serverless function di `api/config.js`,
otomatis terdeteksi Vercel) untuk mengambil nilai ini saat pertama kali
dibuka. Kalau berhasil, dialog sambungan tidak akan muncul sama sekali.

> ⚠️ Ini **bukan** cara menyembunyikan token dari pengguna situs — siapa pun
> yang membuka situs tetap bisa melihat nilainya lewat tab Network browser.
> Manfaatnya: tidak perlu ditulis di kode/GitHub, dan tidak perlu diisi ulang
> di tiap browser. Keamanan sesungguhnya tetap dari `APP_TOKEN` di Apps Script.

### Cara B — manual lewat dialog di UI

1. Buka situs Vercel kamu. Kalau belum ada sambungan (env var atau
   localStorage), dialog **"Sambungan ke Google Sheets"** muncul otomatis.
2. Isi **URL Web App** dan **Token**, klik **Simpan & tes koneksi**.

Pengaturan ini disimpan di `localStorage` browser (per-perangkat/per-browser)
dan akan dipakai duluan daripada env var. Untuk mengubahnya lagi kapan saja,
klik ikon ⚙️ di pojok kanan atas header.

---

## Catatan keamanan

- Web App dengan akses **"Anyone"** artinya siapa pun yang tahu URL-nya bisa
  memanggil API tersebut — karena itu **wajib** isi `APP_TOKEN` untuk pemakaian
  nyata (bukan sekadar coba-coba).
- Token dikirim di dalam body request (bukan header), dan disimpan di
  `localStorage` browser masing-masing — jangan pakai token yang sama dengan
  password akun manapun.
- Request dikirim dengan `Content-Type: text/plain` (bukan `application/json`)
  supaya browser tidak melakukan CORS preflight — Apps Script Web App memang
  tidak mendukung preflight `OPTIONS`. Ini normal, bukan bug.

## Pengembangan lokal

Karena situsnya statis murni, cukup buka `index.html` langsung di browser, atau:

```bash
npx serve .
```

Lalu sambungkan ke Web App Apps Script yang sama seperti di atas.
