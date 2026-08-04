# Absensi Tukang — Versi Web Multi-User (Firebase)

Versi ini React + **Firebase (Firestore)**. Kelebihan dibanding versi Supabase
sebelumnya: **tidak perlu menjalankan skema SQL sama sekali** — Firestore
tidak memakai tabel/SQL, koleksi datanya otomatis terbentuk begitu aplikasi
pertama kali menyimpan data. User admin pertama (`admin` / `admin123`) juga
**dibuat otomatis** saat aplikasi pertama kali dijalankan — tidak perlu
langkah manual apa pun untuk itu.

## 1. Buat project Firebase

1. Buka **https://console.firebase.google.com**, login pakai akun Google.
2. Klik **"Add project" / "Tambahkan project"**.
3. Isi nama project (bebas, misal `absensi-tukang`), lanjutkan (boleh
   matikan Google Analytics kalau tidak perlu), klik **"Create project"**.

## 2. Aktifkan Firestore

1. Di sidebar kiri, klik **"Build" → "Firestore Database"**.
2. Klik **"Create database"**.
3. Pilih **"Start in production mode"**, lanjut, pilih lokasi server
   terdekat (misal `asia-southeast2` untuk Jakarta), klik **"Enable"**.

## 3. Pasang aturan keamanan

1. Masih di halaman Firestore Database, klik tab **"Rules"**.
2. Hapus semua isi kotak kode yang ada.
3. Buka file `firebase/firestore.rules` dari paket ini, **copy semua
   isinya**, **paste** ke kotak tadi.
4. Klik **"Publish"**.

## 4. Ambil kredensial aplikasi

1. Di sidebar kiri, klik ikon gerigi → **"Project settings"**.
2. Scroll ke bagian **"Your apps"**, klik ikon **"</>"  (Web)**.
3. Isi nama app (bebas, misal `absensi-tukang-web`), klik **"Register app"**
   (tidak perlu centang "Firebase Hosting").
4. Akan muncul kode berisi `firebaseConfig = { apiKey: ..., authDomain: ..., ... }`
   — catat semua nilainya, akan dipakai di langkah berikutnya.

## 5. Konfigurasi aplikasi

1. Salin file `.env.example` menjadi `.env`.
2. Isi tiap barisnya dengan nilai dari Langkah 4:
   ```
   VITE_FIREBASE_API_KEY=...
   VITE_FIREBASE_AUTH_DOMAIN=...
   VITE_FIREBASE_PROJECT_ID=...
   VITE_FIREBASE_STORAGE_BUCKET=...
   VITE_FIREBASE_MESSAGING_SENDER_ID=...
   VITE_FIREBASE_APP_ID=...
   ```

## 6. Jalankan

```bash
npm install
npm run dev
```

Buka alamat yang muncul (biasanya `http://localhost:5173`). Login pertama:
`admin` / `admin123` — **segera ganti lewat menu "User"** setelah masuk.

> Kalau menjalankan ini di HP Android lewat Termux: pastikan folder proyek
> ada di dalam folder home Termux sendiri (`~/nama-folder`), **bukan** di
> `/storage/emulated/0/Download/...` — folder penyimpanan bersama Android
> tidak mendukung symlink yang dibutuhkan npm.

## 7. Deploy ke internet (supaya bisa diakses staff)

Sama seperti sebelumnya — lewat **Vercel** atau **Netlify**:
1. Upload folder ini ke repository GitHub.
2. Import di vercel.com / netlify.com.
3. Isi Environment Variables dengan 6 baris yang sama seperti isi `.env`.
4. Deploy.

## 8. Kelola User (multi-admin)

Login sebagai `admin`, tap tombol **"User"** di pojok kanan atas. Dari situ
bisa tambah user baru, ganti password, atau hapus user. Semua user yang
terdaftar punya akses penuh yang sama.

## Catatan Keamanan (penting dibaca)

- Password disimpan dalam bentuk **hash** (bcrypt), tidak pernah dalam
  bentuk teks asli, baik di database maupun saat dikirim antar perangkat.
- Namun, sama seperti versi Supabase sebelumnya: karena tidak memakai
  Firebase Authentication (memakai sistem login sendiri supaya bisa
  dikelola langsung dari dalam aplikasi), **aturan keamanan Firestore
  tidak bisa membedakan siapa yang login**. Siapa pun yang tahu kredensial
  Firebase project kamu (yang memang ada di kode frontend) bisa
  membaca/menulis data langsung lewat API Firebase.
- **Cocok untuk tim kecil yang saling percaya (internal kantor/proyek),
  bukan untuk aplikasi publik.**

## Struktur Data

Firestore menyimpan tiap jenis data sebagai "koleksi" dokumen (mirip folder
berisi file JSON), otomatis terbentuk saat pertama kali dipakai:
`workers`, `projects`, `weeks`, `evidence`, `payments`, `kasbon`,
`kasbonPayments`, `materials`, `suppliers`, `purchases`, `usage`,
`utangPayments`, `peralatan`, `peralatanUsage`, `companyProfile`, `users`.

Tidak ada skema yang perlu dijalankan manual — cukup pastikan Rules di
Langkah 3 sudah terpasang.
