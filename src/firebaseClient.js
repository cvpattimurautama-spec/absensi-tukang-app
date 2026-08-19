import { initializeApp } from 'firebase/app';
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

const looksPlaceholder = (v) => !v || v.includes('xxxxx') || v.includes('your-');

if (looksPlaceholder(firebaseConfig.apiKey) || looksPlaceholder(firebaseConfig.projectId)) {
  throw new Error(
    'File .env belum diisi dengan benar.\n\n' +
    'Langkah perbaikan:\n' +
    '1. Pastikan ada file bernama ".env" (bukan ".env.example") di folder utama proyek ini.\n' +
    '2. Isi baris-barisnya dengan kredensial Firebase asli kamu (dari Firebase Console -> Project Settings -> General -> "Your apps" -> SDK setup and configuration).\n' +
    '3. Setelah .env diisi/diperbaiki, HENTIKAN server (Ctrl+C di terminal) lalu jalankan ulang "npm run dev".\n' +
    '   (Vite tidak otomatis membaca ulang file .env selagi server jalan.)\n\n' +
    'Nilai VITE_FIREBASE_API_KEY saat ini terbaca: ' + JSON.stringify(firebaseConfig.apiKey) + '\n' +
    'Nilai VITE_FIREBASE_PROJECT_ID saat ini terbaca: ' + JSON.stringify(firebaseConfig.projectId)
  );
}

const app = initializeApp(firebaseConfig);

// Offline persistence: data tersimpan di cache lokal (IndexedDB) browser/komputer ini,
// jadi aplikasi tetap bisa dibuka & dipakai walau internet mati. Perubahan yang dibuat
// saat offline otomatis dikirim ke server begitu koneksi internet kembali.
// persistentMultipleTabManager supaya tetap jalan kalau dibuka di beberapa tab sekaligus.
let db;
try {
  db = initializeFirestore(app, {
    localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }),
  });
} catch (err) {
  // Browser lama yang tidak dukung IndexedDB, atau storage penuh -> tetap jalan online biasa
  console.warn('Mode offline tidak bisa diaktifkan, aplikasi tetap jalan online biasa:', err?.message || err);
  db = initializeFirestore(app, {});
}

export { db };
