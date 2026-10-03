# 🚀 Panduan Deployment Aplikasi CERDAS ke Vercel (vercel.app)

Dokumen ini berisi panduan langkah demi langkah untuk mengonlinekan aplikasi **CERDAS (Cerita Digital Anak Sempatik)** ke Vercel secara gratis dan cepat.

---

## 🛠️ Konfigurasi yang Telah Disiapkan

Aplikasi ini sudah dilengkapi dengan konfigurasi otomatis untuk Vercel:
1. `vercel.json` — Mengatur *Single Page Application (SPA) routing* dan mengarahkan endpoint API serverless `/api/*`.
2. `api/index.ts` — Menjalankan Serverless Function Express untuk fitur Gemini AI (`/api/gemini/analyze`, `/api/gemini/transcribe`, dll).
3. `src/firebase.ts` — Membaca konfigurasi dari `firebase-applet-config.json` atau Environment Variables.

---

## 📝 Langkah 1: Hubungkan Repositori ke Vercel

### Opsi A: Menggunakan GitHub (Sangat Direkomendasikan)
1. Unggah (*push*) kode aplikasi ini ke repositori **GitHub** Anda.
2. Buka [Vercel Dashboard](https://vercel.com/dashboard) dan login.
3. Klik tombol **"Add New..."** -> **"Project"**.
4. Pilih repositori GitHub aplikasi CERDAS Anda, lalu klik **Import**.

### Opsi B: Menggunakan Vercel CLI (Via Terminal)
1. Install Vercel CLI jika belum:
   ```bash
   npm i -g vercel
   ```
2. Jalankan perintah deploy di folder proyek:
   ```bash
   vercel
   ```

---

## 🔑 Langkah 2: Pengaturan Environment Variables di Vercel

Pada menu **Environment Variables** di dashboard Vercel proyek Anda, tambahkan variabel berikut:

| Key / Nama Variabel | Nilai / Value | Keterangan |
| :--- | :--- | :--- |
| `GEMINI_API_KEY` | `AIzaSy...` (Kunci API Anda) | **Wajib**. Dapatkan dari [Google AI Studio](https://aistudio.google.com/). |

*(Opsional)* Jika Anda ingin mengganti/menyesuaikan proyek Firebase:
* `VITE_FIREBASE_PROJECT_ID`
* `VITE_FIREBASE_APP_ID`
* `VITE_FIREBASE_API_KEY`
* `VITE_FIREBASE_AUTH_DOMAIN`
* `VITE_FIREBASE_DATABASE_ID`

---

## 🔐 Langkah 3: Tambahkan Domain Vercel ke Firebase Authentication

Agar login Google dan Firebase Auth berfungsi di domain Vercel Anda:
1. Buka [Firebase Console](https://console.firebase.google.com/).
2. Pilih proyek Anda: `gen-lang-client-0131415670`.
3. Buka menu **Build** -> **Authentication** -> Tab **Settings**.
4. Scroll ke bagian **Authorized domains**.
5. Klik **Add domain**, lalu masukkan nama domain Vercel Anda (contoh: `cerdas-app.vercel.app`).
6. Klik **Save**.

---

## ✅ Langkah 4: Selesai!

Aplikasi CERDAS Anda sekarang sudah aktif dan dapat diakses publik di URL `.vercel.app` Anda! 🎉
