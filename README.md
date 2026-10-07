# 📧 Automated Job Application Email Tracker

Sebuah skrip otomatis berbasis Python yang membaca kotak masuk Gmail Anda, mencari email terkait lamaran kerja, lalu mengklasifikasikannya secara cerdas menggunakan AI (Google Gemini). Hasil akhirnya akan diekspor rapi ke dalam sebuah file Excel (beserta pembaruan status lamaran secara otomatis).

## ✨ Fitur Utama
- **Klasifikasi Cerdas:** Menggunakan Gemini AI untuk membaca isi pesan dan menentukan apakah email tersebut adalah "Panggilan Interview" (ON PROGRESS), "Penolakan" (DITOLAK), atau "Offering/Diterima" (DITERIMA).
- **Anti-Duplikat (Smart Update):** Skrip mengenali perusahaan dan posisi yang dilamar. Jika Anda mendapat email penolakan di hari berikutnya, status "ON PROGRESS" Anda di Excel akan otomatis tertimpa menjadi "DITOLAK".
- **Ramah Rate Limit:** Dibekali dengan sistem antrean otomatis dan delay cerdas, sehingga sangat aman digunakan tanpa memicu error pemblokiran dari server Google.

## ⚙️ Persyaratan (Prerequisites)
Sebelum menjalankan skrip ini, pastikan Anda telah memiliki:
1. **Python 3.8+** terinstal di komputer.
2. **Kredensial Gmail API** (`credentials.json`) dari Google Cloud Console.
3. **API Key Google Gemini** dari [Google AI Studio](https://aistudio.google.com/).

## 🚀 Cara Instalasi

**1. Clone Repositori Ini**
Buka terminal Anda dan ketik:
```bash
git clone https://github.com/cescrafli/Job_Email_Checker.git
cd Job_Email_Checker
```

**2. Instal Pustaka Python**
Pastikan Anda menginstal semua dependensi yang diperlukan:
```bash
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib pandas openpyxl google-genai python-dotenv
```

**3. Atur Kredensial (Sangat Penting)**
Demi keamanan, Anda harus membuatnya sendiri di dalam folder proyek ini:
*   Pindahkan file kredensial Gmail OAuth Anda ke folder ini dan beri nama **`credentials.json`**.
*   Buat sebuah file baru bernama **`.env`** dan tempelkan API Key Gemini Anda dengan format berikut:
    ```env
    GEMINI_API_KEY="AIzaSy_API_KEY_ANDA_DISINI"
    ```

## 🏃 Cara Menjalankan

Cukup jalankan file python dari terminal:
```bash
python track_email.py
```
> **Catatan:** Saat pertama kali dijalankan, jendela browser Anda akan terbuka untuk meminta izin akses membaca email dari akun Google Anda. Setelah Anda setuju, file `token.pickle` akan otomatis terbuat agar Anda tidak perlu login ulang di kemudian hari.

## 🤖 Otomatisasi Harian (Windows)
Jika Anda menggunakan Windows, Anda bisa menggunakan file `jalankan_pelacak.bat` dan menyambungkannya ke **Windows Task Scheduler** agar pengecekan email ini bisa berjalan secara otomatis setiap hari tanpa harus menyentuh terminal sama sekali!

---
*Dibuat untuk mempermudah para jobseeker memonitor ribuan lamaran mereka tanpa pusing mencatat manual. Happy hunting!* 💼🚀
