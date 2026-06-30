# 🧠 AI Verbatim Coder

Aplikasi web untuk otomatisasi coding verbatim riset pasar menggunakan Gemini AI.

---

## 🚀 Cara Deploy di Streamlit Cloud (Gratis)

### 1. Persiapan
- Buat akun GitHub di https://github.com (gratis)
- Buat akun Streamlit di https://streamlit.io (gratis, login pakai GitHub)

### 2. Upload file ke GitHub
Buat repository baru di GitHub, lalu upload file-file berikut:
```
app.py
requirements.txt
memory_bank.json
```

### 3. Deploy di Streamlit Cloud
1. Login ke https://share.streamlit.io
2. Klik "New app"
3. Pilih repository GitHub kamu
4. Main file path: `app.py`
5. Klik "Deploy!" → tunggu 2-3 menit
6. Dapat link publik → share ke tim!

---

## 📖 Cara Pakai Aplikasi

### Step 1 – Upload
- Upload file Excel berisi kolom verbatim
- Isi konteks pertanyaan (misal: "D2. Apa alasan Anda memilih kendaraan ini?")
- Pilih bahasa output (Bilingual / ID saja / EN saja)

### Step 2 – Pilih Kolom
- Pilih kolom verbatim yang akan di-coding
- Bisa pilih beberapa kolom sekaligus untuk share 1 codeframe
- Beri nama Kategori/Nett untuk setiap grup kolom

### Step 3 – Review Codeframe
- AI akan generate codeframe berdasarkan verbatim + Memory Bank
- **Baca dan review** codeframe yang muncul
- Edit kode/label/nett jika ada yang perlu diperbaiki
- Klik "Setujui & Lanjut" jika sudah oke

### Step 4 – Autocode
- AI akan assign kode ke setiap baris verbatim
- Multi-code dipisah dengan ";" (contoh: 101;165)
- Progress bar menunjukkan proses per kolom per batch

### Step 5 – Download
- Download Excel berisi 2 sheet:
  - **Codeframe** – daftar kode lengkap dengan Nett & label
  - **Rawdata** – data asli + kolom Code hasil AI (kolom hijau)
- Opsional: simpan studi ke Memory Bank agar AI makin pintar

---

## 🔑 Cara Dapat Gemini API Key (Gratis)

1. Buka https://aistudio.google.com
2. Login dengan akun Google
3. Klik "Get API Key" → "Create API key"
4. Copy key (format: AIza...)
5. Paste di sidebar aplikasi saat pakai

**Limit free tier:** 15 request/menit, 1500 request/hari
→ Cukup untuk 500-1000 verbatim per hari per user

---

## 📚 Memory Bank

Memory Bank (`memory_bank.json`) menyimpan contoh studi sebelumnya.
Semakin banyak studi tersimpan → AI makin akurat dan konsisten.

Cara tambah studi baru ke Memory Bank:
- Via sidebar: upload Excel hasil coding + isi info studi
- Via Step 5 (Done): setelah autocode, langsung save ke Memory Bank

---

## 📁 Struktur File

```
app.py              ← Aplikasi utama Streamlit
requirements.txt    ← Library yang dibutuhkan
memory_bank.json    ← Database contoh studi (makin lama makin kaya)
README.md           ← Panduan ini
```
