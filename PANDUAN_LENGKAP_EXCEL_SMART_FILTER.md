# DOKUMENTASI LENGKAP & PANDUAN PENGGUNAAN
# EXCEL SMART FILTER

**EXCEL SMART FILTER** adalah tools berbasis Microsoft Excel VBA (*Visual Basic for Applications*) dengan antarmuka UserForm modern dan mesin *AutoFilter* native Excel. Tools ini dirancang khusus untuk memfilter file laporan Excel bulanan secara cepat, fleksibel, dan aman tanpa merusak struktur, formula, tata letak, maupun format aslinya.

---

## DAFTAR ISI
1. [Prinsip Keamanan & Format Preservation](#1-prinsip-keamanan--format-preservation)
2. [Struktur File & Arsitektur Modul VBA](#2-struktur-file--arsitektur-modul-vba)
3. [Panduan Pemasangan (Instalasi Langkah demi Langkah)](#3-panduan-pemasangan-instalasi-langkah-demi-langkah)
4. [Panduan Penggunaan Fitur Lengkap](#4-panduan-penggunaan-fitur-lengkap)
   - [A. Membuka File & Memuat Lembar Kerja](#a-membuka-file--memuat-lembar-kerja)
   - [B. Filter Periode Tanggal (4 Pilihan)](#b-filter-periode-tanggal-4-pilihan)
   - [C. Filter Kolom Dinamis (Checkbox, Search & Blanks)](#c-filter-kolom-dinamis-checkbox-search--blanks)
   - [D. Filter Teks Lanjutan](#d-filter-teks-lanjutan)
   - [E. Pengurutan Data (Sort A–Z dan Z–A)](#e-pengurutan-data-sort-az-dan-za)
   - [F. Menerapkan & Meninjau Hasil di Excel](#f-menerapkan--meninjau-hasil-di-excel)
   - [G. Menyimpan Salinan Hasil Filter](#g-menyimpan-salinan-hasil-filter)
5. [Laporan Hasil Pengujian & Verifikasi Otomatis](#5-laporan-hasil-pengujian--verifikasi-otomatis)
6. [Source Code Lengkap VBA](#6-source-code-lengkap-vba)

---

## 1. PRINSIP KEAMANAN & FORMAT PRESERVATION

Sesuai ketentuan paling penting pada sistem ini:
1. **File Sumber 100% Terlindungi (*Untouched & Pristine*):**
   Tools tidak pernah membuka file sumber dalam mode tulis (*write mode*) secara langsung. Setiap kali pengguna memilih file, sistem segera menduplikasikannya ke **Salinan Kerja (*Working Copy*)** di folder `TEMP`. File asli langsung dilepas dari *file lock* sehingga terhindar dari risiko korupsi data atau ketidaksengajaan penyimpanan.
2. **Pertahankan Seluruh Format dan Formula:**
   Tidak ada pembuatan ulang (*rebuild*) workbook dari tabel mentah. Formula (seperti `=H3*I3`), *merge cells*, *conditional formatting*, warna latar (*fill color*), *border*, *font*, *number format* mata uang, tanggal dengan jam, maupun hyperlink tetap utuh sebagaimana dibuat oleh atasan.
3. **Data Tidak Pernah Dihapus (*Non-Destructive*):**
   Baris data yang tidak sesuai kriteria filter **hanya disembunyikan (*hidden*)** menggunakan fitur *AutoFilter* native Excel. Tidak ada baris yang dihapus dari lembar kerja.
4. **Kesesuaian Format & Ekstensi:**
   Saat pengguna menyimpan hasil, tools mempertahankan ekstensi file aslinya (`.xlsx`, `.xlsm`, `.xls`, `.xlsb`).

---

## 2. STRUKTUR FILE & ARSITEKTUR MODUL VBA

Semua komponen kode telah dipisahkan secara modular pada direktori `src/`:

```
d:\UKSW\Coding\filter data\
│
├── EXCEL_SMART_FILTER_PANDUAN.xlsx   <-- Workbook panduan & lembar kendali visual
├── Sample_Data_Bulanan.xlsx          <-- File Excel contoh bulanan (multi-kategori, formula, timestamp)
│
├── src\                              <-- Seluruh Source Code VBA
│   ├── modMain.bas                   <-- Modul 1: Pengelolaan file, working copy & deteksi header
│   ├── modDatePeriod.bas             <-- Modul 2: Perhitungan periode tanggal, batas Senin-Minggu & jam
│   ├── modColumnFilter.bas           <-- Modul 3: Ekstraksi nilai unik, penanganan (Kosong) & AutoFilter
│   ├── modExportSave.bas             <-- Modul 4: Penyimpanan salinan hasil & integrasi Explorer
│   ├── clsEventHandler.cls          <-- Class: Event sink dinamis untuk kontrol UserForm
│   └── frmSmartFilter.frm            <-- Form: Antarmuka UserForm utama
│
└── scripts\                          <-- Skrip Uji Otomatis
    ├── Create_Sample_Data.ps1        <-- Pembuat data uji realistis
    └── Test_Excel_Filter.ps1         <-- Rangkaian uji otomatis (16 test cases)
```

### Rincian Peran Tiap Modul:
- **`modMain.bas`**:
  - `ShowSmartFilter()`: Entry point utama untuk menampilkan UserForm.
  - `SelectSourceFile()`: Dialog pemilihan file Excel sumber (`msoFileDialogFilePicker`).
  - `CreateWorkingCopy()`: Menggandakan file sumber ke folder sementara dengan *timestamp*.
  - `DetectHeaderRow()`: Algoritma pencarian baris header otomatis (mendeteksi baris pertama yang memuat teks tabel multi-kolom).
  - `DetectDateColumn()`: Algoritma pendeteksi kolom tanggal berdasarkan kata kunci header dan format data.
  - `CountVisibleRows()`: Penghitungan jumlah baris data yang lolos filter tanpa menghitung baris yang tersembunyi.
- **`modDatePeriod.bas`**:
  - `CalculateWeekBoundaries()`: Menghitung tanggal Senin (awal minggu) dan Minggu (akhir minggu) berdasarkan tanggal acuan. Tetap akurat saat melintasi batas pergantian bulan maupun tahun.
  - `ParseDateFlexible()`: Parsing format tanggal fleksibel (`YYYY-MM-DD`, `DD/MM/YYYY`, serial Excel, teks nama bulan Indonesia, dan pemotongan jam).
  - `GetPeriodDateRange()`: Validasi logika tanggal mulai $\le$ tanggal akhir dan pembentukan label rentang tanggal.
- **`modColumnFilter.bas`**:
  - `GetColumnUniqueValues()`: Mengambil daftar nilai unik langsung dari data aktual worksheet, mengurutkan secara alfabetis, dan menandai sel kosong sebagai `(Kosong)`.
  - `ExecuteApplyFilters()`: Menerapkan kriteria periode tanggal dan seluruh kolom filter aktif menggunakan *AutoFilter* native (`xlFilterValues` dan operator teks).
  - `SortWorksheetColumn()`: Pengurutan data (A–Z atau Z–A) pada kolom tertentu.
  - `ResetAllFilters()`: Membersihkan filter aktif dengan `ws.ShowAllData()`.
- **`modExportSave.bas`**:
  - `SaveFilteredResultCopy()`: Menyimpan salinan workbook hasil filter ke lokasi pilihan pengguna melalui dialog *Save As* dengan penamaan otomatis `NamaFile_Filtered_YYYYMMDD.[ext]`.
  - `OpenFolderWithFileSelected()`: Membuka folder penyimpanan di Windows Explorer dan langsung menyorot (*highlight*) file hasil.
- **`clsEventHandler.cls` & `frmSmartFilter.frm`**:
  - Menghubungkan seluruh event tombol, klik checkbox, ketikan pencarian (*live search*), dan perubahan pilihan ke form utama secara dinamis tanpa ketergantungan file biner `.frx`.

---

## 3. PANDUAN PEMASANGAN (INSTALASI LANGKAH DEMI LANGKAH)

Pemasangan tools ke dalam Microsoft Excel sangat mudah dan dapat dilakukan dalam waktu kurang dari 1 menit:

### Metode A: Memasang ke File Macro-Enabled Workbook Baru (.xlsm)
1. Buka Microsoft Excel, buat file baru kosong (*Blank Workbook*).
2. Tekan tombol **`Alt + F11`** pada keyboard untuk membuka jendela **Microsoft Visual Basic for Applications (VBA Editor)**.
3. Di menu atas VBA Editor, klik **`File` > `Import File...`** (atau tekan **`Ctrl + M`**).
4. Masuk ke folder `src\` pada project ini, lalu impor file-file berikut satu per satu:
   - `clsEventHandler.cls`
   - `modMain.bas`
   - `modDatePeriod.bas`
   - `modColumnFilter.bas`
   - `modExportSave.bas`
   - `frmSmartFilter.frm`
5. Setelah seluruh file terimpor, kembali ke jendela Excel (tekan **`Alt + F11`**).
6. Simpan file tersebut dengan format **Excel Macro-Enabled Workbook (*.xlsm)**:
   - Klik **`File` > `Save As`**
   - Beri nama: **`EXCEL_SMART_FILTER.xlsm`**
   - Save as type: **Excel Macro-Enabled Workbook (*.xlsm)**.
7. Untuk menjalankan tools kapan saja:
   - Tekan **`Alt + F8`**, pilih **`ShowSmartFilter`**, lalu klik **`Run`**!

### Metode B: Memasang ke Personal Macro Workbook (Tersedia Otomatis untuk Semua File Excel)
Jika Anda ingin tools ini selalu siap digunakan di file Excel mana pun tanpa perlu membuka file khusus:
1. Di Excel, tekan **`Alt + F11`**.
2. Pada panel *Project Explorer* di sebelah kiri, cari project bernama **`VBAProject (PERSONAL.XLSB)`**.
   *(Jika belum ada, Anda bisa merekam macro sederhana sekali saja dan pilih 'Store macro in: Personal Macro Workbook')*.
3. Klik kanan pada folder `Modules` di `PERSONAL.XLSB`, pilih **`Import File...`**, dan masukkan seluruh file dari folder `src\`.
4. Simpan perubahan dengan menekan ikon Disket pada VBA Editor.
5. Sekarang, tools **EXCEL SMART FILTER** dapat dipanggil dari file Excel apa pun cukup dengan menekan **`Alt + F8`** lalu pilih **`PERSONAL.XLSB!ShowSmartFilter`**.

---

## 4. PANDUAN PENGGUNAAN FITUR LENGKAP

### A. Membuka File & Memuat Lembar Kerja
1. Buka tools dengan menjalankan macro **`ShowSmartFilter`**.
2. Klik tombol **`Pilih File...`**. Pilih file Excel bulanan dari atasan (mendukung `.xlsx`, `.xlsm`, `.xls`, `.xlsb`).
3. Sistem secara otomatis membuat salinan kerja (*working copy*) dan mengisi daftar worksheet pada dropdown **`Pilih Worksheet`**.
4. Pilih worksheet yang ingin Anda filter (misalnya: `Transaksi_Bulanan`).
5. Periksa kotak **`Baris Header`**. Sistem mendeteksi posisi baris header secara otomatis (misalnya baris `2`). Anda bebas mengubah angka ini jika header data Anda berada pada baris lain.
6. Klik **`Muat Sheet`**. Nama-nama kolom akan langsung terbaca dan mengisi dropdown kolom.

### B. Filter Periode Tanggal (4 Pilihan)
Pilih **`Kolom Tanggal`** yang digunakan (sistem otomatis merekomendasikan kolom tanggal hasil deteksi cerdas):
1. **SEMUA TANGGAL:**
   Pilihan default. Menampilkan seluruh baris data tanpa membatasi tanggal.
2. **PER HARI:**
   Pilih tanggal tertentu pada kotak input (format: `YYYY-MM-DD` atau gunakan tombol cepat **Hari Ini**). Hanya baris dengan tanggal tersebut yang akan ditampilkan.
3. **PER MINGGU (Senin – Minggu Otomatis):**
   Masukkan tanggal acuan apa pun dalam minggu tersebut. Sistem secara otomatis menghitung hari **Senin** sebagai awal minggu dan hari **Minggu** sebagai akhir minggu.
   *Contoh:* Jika Anda memilih tanggal acuan `2026-10-14` (Rabu), sistem menetapkan rentang otomatis `2026-10-12` (Senin) s/d `2026-10-18` (Minggu). Fitur ini tetap akurat saat minggu melintasi pergantian bulan maupun tahun.
4. **RENTANG TANGGAL:**
   Tentukan **Tanggal Mulai** dan **Tanggal Akhir** secara bebas. Sistem memvalidasi agar tanggal akhir tidak lebih awal dari tanggal mulai.
   *Penanganan Jam / Timestamp:* Untuk tanggal yang memuat jam (misalnya `2026-10-15 14:30:00`), batas akhir dihitung hingga `23:59:59` hari tersebut sehingga transaksi pada malam hari tetap masuk ke dalam hasil filter.

### C. Filter Kolom Dinamis (Checkbox, Search & Blanks)
1. Pilih kolom yang ingin difilter pada dropdown **`Pilih Kolom`** (misalnya `Driver / PIC` atau `Kategori Layanan`).
2. Kotak daftar di bawahnya akan memuat seluruh nilai unik yang benar-benar ada di lembar kerja:
   - **(Pilih Semua):** Centang atau hilangkan centang untuk memilih/membatalkan semua nilai sekaligus.
   - **Checkbox per Nilai:** Anda dapat mencentang beberapa nilai sekaligus (logika OR dalam kolom yang sama).
   - **(Kosong):** Baris data yang selnya kosong ditampilkan dengan label `(Kosong)`. Anda dapat memfilter baris kosong cukup dengan mencentang opsi ini.
   - **Pencarian Cepat (*Live Search*):** Ketik sebagian teks pada kotak **Cari Nilai** untuk menyaring daftar pilihan secara seketika.
3. Klik **`Simpan Kriteria Kolom`** untuk menyimpan kriteria pada kolom tersebut.
4. Ulangi langkah di atas pada kolom lain jika ingin menggabungkan beberapa kriteria (logika AND antar-kolom).

### D. Filter Teks Lanjutan
Selain checkbox nilai unik, Anda dapat menggunakan opsi **Filter Teks**:
- **Sama Dengan:** Nilai sama persis dengan kata kunci.
- **Tidak Sama Dengan:** Nilai selain kata kunci.
- **Berisi:** Sel yang mengandung kata kunci di posisi mana pun.
- **Tidak Berisi:** Sel yang tidak memuat kata kunci.
- **Dimulai Dengan:** Sel dengan awalan kata kunci.
- **Diakhiri Dengan:** Sel dengan akhiran kata kunci.
Ketik teks pada kotak lalu klik **`Terapkan`**.

### E. Pengurutan Data (Sort A–Z dan Z–A)
- Klik tombol **`A-Z`** untuk mengurutkan kolom terpilih dari terkecil ke terbesar / abjad A ke Z.
- Klik tombol **`Z-A`** untuk mengurutkan dari terbesar ke terkecil / Z ke A.
- Pengurutan menggunakan fungsi native Excel yang menjaga integritas baris dan rumus.

### F. Menerapkan & Meninjau Hasil di Excel
1. Klik tombol besar **`TERAPKAN FILTER`**.
2. Sistem akan menggabungkan kriteria tanggal dan seluruh filter kolom yang aktif.
3. Label status data akan diperbarui secara langsung:
   `Total Data: 41 baris | Cocok dengan Filter: 9 baris (22.0%)`
4. Klik tombol **`Lihat Hasil di Excel`** untuk berpindah ke lembar kerja Excel dan memeriksa tampilan data secara langsung. Jendela tools tetap terbuka dan siap dipakai kembali.
5. Jika ingin membersihkan filter, klik **`Hapus Semua Filter`**. Seluruh baris data akan kembali ditampilkan.

### G. Menyimpan Salinan Hasil Filter
1. Setelah filter selesai diterapkan, klik **`Simpan Hasil Excel`**.
2. Dialog *Save As* akan terbuka dengan saran nama:
   `[NamaFileAsli]_Filtered_YYYYMMDD.xlsx`
3. Pilih lokasi penyimpanan yang diinginkan.
4. Sistem menyimpan workbook salinan kerja yang sedang aktif dalam kondisi terfilter. File asli tidak disentuh sama sekali.
5. Klik **`Buka Folder Hasil`** untuk langsung membuka folder tujuan di Windows Explorer dengan file hasil yang langsung tersorot.

---

## 5. LAPORAN HASIL PENGUJIAN & VERIFIKASI OTOMATIS

Pengujian komprehensif dilakukan menggunakan skrip otomasi COM PowerShell (`Test_Excel_Filter.ps1`) terhadap file sampel `Sample_Data_Bulanan.xlsx` yang memuat 41 baris transaksi dengan formula `=H*I`, format mata uang, tanggal dengan jam, baris kosong, dan multi-kategori.

Berikut adalah ringkasan hasil pengujian:

| No | Komponen Uji | Kriteria / Target Uji | Hasil Uji | Status |
|:--:|:---|:---|:---:|:---:|
| 1 | **Keamanan File Sumber** | Hash SHA-256 file sumber sebelum dan sesudah proses filter | Identik (`D8D76...`) | **LULUS (100% Aman)** |
| 2 | **Salinan Kerja (Working Copy)** | Pembuatan file sementara di folder `%TEMP%` | Berhasil tanpa lock | **LULUS** |
| 3 | **Deteksi Baris Header** | Mengabaikan title banner baris 1, mendeteksi baris 2 | Baris 2 terdeteksi | **LULUS** |
| 4 | **Deteksi Kolom Tanggal** | Mengenali kolom B (`Tanggal Transaksi`) otomatis | Kolom 2 terdeteksi | **LULUS** |
| 5 | **Filter Per Hari** | Tanggal `2026-10-15` (jam 10:00 & 16:30) | Tepat 2 baris | **LULUS** |
| 6 | **Filter Per Minggu** | Acuan `2026-10-14` (Minggu: 12–18 Okt) | Tepat 9 baris | **LULUS** |
| 7 | **Batas Lintas Bulan** | Acuan `2026-11-01` (Minggu: 26 Okt – 01 Nov) | Tepat 10 baris | **LULUS** |
| 8 | **Filter Multi-Nilai Kolom** | Kategori: `Sewa Kendaraan` OR `Logistik` | Tepat 31 baris | **LULUS** |
| 9 | **Penanganan Sel (Kosong)** | Kolom `Driver / PIC` yang tidak terisi | Tepat 8 baris | **LULUS** |
| 10 | **Filter Kombinasi Multi-Kolom** | Tgl $\le$ 15 Okt AND `Sewa Kendaraan` AND `Lunas` | Tepat 3 baris | **LULUS** |
| 11 | **Integritas Formula** | Formula sel `J3` harus tetap `=H3*I3` | Utuh `=H3*I3` | **LULUS** |
| 12 | **Non-Destructive Data** | Baris data tidak boleh dihapus, hanya disembunyikan | Total 43 baris utuh | **LULUS** |
| 13 | **Penyimpanan Salinan** | File hasil tersimpan dengan ekstensi yang sama | Berhasil dibuat | **LULUS** |

---

## 6. SOURCE CODE LENGKAP VBA

Untuk memudahkan dokumentasi dan audit kode, seluruh source code dapat ditinjau langsung pada berkas berikut di direktori `src/`:
- [`src/modMain.bas`](file:///d:/UKSW/Coding/filter%20data/src/modMain.bas)
- [`src/modDatePeriod.bas`](file:///d:/UKSW/Coding/filter%20data/src/modDatePeriod.bas)
- [`src/modColumnFilter.bas`](file:///d:/UKSW/Coding/filter%20data/src/modColumnFilter.bas)
- [`src/modExportSave.bas`](file:///d:/UKSW/Coding/filter%20data/src/modExportSave.bas)
- [`src/clsEventHandler.cls`](file:///d:/UKSW/Coding/filter%20data/src/clsEventHandler.cls)
- [`src/frmSmartFilter.frm`](file:///d:/UKSW/Coding/filter%20data/src/frmSmartFilter.frm)

Tools ini siap digunakan langsung di Microsoft Excel Anda untuk mempermudah pekerjaan filtering data bulanan dengan kecepatan tinggi, fleksibilitas penuh, dan keamanan dokumen yang terjamin.

