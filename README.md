# Zotero Citation Automation (Universal Edition)

[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Platform Windows](https://img.shields.io/badge/Platform-Windows-green.svg)]()
[![Zotero 7+](https://img.shields.io/badge/Zotero-7%2B-red.svg)](https://www.zotero.org/)
[![Self-test 41/41](https://img.shields.io/badge/self--test-41%2F41-brightgreen.svg)]()

Tool otomatisasi untuk mengonversi **placeholder sitasi teks** (format apa pun: `NR007`, `S12`,
`REF-1`, `[1]`, `[1-5]`, `{A3}`, dll.) di dokumen Word (`.docx`) menjadi **field sitasi resmi
Zotero (`ADDIN ZOTERO_ITEM CSL_CITATION`)**.

Hasil konversi berupa dokumen baru **`manuscript - Zotero.docx`** yang langsung terhubung ke
plugin Zotero di Word, buku kerja Excel **`laporan_sitasi.xlsx`** yang memuat audit sitasi secara
lengkap, dan **`referensi_belum_ada.ris`** untuk referensi yang belum ada di library.

Dokumen Word asli Anda **aman dan tidak pernah ditimpa**.

---

## Daftar Isi

- [Latar Belakang](#latar-belakang)
- [Fitur Utama](#fitur-utama)
- [Cara Kerja Singkat](#cara-kerja-singkat)
- [Prasyarat Sistem](#prasyarat-sistem)
- [Instalasi](#instalasi)
- [Tutorial Lengkap Penggunaan](#tutorial-lengkap-penggunaan)
  - [Langkah 0: Siapkan Zotero & Local API](#langkah-0-siapkan-zotero--local-api)
  - [Langkah 1: Siapkan File Word dan Excel](#langkah-1-siapkan-file-word-dan-excel)
  - [Langkah 2: Jalankan Otomatisasi](#langkah-2-jalankan-otomatisasi)
  - [Langkah 3: Periksa Laporan Excel](#langkah-3-periksa-laporan-excel)
  - [Langkah 4: Tindak Lanjuti Referensi yang Kurang](#langkah-4-tindak-lanjuti-referensi-yang-kurang)
  - [Langkah 5: Finalisasi di Microsoft Word](#langkah-5-finalisasi-di-microsoft-word)
- [Format Excel yang Didukung](#format-excel-yang-didukung)
- [Penjelasan Laporan Excel](#penjelasan-laporan-excel-laporan_sitasixlsx)
- [Opsi Perintah CLI Lanjutan](#opsi-perintah-cli-lanjutan)
- [Mode Koneksi Library](#mode-koneksi-library)
- [Kalibrasi Properti ZOTERO_PREF_1](#kalibrasi-properti-zotero_pref_1)
- [Struktur Repositori](#struktur-repositori)
- [Self-Test](#self-test)
- [Troubleshooting & Solusi Masalah](#troubleshooting--solusi-masalah)
- [Perbedaan dengan Versi Mendeley](#perbedaan-dengan-versi-mendeley)
- [Roadmap](#roadmap)
- [Lisensi](#lisensi)

---

## Latar Belakang

Saat menulis naskah ilmiah, artikel jurnal, atau skripsi, sitasi sementara sering ditulis dalam
bentuk kode sederhana seperti `(NR007)`, `[S12]`, atau `[1, 2, 5]`, sementara data pustakanya
dicatat di spreadsheet Excel.

Plugin **Zotero** di Word tidak dapat langsung mengubah teks biasa tersebut menjadi sitasi aktif.
Menghubungkan satu per satu secara manual ke ratusan sitasi membutuhkan waktu berhari-hari dan
rawan salah.

**Tool ini menyelesaikan masalah tersebut dalam hitungan detik:**
Secara langsung pada level OpenXML (`word/document.xml`), tool menyisipkan struktur field Zotero
yang valid dengan **item key** dari database Zotero Anda, plus properti dokumen `ZOTERO_PREF_1`
yang membuat Word mengenali dokumen sebagai dokumen Zotero.

Versi ini adalah kembaran dari
[Mendeley Citation Automation](https://github.com/Nashhansyah/MendeleycAutomation) — alur
pemakaiannya sama, hanya target aplikasinya Zotero.

---

## Fitur Utama

- **Universal Placeholder Support** — mendeteksi pola apa pun: huruf+angka (`NR001`, `S1`,
  `REF-01`), angka murni (`[1]`, `(1)`), kurung bulat `()`, siku `[]`, maupun kurawal `{}`.
- **Auto Range Expansion** — `[1-5]` menjadi `1, 2, 3, 4, 5`; `(NR001-NR004)` menjadi empat kode.
- **Deteksi Cerdas Kolom Excel** — otomatis mengenali sheet, baris header, dan kolom ID, Judul,
  DOI, Penulis, Tahun, Jurnal tanpa mengubah format Excel.
- **Pencocokan Bertingkat ke Zotero**:
  1. Kecocokan persis nomor **DOI**
  2. Kecocokan **Judul Persis** (dinormalisasi)
  3. Kecocokan **Fuzzy Judul** (≥ 90%) dengan `difflib`
- **Pengayaan Metadata via Crossref** — volume, issue, dan halaman dilengkapi otomatis, dengan
  cache lokal (`crossref_cache.json`) agar hemat kuota.
- **Empat Sumber Library** — Local API (default), SQLite langsung, Web API, dan fixture JSON
  untuk pengujian tanpa Zotero.
- **Audit Komprehensif di Excel** — `laporan_sitasi.xlsx` 4 sheet dengan indikator warna status
  (Berhasil, Gagal Tidak Ada di Zotero, Gagal Tidak Ada di Excel).
- **Generator RIS Otomatis** — `referensi_belum_ada.ris` siap di-*drag & drop* ke Zotero.
- **Aman** — dokumen asli tidak pernah ditimpa; mode SQLite membaca lewat salinan temporer
  sehingga tidak bentrok saat Zotero terbuka.
- **Tahan Run Terpotong** — placeholder yang terpecah antar-`<w:r>` (mis. `NR0` + `01`) tetap
  terdeteksi karena teks dipetakan ulang dengan offset sebelum dipatch.

---

## Cara Kerja Singkat

```text
naskah.docx (placeholder)  +  referensi.xlsx
            ↓  6 tahap
    [1] baca teks & token Word (tahan run terpotong)
    [2] pilih sheet + petakan kolom Excel
    [3] deteksi pola placeholder dari kamus ID Excel
    [4] susun metadata + tambal dari Crossref
    [5] cocokkan ke library Zotero (DOI → Judul → Fuzzy)
    [6] sisipkan field Zotero + ZOTERO_PREF_1
            ↓
manuscript - Zotero.docx → tab Zotero → Refresh → sitasi jadi
```

Field yang ditulis ke dokumen:

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText>ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"kEVov4ts", ...}</w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>(Smith &amp; Doe, 2021)</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

Teks `(Smith & Doe, 2021)` hanyalah tampilan sementara. Setelah **Refresh**, Zotero menggantinya
dengan hasil format CSL yang sebenarnya (APA/IEEE/Nature/…), termasuk tahun-suffix otomatis
(`2023a`, `2023b`).

---

## Prasyarat Sistem

1. **Sistem Operasi**: Windows 10 / 11
2. **Python**: versi 3.8 ke atas (teruji pada 3.11)
3. **Microsoft Word** dengan plugin **Zotero** terpasang (tab **Zotero** muncul di Ribbon)
4. **Zotero Desktop** terpasang (versi 7 atau lebih baru) dan library sudah berisi referensi
5. **Local API Zotero** diaktifkan (lihat Langkah 0)

---

## Instalasi

```powershell
git clone https://github.com/Nashhansyah/Zetero-Automation.git
cd Zetero-Automation
pip install -r requirements.txt
```

Dependensi: `lxml`, `pandas`, `openpyxl`.

---

## Tutorial Lengkap Penggunaan

### Langkah 0: Siapkan Zotero & Local API

Tool membaca library Zotero lewat **Local API** — lebih cepat, tanpa API key, dan tidak mengunci
database.

1. Buka **Zotero**.
2. Masuk **Edit → Settings** (Windows) atau **Zotero → Settings**.
3. Pilih tab **Advanced**.
4. Centang **"Allow other applications on this computer to communicate with Zotero"**.
5. Uji dari terminal:

```powershell
Invoke-RestMethod "http://localhost:23119/api/users/0/items/top?format=json&limit=1"
```

Kalau muncul JSON berisi item, koneksi siap.

> **Zotero ditutup?** Tool otomatis jatuh ke mode `sqlite` (membaca `zotero.sqlite` lewat salinan
> temporer). Tidak perlu mengubah apa pun.

> **Library belum pernah disinkronkan?** URI item akan memakai bentuk
> `http://zotero.org/users/local/<libraryID>/items/<key>`. Bisa dipaksa lewat
> `--uri-mode local`.

---

### Langkah 1: Siapkan File Word dan Excel

1. Siapkan **1 file naskah Word (`.docx`)** yang berisi placeholder sitasi, misalnya:

   ```text
   ... metode ini terbukti efektif (NR007; NR012).
   ... menurut beberapa studi [1-3] ...
   ... hasil serupa dilaporkan {A3} ...
   ```

2. Siapkan **1 file Excel (`.xlsx`)** berisi daftar referensi. Minimal ada kolom kode/ID dan
   judul; kolom DOI/Penulis/Tahun/Jurnal sangat dianjurkan agar pencocokan akurat:

   | Study_ID | Judul | DOI | Penulis | Tahun | Jurnal |
   |---|---|---|---|---|---|
   | NR001 | Deep Learning for Medical Imaging | 10.1000/demo001 | Smith, John; Doe, Jane | 2021 | Nature Machine Intelligence |
   | NR002 | Transformer Models in NLP | 10.1000/demo002 | Lee, Min; Park, Soo | 2022 | Journal of AI |

3. Salin kedua file ke folder repositori ini.

> **Catatan:** nama file bebas. Script otomatis memilih file `.docx` dan `.xlsx` terbaru di
> folder. Kalau ada lebih dari satu, gunakan `--docx` dan `--workbook` untuk menunjuk file
> tertentu.

---

### Langkah 2: Jalankan Otomatisasi

**Cara 1 (Paling Mudah — Tanpa Terminal):**

- Double-click berkas **`Run.bat`**.

**Cara 2 (Melalui Terminal):**

```powershell
python zotero_cite.py
```

Tunggu beberapa detik hingga proses selesai:

```text
========================================================================
  WORD  ->  ZOTERO CITATION AUTOMATION (UNIVERSAL EDITION)
========================================================================
  Folder Kerja : C:\ZOTERO AUTOMATION
  Naskah Word  : naskah_penelitian.docx
  Data Excel   : data_referensi.xlsx

[1/6] Membaca naskah dokumen Word ...
      84 paragraf, 1 bagian dokumen
      126 token kandidat sitasi terdeteksi

[2/6] Menganalisis file Excel & mencocokkan kolom ...
      Sheet terpilih: 'Data Referensi' (baris header: 1)
      Kolom ID      : 'Study_ID' (42 nilai langsung cocok dengan dokumen)

[3/6] Mendeteksi pola placeholder sitasi di Word ...
      Pola Terpilih: prefix 'NR' + angka
      30 ID Excel terdeteksi disitasi di Word

[4/6] Menyusun metadata referensi ...
      Kolom terpetakan: id='Study_ID', doi='DOI', year='Tahun', ...
      42 referensi siap dipadankan

[5/6] Mencocokkan dengan library Zotero ...
      Sumber Library : localapi
      User ID        : 1234567
      Isi Library    : 318 dokumen | 290 ber-DOI | 315 ber-Judul
      Berhasil dicocokkan ke Zotero: 40 dari 42 referensi

[6/6] Menyisipkan field Zotero & menulis dokumen baru ...
      Properti dokumen : ZOTERO_PREF_1 (disuntikkan)

========================================================================
  HASIL EKSEKUSI:
  * 42 kelompok sitasi berhasil diubah menjadi field Zotero
  * 40 ID unik sitasi sukses terhubung ke Zotero
  * 2 ID belum ada di Zotero : NR015, NR028

  Berkas Hasil Word   : manuscript - Zotero.docx
  Berkas Laporan Excel: laporan_sitasi.xlsx
========================================================================
  Berkas RIS          : referensi_belum_ada.ris  (drag & drop langsung ke Zotero)
  Catatan Perbaikan   : PERLU_DIPERIKSA.txt
```

**Cara 3 (dari dalam Word, opsional):**

1. Di Word tekan `Alt+F11`.
2. **File → Import File** → pilih `WordMacro.bas`.
3. **Developer → Macros → `UpdateZoteroCitations`**.

Macro menyimpan dokumen aktif, menjalankan script, lalu membuka hasilnya sebagai dokumen baru.

---

### Langkah 3: Periksa Laporan Excel

Buka **`laporan_sitasi.xlsx`** yang baru dihasilkan.

- Periksa sheet **Ringkasan** untuk melihat persentase keberhasilan.
- Periksa sheet **Perlu Tindakan** jika ada sitasi yang berstatus `GAGAL`.
- Periksa sheet **Excel Belum Digunakan** untuk melihat referensi yang tidak pernah dikutip.

---

### Langkah 4: Tindak Lanjuti Referensi yang Kurang

**Jika ada referensi berstatus `GAGAL: TIDAK ADA DI ZOTERO`:**

1. Tool otomatis membuat **`referensi_belum_ada.ris`**.
2. Buka aplikasi **Zotero**.
3. Tarik (*drag & drop*) file `referensi_belum_ada.ris` langsung ke jendela Zotero.
4. Jalankan `Run.bat` lagi. Sitasi tersebut sekarang berstatus `BERHASIL`.

**Jika ada referensi berstatus `GAGAL: TIDAK ADA DI EXCEL`:**

Kode tersebut ada di Word tetapi belum terdaftar di Excel. Tambahkan barisnya, lalu jalankan ulang.

---

### Langkah 5: Finalisasi di Microsoft Word

1. Buka file hasil: **`manuscript - Zotero.docx`**.
2. Di Microsoft Word, buka tab **Zotero** (atau **References → Zotero**).
3. Di panel Zotero, pilih:
   - **Document Preferences** untuk memilih gaya sitasi (APA 7th, IEEE, Nature, Harvard, dll.)
   - **Refresh** untuk memperbarui seluruh sitasi di dokumen
4. Letakkan kursor di bagian akhir naskah (bab *References / Daftar Pustaka*), lalu klik
   **Add/Edit Bibliography**.
5. Hapus daftar pustaka lama yang sebelumnya diketik manual agar tidak ganda.

> **Penting:** langkah **Refresh** wajib. Tool hanya menulis payload mentah; format akhir
> ditentukan oleh CSL di sisi Zotero.

---

## Format Excel yang Didukung

Tool tidak mewajibkan nama kolom tertentu — ia menebak dari isi dan nama header, lalu
memverifikasi lewat tingkat kecocokan dengan token di Word.

| Peran | Contoh nama kolom yang dikenali |
|---|---|
| ID / kode | `Study_ID`, `ID`, `Kode`, `Ref_ID`, `No`, `Citation Key`, `NR` |
| Judul | `Judul`, `Title`, `Artikel`, `Paper`, `Document Title` |
| DOI | `DOI`, `Doi`, `Identifier`, `Link DOI` |
| Penulis | `Penulis`, `Author(s)`, `Creator`, `Author Full Names` |
| Tahun | `Tahun`, `Year`, `Date`, `Publication Year` |
| Jurnal | `Jurnal`, `Journal`, `Venue`, `Source`, `Container` |
| Volume / Issue / Halaman | `Volume`, `Issue`, `Pages`, `Pagination` |

Kalau penebakan gagal, paksa secara manual:

```powershell
python zotero_cite.py --id-column "Study_ID" --sheet "Data Referensi"
```

Format penulis yang didukung: `Smith, John; Doe, Jane`, `Smith J.; Doe A.`,
`John Smith, Jane Doe`, atau `Smith et al. (2021)`.

---

## Penjelasan Laporan Excel (`laporan_sitasi.xlsx`)

| Sheet | Penjelasan |
|---|---|
| **`Ringkasan`** | Dashboard metrik: total sitasi di Word, jumlah & persentase berhasil, jumlah gagal (tidak ada di Zotero vs tidak ada di Excel), berkas yang diproses, dan sumber library yang digunakan. |
| **`Detail Sitasi Word`** | Tabel audit lengkap per kode sitasi: ID, Status (`BERHASIL` hijau, `GAGAL: TIDAK ADA DI ZOTERO` oranye, `GAGAL: TIDAK ADA DI EXCEL` kuning), alasan, jumlah kemunculan, metode pencocokan, **item key Zotero**, metadata artikel, dan cuplikan kalimat paragraf di Word. |
| **`Perlu Tindakan`** | Filter khusus sitasi yang gagal beserta petunjuk penyelesaian langkah-demi-langkah. |
| **`Excel Belum Digunakan`** | Daftar referensi yang ada di file Excel tetapi tidak pernah dikutip di naskah Word. |

File pendamping lain:

| File | Isi |
|---|---|
| `mapping_report.csv` | rekap ringkas ID → status → item key |
| `referensi_belum_ada.ris` | referensi yang belum ada di library, siap diimpor |
| `PERLU_DIPERIKSA.txt` | panduan perbaikan + cuplikan paragraf yang masih berisi placeholder |

---

## Opsi Perintah CLI Lanjutan

```powershell
python zotero_cite.py [opsi]
```

| Opsi | Contoh Argumen | Keterangan |
|---|---|---|
| `--folder` | `--folder "D:\Riset"` | Folder tempat file Word & Excel berada |
| `--docx` | `--docx "draft.docx"` | Menunjuk file Word tertentu |
| `--workbook` | `--workbook "refs.xlsx"` | Menunjuk file Excel tertentu |
| `--out` | `--out "final_zotero.docx"` | Nama file Word hasil |
| `--style` | `--style numeric` | Gaya tampilan sementara (`author-year` / `numeric`) |
| `--id-regex` | `--id-regex "NR\d{3}"` | Memaksa pola regex ID |
| `--id-column` | `--id-column "Study_ID"` | Memaksa nama kolom ID |
| `--sheet` | `--sheet "Data Referensi"` | Memaksa nama sheet Excel |
| `--no-enrich` | *(flag)* | Menonaktifkan pengambilan metadata dari Crossref (offline) |
| `--dry-run` | *(flag)* | Hanya analisis & pemetaan, tanpa menulis file |
| `--adapter` | `--adapter sqlite` | `auto` / `localapi` / `sqlite` / `webapi` / `fixture` |
| `--fixture` | `--fixture samples\library_fixture.json` | File JSON library (mode uji) |
| `--zotero-dir` | `--zotero-dir "D:\Zotero"` | Folder data Zotero (berisi `zotero.sqlite`) |
| `--user-id` | `--user-id 1234567` | ID akun Zotero untuk URI field |
| `--uri-mode` | `--uri-mode local` | Bentuk URI: `auto` / `user` / `local` |
| `--library-id` | `--library-id 1` | ID library (default 1) |
| `--citation-style` | `--citation-style "http://www.zotero.org/styles/ieee"` | CSL style untuk `ZOTERO_PREF_1` |
| `--locale` | `--locale "id-ID"` | Locale CSL |
| `--make-bibl` | *(flag)* | Sekalian menulis field bibliography di akhir dokumen |
| `--no-prefs` | *(flag)* | Jangan suntikkan `ZOTERO_PREF_1` |

Contoh analisis tanpa menulis apa pun:

```powershell
python zotero_cite.py --dry-run
```

---

## Mode Koneksi Library

| Adapter | Kapan dipakai | Butuh |
|---|---|---|
| `localapi` **(default)** | Zotero berjalan. Cepat, tanpa auth, tanpa lock DB | Izin komunikasi di Settings → Advanced |
| `sqlite` | Zotero ditutup / Local API diblokir | `zotero.sqlite` (dibaca lewat salinan temporer) |
| `webapi` | Perlu library group / lintas mesin | API key + `--user-id` |
| `fixture` | Uji coba & pengembangan tanpa Zotero | file JSON |

`adapter: "auto"` mencoba `localapi` → `sqlite`. Adapter `fixture` **tidak** ikut dalam rantai
otomatis — ia hanya aktif kalau diminta eksplisit, supaya mesin tanpa Zotero tidak diam-diam
memakai library contoh.

---

## Kalibrasi Properti `ZOTERO_PREF_1`

Zotero menandai sebuah dokumen sebagai "dokumen Zotero" lewat properti `ZOTERO_PREF_1` di
`docProps/custom.xml`. Tanpa properti ini, tombol **Refresh** bisa menolak memproses dokumen.
Tool sudah menyuntikkan versi kandidat; kalibrasi memastikan formatnya sama persis dengan yang
dibuat Zotero.

1. Buat `sampel.docx` kosong, ketik satu kalimat, sisipkan **1 sitasi asli** lewat tab Zotero,
   lalu simpan.
2. Dump propertinya:

```powershell
python -c "import zipfile,pathlib; z=zipfile.ZipFile('sampel.docx'); pathlib.Path('dump_custom.xml').write_bytes(z.read('docProps/custom.xml')); pathlib.Path('dump_document.xml').write_bytes(z.read('word/document.xml')); print('ok')"
```

3. Bandingkan `dump_custom.xml` dengan output tool. Kalau strukturnya berbeda, sesuaikan fungsi
   `build_prefs()` di `zotero_field.py` — hanya satu fungsi.
4. Periksa juga URI di `dump_document.xml`: `users/<ID>/items/<KEY>` atau
   `users/local/<libraryID>/items/<KEY>`. Sesuaikan lewat `--uri-mode`.
5. Jalankan ulang tool, buka hasilnya di Word, klik **Refresh**. Kalau sitasi berubah mengikuti
   gaya CSL, kalibrasi selesai.

---

## Struktur Repositori

```text
Zetero-Automation/
├── zotero_cite.py        # Pipeline utama (6 tahap)
├── zotero_library.py     # Adapter localapi | sqlite | webapi | fixture
├── zotero_field.py       # Pembangun field + ZOTERO_PREF_1
├── selftest.py           # Uji end-to-end tanpa Zotero
├── Run.bat               # Launcher Windows 1-klik
├── config.json           # Konfigurasi default
├── requirements.txt      # lxml, pandas, openpyxl
├── WordMacro.bas         # Modul Macro VBA Word (opsional)
├── FLOW.md              # Rancangan alur & keputusan desain
├── README.md
├── templates/
│   └── zotero_prefs.json # Template properti dokumen
├── samples/
│   ├── naskah_uji.docx        # contoh naskah
│   ├── referensi_uji.xlsx     # contoh daftar referensi
│   └── library_fixture.json   # contoh library untuk self-test
└── crossref_cache.json   # cache Crossref (dibuat otomatis)
```

---

## Self-Test

Uji seluruh pipeline tanpa perlu Zotero terpasang:

```powershell
python selftest.py
```

Self-test akan:

1. Membuat sampel Word + Excel + library fixture di `samples/`
2. Menjalankan pipeline dengan adapter `fixture`
3. Memeriksa 41 hal: jumlah field, `citationID`, URI, `itemData`, teks tampilan,
   validitas docx, `ZOTERO_PREF_1`, 4 sheet laporan, status per ID, isi RIS, dan CSV

Hasil saat ini: **41/41 lulus**.

---

## Troubleshooting & Solusi Masalah

| Kendala | Kemungkinan Penyebab | Solusi |
|---|---|---|
| `403` dari `localhost:23119` | Local API belum diaktifkan | Settings → Advanced → centang izin komunikasi |
| `Connection refused` | Zotero tidak berjalan | Buka Zotero, atau jalankan `--adapter sqlite` |
| Semua berstatus "tidak ada di Zotero" | Adapter membaca library kosong | Cek jumlah dokumen di langkah `[5/6]` |
| Sitasi jadi teks abu-abu, Refresh tidak berubah | `ZOTERO_PREF_1` / `userID` salah | Ikuti bagian Kalibrasi |
| Tab Zotero bilang "bukan dokumen Zotero" | Properti prefs tidak terdeteksi | Bandingkan dengan `dump_custom.xml` |
| Item ada di Zotero tapi tidak cocok | DOI kosong di Zotero + judul beda tipis | Lengkapi DOI, atau turunkan ambang fuzzy |
| Word rusak / tidak bisa dibuka | XML tidak valid | Tool selalu menulis file baru; buka file asli dan laporkan |
| `Database is locked` (mode sqlite) | Zotero memegang lock | Tutup Zotero atau kembali ke `localapi` |
| 0 sitasi terkonversi | Pola kurung tidak standar / ID beda | Coba `--id-regex "POLA"` atau periksa kesamaan kode Word vs Excel |

---

## Perbedaan dengan Versi Mendeley

| Aspek | Mendeley | Zotero |
|---|---|---|
| Kode field | `ADDIN CSL_CITATION` | `ADDIN ZOTERO_ITEM CSL_CITATION` |
| ID item | UUID | 8 karakter item key |
| URI | `mendeley.com/documents/?uuid=` | `zotero.org/users/<id>/items/<key>` |
| Properti dokumen | tidak perlu | **wajib** `ZOTERO_PREF_1` |
| Sumber library | SQLite MRM | Local API / SQLite / Web API |
| Kunci field | `fldLock=1` | **tidak dikunci** agar bisa di-Refresh |

### Bug yang ditemukan saat porting (dan diperbaiki di sini)

Regex pemisah pada `_make_pattern()` tidak dibungkus grup non-capturing:

```python
sep = r"\s*[,;&]\s*|\s+and\s+|\s+dan\s+"        # SALAH
sep = r"(?:\s*[,;&]\s*|\s+and\s+|\s+dan\s+)"    # BENAR
```

Akibatnya `(NR001; NR002)` gagal dikenali sebagai **satu** kelompok dan dipecah menjadi dua field
terpisah. Gejala di Word: `((Smith, 2021); (Lee, 2022))` dengan dua field bertumpuk.

Perbaikan ini **juga berlaku untuk versi Mendeley** dan sebaiknya di-backport.

---

## Roadmap

- [ ] Kalibrasi `ZOTERO_PREF_1` dengan dokumen Zotero asli
- [ ] Verifikasi perilaku Refresh untuk library yang belum sinkron
- [ ] Dukungan library group (`--group-id`)
- [ ] GUI review sederhana
- [ ] Backport perbaikan regex ke versi Mendeley

---

## Lisensi

Didistribusikan di bawah Lisensi MIT. Bebas digunakan, dimodifikasi, dan dikembangkan untuk
keperluan akademik maupun komersial.
