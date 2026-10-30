# Rencana Implementasi: SiBOS V1 — Aplikasi Pengelolaan Dana BOS & Belanja SI (Cloud & Spreadsheet Sync)

## 1. Gambaran Umum Aplikasi
Membangun aplikasi web **SiBOS V1 — Aplikasi Pengelolaan Dana BOS & Belanja SI** yang mereplikasi **100% tampilan visual, struktur tabel, tata letak cetak (A4 Landscape & Portrait), kalkulasi otomatis, dan seluruh fitur asli** dari kode referensi (`Index.html` dan `Kode.gs`) tanpa mengurangi atau mengubah fungsionalitas yang ada, serta memastikan **seluruh perubahan tersimpan otomatis di cloud dan sinkron di seluruh perangkat pengguna**.

---

## 2. Keamanan Password & Sheet `Role Pengguna` (Tanpa Hardcode di Script Mana Pun)
Sesuai instruksi Anda:
- **Zero Hardcoded Password di Seluruh Script**:
  - **Tidak ada password apa pun yang ditulis di dalam kode script** (baik di script frontend React, backend server, maupun di `Kode.gs` pada fungsi `ensureAllSheets` / `apiSaveSchoolData`).
  - Saat Sheet 19 (`Role Pengguna`) diinisialisasi pertama kali oleh `ensureAllSheets()`, kolom `PASSWORD` (kolom D) dikosongkan (`""`) dan tidak pernah ditulis di dalam kode program.
- **Verifikasi Murni dari Sheet `Role Pengguna` di Spreadsheet**:
  - Ketika pengguna mengklik ikon gir (`⚙️`) di pojok kanan atas untuk masuk ke **Mode Admin**, sistem memanggil fungsi verifikasi (`apiVerifyAdminLogin`) yang membaca **langsung dari kolom `PASSWORD` pada Sheet `Role Pengguna` di Spreadsheet / Database**.
  - Saat penyimpanan data sekolah otomatis (`apiSaveSchoolData`) berjalan, kolom `ROLE`, `USER / NPSN`, dan `PASSWORD` pada Sheet `Role Pengguna` dijaga agar **tidak pernah tertimpa**, sehingga password yang Anda tulis langsung di Sheet `Role Pengguna` pada Spreadsheet tetap utuh 100%.
  - Perubahan kredensial melalui menu **Pengaturan Sekolah** hanya memperbarui sel kolom `USER / NPSN` dan `PASSWORD` di Sheet `Role Pengguna` melalui fungsi `apiUpdateAdminCredentials`.

---

## 3. Arsitektur Penyimpanan Cloud Otomatis & Sinkronisasi 20 Sheet Spreadsheet
- **Kecepatan Instan (`localStorage` Cache)**: Saat aplikasi dibuka di perangkat mana pun, halaman langsung tampil dalam 0 detik menggunakan cache lokal sembari menyinkronkan data terbaru di latar belakang.
- **Sinkronisasi Otomatis 2 Arah (Cloud Server & Google Spreadsheet 20 Sheet)**:
  - Mendukung penyimpanan otomatis ke Cloud Database dan jembatan sinkronisasi langsung dengan **20 Sheet Google Spreadsheet** (`Database Master RKAS`, `Database Belanja`, `SI`, `Surat Transfer`, `Rekap Serapan`, `Rekening Koran`, `Januari` s/d `Desember`, `Role Pengguna`, dan `Realisasi Distribusi`).
  - Indikator status di pojok kanan atas (`🟢 Tersimpan Otomatis`, `🔄 Menyimpan Data...`) bekerja otomatis setiap ada perubahan data (`auto-save` 500ms + `auto-pull` saat ada pembaruan dari perangkat lain).
- **Kondisi Awal Data**:
  - Tabel **Database Master RKAS**, **Belanja**, **Per-Bulan**, dan **Distribusi** dimulai dalam keadaan **kosong** hingga Anda mengimpor file RKAS (PDF/Excel/JSON) atau memuat dari Spreadsheet.
  - Identitas sekolah default: **SDIT ANNISA** (NPSN: `20231556`, Tahun Anggaran: `2026`, Sumber Dana: `BOS Reguler`, Nama Bank: `Bank BJB Kantor Cabang Jatiasih`).

---

## 4. Modul & Fitur Lengkap yang Diimplementasikan (100% Sesuai Referensi)

### A. Navigasi & Tata Letak Responsif
- **Sidebar Desktop (`#0f172a`)**: Navigasi lengkap ke seluruh 10 menu utama + *accordion* Rincian Kertas Kerja Per-Bulan (`Januari` s/d `Desember`) dengan lencana (*badge*) jumlah item aktif per bulan.
- **Header & Mobile Bar**: Header atas dengan identitas sekolah, status penyimpanan otomatis, dan tombol gir Admin (`⚙️`), dilengkapi navigasi bawah (*bottom bar*) khusus HP/Tablet (`Dashboard`, `Distribusi Stok`, `Menu`) serta 10 tombol pintasan modul di Dashboard tampilan mobile.

### B. 10 Modul Utama & Fungsionalitas Asli
1. **Dashboard Ringkasan Keuangan BOS**:
   - 4 Kartu KPI gradien: *Saldo Awal Tahun*, *Pendapatan (Dana Masuk)*, *Realisasi Belanja*, dan *Sisa Kas Tersedia*.
   - Tabel *Anggaran & Realisasi per Jenis Belanja* (Belanja Pegawai `5.1.01`, Belanja Barang & Jasa, Belanja Modal `5.2`) lengkap dengan persentase serapan dan sisa pagu.
   - Kartu progres *Penggunaan Dana* (Dana Terpakai vs Sisa Dana).
2. **Database Master RKAS**:
   - Impor otomatis dari file **PDF RKAS/ARKAS**, **Excel (`.xlsx`/`.xls`)**, atau **Backup JSON (`.json`)**.
   - Dilengkapi kamus pelengkap otomatis nomenklatur ARKAS (`ARKAS_KEGIATAN_FULL_MAP`, `ARKAS_REKENING_FULL_MAP`, `ARKAS_TRUNCATED_ITEM_MAP`) agar teks kegiatan/rekening/uraian yang terpotong di PDF tersambung utuh.
   - Hirarki baris berwarna (*Standar Nasional* merah muda/biru, *Kegiatan* hijau, *Rekening Belanja* kuning muda) dan perhitungan subtotal otomatis.
   - Fitur **Ceklis Item Terbeli** (baris berubah menjadi kuning terang `bg-purchased-yellow`) yang terhubung otomatis dengan transaksi belanja maupun ceklis manual Admin.
   - Kolom **Keterangan Manual** yang dapat diketik langsung dan tersimpan otomatis.
   - Proteksi **Wajib Backup JSON sebelum Reset Database**, serta tombol **Unduh Excel** dan **Unduh PDF**.
3. **Database Transaksi Belanja Sekolah**:
   - Filter berdasarkan *Tanggal SI* dan *Bulan Anggaran*.
   - Form Input Transaksi Belanja (khusus Admin) dengan *sticky form* saat memilih rincian sub-item RKAS menggunakan *checkbox* pintar dan kalkulasi otomatis *Sisa Anggaran Item*.
   - Fitur Edit & Hapus transaksi, Backup/Restore JSON, serta Ekspor Excel & PDF.
4. **Standing Instruction (SI) Cetak**:
   - Filter *Tanggal SI* dan *Bulan*, pratinjau dokumen siap cetak ukuran **A4 Landscape**, tabel transaksi terstruktur, serta blok tanda tangan Kepala Sekolah & Bendahara.
5. **Surat Perintah Pemindahbukuan (Surat Transfer)**:
   - Tata letak cetak **A4 Portrait** dengan **Kop Surat Sekolah** di bagian atas.
   - Input *Tanggal Cetak* dan *Nomor Surat* yang dapat disesuaikan, tabel rekapitulasi transfer per rekening/toko, serta konversi otomatis nominal ke kalimat **Terbilang Rupiah**.
6. **Laporan Rekapitulasi Realisasi & Serapan Dana BOS**:
   - Filter *Semua Bulan (1 Tahun)* atau *Per Bulan*.
   - 4 kartu ringkasan (Total Pagu Anggaran, Realisasi Belanja, Sisa Pagu Anggaran, Persentase Serapan) dan tabel perbandingan detail antara Pagu RKAS, Realisasi, Sisa Pagu, dan Status Serapan.
7. **Laporan Rekening Koran Transaksi Bank**:
   - Buku kas/passbook 4 kolom (`TGL`, `DEBIT`, `KREDIT`, `TOTAL` saldo berjalan) yang mengurutkan transaksi secara kronologis.
   - Mode tampilan **Transaksi Per-Tanggal** atau **Total Belanja Per-Bulan**.
   - Fitur Admin untuk mengubah *Saldo Awal*, menambah/mengedit/menghapus *Dana Masuk (Kredit)*, serta unduh Excel & PDF.
8. **Rincian Kertas Kerja Per-Bulan (Januari – Desember)**:
   - Impor PDF/Excel/JSON RKAS bulanan (otomatis mendeteksi kolom bulan atau mendistribusikan ke bulan aktif).
   - Fitur **Generasi Rincian** (membagi rata pagu Master RKAS ke 12 bulan), **Hapus Anggaran Bulan**, ceklis kuning terintegrasi dengan transaksi belanja bulan terkait, serta unduh Excel & PDF.
9. **Realisasi Distribusi — Pengeluaran & Stok Barang**:
   - Mengekstrak stok dan satuan barang secara otomatis dari Kertas Kerja Per-Bulan & Master RKAS.
   - Form pencatatan pengeluaran barang dengan pencarian item, validasi sisa stok agar tidak melebihi stok tersedia, tabel riwayat distribusi, pencarian, edit/hapus, dan ekspor Excel.
10. **Pengaturan Sekolah, Logo, Kop Surat & Kredensial Sheet Role Pengguna**:
    - Formulir lengkap profil sekolah (NPSN, Nama Sekolah, Tahun Anggaran, Sumber Dana, Nama Bank, Kepala Sekolah + NIP, Bendahara + NIP).
    - Unggah **Logo Sekolah** dan **Kop Surat** (JPG/PNG) dengan kompresi kanvas otomatis.
    - Pengaturan pembaruan **User & Password Admin** yang langsung disimpan ke Sheet `Role Pengguna` serta salinan skrip `Kode.gs` bebas-hardcode untuk Spreadsheet Anda.
