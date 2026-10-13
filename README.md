# FIXZY — Kerangka Pembukuan Servis & Penjualan

Kerangka awal untuk penggunaan pribadi owner dan karyawan. Google Spreadsheet adalah database utama. Paket ini belum mengandung ID Spreadsheet atau URL deployment karena harus dibuat pada akun Google milik pemilik.

## Isi ZIP
- `Code.gs` — backend Google Apps Script dan pembuatan struktur sheet.
- `index.html` — dashboard dan modul dasar satu halaman.
- `app.js` — penghubung UI ke API Apps Script.
- `style.css` — tema FIXZY hitam, putih, kuning.
- `config.js` — tempat URL Web App Apps Script.
- `SPREADSHEET_HEADERS.csv` — referensi nama sheet dan kolom.

## Setup
1. Buat Google Spreadsheet baru khusus FIXZY. Jangan gunakan Spreadsheet TESPOIN.
2. Buka Extensions → Apps Script.
3. Ganti isi `Code.gs` dengan file dari ZIP.
4. Isi `SPREADSHEET_ID` di bagian atas `Code.gs` dengan ID Spreadsheet baru.
5. Simpan, pilih fungsi `setupFixzy`, lalu Run dan berikan izin yang diminta. Fungsi akan membuat sheet dan header. Jika header sheet yang sudah ada berbeda, fungsi berhenti agar tidak menimpa data.
6. Deploy → New deployment → Web app. Execute as: Me. Atur akses sesuai kebutuhan akun dan keamanan. Jangan menganggap URL yang dapat diakses publik aman untuk data keuangan; autentikasi dan otorisasi server-side perlu diselesaikan sebelum penggunaan nyata.
7. Salin URL `/exec` ke `FIXZY_API_URL` pada `config.js`.
8. Jalankan `index.html` melalui hosting HTTPS (misalnya GitHub Pages). Buka dashboard dan periksa status API.

## Catatan penting untuk versi kerangka
- Belum ada sistem login/otorisasi owner-karyawan yang aman. Jangan deploy untuk penggunaan data nyata sebelum autentikasi server-side ditambahkan.
- Form servis saat ini menyimpan jasa dan data perangkat; pemilihan sparepart di form belum dibuat. Endpoint backend menerima array `spareparts` untuk tahap berikutnya.
- Form penjualan kerangka mendukung satu jenis produk per nota. Keranjang multi-produk, diskon per item, nota PDF/WA, laporan periode, payroll lengkap, dan pembatalan dengan mutasi balik adalah tahap berikutnya.
- Pembelian stok perlu dilakukan melalui alur `stockin` dengan `jenisMutasi: "PEMBELIAN"` dan `totalBiaya`; hindari mencatat pembelian yang sama lagi melalui pengeluaran agar arus kas tidak terhitung ganda.
- `setupFixzy()` membuat sheet dan header, bukan memigrasikan data lama.
- Uji di Spreadsheet percobaan terlebih dahulu. Backup sebelum memasukkan data asli.

## Endpoint dasar
GET: `ping`, `dashboard`, `servis`, `produk`, `stok`, `penjualan`, `pengeluaran`, `pembayaran`.
POST: `createservis`, `payment`, `updatestatus`, `createproduk`, `stockin`, `createsale`, `createexpense`, `checkin`, `checkout`.

## Kategori pengeluaran
Modal Belanja Stok, Transportasi, Makan / Konsumsi, Operasional Konter, Upah Karyawan, Peralatan Servis, Cicilan / Utang, Pengeluaran Lainnya.

Catatan akuntansi: pembelian persediaan adalah arus kas keluar dan penambah persediaan. Biaya pokok baru dibebankan saat barang terjual atau digunakan. Ini membantu mencegah penghitungan ganda pada laporan laba.
