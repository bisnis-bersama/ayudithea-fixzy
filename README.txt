FIXZY - Spreadsheet + Apps Script

File HTML: index.html, servis.html, penjualan.html, stok.html, pengeluaran.html, data.html, pengaturan.html, karyawan.html, absensi.html, upah.html.
1. Buat Google Spreadsheet kosong.
2. Extensions > Apps Script, tempel Code.gs. Jika Apps Script tidak terikat ke spreadsheet, isi Script Properties SPREADSHEET_ID dengan ID spreadsheet.
3. Jalankan setupSheets() satu kali dan beri izin.
4. Deploy > New deployment > Web app; Execute as: Me; akses: sesuai kebutuhan aplikasi. Salin URL /exec.
5. Di setiap file HTML, ganti const FIXZY_API_URL=''; dengan URL /exec yang sama.
6. Upload semua HTML ke folder/repo yang sama.

Catatan: Ini backend baru kompatibel dengan action utama halaman saat ini; backup spreadsheet sebelum menghubungkan data produksi. Tes alur di salinan spreadsheet dahulu. Upah menggunakan hari berstatus Hadir saja; Libur/Izin/Sakit tidak dihitung. Pembayaran upah dan kasbon dipisahkan agar tidak terpotong ganda.
