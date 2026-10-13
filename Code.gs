/**
 * FIXZY BACKEND
 * Google Apps Script + Google Spreadsheet
 *
 * Aturan:
 * 1. Pendapatan servis final hanya diakui saat Sudah Diambil.
 * 2. Selesai masih berupa tagihan/piutang sementara.
 * 3. Dibatalkan tidak menghasilkan pendapatan.
 * 4. DP wajib dikembalikan saat pembatalan.
 * 5. Status pembayaran dan status pengerjaan terpisah.
 * 6. Data dan riwayat lama dipertahankan.
 */

const FIXZY_SPREADSHEET_ID =
  '1Dx5wRs39KLFaUStd9NynKJX7j80Zswj66P_dkD_N_Ao';

const SHEETS = {
  Produk: [
    'Kode Produk',
    'Nama Produk',
    'Kategori',
    'Merek',
    'Supplier',
    'Stok',
    'Stok Minimum',
    'Modal Satuan',
    'Harga Jual'
  ],

  Servis: [
    'ID Servis',
    'Tanggal',
    'Nama Pelanggan',
    'WhatsApp',
    'Model HP',
    'IMEI',
    'Keluhan',
    'Biaya Jasa',
    'Teknisi',
    'Karyawan ID',
    'Catatan Internal',
    'Sparepart JSON',
    'Total Tagihan',
    'Dibayar',
    'Sisa Piutang',
    'Status',
    'Status Pengerjaan',
    'Catatan Pengerjaan',
    'Tanggal Diambil',
    'Pendapatan Final',
    'Status Pengembalian DP'
  ],

  Penjualan: [
    'ID Penjualan',
    'Tanggal',
    'Nama Pelanggan',
    'WhatsApp',
    'Karyawan ID',
    'Items JSON',
    'Diskon',
    'Total',
    'Dibayar',
    'Sisa Piutang',
    'Metode'
  ],

  Pengeluaran: [
    'ID Pengeluaran',
    'Tanggal',
    'Kategori',
    'Nominal',
    'Metode',
    'Supplier',
    'Karyawan ID',
    'Keterangan'
  ],

  Karyawan: [
    'ID',
    'Nama',
    'Telepon',
    'Jabatan',
    'Upah Harian',
    'Status',
    'Dibuat'
  ],

  Absensi: [
    'ID',
    'Tanggal',
    'Karyawan ID',
    'Nama',
    'Status',
    'Catatan'
  ],

  Kasbon: [
    'ID',
    'Tanggal',
    'Karyawan ID',
    'Nama',
    'Jenis',
    'Nominal',
    'Keterangan'
  ],

  Pengaturan: [
    'Kunci',
    'Nilai'
  ],

  PembayaranServis: [
    'ID Pembayaran',
    'Tanggal',
    'ID Servis',
    'Nama Pelanggan',
    'Nominal',
    'Metode',
    'Catatan'
  ],

  PengembalianDP: [
    'ID Pengembalian',
    'Tanggal',
    'ID Servis',
    'Nama Pelanggan',
    'Nominal',
    'Metode',
    'Status Pengembalian',
    'Catatan'
  ]
};

const STATUS_PENGERJAAN = [
  'Menunggu',
  'Sedang Dikerjakan',
  'Menunggu Sparepart',
  'Selesai',
  'Sudah Diambil',
  'Dibatalkan'
];


/* =====================================
   DASAR SPREADSHEET
===================================== */

function ss() {
  return SpreadsheetApp.openById(FIXZY_SPREADSHEET_ID);
}

function sheet(nama) {
  const book = ss();
  let sh = book.getSheetByName(nama);

  if (!sh) {
    sh = book.insertSheet(nama);
  }

  return sh;
}

/**
 * Menambahkan header yang belum tersedia tanpa
 * menghapus kolom atau data yang sudah ada.
 */
function ensureHeaders(nama, headers) {
  const sh = sheet(nama);

  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    return sh;
  }

  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existing = sh
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(v => String(v || '').trim());

  headers.forEach(header => {
    if (!existing.includes(header)) {
      const nextCol = sh.getLastColumn() + 1;
      sh.getRange(1, nextCol).setValue(header);
      existing.push(header);
    }
  });

  return sh;
}

function setupSheets() {
  Object.keys(SHEETS).forEach(nama => {
    ensureHeaders(nama, SHEETS[nama]);
  });

  return {
    ok: true,
    message: 'Struktur Spreadsheet FIXZY telah diperiksa.'
  };
}

function headersOf(sh) {
  return sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(String);
}

function rows(nama) {
  const sh = sheet(nama);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);

  return values.slice(1)
    .filter(row => row.some(v => v !== '' && v !== null))
    .map(row => {
      const obj = {};

      headers.forEach((h, i) => {
        obj[h] = row[i];
      });

      return obj;
    });
}

function append(nama, obj) {
  const sh = sheet(nama);
  const headers = headersOf(sh);

  sh.appendRow(
    headers.map(h =>
      obj[h] === undefined ? '' : obj[h]
    )
  );
}

function id(prefix) {
  return prefix + '-' +
    Utilities.getUuid().slice(0, 8).toUpperCase();
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function num(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function fmtDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  return String(value || '').slice(0, 10);
}

function findRowByValue(sh, columnName, value) {
  const headers = headersOf(sh);
  const col = headers.indexOf(columnName);

  if (col < 0) {
    throw new Error('Kolom tidak ditemukan: ' + columnName);
  }

  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][col]) === String(value)) {
      return {
        rowNumber: i + 1,
        headers: headers,
        values: values[i],
        col: col
      };
    }
  }

  return null;
}

function setCell(sh, rowNumber, columnName, value) {
  const headers = headersOf(sh);
  const col = headers.indexOf(columnName);

  if (col < 0) {
    throw new Error('Kolom tidak ditemukan: ' + columnName);
  }

  sh.getRange(rowNumber, col + 1).setValue(value);
}

function serviceRemaining(service) {
  return Math.max(
    0,
    num(service['Total Tagihan']) -
    num(service.Dibayar)
  );
}


/* =====================================
   API GET
===================================== */

function doGet(e) {
  try {
    setupSheets();

    const q = e && e.parameter ? e.parameter : {};
    const action = String(q.action || 'ping');

    switch (action) {
      case 'ping':
        return json({
          ok: true,
          message: 'FIXZY API siap'
        });

      case 'stok':
        return json({
          ok: true,
          data: rows('Produk')
        });

      case 'servis':
        return json({
          ok: true,
          data: rows('Servis')
        });

      case 'penjualan':
        return json({
          ok: true,
          data: rows('Penjualan')
        });

      case 'pengeluaran':
        return json({
          ok: true,
          data: rows('Pengeluaran')
        });

      case 'karyawan':
        return json({
          ok: true,
          data: rows('Karyawan')
        });

      case 'pengaturan': {
        const result = {};

        rows('Pengaturan').forEach(x => {
          result[x.Kunci] = x.Nilai;
        });

        return json({
          ok: true,
          data: result
        });
      }

      case 'pembayaran':
        return json({
          ok: true,
          data: rows('Servis').filter(
            x => serviceRemaining(x) > 0 &&
              x['Status Pengerjaan'] !== 'Dibatalkan'
          )
        });

      case 'riwayatpembayaran':
        return json({
          ok: true,
          data: rows('PembayaranServis')
        });

      case 'riwayatpengembalian':
        return json({
          ok: true,
          data: rows('PengembalianDP')
        });

      case 'absensi': {
        const bulan = q.bulan ||
          Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            'yyyy-MM'
          );

        const data = rows('Absensi').filter(
          x => fmtDate(x.Tanggal).slice(0, 7) === bulan
        );

        return json({
          ok: true,
          data: data
        });
      }

      case 'upah': {
        const bulan = q.bulan ||
          Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            'yyyy-MM'
          );

        return json({
          ok: true,
          data: calcPayroll(bulan)
        });
      }

      case 'dashboard': {
        const servis = rows('Servis');
        const penjualan = rows('Penjualan');
        const produk = rows('Produk');

        const pendapatanServis = servis.reduce(
          (sum, x) =>
            sum + num(x['Pendapatan Final']),
          0
        );

        const piutangSementara = servis.reduce(
          (sum, x) => {
            if (
              x['Status Pengerjaan'] === 'Dibatalkan' ||
              x['Status Pengerjaan'] === 'Sudah Diambil'
            ) {
              return sum;
            }

            return sum + serviceRemaining(x);
          },
          0
        );

        const totalPenjualan = penjualan.reduce(
          (sum, x) => {
            if (String(x.Status || '').toLowerCase() === 'batal') {
              return sum;
            }

            return sum + num(x.Total);
          },
          0
        );

        return json({
          ok: true,
          ringkasan: {
            jumlahServis: servis.length,
            pendapatanFinalServis: pendapatanServis,
            piutangSementaraServis: piutangSementara,
            totalPenjualan: totalPenjualan,
            stokMenipis: produk.filter(
              x => num(x.Stok) <= num(x['Stok Minimum'])
            ).length
          }
        });
      }

      default:
        return json({
          ok: false,
          message: 'Action GET tidak dikenal: ' + action
        });
    }
  } catch (err) {
    return json({
      ok: false,
      message: err.message
    });
  }
}


/* =====================================
   API POST
===================================== */

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    setupSheets();

    const raw = e && e.postData
      ? e.postData.contents
      : '{}';

    const d = JSON.parse(raw || '{}');
    const action = String(d.action || '');

    switch (action) {
      case 'ubahstatusservis':
        return json(
          ubahStatusServis(d, lock)
        );

      case 'bayarservis':
      case 'pembayaranservis':
        return json(
          bayarServis(d, lock)
        );

      case 'kembalikanDP':
        return json(
          kembalikanDP(d, lock)
        );

      case 'createproduk':
        return json(createProduk(d));

      case 'createservis':
        return json(createServis(d));

      case 'createsale':
        return json(createSale(d));

      case 'createexpense':
        return json(createExpense(d));

      case 'savekaryawan':
        return json(saveKaryawan(d));

      case 'saveabsensi':
        return json(saveAbsensi(d));

      case 'savekasbon':
        return json(saveKasbon(d));

      case 'savepengaturan':
        return json(savePengaturan(d));

      default:
        return json({
          ok: false,
          message: 'Action POST tidak dikenal: ' + action
        });
    }
  } catch (err) {
    return json({
      ok: false,
      message: err.message
    });
  }
}


/* =====================================
   UBAH STATUS PENGERJAAN
===================================== */

function ubahStatusServis(d, lock) {
  const idServis = String(d.idServis || '').trim();
  const statusBaru = String(d.statusPengerjaan || '').trim();
  const catatan = String(d.catatanPengerjaan || '').trim();

  if (!idServis) {
    throw new Error('ID servis wajib diisi.');
  }

  if (!STATUS_PENGERJAAN.includes(statusBaru)) {
    throw new Error('Status pengerjaan tidak valid.');
  }

  lock.waitLock(15000);

  try {
    const sh = sheet('Servis');
    const found = findRowByValue(
      sh,
      'ID Servis',
      idServis
    );

    if (!found) {
      throw new Error('Data servis tidak ditemukan.');
    }

    const obj = {};
    found.headers.forEach((h, i) => {
      obj[h] = found.values[i];
    });

    const statusLama = String(
      obj['Status Pengerjaan'] || 'Menunggu'
    );

    if (statusLama === 'Dibatalkan' &&
        statusBaru !== 'Dibatalkan') {
      throw new Error(
        'Servis yang dibatalkan tidak bisa diaktifkan kembali melalui perubahan status biasa.'
      );
    }

    if (statusLama === 'Sudah Diambil' &&
        statusBaru !== 'Sudah Diambil') {
      throw new Error(
        'Servis yang sudah diambil tidak bisa dibuka kembali melalui perubahan status biasa.'
      );
    }

    if (statusBaru === 'Sudah Diambil') {
      const sisa = serviceRemaining(obj);

      if (sisa > 0) {
        throw new Error(
          'HP belum bisa ditandai Sudah Diambil. Sisa tagihan Rp' +
          sisa.toLocaleString('id-ID') +
          '. Catat pelunasan terlebih dahulu.'
        );
      }

      if (statusLama === 'Dibatalkan') {
        throw new Error('Servis sudah dibatalkan.');
      }

      setCell(
        sh,
        found.rowNumber,
        'Tanggal Diambil',
        new Date()
      );

      /*
       * Pendapatan final diakui sekali saat diambil.
       * Jangan menambahkan DP lagi karena DP sudah
       * tercatat dalam pembayaran.
       */
      setCell(
        sh,
        found.rowNumber,
        'Pendapatan Final',
        num(obj['Total Tagihan'])
      );
    }

    if (statusBaru === 'Dibatalkan') {
      const dp = num(obj.Dibayar);

      setCell(
        sh,
        found.rowNumber,
        'Pendapatan Final',
        0
      );

      setCell(
        sh,
        found.rowNumber,
        'Sisa Piutang',
        0
      );

      setCell(
        sh,
        found.rowNumber,
        'Status',
        'Dibatalkan'
      );

      setCell(
        sh,
        found.rowNumber,
        'Status Pengembalian DP',
        dp > 0 ? 'Menunggu Pengembalian' : 'Tidak Ada DP'
      );

      /*
       * Jika DP > 0, sistem belum menyatakan uang
       * sudah dikembalikan. Catat pengembaliannya
       * lewat action kembalikanDP setelah uang benar
       * benar diberikan kepada pelanggan.
       */
    } else if (statusBaru === 'Selesai') {
      setCell(
        sh,
        found.rowNumber,
        'Pendapatan Final',
        0
      );
    } else if (statusBaru !== 'Sudah Diambil') {
      setCell(
        sh,
        found.rowNumber,
        'Pendapatan Final',
        0
      );
    }

    setCell(
      sh,
      found.rowNumber,
      'Status Pengerjaan',
      statusBaru
    );

    setCell(
      sh,
      found.rowNumber,
      'Catatan Pengerjaan',
      catatan
    );

    return {
      ok: true,
      message: 'Status pengerjaan berhasil diperbarui.',
      idServis: idServis,
      statusPengerjaan: statusBaru,
      pendapatanFinal:
        statusBaru === 'Sudah Diambil'
          ? num(obj['Total Tagihan'])
          : 0,
      sisaPiutang:
        statusBaru === 'Dibatalkan' ||
        statusBaru === 'Sudah Diambil'
          ? 0
          : serviceRemaining(obj),
      dpWajibDikembalikan:
        statusBaru === 'Dibatalkan'
          ? num(obj.Dibayar)
          : 0
    };
  } finally {
    lock.releaseLock();
  }
}


/* =====================================
   PEMBAYARAN SERVIS
===================================== */

function bayarServis(d, lock) {
  const idServis = String(d.idServis || '').trim();
  const nominal = num(d.nominal);

  if (!idServis) {
    throw new Error('ID servis wajib diisi.');
  }

  if (nominal <= 0) {
    throw new Error('Nominal pembayaran harus lebih dari nol.');
  }

  lock.waitLock(15000);

  try {
    const sh = sheet('Servis');
    const found = findRowByValue(
      sh,
      'ID Servis',
      idServis
    );

    if (!found) {
      throw new Error('Data servis tidak ditemukan.');
    }

    const obj = {};
    found.headers.forEach((h, i) => {
      obj[h] = found.values[i];
    });

    const statusKerja = String(
      obj['Status Pengerjaan'] || 'Menunggu'
    );

    if (statusKerja === 'Dibatalkan') {
      throw new Error(
        'Servis dibatalkan. Gunakan proses pengembalian DP, bukan pembayaran baru.'
      );
    }

    if (statusKerja === 'Sudah Diambil') {
      throw new Error(
        'Servis sudah diambil. Pembayaran tambahan harus diproses melalui koreksi transaksi.'
      );
    }

    const total = num(obj['Total Tagihan']);
    const dibayarLama = num(obj.Dibayar);
    const sisa = Math.max(0, total - dibayarLama);

    if (sisa <= 0) {
      throw new Error('Tagihan sudah lunas.');
    }

    if (nominal > sisa) {
      throw new Error(
        'Nominal melebihi sisa tagihan Rp' +
        sisa.toLocaleString('id-ID')
      );
    }

    const dibayarBaru = dibayarLama + nominal;
    const sisaBaru = Math.max(0, total - dibayarBaru);

    const log = sheet('PembayaranServis');

    append('PembayaranServis', {
      'ID Pembayaran': id('BYR'),
      'Tanggal': new Date(),
      'ID Servis': idServis,
      'Nama Pelanggan': obj['Nama Pelanggan'] || '',
      'Nominal': nominal,
      'Metode': d.metode || 'Tunai',
      'Catatan': d.catatan || ''
    });

    setCell(sh, found.rowNumber, 'Dibayar', dibayarBaru);
    setCell(sh, found.rowNumber, 'Sisa Piutang', sisaBaru);

    setCell(
      sh,
      found.rowNumber,
      'Status',
      sisaBaru === 0 ? 'Lunas' : 'Belum Lunas'
    );

    return {
      ok: true,
      message: 'Pembayaran berhasil dicatat.',
      idServis: idServis,
      totalTagihan: total,
      pembayaranTambahan: nominal,
      totalDibayar: dibayarBaru,
      sisaPiutang: sisaBaru,
      status: sisaBaru === 0 ? 'Lunas' : 'Belum Lunas'
    };
  } finally {
    lock.releaseLock();
  }
}


/* =====================================
   PENGEMBALIAN DP
===================================== */

function kembalikanDP(d, lock) {
  const idServis = String(d.idServis || '').trim();
  const nominal = num(d.nominal);
  const metode = String(d.metode || 'Tunai');
  const catatan = String(d.catatan || '');

  if (!idServis) {
    throw new Error('ID servis wajib diisi.');
  }

  lock.waitLock(15000);

  try {
    const sh = sheet('Servis');
    const found = findRowByValue(
      sh,
      'ID Servis',
      idServis
    );

    if (!found) {
      throw new Error('Data servis tidak ditemukan.');
    }

    const obj = {};
    found.headers.forEach((h, i) => {
      obj[h] = found.values[i];
    });

    if (obj['Status Pengerjaan'] !== 'Dibatalkan') {
      throw new Error(
        'Pengembalian DP hanya dapat dilakukan untuk servis yang dibatalkan.'
      );
    }

    const statusPengembalian = String(
      obj['Status Pengembalian DP'] || ''
    );

    if (statusPengembalian === 'Sudah Dikembalikan') {
      throw new Error('DP servis ini sudah dikembalikan.');
    }

    const dp = num(obj.Dibayar);

    if (dp <= 0) {
      throw new Error('Tidak ada DP yang perlu dikembalikan.');
    }

    if (nominal !== dp) {
      throw new Error(
        'DP harus dikembalikan penuh sebesar Rp' +
        dp.toLocaleString('id-ID')
      );
    }

    /*
     * Action ini dipanggil setelah pengembalian uang
     * benar-benar dilakukan oleh petugas.
     */
    append('PengembalianDP', {
      'ID Pengembalian': id('RFD'),
      'Tanggal': new Date(),
      'ID Servis': idServis,
      'Nama Pelanggan': obj['Nama Pelanggan'] || '',
      'Nominal': nominal,
      'Metode': metode,
      'Status Pengembalian': 'Sudah Dikembalikan',
      'Catatan': catatan
    });

    append('Pengeluaran', {
      'ID Pengeluaran': id('EXP'),
      'Tanggal': new Date(),
      'Kategori': 'Pengembalian DP Servis',
      'Nominal': nominal,
      'Metode': metode,
      'Supplier': '',
      'Karyawan ID': '',
      'Keterangan':
        'Pengembalian DP servis ' + idServis +
        '. ' + catatan
    });

    setCell(
      sh,
      found.rowNumber,
      'Status Pengembalian DP',
      'Sudah Dikembalikan'
    );

    /*
     * Dibayar tetap menyimpan riwayat jumlah DP yang
     * pernah diterima. Jangan menghapus pembayaran lama.
     */
    return {
      ok: true,
      message: 'Pengembalian DP berhasil dicatat.',
      idServis: idServis,
      nominalDikembalikan: nominal,
      statusPengembalian: 'Sudah Dikembalikan'
    };
  } finally {
    lock.releaseLock();
  }
}


/* =====================================
   PRODUK
===================================== */

function createProduk(d) {
  const kode = d.kodeProduk || id('PRD');

  append('Produk', {
    'Kode Produk': kode,
    'Nama Produk': d.namaProduk || '',
    'Kategori': d.kategori || '',
    'Merek': d.merek || '',
    'Supplier': d.supplier || '',
    'Stok': num(d.stokAwal),
    'Stok Minimum': num(d.stokMinimum),
    'Modal Satuan': num(d.modal),
    'Harga Jual': num(d.hargaJual)
  });

  return {
    ok: true,
    message: 'Produk tersimpan.',
    kodeProduk: kode
  };
}

function updateStock(kode, delta) {
  const sh = sheet('Produk');
  const found = findRowByValue(sh, 'Kode Produk', kode);

  if (!found) {
    throw new Error('Produk tidak ditemukan: ' + kode);
  }

  const col = found.headers.indexOf('Stok');
  const stokBaru = num(found.values[col]) + num(delta);

  if (stokBaru < 0) {
    throw new Error('Stok tidak mencukupi.');
  }

  sh.getRange(found.rowNumber, col + 1).setValue(stokBaru);
}


/* =====================================
   SERVIS BARU
===================================== */

function createServis(d) {
  const parts = Array.isArray(d.spareparts)
    ? d.spareparts
    : [];

  const products = rows('Produk');
  let totalParts = 0;

  /*
   * Validasi seluruh sparepart sebelum mengubah stok.
   */
  parts.forEach(p => {
    const product = products.find(
      x => String(x['Kode Produk']) === String(p.kode)
    );

    if (!product) {
      throw new Error('Sparepart tidak ditemukan: ' + p.kode);
    }

    const qty = num(p.qty);

    if (qty <= 0) {
      throw new Error('Jumlah sparepart tidak valid.');
    }

    if (num(product.Stok) < qty) {
      throw new Error(
        'Stok tidak cukup untuk ' + product['Nama Produk']
      );
    }

    totalParts += qty * num(product['Harga Jual']);
  });

  const total = num(d.biayaJasa) + totalParts;
  const paid = num(d.dibayar);

  if (paid < 0 || paid > total) {
    throw new Error('Nominal DP tidak valid.');
  }

  const sid = id('SRV');

  append('Servis', {
    'ID Servis': sid,
    'Tanggal': new Date(),
    'Nama Pelanggan': d.namaPelanggan || '',
    'WhatsApp': d.wa || '',
    'Model HP': d.model || '',
    'IMEI': d.imei || '',
    'Keluhan': d.keluhan || '',
    'Biaya Jasa': num(d.biayaJasa),
    'Teknisi': d.teknisi || '',
    'Karyawan ID': d.karyawanId || '',
    'Catatan Internal': d.catatan || '',
    'Sparepart JSON': JSON.stringify(parts),
    'Total Tagihan': total,
    'Dibayar': paid,
    'Sisa Piutang': Math.max(0, total - paid),
    'Status': paid >= total ? 'Lunas' : 'Belum Lunas',
    'Status Pengerjaan': 'Menunggu',
    'Catatan Pengerjaan': '',
    'Tanggal Diambil': '',
    'Pendapatan Final': 0,
    'Status Pengembalian DP': paid > 0
      ? 'Belum Dikembalikan'
      : 'Tidak Ada DP'
  });

  /*
   * Stok baru dikurangi setelah data servis berhasil
   * disimpan, mengurangi risiko stok berubah jika
   * penyimpanan servis gagal.
   */
  parts.forEach(p => {
    updateStock(p.kode, -num(p.qty));
  });

  if (paid > 0) {
    append('PembayaranServis', {
      'ID Pembayaran': id('BYR'),
      'Tanggal': new Date(),
      'ID Servis': sid,
      'Nama Pelanggan': d.namaPelanggan || '',
      'Nominal': paid,
      'Metode': d.metodeDP || 'Tunai',
      'Catatan': 'DP awal servis'
    });
  }

  return {
    ok: true,
    message: 'Servis berhasil disimpan.',
    idServis: sid,
    total: total,
    dp: paid,
    sisaPiutang: Math.max(0, total - paid),
    pendapatanFinal: 0
  };
}


/* =====================================
   PENJUALAN
===================================== */

function createSale(d) {
  const items = Array.isArray(d.items) ? d.items : [];
  const products = rows('Produk');
  let total = 0;

  items.forEach(item => {
    const product = products.find(
      x => String(x['Kode Produk']) === String(item.kode)
    );

    if (!product) {
      throw new Error('Produk tidak ditemukan: ' + item.kode);
    }

    const qty = num(item.qty);

    if (qty <= 0 || num(product.Stok) < qty) {
      throw new Error(
        'Stok tidak cukup: ' + product['Nama Produk']
      );
    }

    total += num(product['Harga Jual']) * qty;
  });

  total = Math.max(0, total - num(d.diskon));

  const paid = num(d.dibayar);

  if (paid < 0 || paid > total) {
    throw new Error('Pembayaran penjualan tidak valid.');
  }

  const sid = id('JUAL');

  append('Penjualan', {
    'ID Penjualan': sid,
    'Tanggal': new Date(),
    'Nama Pelanggan': d.namaPelanggan || '',
    'WhatsApp': d.wa || '',
    'Karyawan ID': d.karyawanId || '',
    'Items JSON': JSON.stringify(items),
    'Diskon': num(d.diskon),
    'Total': total,
    'Dibayar': paid,
    'Sisa Piutang': Math.max(0, total - paid),
    'Metode': d.metode || ''
  });

  items.forEach(item => {
    updateStock(item.kode, -num(item.qty));
  });

  return {
    ok: true,
    message: 'Penjualan tersimpan.',
    idPenjualan: sid,
    total: total,
    sisaPiutang: Math.max(0, total - paid)
  };
}


/* =====================================
   PENGELUARAN
===================================== */

function createExpense(d) {
  const nominal = num(d.nominal);

  if (nominal <= 0) {
    throw new Error('Nominal pengeluaran harus lebih dari nol.');
  }

  const eid = id('EXP');

  append('Pengeluaran', {
    'ID Pengeluaran': eid,
    'Tanggal': new Date(),
    'Kategori': d.kategori || '',
    'Nominal': nominal,
    'Metode': d.metode || 'Tunai',
    'Supplier': d.supplier || '',
    'Karyawan ID': d.karyawanId || '',
    'Keterangan': d.keterangan || ''
  });

  if (d.kategori === 'Kasbon Karyawan' && d.karyawanId) {
    const emp = rows('Karyawan').find(
      x => String(x.ID) === String(d.karyawanId)
    );

    append('Kasbon', {
      'ID': id('KAS'),
      'Tanggal': new Date(),
      'Karyawan ID': d.karyawanId,
      'Nama': emp ? emp.Nama : '',
      'Jenis': 'Kasbon',
      'Nominal': nominal,
      'Keterangan': d.keterangan || ''
    });
  }

  return {
    ok: true,
    message: 'Pengeluaran tersimpan.',
    id: eid
  };
}


/* =====================================
   KARYAWAN
===================================== */

function saveKaryawan(d) {
  const sh = sheet('Karyawan');
  const kid = String(d.id || id('KRY'));

  const obj = {
    'ID': kid,
    'Nama': d.nama || '',
    'Telepon': d.telepon || '',
    'Jabatan': d.jabatan || '',
    'Upah Harian': num(d.upahHarian),
    'Status': d.status || 'Aktif',
    'Dibuat': new Date()
  };

  const found = findRowByValue(sh, 'ID', kid);

  if (found) {
    found.headers.forEach((h, i) => {
      if (obj[h] !== undefined) {
        sh.getRange(found.rowNumber, i + 1).setValue(obj[h]);
      }
    });
  } else {
    append('Karyawan', obj);
  }

  return {
    ok: true,
    message: 'Data karyawan tersimpan.',
    id: kid
  };
}

function saveAbsensi(d) {
  const emp = rows('Karyawan').find(
    x => String(x.ID) === String(d.karyawanId)
  );

  if (!emp) {
    throw new Error('Karyawan tidak ditemukan.');
  }

  const sh = sheet('Absensi');
  const all = rows('Absensi');

  const index = all.findIndex(x =>
    fmtDate(x.Tanggal) === String(d.tanggal) &&
    String(x['Karyawan ID']) === String(d.karyawanId)
  );

  const obj = {
    'ID': id('ABS'),
    'Tanggal': d.tanggal,
    'Karyawan ID': d.karyawanId,
    'Nama': emp.Nama,
    'Status': d.status || '',
    'Catatan': d.catatan || ''
  };

  if (index >= 0) {
    const rowNumber = index + 2;
    const headers = headersOf(sh);

    headers.forEach((h, i) => {
      if (obj[h] !== undefined) {
        sh.getRange(rowNumber, i + 1).setValue(obj[h]);
      }
    });
  } else {
    append('Absensi', obj);
  }

  return {
    ok: true,
    message: 'Absensi tersimpan.'
  };
}

function saveKasbon(d) {
  const emp = rows('Karyawan').find(
    x => String(x.ID) === String(d.karyawanId)
  );

  if (!emp) {
    throw new Error('Karyawan tidak ditemukan.');
  }

  const nominal = num(d.nominal);

  if (nominal <= 0) {
    throw new Error('Nominal harus lebih dari nol.');
  }

  append('Kasbon', {
    'ID': id('KAS'),
    'Tanggal': d.tanggal || new Date(),
    'Karyawan ID': d.karyawanId,
    'Nama': emp.Nama,
    'Jenis': d.jenis || 'Kasbon',
    'Nominal': nominal,
    'Keterangan': d.keterangan || ''
  });

  return {
    ok: true,
    message: 'Kasbon/pembayaran tersimpan.'
  };
}

function savePengaturan(d) {
  const sh = sheet('Pengaturan');
  const existing = rows('Pengaturan');

  Object.keys(d)
    .filter(k => k !== 'action')
    .forEach(key => {
      const index = existing.findIndex(
        x => String(x.Kunci) === key
      );

      if (index >= 0) {
        const rowNumber = index + 2;
        sh.getRange(rowNumber, 2).setValue(d[key]);
      } else {
        append('Pengaturan', {
          'Kunci': key,
          'Nilai': d[key]
        });
      }
    });

  return {
    ok: true,
    message: 'Pengaturan tersimpan.'
  };
}


/* =====================================
   PENGGAJIAN
===================================== */

function calcPayroll(month) {
  const employees = rows('Karyawan');

  const attendance = rows('Absensi').filter(
    x => fmtDate(x.Tanggal).slice(0, 7) === month
  );

  const kasbon = rows('Kasbon').filter(
    x => fmtDate(x.Tanggal).slice(0, 7) === month
  );

  return employees.map(emp => {
    const hadir = attendance.filter(
      x =>
        String(x['Karyawan ID']) === String(emp.ID) &&
        x.Status === 'Hadir'
    ).length;

    const gross = hadir * num(emp['Upah Harian']);

    const totalKasbon = kasbon.filter(
      x =>
        String(x['Karyawan ID']) === String(emp.ID) &&
        x.Jenis === 'Kasbon'
    ).reduce((sum, x) => sum + num(x.Nominal), 0);

    const dibayar = kasbon.filter(
      x =>
        String(x['Karyawan ID']) === String(emp.ID) &&
        x.Jenis === 'Pembayaran Upah'
    ).reduce((sum, x) => sum + num(x.Nominal), 0);

    return {
      ID: emp.ID,
      Nama: emp.Nama,
      HariHadir: hadir,
      UpahKotor: gross,
      Kasbon: totalKasbon,
      SudahDibayar: dibayar,
      SisaUpah: Math.max(0, gross - totalKasbon - dibayar)
    };
  });
}