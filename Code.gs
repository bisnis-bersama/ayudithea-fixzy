/**
 * ============================================================
 * FIXZY BACKEND
 * Google Apps Script + Google Spreadsheet
 * ============================================================
 * Mempertahankan struktur kolom lama.
 * Kolom tambahan akan dibuat di sebelah kanan.
 */

const FIXZY_SPREADSHEET_ID =
  '1Dx5wRs39KLFaUStd9NynKJX7j80Zswj66P_dkD_N_Ao';

const FIXZY_SHEET_HEADERS = {
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
    'Jenis Kunci',
    'PIN/Pola',
    'Biaya Risiko',
    'Riwayat Status',
    'Status Pengerjaan'
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
    'Keterangan',
    'Referensi ID'
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
    'Keterangan',
    'Status'
  ],

  Pengaturan: [
    'Kunci',
    'Nilai'
  ],

  RiwayatStok: [
    'ID',
    'Tanggal',
    'Kode Produk',
    'Nama Produk',
    'Jenis',
    'Jumlah',
    'Stok Sebelum',
    'Stok Sesudah',
    'Alasan',
    'Referensi',
    'Petugas'
  ],

  BiayaRisiko: [
    'ID',
    'Tanggal',
    'ID Servis',
    'Karyawan ID',
    'Nama Karyawan',
    'Nominal',
    'Keterangan',
    'Status'
  ]
};

const FIXZY_STATUS_PENGERJAAN = [
  'Masuk Antrian',
  'Menunggu Part',
  'Proses',
  'Selesai',
  'Sudah Diambil',
  'Batal'
];


/* ============================================================
   UTILITAS DASAR
   ============================================================ */

function fixzySS() {
  return SpreadsheetApp.openById(FIXZY_SPREADSHEET_ID);
}

function fixzyJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function fixzyID(prefix) {
  return prefix + '-' +
    Utilities.getUuid().substring(0, 8).toUpperCase();
}

function fixzyNum(value) {
  if (typeof value === 'number') {
    return isFinite(value) ? value : 0;
  }

  if (value === null || value === undefined || value === '') {
    return 0;
  }

  let s = String(value).trim();

  // Menangani format angka Indonesia, misalnya 1.500 atau 1.500,50.
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else {
    s = s.replace(/,/g, '');
  }

  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function fixzyText(value) {
  return value === null || value === undefined
    ? ''
    : String(value).trim();
}

function fixzyDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  return String(value || '').substring(0, 10);
}

function fixzyNow() {
  return new Date();
}


/* ============================================================
   SHEET DAN HEADER
   Tidak menghapus kolom/data lama.
   ============================================================ */

function setupSheets() {
  const ss = fixzySS();

  Object.keys(FIXZY_SHEET_HEADERS).forEach(function(name) {
    let sh = ss.getSheetByName(name);

    if (!sh) {
      sh = ss.insertSheet(name);
    }

    if (sh.getLastRow() === 0) {
      sh.getRange(
        1,
        1,
        1,
        FIXZY_SHEET_HEADERS[name].length
      ).setValues([FIXZY_SHEET_HEADERS[name]]);

      sh.setFrozenRows(1);
      return;
    }

    const lastCol = Math.max(sh.getLastColumn(), 1);

    const existing = sh
      .getRange(1, 1, 1, lastCol)
      .getDisplayValues()[0]
      .map(function(x) {
        return String(x).trim();
      });

    FIXZY_SHEET_HEADERS[name].forEach(function(header) {
      if (existing.indexOf(header) === -1) {
        const newCol = sh.getLastColumn() + 1;
        sh.getRange(1, newCol).setValue(header);
        existing.push(header);
      }
    });

    sh.setFrozenRows(1);
  });

  return {
    ok: true,
    message: 'Pemeriksaan sheet selesai. Data lama tidak dihapus.'
  };
}

function fixzySheet(name) {
  let sh = fixzySS().getSheetByName(name);

  if (!sh) {
    setupSheets();
    sh = fixzySS().getSheetByName(name);
  }

  if (!sh) {
    throw new Error('Sheet tidak ditemukan: ' + name);
  }

  return sh;
}

function fixzyHeaders(name) {
  const sh = fixzySheet(name);

  if (sh.getLastColumn() === 0) {
    return [];
  }

  return sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(function(x) {
      return String(x).trim();
    });
}

function fixzyRows(name) {
  const sh = fixzySheet(name);
  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(function(x) {
    return String(x).trim();
  });

  return values.slice(1)
    .filter(function(row) {
      return row.some(function(value) {
        return value !== '' && value !== null;
      });
    })
    .map(function(row, index) {
      const obj = {};

      headers.forEach(function(header, i) {
        if (header) {
          obj[header] = row[i];
        }
      });

      obj._row = index + 2;
      return obj;
    });
}

function fixzyAppend(name, obj) {
  const sh = fixzySheet(name);
  const headers = fixzyHeaders(name);

  sh.appendRow(headers.map(function(header) {
    return obj[header] === undefined ? '' : obj[header];
  }));
}

function fixzyFind(name, header, value) {
  return fixzyRows(name).find(function(row) {
    return String(row[header]) === String(value);
  }) || null;
}

function fixzyUpdate(name, rowNumber, obj) {
  const sh = fixzySheet(name);
  const headers = fixzyHeaders(name);

  const current = sh
    .getRange(rowNumber, 1, 1, headers.length)
    .getValues()[0];

  const mapped = {};

  headers.forEach(function(header, i) {
    mapped[header] = current[i];
  });

  Object.keys(obj).forEach(function(key) {
    if (headers.indexOf(key) !== -1) {
      mapped[key] = obj[key];
    }
  });

  sh.getRange(rowNumber, 1, 1, headers.length)
    .setValues([
      headers.map(function(header) {
        return mapped[header] === undefined
          ? ''
          : mapped[header];
      })
    ]);
}

function fixzyDeleteRows(name, rowNumbers) {
  const sh = fixzySheet(name);

  // Hapus dari baris terbesar agar nomor baris tidak bergeser.
  rowNumbers
    .filter(function(n) {
      return n > 1;
    })
    .sort(function(a, b) {
      return b - a;
    })
    .forEach(function(n) {
      sh.deleteRow(n);
    });
}


/* ============================================================
   WEB APP
   ============================================================ */

function doGet(e) {
  try {
    setupSheets();

    const p = e && e.parameter ? e.parameter : {};
    const action = String(p.action || 'ping');

    switch (action) {
      case 'ping':
        return fixzyJSON({
          ok: true,
          message: 'FIXZY API aktif'
        });

      case 'stok':
        return fixzyJSON({
          ok: true,
          data: fixzyRows('Produk')
        });

      case 'servis':
        return fixzyJSON({
          ok: true,
          data: fixzyRows('Servis')
        });

      case 'penjualan':
        return fixzyJSON({
          ok: true,
          data: fixzyRows('Penjualan')
        });

      case 'pengeluaran':
        return fixzyJSON({
          ok: true,
          data: fixzyRows('Pengeluaran')
        });

      case 'karyawan':
        return fixzyJSON({
          ok: true,
          data: fixzyRows('Karyawan')
        });

      case 'pengaturan': {
        const settings = {};

        fixzyRows('Pengaturan').forEach(function(row) {
          settings[row.Kunci] = row.Nilai;
        });

        return fixzyJSON({
          ok: true,
          data: settings
        });
      }

      case 'absensi': {
        const month = p.bulan || Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'yyyy-MM'
        );

        const data = fixzyRows('Absensi').filter(function(row) {
          return fixzyDate(row.Tanggal).substring(0, 7) === month;
        });

        return fixzyJSON({
          ok: true,
          data: data
        });
      }

      case 'kasbon':
        return fixzyJSON({
          ok: true,
          data: fixzyRows('Kasbon')
        });

      case 'riwayatstok':
        return fixzyJSON({
          ok: true,
          data: fixzyRows('RiwayatStok')
        });

      case 'biayarisiko':
        return fixzyJSON({
          ok: true,
          data: fixzyRows('BiayaRisiko')
        });

      case 'upah':
        return fixzyJSON({
          ok: true,
          data: fixzyPayroll(
            p.bulan || Utilities.formatDate(
              new Date(),
              Session.getScriptTimeZone(),
              'yyyy-MM'
            )
          )
        });

      case 'dashboard':
        return fixzyJSON({
          ok: true,
          data: fixzyDashboard()
        });

      case 'pembayaran':
        return fixzyJSON({
          ok: true,
          data: fixzyRows('Servis').filter(function(row) {
            return fixzyNum(row['Sisa Piutang']) > 0;
          })
        });

      default:
        return fixzyJSON({
          ok: false,
          message: 'Action GET tidak dikenal: ' + action
        });
    }
  } catch (err) {
    return fixzyJSON({
      ok: false,
      message: err.message
    });
  }
}


/* ============================================================
   POST ROUTER
   ============================================================ */

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);
    setupSheets();

    const body = e && e.postData
      ? JSON.parse(e.postData.contents || '{}')
      : {};

    const action = String(body.action || '');

    switch (action) {
      case 'createproduk':
        return fixzyJSON(fixzyCreateProduct(body));

      case 'editstok':
      case 'adjuststock':
        return fixzyJSON(fixzyAdjustStock(body));

      case 'returstok':
      case 'returnstock':
        return fixzyJSON(fixzyReturnStock(body));

      case 'savekaryawan':
        return fixzyJSON(fixzySaveEmployee(body));

      case 'saveabsensi':
        return fixzyJSON(fixzySaveAttendance(body));

      case 'savekasbon':
        return fixzyJSON(fixzySaveKasbon(body));

      case 'savepengaturan':
        return fixzyJSON(fixzySaveSettings(body));

      case 'createexpense':
        return fixzyJSON(fixzyCreateExpense(body));

      case 'createservis':
        return fixzyJSON(fixzyCreateService(body));

      case 'updateservis':
        return fixzyJSON(fixzyUpdateService(body));

      case 'updateservicestatus':
      case 'ubahstatusservis':
        return fixzyJSON(fixzyUpdateServiceStatus(body));

      case 'deleteservis':
        return fixzyJSON(fixzyDeleteService(body));

      case 'deleteservisbanyak':
        return fixzyJSON(fixzyDeleteManyServices(body));

      case 'createbiayarisiko':
      case 'savebiayarisiko':
        return fixzyJSON(fixzyCreateRiskCost(body));

      case 'createsale':
        return fixzyJSON(fixzyCreateSale(body));

      default:
        return fixzyJSON({
          ok: false,
          message: 'Action POST tidak dikenal: ' + action
        });
    }
  } catch (err) {
    return fixzyJSON({
      ok: false,
      message: err.message
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/* ============================================================
   PRODUK DAN STOK
   ============================================================ */

function fixzyCreateProduct(d) {
  const code = fixzyText(d.kodeProduk) || fixzyID('PRD');

  if (fixzyFind('Produk', 'Kode Produk', code)) {
    throw new Error('Kode produk sudah digunakan.');
  }

  const obj = {
    'Kode Produk': code,
    'Nama Produk': fixzyText(d.namaProduk),
    'Kategori': fixzyText(d.kategori),
    'Merek': fixzyText(d.merek),
    'Supplier': fixzyText(d.supplier),
    'Stok': Math.max(0, fixzyNum(d.stokAwal)),
    'Stok Minimum': Math.max(0, fixzyNum(d.stokMinimum)),
    'Modal Satuan': fixzyNum(d.modal),
    'Harga Jual': fixzyNum(d.hargaJual)
  };

  if (!obj['Nama Produk']) {
    throw new Error('Nama produk wajib diisi.');
  }

  fixzyAppend('Produk', obj);

  return {
    ok: true,
    message: 'Produk berhasil disimpan.',
    kodeProduk: code
  };
}

function fixzyStockChange(code, delta, type, reason, ref, petugas) {
  const product = fixzyFind('Produk', 'Kode Produk', code);

  if (!product) {
    throw new Error('Produk tidak ditemukan: ' + code);
  }

  const before = fixzyNum(product.Stok);
  const after = before + delta;

  if (after < 0) {
    throw new Error('Stok tidak mencukupi untuk ' + product['Nama Produk']);
  }

  const sh = fixzySheet('Produk');
  const headers = fixzyHeaders('Produk');
  const stockColumn = headers.indexOf('Stok') + 1;

  if (stockColumn < 1) {
    throw new Error('Kolom Stok tidak ditemukan.');
  }

  sh.getRange(product._row, stockColumn).setValue(after);

  fixzyAppend('RiwayatStok', {
    ID: fixzyID('STK'),
    Tanggal: fixzyNow(),
    'Kode Produk': code,
    'Nama Produk': product['Nama Produk'],
    Jenis: type,
    Jumlah: Math.abs(delta),
    'Stok Sebelum': before,
    'Stok Sesudah': after,
    Alasan: reason || '',
    Referensi: ref || '',
    Petugas: petugas || ''
  });

  return {
    ok: true,
    sebelum: before,
    sesudah: after
  };
}

function fixzyAdjustStock(d) {
  const delta = fixzyNum(d.perubahan);

  if (delta === 0) {
    throw new Error('Isi jumlah perubahan stok selain nol.');
  }

  return {
    ...fixzyStockChange(
      d.kodeProduk,
      delta,
      delta > 0 ? 'Penambahan' : 'Pengurangan',
      d.alasan || 'Penyesuaian stok',
      d.referensi || '',
      d.petugas || ''
    ),
    message: 'Stok berhasil diperbarui.'
  };
}

function fixzyReturnStock(d) {
  const qty = Math.abs(fixzyNum(d.jumlah));

  if (qty <= 0) {
    throw new Error('Jumlah pengembalian harus lebih dari nol.');
  }

  return {
    ...fixzyStockChange(
      d.kodeProduk,
      qty,
      'Retur',
      d.alasan || 'Pengembalian stok',
      d.referensi || '',
      d.petugas || ''
    ),
    message: 'Pengembalian stok berhasil dicatat.'
  };
}


/* ============================================================
   KARYAWAN
   ============================================================ */

function fixzySaveEmployee(d) {
  const name = fixzyText(d.nama);

  if (!name) {
    throw new Error('Nama karyawan wajib diisi.');
  }

  const employeeId = fixzyText(d.id) || fixzyID('KRY');

  const obj = {
    ID: employeeId,
    Nama: name,
    Telepon: fixzyText(d.telepon),
    Jabatan: fixzyText(d.jabatan),
    'Upah Harian': fixzyNum(d.upahHarian),
    Status: fixzyText(d.status) || 'Aktif',
    Dibuat: fixzyNow()
  };

  const old = fixzyFind('Karyawan', 'ID', employeeId);

  if (old) {
    fixzyUpdate('Karyawan', old._row, obj);
  } else {
    fixzyAppend('Karyawan', obj);
  }

  return {
    ok: true,
    message: 'Data karyawan berhasil disimpan.',
    id: employeeId
  };
}

function fixzySaveAttendance(d) {
  const employee = fixzyFind('Karyawan', 'ID', d.karyawanId);

  if (!employee) {
    throw new Error('Karyawan tidak ditemukan.');
  }

  const date = fixzyText(d.tanggal) || fixzyDate(fixzyNow());
  const all = fixzyRows('Absensi');

  const old = all.find(function(row) {
    return fixzyDate(row.Tanggal) === date &&
      String(row['Karyawan ID']) === String(d.karyawanId);
  });

  const obj = {
    ID: old ? old.ID : fixzyID('ABS'),
    Tanggal: date,
    'Karyawan ID': d.karyawanId,
    Nama: employee.Nama,
    Status: fixzyText(d.status) || 'Hadir',
    Catatan: fixzyText(d.catatan)
  };

  if (old) {
    fixzyUpdate('Absensi', old._row, obj);
  } else {
    fixzyAppend('Absensi', obj);
  }

  return {
    ok: true,
    message: 'Absensi berhasil disimpan.'
  };
}

function fixzySaveKasbon(d) {
  const employee = fixzyFind('Karyawan', 'ID', d.karyawanId);

  if (!employee) {
    throw new Error('Karyawan tidak ditemukan.');
  }

  const nominal = fixzyNum(d.nominal);

  if (nominal <= 0) {
    throw new Error('Nominal harus lebih dari nol.');
  }

  fixzyAppend('Kasbon', {
    ID: fixzyID('KAS'),
    Tanggal: d.tanggal || fixzyNow(),
    'Karyawan ID': d.karyawanId,
    Nama: employee.Nama,
    Jenis: d.jenis || 'Kasbon',
    Nominal: nominal,
    Keterangan: fixzyText(d.keterangan),
    Status: 'Aktif'
  });

  return {
    ok: true,
    message: 'Kasbon berhasil disimpan.'
  };
}


/* ============================================================
   PENGATURAN DAN PENGELUARAN
   ============================================================ */

function fixzySaveSettings(d) {
  const sh = fixzySheet('Pengaturan');
  const all = fixzyRows('Pengaturan');

  Object.keys(d).forEach(function(key) {
    if (key === 'action') return;

    const old = all.find(function(row) {
      return String(row.Kunci) === key;
    });

    if (old) {
      fixzyUpdate('Pengaturan', old._row, {
        Nilai: d[key]
      });
    } else {
      fixzyAppend('Pengaturan', {
        Kunci: key,
        Nilai: d[key]
      });
    }
  });

  return {
    ok: true,
    message: 'Pengaturan berhasil disimpan.'
  };
}

function fixzyCreateExpense(d) {
  const id = fixzyID('EXP');

  fixzyAppend('Pengeluaran', {
    'ID Pengeluaran': id,
    Tanggal: d.tanggal || fixzyNow(),
    Kategori: fixzyText(d.kategori),
    Nominal: fixzyNum(d.nominal),
    Metode: fixzyText(d.metode),
    Supplier: fixzyText(d.supplier),
    'Karyawan ID': fixzyText(d.karyawanId),
    Keterangan: fixzyText(d.keterangan),
    'Referensi ID': fixzyText(d.referensiId)
  });

  if (
    d.kategori === 'Kasbon Karyawan' &&
    d.karyawanId
  ) {
    const employee = fixzyFind(
      'Karyawan',
      'ID',
      d.karyawanId
    );

    fixzyAppend('Kasbon', {
      ID: fixzyID('KAS'),
      Tanggal: d.tanggal || fixzyNow(),
      'Karyawan ID': d.karyawanId,
      Nama: employee ? employee.Nama : '',
      Jenis: 'Kasbon',
      Nominal: fixzyNum(d.nominal),
      Keterangan: fixzyText(d.keterangan),
      Status: 'Aktif'
    });
  }

  return {
    ok: true,
    message: 'Pengeluaran berhasil disimpan.',
    id: id
  };
}


/* ============================================================
   SERVIS
   ============================================================ */

function fixzyCreateService(d) {
  const parts = Array.isArray(d.spareparts)
    ? d.spareparts
    : [];

  const serviceId = fixzyID('SRV');

  // Validasi semua sparepart sebelum stok diubah.
  const validated = [];
  let partTotal = 0;

  parts.forEach(function(part) {
    const qty = fixzyNum(part.qty);

    if (qty <= 0) {
      throw new Error('Jumlah sparepart harus lebih dari nol.');
    }

    const product = fixzyFind(
      'Produk',
      'Kode Produk',
      part.kode
    );

    if (!product) {
      throw new Error('Sparepart tidak ditemukan: ' + part.kode);
    }

    if (fixzyNum(product.Stok) < qty) {
      throw new Error(
        'Stok tidak cukup untuk ' + product['Nama Produk']
      );
    }

    partTotal += qty * fixzyNum(product['Harga Jual']);

    validated.push({
      kode: part.kode,
      qty: qty,
      nama: product['Nama Produk']
    });
  });

  const jasa = fixzyNum(d.biayaJasa);
  const riskCost = fixzyNum(d.biayaRisiko);
  const total = jasa + partTotal;
  const paid = Math.min(total, Math.max(0, fixzyNum(d.dibayar)));

  const history = [{
    status: 'Masuk Antrian',
    tanggal: fixzyNow()
  }];

  const obj = {
    'ID Servis': serviceId,
    Tanggal: fixzyNow(),
    'Nama Pelanggan': fixzyText(d.namaPelanggan),
    WhatsApp: fixzyText(d.wa),
    'Model HP': fixzyText(d.model),
    IMEI: fixzyText(d.imei),
    Keluhan: fixzyText(d.keluhan),
    'Biaya Jasa': jasa,
    Teknisi: fixzyText(d.teknisi),
    'Karyawan ID': fixzyText(d.karyawanId),
    'Catatan Internal': fixzyText(d.catatan),
    'Sparepart JSON': JSON.stringify(validated),
    'Total Tagihan': total,
    Dibayar: paid,
    'Sisa Piutang': Math.max(0, total - paid),

    // Kolom Status lama tetap dipakai untuk status pembayaran.
    Status: paid >= total ? 'Lunas' : 'Belum Lunas',

    'Jenis Kunci': fixzyText(d.jenisKunci),
    'PIN/Pola': fixzyText(d.pinPola),
    'Biaya Risiko': riskCost,
    'Riwayat Status': JSON.stringify(history),
    'Status Pengerjaan': 'Masuk Antrian'
  };

  // Simpan servis terlebih dahulu setelah validasi berhasil.
  fixzyAppend('Servis', obj);

  // Stok dikurangi dan setiap perubahan dicatat.
  validated.forEach(function(part) {
    fixzyStockChange(
      part.kode,
      -part.qty,
      'Servis',
      'Pemakaian sparepart servis',
      serviceId,
      d.teknisi || ''
    );
  });

  return {
    ok: true,
    message: 'Servis berhasil disimpan.',
    idServis: serviceId,
    total: total,
    sisaPiutang: Math.max(0, total - paid)
  };
}

function fixzyUpdateServiceStatus(d) {
  const serviceId = fixzyText(d.idServis);
  const newStatus = fixzyText(d.status);

  if (FIXZY_STATUS_PENGERJAAN.indexOf(newStatus) === -1) {
    throw new Error('Status pengerjaan tidak valid.');
  }

  const old = fixzyFind('Servis', 'ID Servis', serviceId);

  if (!old) {
    throw new Error('Data servis tidak ditemukan.');
  }

  let history = [];

  try {
    history = JSON.parse(old['Riwayat Status'] || '[]');
  } catch (ignore) {
    history = [];
  }

  history.push({
    status: newStatus,
    tanggal: fixzyNow(),
    catatan: fixzyText(d.catatan)
  });

  fixzyUpdate('Servis', old._row, {
    'Status Pengerjaan': newStatus,
    'Riwayat Status': JSON.stringify(history)
  });

  return {
    ok: true,
    message: 'Status pengerjaan diperbarui.',
    status: newStatus
  };
}

function fixzyUpdateService(d) {
  const serviceId = fixzyText(d.idServis);
  const old = fixzyFind('Servis', 'ID Servis', serviceId);

  if (!old) {
    throw new Error('Data servis tidak ditemukan.');
  }

  // Fungsi ini mengubah data dasar, bukan mengganti stok sparepart.
  // Perubahan sparepart perlu diproses dengan mekanisme retur dan
  // pemakaian stok terpisah agar stok tidak menjadi tidak akurat.
  const fields = {
    'Nama Pelanggan': fixzyText(d.namaPelanggan),
    WhatsApp: fixzyText(d.wa),
    'Model HP': fixzyText(d.model),
    IMEI: fixzyText(d.imei),
    Keluhan: fixzyText(d.keluhan),
    'Biaya Jasa': fixzyNum(d.biayaJasa),
    Teknisi: fixzyText(d.teknisi),
    'Karyawan ID': fixzyText(d.karyawanId),
    'Catatan Internal': fixzyText(d.catatan),
    'Jenis Kunci': fixzyText(d.jenisKunci),
    'PIN/Pola': fixzyText(d.pinPola),
    'Biaya Risiko': fixzyNum(d.biayaRisiko)
  };

  const total = fields['Biaya Jasa'] +
    fixzyPartsTotal(old['Sparepart JSON']);

  const paid = fixzyNum(old.Dibayar);

  fields['Total Tagihan'] = total;
  fields['Sisa Piutang'] = Math.max(0, total - paid);
  fields.Status = paid >= total ? 'Lunas' : 'Belum Lunas';

  fixzyUpdate('Servis', old._row, fields);

  return {
    ok: true,
    message: 'Data servis berhasil diperbarui.',
    total: total
  };
}

function fixzyPartsTotal(jsonText) {
  let parts = [];

  try {
    parts = JSON.parse(jsonText || '[]');
  } catch (ignore) {
    return 0;
  }

  return parts.reduce(function(sum, part) {
    const product = fixzyFind(
      'Produk',
      'Kode Produk',
      part.kode
    );

    return sum + fixzyNum(part.qty) *
      (product ? fixzyNum(product['Harga Jual']) : 0);
  }, 0);
}

function fixzyDeleteService(d) {
  const serviceId = fixzyText(d.idServis);
  const old = fixzyFind('Servis', 'ID Servis', serviceId);

  if (!old) {
    throw new Error('Data servis tidak ditemukan.');
  }

  if (fixzyNum(old.Dibayar) > 0) {
    throw new Error(
      'Servis sudah memiliki pembayaran. Selesaikan koreksi pembayaran terlebih dahulu.'
    );
  }

  // Mengembalikan sparepart yang sebelumnya dipakai.
  let parts = [];

  try {
    parts = JSON.parse(old['Sparepart JSON'] || '[]');
  } catch (ignore) {
    parts = [];
  }

  parts.forEach(function(part) {
    fixzyStockChange(
      part.kode,
      fixzyNum(part.qty),
      'Retur Pembatalan Servis',
      'Pengembalian stok dari servis yang dihapus',
      serviceId,
      ''
    );
  });

  fixzyDeleteRows('Servis', [old._row]);

  return {
    ok: true,
    message: 'Servis dihapus dan stok sparepart dikembalikan.'
  };
}

function fixzyDeleteManyServices(d) {
  const ids = Array.isArray(d.ids)
    ? d.ids.map(String)
    : [];

  if (!ids.length) {
    throw new Error('Pilih minimal satu servis.');
  }

  const all = fixzyRows('Servis');
  const selected = all.filter(function(row) {
    return ids.indexOf(String(row['ID Servis'])) !== -1;
  });

  const blocked = selected.filter(function(row) {
    return fixzyNum(row.Dibayar) > 0;
  });

  if (blocked.length) {
    throw new Error(
      'Ada servis yang sudah dibayar. Data tersebut tidak dihapus.'
    );
  }

  selected.forEach(function(row) {
    let parts = [];

    try {
      parts = JSON.parse(row['Sparepart JSON'] || '[]');
    } catch (ignore) {
      parts = [];
    }

    parts.forEach(function(part) {
      fixzyStockChange(
        part.kode,
        fixzyNum(part.qty),
        'Retur Pembatalan Servis',
        'Pengembalian stok dari servis yang dihapus',
        row['ID Servis'],
        ''
      );
    });
  });

  fixzyDeleteRows(
    'Servis',
    selected.map(function(row) {
      return row._row;
    })
  );

  return {
    ok: true,
    message: selected.length + ' servis berhasil dihapus.'
  };
}


/* ============================================================
   BIAYA RISIKO SERVIS
   ============================================================ */

function fixzyCreateRiskCost(d) {
  const service = fixzyFind(
    'Servis',
    'ID Servis',
    d.idServis
  );

  if (!service) {
    throw new Error('ID servis tidak ditemukan.');
  }

  const nominal = fixzyNum(d.nominal);

  if (nominal < 0) {
    throw new Error('Nominal tidak boleh negatif.');
  }

  const employee = d.karyawanId
    ? fixzyFind('Karyawan', 'ID', d.karyawanId)
    : null;

  fixzyAppend('BiayaRisiko', {
    ID: fixzyID('RSK'),
    Tanggal: fixzyNow(),
    'ID Servis': d.idServis,
    'Karyawan ID': d.karyawanId || '',
    'Nama Karyawan': employee ? employee.Nama : '',
    Nominal: nominal,
    Keterangan: fixzyText(d.keterangan),
    Status: 'Dicatat'
  });

  return {
    ok: true,
    message: 'Biaya risiko berhasil dicatat.'
  };
}


/* ============================================================
   PENJUALAN
   ============================================================ */

function fixzyCreateSale(d) {
  const items = Array.isArray(d.items) ? d.items : [];

  if (!items.length) {
    throw new Error('Item penjualan belum diisi.');
  }

  let total = 0;
  const validated = [];

  items.forEach(function(item) {
    const product = fixzyFind(
      'Produk',
      'Kode Produk',
      item.kode
    );

    if (!product) {
      throw new Error('Produk tidak ditemukan: ' + item.kode);
    }

    const qty = fixzyNum(item.qty);

    if (qty <= 0) {
      throw new Error('Jumlah produk harus lebih dari nol.');
    }

    if (fixzyNum(product.Stok) < qty) {
      throw new Error('Stok tidak cukup: ' + product['Nama Produk']);
    }

    total += qty * fixzyNum(product['Harga Jual']);

    validated.push({
      kode: item.kode,
      qty: qty
    });
  });

  total = Math.max(0, total - fixzyNum(d.diskon));

  const paid = Math.min(
    total,
    Math.max(0, fixzyNum(d.dibayar))
  );

  const saleId = fixzyID('JUAL');

  fixzyAppend('Penjualan', {
    'ID Penjualan': saleId,
    Tanggal: fixzyNow(),
    'Nama Pelanggan': fixzyText(d.namaPelanggan),
    WhatsApp: fixzyText(d.wa),
    'Karyawan ID': fixzyText(d.karyawanId),
    'Items JSON': JSON.stringify(validated),
    Diskon: fixzyNum(d.diskon),
    Total: total,
    Dibayar: paid,
    'Sisa Piutang': Math.max(0, total - paid),
    Metode: fixzyText(d.metode)
  });

  validated.forEach(function(item) {
    fixzyStockChange(
      item.kode,
      -item.qty,
      'Penjualan',
      'Stok terjual',
      saleId,
      d.karyawanId || ''
    );
  });

  return {
    ok: true,
    message: 'Penjualan berhasil disimpan.',
    idPenjualan: saleId,
    total: total,
    sisaPiutang: Math.max(0, total - paid)
  };
}


/* ============================================================
   LAPORAN UPAH DAN DASHBOARD
   ============================================================ */

function fixzyPayroll(month) {
  const employees = fixzyRows('Karyawan');

  const attendance = fixzyRows('Absensi').filter(function(row) {
    return fixzyDate(row.Tanggal).substring(0, 7) === month;
  });

  const advances = fixzyRows('Kasbon').filter(function(row) {
    return fixzyDate(row.Tanggal).substring(0, 7) === month;
  });

  return employees.map(function(employee) {
    const present = attendance.filter(function(row) {
      return String(row['Karyawan ID']) === String(employee.ID) &&
        String(row.Status).toLowerCase() === 'hadir';
    }).length;

    const gross = present * fixzyNum(employee['Upah Harian']);

    const loan = advances.filter(function(row) {
      return String(row['Karyawan ID']) === String(employee.ID) &&
        String(row.Jenis).toLowerCase() === 'kasbon';
    }).reduce(function(sum, row) {
      return sum + fixzyNum(row.Nominal);
    }, 0);

    const paid = advances.filter(function(row) {
      return String(row['Karyawan ID']) === String(employee.ID) &&
        String(row.Jenis).toLowerCase() === 'pembayaran upah';
    }).reduce(function(sum, row) {
      return sum + fixzyNum(row.Nominal);
    }, 0);

    return {
      ID: employee.ID,
      Nama: employee.Nama,
      HariHadir: present,
      UpahKotor: gross,
      Kasbon: loan,
      SudahDibayar: paid,
      SisaUpah: Math.max(0, gross - loan - paid)
    };
  });
}

function fixzyDashboard() {
  const services = fixzyRows('Servis');
  const sales = fixzyRows('Penjualan');
  const products = fixzyRows('Produk');
  const expenses = fixzyRows('Pengeluaran');

  return {
    jumlahServis: services.length,

    totalPiutangServis: services.reduce(function(sum, row) {
      return sum + fixzyNum(row['Sisa Piutang']);
    }, 0),

    totalPenjualan: sales.reduce(function(sum, row) {
      return sum + fixzyNum(row.Total);
    }, 0),

    totalPengeluaran: expenses.reduce(function(sum, row) {
      return sum + fixzyNum(row.Nominal);
    }, 0),

    stokMenipis: products.filter(function(row) {
      return fixzyNum(row.Stok) <=
        fixzyNum(row['Stok Minimum']);
    }).length
  };
}