/**
 * FIXZY - BACKEND GOOGLE APPS SCRIPT
 * Database: Google Spreadsheet
 *
 * Absensi:
 * - Jam kerja standar 08.00–17.00 = 9 jam.
 * - Upah harian dihitung proporsional berdasarkan jam kerja.
 * - Satu karyawan hanya memiliki satu catatan absensi per tanggal.
 */

const SHEETS = {
  Produk: [
    'Kode Produk', 'Nama Produk', 'Kategori', 'Merek', 'Supplier',
    'Stok', 'Stok Minimum', 'Modal Satuan', 'Harga Jual'
  ],

  Servis: [
    'ID Servis', 'Tanggal', 'Nama Pelanggan', 'WhatsApp',
    'Model HP', 'IMEI', 'Keluhan', 'Biaya Jasa', 'Teknisi',
    'Karyawan ID', 'Catatan Internal', 'Sparepart JSON',
    'Total Tagihan', 'Dibayar', 'Sisa Piutang', 'Status'
  ],

  Penjualan: [
    'ID Penjualan', 'Tanggal', 'Nama Pelanggan', 'WhatsApp',
    'Karyawan ID', 'Items JSON', 'Diskon', 'Total',
    'Dibayar', 'Sisa Piutang', 'Metode'
  ],

  Pengeluaran: [
    'ID Pengeluaran', 'Tanggal', 'Kategori', 'Nominal',
    'Metode', 'Supplier', 'Karyawan ID', 'Keterangan'
  ],

  Karyawan: [
    'ID', 'Nama', 'Telepon', 'Jabatan',
    'Upah Harian', 'Status', 'Dibuat'
  ],

  Absensi: [
    'ID', 'Tanggal', 'Karyawan ID', 'Nama', 'Status',
    'Jam Masuk', 'Jam Pulang', 'Jam Kerja',
    'Rasio Upah', 'Catatan'
  ],

  Kasbon: [
    'ID', 'Tanggal', 'Karyawan ID', 'Nama',
    'Jenis', 'Nominal', 'Keterangan'
  ],

  Pengaturan: [
    'Kunci', 'Nilai'
  ]
};


// =====================================================
// KONFIGURASI SPREADSHEET
// =====================================================

const FIXZY_SPREADSHEET_ID =
  '1Dx5wRs39KLFaUStd9NynKJX7j80Zswj66P_dkD_N_Ao';

const JAM_KERJA_STANDAR = 9;


function ss() {
  return SpreadsheetApp.openById(FIXZY_SPREADSHEET_ID);
}


function setupSheets() {
  const book = ss();

  Object.keys(SHEETS).forEach(function(namaSheet) {
    let sh = book.getSheetByName(namaSheet);

    if (!sh) {
      sh = book.insertSheet(namaSheet);
    }

    if (sh.getLastRow() === 0) {
      sh.appendRow(SHEETS[namaSheet]);
      return;
    }

    // Tambahkan kolom baru tanpa menghapus data lama.
    let headers = sh
      .getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
      .getValues()[0]
      .map(String);

    SHEETS[namaSheet].forEach(function(header) {
      if (!headers.includes(header)) {
        sh.getRange(1, sh.getLastColumn() + 1)
          .setValue(header);

        headers.push(header);
      }
    });
  });

  return {
    ok: true,
    message: 'Sheet FIXZY siap digunakan.'
  };
}


function sheet(nama) {
  let sh = ss().getSheetByName(nama);

  if (!sh) {
    setupSheets();
    sh = ss().getSheetByName(nama);
  }

  return sh;
}


function rows(nama) {
  const sh = sheet(nama);
  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0];

  return values.slice(1)
    .filter(function(row) {
      return row.some(function(v) {
        return v !== '' && v !== null;
      });
    })
    .map(function(row) {
      return Object.fromEntries(
        headers.map(function(header, index) {
          return [header, row[index]];
        })
      );
    });
}


function append(nama, obj) {
  const sh = sheet(nama);
  const headers = SHEETS[nama];

  sh.appendRow(
    headers.map(function(header) {
      return obj[header] === undefined
        ? ''
        : obj[header];
    })
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
  return isFinite(n) ? n : 0;
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


// =====================================================
// PERHITUNGAN JAM KERJA
// =====================================================

function hitungJamKerja(jamMasuk, jamPulang) {
  if (!jamMasuk || !jamPulang) {
    return 0;
  }

  const masuk = String(jamMasuk).split(':');
  const pulang = String(jamPulang).split(':');

  if (masuk.length < 2 || pulang.length < 2) {
    return 0;
  }

  const menitMasuk =
    Number(masuk[0]) * 60 + Number(masuk[1]);

  const menitPulang =
    Number(pulang[0]) * 60 + Number(pulang[1]);

  if (
    !Number.isFinite(menitMasuk) ||
    !Number.isFinite(menitPulang) ||
    menitPulang <= menitMasuk
  ) {
    return 0;
  }

  return (menitPulang - menitMasuk) / 60;
}


function hitungRasioUpah(jamKerja) {
  return Math.max(
    0,
    Math.min(1, num(jamKerja) / JAM_KERJA_STANDAR)
  );
}


// =====================================================
// GET API
// =====================================================

function doGet(e) {
  try {
    setupSheets();

    const action = String(
      (e.parameter.action || 'ping')
    ).split('&')[0];

    const q = e.parameter;

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
        const data = {};

        rows('Pengaturan').forEach(function(item) {
          data[item.Kunci] = item.Nilai;
        });

        return json({
          ok: true,
          data: data
        });
      }


      case 'absensi': {
        const bulan = q.bulan ||
          Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            'yyyy-MM'
          );

        const data = rows('Absensi')
          .filter(function(item) {
            return fmtDate(item.Tanggal).slice(0, 7) === bulan;
          })
          .map(function(item) {
            item.Tanggal = fmtDate(item.Tanggal);
            return item;
          });

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

        return json({
          ok: true,
          ringkasan: {
            jumlahServis: servis.length,

            totalPiutangServis: servis.reduce(
              function(total, item) {
                return total + num(item['Sisa Piutang']);
              },
              0
            ),

            totalPenjualan: penjualan.reduce(
              function(total, item) {
                return total + num(item.Total);
              },
              0
            ),

            stokMenipis: produk.filter(
              function(item) {
                return num(item.Stok) <=
                  num(item['Stok Minimum']);
              }
            ).length
          }
        });
      }


      case 'pembayaran': {
        const data = rows('Servis').filter(
          function(item) {
            return num(item['Sisa Piutang']) > 0;
          }
        );

        return json({
          ok: true,
          data: data
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


// =====================================================
// POST API
// =====================================================

function doPost(e) {
  try {
    setupSheets();

    const d = JSON.parse(
      e.postData.contents || '{}'
    );

    const action = d.action || '';


    // -------------------------------------------------
    // TAMBAH PRODUK
    // -------------------------------------------------

    if (action === 'createproduk') {
      const kode = d.kodeProduk || id('PRD');

      append('Produk', {
        'Kode Produk': kode,
        'Nama Produk': d.namaProduk,
        'Kategori': d.kategori,
        'Merek': d.merek,
        'Supplier': d.supplier,
        'Stok': num(d.stokAwal),
        'Stok Minimum': num(d.stokMinimum),
        'Modal Satuan': num(d.modal),
        'Harga Jual': num(d.hargaJual)
      });

      return json({
        ok: true,
        message: 'Produk tersimpan',
        kodeProduk: kode
      });
    }


    // -------------------------------------------------
    // TAMBAH / EDIT KARYAWAN
    // -------------------------------------------------

    if (action === 'savekaryawan') {
      const semua = rows('Karyawan');
      const kid = d.id || id('KRY');

      const obj = {
        ID: kid,
        Nama: d.nama,
        Telepon: d.telepon,
        Jabatan: d.jabatan,
        'Upah Harian': num(d.upahHarian),
        Status: d.status || 'Aktif',
        Dibuat: new Date()
      };

      const index = semua.findIndex(function(item) {
        return String(item.ID) === String(kid);
      });

      if (index >= 0) {
        const sh = sheet('Karyawan');
        const headers = sh
          .getRange(1, 1, 1, sh.getLastColumn())
          .getValues()[0]
          .map(String);

        const row = headers.map(function(header) {
          return obj[header] === undefined
            ? semua[index][header] || ''
            : obj[header];
        });

        sh.getRange(index + 2, 1, 1, row.length)
          .setValues([row]);

      } else {
        append('Karyawan', obj);
      }

      return json({
        ok: true,
        message: 'Data karyawan tersimpan',
        id: kid
      });
    }


    // -------------------------------------------------
    // ABSENSI DENGAN JAM MASUK DAN JAM PULANG
    // -------------------------------------------------

    if (action === 'saveabsensi') {
      const karyawan = rows('Karyawan').find(
        function(item) {
          return String(item.ID) === String(d.karyawanId);
        }
      );

      if (!karyawan) {
        throw Error('Karyawan tidak ditemukan.');
      }

      if (!d.tanggal) {
        throw Error('Tanggal absensi wajib diisi.');
      }

      const status = String(d.status || '');

      if (!['Hadir', 'Libur', 'Izin', 'Sakit', 'Alpa'].includes(status)) {
        throw Error('Status absensi tidak valid.');
      }

      let jamKerja = 0;
      let rasioUpah = 0;

      if (status === 'Hadir') {
        if (!d.jamMasuk || !d.jamPulang) {
          throw Error(
            'Jam masuk dan jam pulang wajib diisi.'
          );
        }

        // Hitung dari jam masuk dan jam pulang.
        jamKerja = hitungJamKerja(
          d.jamMasuk,
          d.jamPulang
        );

        if (jamKerja <= 0) {
          throw Error(
            'Jam pulang harus lebih akhir daripada jam masuk.'
          );
        }

        // Rasio upah maksimal 100%.
        rasioUpah = hitungRasioUpah(jamKerja);
      }

      const sh = sheet('Absensi');
      const semua = rows('Absensi');

      // Cari absensi karyawan pada tanggal yang sama.
      const index = semua.findIndex(function(item) {
        return fmtDate(item.Tanggal) === String(d.tanggal) &&
          String(item['Karyawan ID']) === String(d.karyawanId);
      });

      const obj = {
        'ID': index >= 0
          ? semua[index].ID
          : id('ABS'),

        'Tanggal': d.tanggal,
        'Karyawan ID': d.karyawanId,
        'Nama': karyawan.Nama,
        'Status': status,
        'Jam Masuk': status === 'Hadir'
          ? d.jamMasuk
          : '',

        'Jam Pulang': status === 'Hadir'
          ? d.jamPulang
          : '',

        'Jam Kerja': jamKerja,
        'Rasio Upah': rasioUpah,
        'Catatan': d.catatan || ''
      };

      const headers = sh
        .getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map(String);

      const row = headers.map(function(header) {
        return obj[header] === undefined
          ? ''
          : obj[header];
      });

      if (index >= 0) {
        sh.getRange(index + 2, 1, 1, row.length)
          .setValues([row]);
      } else {
        sh.appendRow(row);
      }

      return json({
        ok: true,
        message: 'Absensi dan jam kerja tersimpan.',
        nama: karyawan.Nama,
        status: status,
        jamMasuk: obj['Jam Masuk'],
        jamPulang: obj['Jam Pulang'],
        jamKerja: jamKerja,
        rasioUpah: rasioUpah,
        upahHarian: num(karyawan['Upah Harian']),
        perkiraanUpah: Math.round(
          num(karyawan['Upah Harian']) * rasioUpah
        )
      });
    }


    // -------------------------------------------------
    // KASBON / PEMBAYARAN UPAH
    // -------------------------------------------------

    if (action === 'savekasbon') {
      const karyawan = rows('Karyawan').find(
        function(item) {
          return String(item.ID) === String(d.karyawanId);
        }
      );

      if (!karyawan) {
        throw Error('Karyawan tidak ditemukan.');
      }

      const jenis = String(d.jenis || '');

      if (!['Kasbon', 'Pembayaran Upah'].includes(jenis)) {
        throw Error(
          'Jenis harus Kasbon atau Pembayaran Upah.'
        );
      }

      if (num(d.nominal) <= 0) {
        throw Error('Nominal harus lebih dari nol.');
      }

      append('Kasbon', {
        'ID': id('KAS'),
        'Tanggal': d.tanggal || new Date(),
        'Karyawan ID': d.karyawanId,
        'Nama': karyawan.Nama,
        'Jenis': jenis,
        'Nominal': num(d.nominal),
        'Keterangan': d.keterangan || ''
      });

      return json({
        ok: true,
        message: 'Kasbon/pembayaran upah tersimpan.'
      });
    }


    // -------------------------------------------------
    // SIMPAN PENGATURAN
    // -------------------------------------------------

    if (action === 'savepengaturan') {
      const sh = sheet('Pengaturan');

      sh.clearContents();
      sh.appendRow(['Kunci', 'Nilai']);

      Object.keys(d)
        .filter(function(key) {
          return key !== 'action';
        })
        .forEach(function(key) {
          sh.appendRow([key, d[key]]);
        });

      return json({
        ok: true,
        message: 'Pengaturan tersimpan.'
      });
    }


    // -------------------------------------------------
    // TAMBAH PENGELUARAN
    // -------------------------------------------------

    if (action === 'createexpense') {
      const eid = id('EXP');

      append('Pengeluaran', {
        'ID Pengeluaran': eid,
        'Tanggal': d.tanggal || new Date(),
        'Kategori': d.kategori,
        'Nominal': num(d.nominal),
        'Metode': d.metode,
        'Supplier': d.supplier,
        'Karyawan ID': d.karyawanId,
        'Keterangan': d.keterangan
      });

      // Jika kategori kasbon karyawan,
      // catat juga pada tab Kasbon.
      if (
        d.kategori === 'Kasbon Karyawan' &&
        d.karyawanId
      ) {
        const karyawan = rows('Karyawan').find(
          function(item) {
            return String(item.ID) === String(d.karyawanId);
          }
        );

        append('Kasbon', {
          'ID': id('KAS'),
          'Tanggal': d.tanggal || new Date(),
          'Karyawan ID': d.karyawanId,
          'Nama': karyawan ? karyawan.Nama : '',
          'Jenis': 'Kasbon',
          'Nominal': num(d.nominal),
          'Keterangan': d.keterangan || ''
        });
      }

      return json({
        ok: true,
        message: 'Pengeluaran tersimpan.',
        id: eid
      });
    }


    // -------------------------------------------------
    // PEMBAYARAN TAGIHAN SERVIS
    // -------------------------------------------------

    if (
      action === 'bayarservis' ||
      action === 'pembayaranservis'
    ) {
      const idServis = String(d.idServis || '').trim();
      const nominal = num(d.nominal);

      if (!idServis) {
        throw Error('ID servis wajib diisi.');
      }

      if (nominal <= 0) {
        throw Error(
          'Nominal pembayaran harus lebih dari nol.'
        );
      }

      const lock = LockService.getScriptLock();
      lock.waitLock(15000);

      try {
        const sh = sheet('Servis');
        const values = sh.getDataRange().getValues();

        if (values.length < 2) {
          throw Error('Data servis belum tersedia.');
        }

        const headers = values[0].map(String);
        const col = {};

        headers.forEach(function(header, index) {
          col[header] = index;
        });

        [
          'ID Servis',
          'Total Tagihan',
          'Dibayar',
          'Sisa Piutang',
          'Status'
        ].forEach(function(header) {
          if (col[header] === undefined) {
            throw Error(
              'Kolom servis tidak ditemukan: ' + header
            );
          }
        });

        let rowIndex = -1;
        let service = null;

        for (let i = 1; i < values.length; i++) {
          if (
            String(values[i][col['ID Servis']]) === idServis
          ) {
            rowIndex = i + 1;
            service = values[i];
            break;
          }
        }

        if (rowIndex < 0) {
          throw Error(
            'Data servis ' + idServis + ' tidak ditemukan.'
          );
        }

        const total = num(
          service[col['Total Tagihan']]
        );

        const dibayarLama = num(
          service[col['Dibayar']]
        );

        const sisaLama = Math.max(
          0,
          total - dibayarLama
        );

        if (sisaLama <= 0) {
          throw Error('Tagihan servis ini sudah lunas.');
        }

        if (nominal > sisaLama) {
          throw Error(
            'Pembayaran melebihi sisa tagihan Rp' +
            sisaLama.toLocaleString('id-ID') + '.'
          );
        }

        const dibayarBaru = dibayarLama + nominal;
        const sisaBaru = Math.max(
          0,
          total - dibayarBaru
        );

        const statusBaru = sisaBaru === 0
          ? 'Lunas'
          : 'Belum Lunas';

        const book = ss();

        let log = book.getSheetByName(
          'PembayaranServis'
        );

        const logHeaders = [
          'ID Pembayaran',
          'Tanggal',
          'ID Servis',
          'Nama Pelanggan',
          'Nominal',
          'Metode',
          'Catatan'
        ];

        if (!log) {
          log = book.insertSheet('PembayaranServis');
          log.appendRow(logHeaders);
        } else if (log.getLastRow() === 0) {
          log.appendRow(logHeaders);
        }

        let existing = log
          .getRange(1, 1, 1, log.getLastColumn())
          .getValues()[0]
          .map(String);

        logHeaders.forEach(function(header) {
          if (!existing.includes(header)) {
            log.getRange(
              1,
              log.getLastColumn() + 1
            ).setValue(header);

            existing.push(header);
          }
        });

        const logMap = {};

        log.getRange(
          1,
          1,
          1,
          log.getLastColumn()
        ).getValues()[0].forEach(function(header, index) {
          logMap[String(header)] = index;
        });

        const logRow = new Array(
          log.getLastColumn()
        ).fill('');

        const logData = {
          'ID Pembayaran': id('BYR'),
          'Tanggal': new Date(),
          'ID Servis': idServis,

          'Nama Pelanggan':
            col['Nama Pelanggan'] !== undefined
              ? service[col['Nama Pelanggan']]
              : '',

          'Nominal': nominal,
          'Metode': d.metode || 'Tunai',
          'Catatan': d.catatan || ''
        };

        Object.keys(logData).forEach(function(key) {
          if (logMap[key] !== undefined) {
            logRow[logMap[key]] = logData[key];
          }
        });

        log.appendRow(logRow);

        sh.getRange(
          rowIndex,
          col['Dibayar'] + 1
        ).setValue(dibayarBaru);

        sh.getRange(
          rowIndex,
          col['Sisa Piutang'] + 1
        ).setValue(sisaBaru);

        sh.getRange(
          rowIndex,
          col['Status'] + 1
        ).setValue(statusBaru);

        return json({
          ok: true,
          message: 'Pembayaran servis berhasil dicatat.',
          idServis: idServis,
          totalTagihan: total,
          pembayaranTambahan: nominal,
          totalDibayar: dibayarBaru,
          sisaPiutang: sisaBaru,
          status: statusBaru
        });

      } finally {
        lock.releaseLock();
      }
    }


    // -------------------------------------------------
    // INPUT SERVIS
    // -------------------------------------------------

    if (action === 'createservis') {
      const parts = Array.isArray(d.spareparts)
        ? d.spareparts
        : [];

      const products = rows('Produk');
      let partTotal = 0;

      parts.forEach(function(part) {
        const product = products.find(function(item) {
          return String(item['Kode Produk']) ===
            String(part.kode);
        });

        if (!product) {
          throw Error(
            'Sparepart tidak ditemukan: ' + part.kode
          );
        }

        const qty = num(part.qty);

        if (qty <= 0) {
          throw Error('Jumlah sparepart tidak valid.');
        }

        if (num(product.Stok) < qty) {
          throw Error(
            'Stok tidak cukup untuk ' +
            product['Nama Produk']
          );
        }

        partTotal += qty * num(product['Harga Jual']);

        updateStock(part.kode, -qty);
      });

      const total = num(d.biayaJasa) + partTotal;
      const paid = num(d.dibayar);
      const sid = id('SRV');

      append('Servis', {
        'ID Servis': sid,
        'Tanggal': new Date(),
        'Nama Pelanggan': d.namaPelanggan,
        'WhatsApp': d.wa,
        'Model HP': d.model,
        'IMEI': d.imei,
        'Keluhan': d.keluhan,
        'Biaya Jasa': num(d.biayaJasa),
        'Teknisi': d.teknisi,
        'Karyawan ID': d.karyawanId,
        'Catatan Internal': d.catatan,
        'Sparepart JSON': JSON.stringify(parts),
        'Total Tagihan': total,
        'Dibayar': paid,
        'Sisa Piutang': Math.max(0, total - paid),
        'Status': paid >= total
          ? 'Lunas'
          : 'Belum Lunas'
      });

      return json({
        ok: true,
        message: 'Servis tersimpan.',
        idServis: sid,
        total: total,
        sisaPiutang: Math.max(0, total - paid)
      });
    }


    // -------------------------------------------------
    // INPUT PENJUALAN
    // -------------------------------------------------

    if (action === 'createsale') {
      const items = Array.isArray(d.items)
        ? d.items
        : [];

      let total = 0;

      items.forEach(function(item) {
        const product = rows('Produk').find(
          function(p) {
            return String(p['Kode Produk']) ===
              String(item.kode);
          }
        );

        if (!product) {
          throw Error('Produk tidak ditemukan.');
        }

        const qty = num(item.qty);

        if (qty <= 0) {
          throw Error('Jumlah produk tidak valid.');
        }

        if (num(product.Stok) < qty) {
          throw Error(
            'Stok tidak cukup: ' +
            product['Nama Produk']
          );
        }

        total += num(product['Harga Jual']) * qty;

        updateStock(item.kode, -qty);
      });

      total = Math.max(
        0,
        total - num(d.diskon)
      );

      const paid = num(d.dibayar);
      const saleId = id('JUAL');

      append('Penjualan', {
        'ID Penjualan': saleId,
        'Tanggal': new Date(),
        'Nama Pelanggan': d.namaPelanggan,
        'WhatsApp': d.wa,
        'Karyawan ID': d.karyawanId,
        'Items JSON': JSON.stringify(items),
        'Diskon': num(d.diskon),
        'Total': total,
        'Dibayar': paid,
        'Sisa Piutang': Math.max(0, total - paid),
        'Metode': d.metode
      });

      return json({
        ok: true,
        message: 'Penjualan tersimpan.',
        idPenjualan: saleId,
        total: total,
        sisaPiutang: Math.max(0, total - paid)
      });
    }


    // -------------------------------------------------
    // ACTION TIDAK DIKENAL
    // -------------------------------------------------

    return json({
      ok: false,
      message: 'Action POST tidak dikenal: ' + action
    });

  } catch (err) {
    return json({
      ok: false,
      message: err.message
    });
  }
}


// =====================================================
// UPDATE STOK
// =====================================================

function updateStock(kode, delta) {
  const sh = sheet('Produk');
  const semua = rows('Produk');

  const index = semua.findIndex(function(item) {
    return String(item['Kode Produk']) === String(kode);
  });

  if (index < 0) {
    throw Error('Produk tidak ditemukan.');
  }

  sh.getRange(index + 2, 6)
    .setValue(num(semua[index].Stok) + delta);
}


// =====================================================
// PERHITUNGAN UPAH BULANAN
// =====================================================

function calcPayroll(bulan) {
  const karyawan = rows('Karyawan');

  const absensi = rows('Absensi').filter(
    function(item) {
      return fmtDate(item.Tanggal).slice(0, 7) === bulan;
    }
  );

  const kasbon = rows('Kasbon').filter(
    function(item) {
      return fmtDate(item.Tanggal).slice(0, 7) === bulan;
    }
  );

  return karyawan.map(function(emp) {

    const absensiHadir = absensi.filter(
      function(item) {
        return String(item['Karyawan ID']) ===
          String(emp.ID) &&
          item.Status === 'Hadir';
      }
    );

    const jumlahHadir = absensiHadir.length;

    const totalJamKerja = absensiHadir.reduce(
      function(total, item) {
        return total + num(item['Jam Kerja']);
      },
      0
    );

    const setaraHariKerja = absensiHadir.reduce(
      function(total, item) {
        const rasio = item['Rasio Upah'] === '' ||
          item['Rasio Upah'] === null ||
          item['Rasio Upah'] === undefined
            ? hitungRasioUpah(item['Jam Kerja'])
            : num(item['Rasio Upah']);

        return total + Math.max(0, Math.min(1, rasio));
      },
      0
    );

    const upahHarian = num(emp['Upah Harian']);

    const upahKotor = Math.round(
      setaraHariKerja * upahHarian
    );

    const totalKasbon = kasbon
      .filter(function(item) {
        return String(item['Karyawan ID']) ===
          String(emp.ID) &&
          item.Jenis === 'Kasbon';
      })
      .reduce(function(total, item) {
        return total + num(item.Nominal);
      }, 0);

    const sudahDibayar = kasbon
      .filter(function(item) {
        return String(item['Karyawan ID']) ===
          String(emp.ID) &&
          item.Jenis === 'Pembayaran Upah';
      })
      .reduce(function(total, item) {
        return total + num(item.Nominal);
      }, 0);

    const sisaUpah = Math.max(
      0,
      upahKotor - totalKasbon - sudahDibayar
    );

    return {
      ID: emp.ID,
      Nama: emp.Nama,
      UpahHarian: upahHarian,
      HariHadir: jumlahHadir,
      TotalJamKerja: Math.round(totalJamKerja * 100) / 100,
      SetaraHariKerja:
        Math.round(setaraHariKerja * 100) / 100,
      UpahKotor: upahKotor,
      Kasbon: totalKasbon,
      SudahDibayar: sudahDibayar,
      SisaUpah: sisaUpah
    };
  });
}