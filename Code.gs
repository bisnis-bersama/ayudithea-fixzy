/**
 * FIXZY - BACKEND GOOGLE APPS SCRIPT
 * Database: Google Spreadsheet
 *
 * ATURAN SERVIS:
 * 1. Status pembayaran terpisah dari status pengerjaan.
 * 2. Selesai = servis selesai dikerjakan, belum tentu lunas.
 * 3. Sudah Diambil = pembayaran otomatis Lunas.
 * 4. Pendapatan final dicatat satu kali.
 * 5. Dibatalkan = tidak dihitung sebagai pendapatan final.
 * 6. DP servis yang dibatalkan wajib dikembalikan 100%.
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
  ]
};


/* =====================================================
   DASAR SPREADSHEET
===================================================== */

function ss() {
  return SpreadsheetApp.openById(FIXZY_SPREADSHEET_ID);
}


function setupSheets() {
  const book = ss();

  Object.keys(SHEETS).forEach(function(name) {
    let sh = book.getSheetByName(name);

    if (!sh) {
      sh = book.insertSheet(name);
    }

    const headers = SHEETS[name];

    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      return;
    }

    const current = sh
      .getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
      .getValues()[0]
      .map(String);

    headers.forEach(function(header) {
      if (!current.includes(header)) {
        sh.getRange(1, sh.getLastColumn() + 1)
          .setValue(header);

        current.push(header);
      }
    });
  });

  ensureExtraSheet('PembayaranServis', [
    'ID Pembayaran',
    'Tanggal',
    'ID Servis',
    'Nama Pelanggan',
    'Nominal',
    'Metode',
    'Catatan'
  ]);

  ensureExtraSheet('PengembalianDP', [
    'ID Pengembalian',
    'Tanggal',
    'ID Servis',
    'Nama Pelanggan',
    'Nominal',
    'Metode',
    'Catatan'
  ]);

  return {
    ok: true,
    message: 'Sheet FIXZY berhasil diperiksa.'
  };
}


function ensureExtraSheet(name, headers) {
  const book = ss();
  let sh = book.getSheetByName(name);

  if (!sh) {
    sh = book.insertSheet(name);
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    return;
  }

  const current = sh
    .getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
    .getValues()[0]
    .map(String);

  headers.forEach(function(header) {
    if (!current.includes(header)) {
      sh.getRange(1, sh.getLastColumn() + 1)
        .setValue(header);

      current.push(header);
    }
  });
}


function sheet(name) {
  let sh = ss().getSheetByName(name);

  if (!sh) {
    setupSheets();
    sh = ss().getSheetByName(name);
  }

  return sh;
}


function headersOf(sh) {
  return sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(String);
}


function rows(name) {
  const sh = sheet(name);
  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(String);

  return values.slice(1)
    .filter(function(row) {
      return row.some(function(value) {
        return value !== '' && value !== null;
      });
    })
    .map(function(row) {
      const result = {};

      headers.forEach(function(header, index) {
        result[header] = row[index];
      });

      return result;
    });
}


function append(name, obj) {
  const sh = sheet(name);
  const headers = headersOf(sh);

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
  if (typeof value === 'number') {
    return isFinite(value) ? value : 0;
  }

  const result = Number(
    String(value === undefined || value === null ? '' : value)
      .replace(/,/g, '')
  );

  return isFinite(result) ? result : 0;
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


function indexByHeader(headers) {
  const result = {};

  headers.forEach(function(header, index) {
    result[header] = index;
  });

  return result;
}


/* =====================================================
   PENCARIAN DATA SERVIS
===================================================== */

function findServis(idServis) {
  const sh = sheet('Servis');
  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    throw Error('Data servis belum tersedia.');
  }

  const headers = values[0].map(String);
  const col = indexByHeader(headers);

  if (col['ID Servis'] === undefined) {
    throw Error('Kolom ID Servis tidak ditemukan.');
  }

  for (let i = 1; i < values.length; i++) {
    if (
      String(values[i][col['ID Servis']]) ===
      String(idServis)
    ) {
      return {
        sh: sh,
        values: values,
        headers: headers,
        col: col,
        row: values[i],
        rowIndex: i + 1
      };
    }
  }

  throw Error(
    'Data servis ' + idServis + ' tidak ditemukan.'
  );
}


function writeCell(found, header, value) {
  if (found.col[header] === undefined) {
    throw Error('Kolom tidak ditemukan: ' + header);
  }

  found.sh
    .getRange(found.rowIndex, found.col[header] + 1)
    .setValue(value);
}


/* =====================================================
   RIWAYAT PEMBAYARAN
===================================================== */

function appendPaymentLog(
  idServis,
  nama,
  nominal,
  metode,
  catatan
) {
  append('PembayaranServis', {
    'ID Pembayaran': id('BYR'),
    'Tanggal': new Date(),
    'ID Servis': idServis,
    'Nama Pelanggan': nama,
    'Nominal': nominal,
    'Metode': metode || 'Tunai',
    'Catatan': catatan || ''
  });
}


function appendRefundLog(
  idServis,
  nama,
  nominal,
  metode,
  catatan
) {
  append('PengembalianDP', {
    'ID Pengembalian': id('KEM'),
    'Tanggal': new Date(),
    'ID Servis': idServis,
    'Nama Pelanggan': nama,
    'Nominal': nominal,
    'Metode': metode || 'Tunai',
    'Catatan': catatan || ''
  });
}


/* =====================================================
   GET API
===================================================== */

function doGet(e) {
  try {
    setupSheets();

    const q = (e && e.parameter) || {};
    const action = String(q.action || 'ping').split('&')[0];

    switch (action) {

      case 'ping':
        return json({
          ok: true,
          message: 'FIXZY API siap'
        });

      case 'stok':
      case 'produk':
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

        rows('Pengaturan').forEach(function(item) {
          result[item.Kunci] = item.Nilai;
        });

        return json({
          ok: true,
          data: result
        });
      }

      case 'pembayaran':
        return json({
          ok: true,
          data: rows('Servis').filter(function(item) {
            return (
              num(item['Sisa Piutang']) > 0 &&
              String(item['Status Pengerjaan'] || '') !== 'Dibatalkan'
            );
          })
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
        const month = q.bulan ||
          Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            'yyyy-MM'
          );

        return json({
          ok: true,
          data: rows('Absensi').filter(function(item) {
            return fmtDate(item.Tanggal).slice(0, 7) === month;
          })
        });
      }

      case 'upah':
        return json({
          ok: true,
          data: calcPayroll(
            q.bulan ||
            Utilities.formatDate(
              new Date(),
              Session.getScriptTimeZone(),
              'yyyy-MM'
            )
          )
        });

      case 'dashboard': {
        const servis = rows('Servis');
        const penjualan = rows('Penjualan');
        const produk = rows('Produk');

        return json({
          ok: true,
          ringkasan: {
            jumlahServis: servis.length,

            totalPiutangServis: servis
              .filter(function(item) {
                return String(
                  item['Status Pengerjaan'] || ''
                ) !== 'Dibatalkan';
              })
              .reduce(function(total, item) {
                return total + num(item['Sisa Piutang']);
              }, 0),

            pendapatanFinalServis: servis.reduce(
              function(total, item) {
                return total + num(item['Pendapatan Final']);
              },
              0
            ),

            totalPenjualan: penjualan.reduce(
              function(total, item) {
                return total + num(item.Total);
              },
              0
            ),

            stokMenipis: produk.filter(function(item) {
              return num(item.Stok) <= num(item['Stok Minimum']);
            }).length
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
      message: err.message || String(err)
    });
  }
}


/* =====================================================
   POST API
===================================================== */

function doPost(e) {
  try {
    setupSheets();

    const d = JSON.parse(
      (e && e.postData && e.postData.contents) || '{}'
    );

    const action = String(d.action || '');


    /* -----------------------------------------------
       TAMBAH PRODUK
    ------------------------------------------------ */

    if (action === 'createproduk') {
      const code = d.kodeProduk || id('PRD');

      append('Produk', {
        'Kode Produk': code,
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
        kodeProduk: code
      });
    }


    /* -----------------------------------------------
       DATA KARYAWAN
    ------------------------------------------------ */

    if (action === 'savekaryawan') {
      const all = rows('Karyawan');
      const kid = String(d.id || id('KRY'));

      const obj = {
        'ID': kid,
        'Nama': d.nama,
        'Telepon': d.telepon,
        'Jabatan': d.jabatan,
        'Upah Harian': num(d.upahHarian),
        'Status': d.status || 'Aktif',
        'Dibuat': new Date()
      };

      const found = all.findIndex(function(item) {
        return String(item.ID) === kid;
      });

      if (found >= 0) {
        const sh = sheet('Karyawan');
        const headers = headersOf(sh);

        sh.getRange(
          found + 2,
          1,
          1,
          headers.length
        ).setValues([
          headers.map(function(header) {
            return obj[header] === undefined
              ? ''
              : obj[header];
          })
        ]);

      } else {
        append('Karyawan', obj);
      }

      return json({
        ok: true,
        message: 'Data karyawan tersimpan',
        id: kid
      });
    }


    /* -----------------------------------------------
       ABSENSI
    ------------------------------------------------ */

    if (action === 'saveabsensi') {
      const emp = rows('Karyawan').find(function(item) {
        return String(item.ID) === String(d.karyawanId);
      });

      if (!emp) {
        throw Error('Karyawan tidak ditemukan.');
      }

      const sh = sheet('Absensi');
      const all = rows('Absensi');

      const idx = all.findIndex(function(item) {
        return (
          fmtDate(item.Tanggal) === String(d.tanggal) &&
          String(item['Karyawan ID']) === String(d.karyawanId)
        );
      });

      const obj = {
        'ID': id('ABS'),
        'Tanggal': d.tanggal,
        'Karyawan ID': d.karyawanId,
        'Nama': emp.Nama,
        'Status': d.status,
        'Catatan': d.catatan
      };

      if (idx >= 0) {
        const headers = headersOf(sh);

        sh.getRange(
          idx + 2,
          1,
          1,
          headers.length
        ).setValues([
          headers.map(function(header) {
            return obj[header] === undefined
              ? ''
              : obj[header];
          })
        ]);

      } else {
        append('Absensi', obj);
      }

      return json({
        ok: true,
        message: 'Absensi tersimpan'
      });
    }


    /* -----------------------------------------------
       KASBON / PEMBAYARAN UPAH
    ------------------------------------------------ */

    if (action === 'savekasbon') {
      const emp = rows('Karyawan').find(function(item) {
        return String(item.ID) === String(d.karyawanId);
      });

      if (!emp) {
        throw Error('Karyawan tidak ditemukan.');
      }

      append('Kasbon', {
        'ID': id('KAS'),
        'Tanggal': d.tanggal || new Date(),
        'Karyawan ID': d.karyawanId,
        'Nama': emp.Nama,
        'Jenis': d.jenis,
        'Nominal': num(d.nominal),
        'Keterangan': d.keterangan
      });

      return json({
        ok: true,
        message: 'Kasbon/pembayaran tersimpan'
      });
    }


    /* -----------------------------------------------
       PENGATURAN
    ------------------------------------------------ */

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
        message: 'Pengaturan tersimpan'
      });
    }


    /* -----------------------------------------------
       PENGELUARAN
    ------------------------------------------------ */

    if (action === 'createexpense') {
      const eid = id('EXP');

      append('Pengeluaran', {
        'ID Pengeluaran': eid,
        'Tanggal': new Date(),
        'Kategori': d.kategori,
        'Nominal': num(d.nominal),
        'Metode': d.metode,
        'Supplier': d.supplier,
        'Karyawan ID': d.karyawanId,
        'Keterangan': d.keterangan
      });

      if (
        d.kategori === 'Kasbon Karyawan' &&
        d.karyawanId
      ) {
        const emp = rows('Karyawan').find(function(item) {
          return String(item.ID) === String(d.karyawanId);
        });

        append('Kasbon', {
          'ID': id('KAS'),
          'Tanggal': new Date(),
          'Karyawan ID': d.karyawanId,
          'Nama': emp ? emp.Nama : '',
          'Jenis': 'Kasbon',
          'Nominal': num(d.nominal),
          'Keterangan': d.keterangan
        });
      }

      return json({
        ok: true,
        message: 'Pengeluaran tersimpan',
        id: eid
      });
    }


    /* -----------------------------------------------
       INPUT SERVIS
    ------------------------------------------------ */

    if (action === 'createservis') {
      const lock = LockService.getScriptLock();
      lock.waitLock(15000);

      try {
        const parts = Array.isArray(d.spareparts)
          ? d.spareparts
          : [];

        const products = rows('Produk');
        const demand = {};
        let partTotal = 0;

        parts.forEach(function(part) {
          const code = String(part.kode);
          const qty = num(part.qty);

          if (qty <= 0) {
            throw Error(
              'Jumlah sparepart harus lebih dari nol.'
            );
          }

          demand[code] = (demand[code] || 0) + qty;
        });

        Object.keys(demand).forEach(function(code) {
          const product = products.find(function(item) {
            return String(item['Kode Produk']) === code;
          });

          if (!product) {
            throw Error(
              'Sparepart tidak ditemukan: ' + code
            );
          }

          if (num(product.Stok) < demand[code]) {
            throw Error(
              'Stok tidak cukup untuk ' +
              product['Nama Produk']
            );
          }
        });

        parts.forEach(function(part) {
          const product = products.find(function(item) {
            return String(item['Kode Produk']) ===
              String(part.kode);
          });

          partTotal +=
            num(part.qty) * num(product['Harga Jual']);
        });

        const total = num(d.biayaJasa) + partTotal;
        const paid = num(d.dibayar);

        if (paid < 0 || paid > total) {
          throw Error('Nominal DP tidak valid.');
        }

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
          'Status': paid >= total ? 'Lunas' : 'Belum Lunas',
          'Status Pengerjaan': 'Masuk',
          'Catatan Pengerjaan': '',
          'Tanggal Diambil': '',
          'Pendapatan Final': 0,
          'Status Pengembalian DP':
            paid > 0 ? 'Belum Dikembalikan' : 'Tidak Ada DP'
        });

        parts.forEach(function(part) {
          updateStock(part.kode, -num(part.qty));
        });

        if (paid > 0) {
          appendPaymentLog(
            sid,
            d.namaPelanggan,
            paid,
            d.metodeDP || 'Tunai',
            'DP awal servis'
          );
        }

        return json({
          ok: true,
          message: 'Servis tersimpan',
          idServis: sid,
          total: total,
          sisaPiutang: Math.max(0, total - paid)
        });

      } finally {
        lock.releaseLock();
      }
    }


    /* -----------------------------------------------
       UBAH STATUS PENGERJAAN SERVIS

       Sudah Diambil:
       - Dibayar = Total Tagihan
       - Sisa Piutang = 0
       - Status pembayaran = Lunas
       - Pendapatan Final dicatat satu kali
       - Pelunasan selisih masuk riwayat pembayaran

       Dibatalkan:
       - Pendapatan final = 0
       - Piutang servis = 0
       - Status pembayaran tidak ditimpa
    ------------------------------------------------ */

    if (action === 'ubahstatusservis') {
      const sid = String(d.idServis || '').trim();

      const status = String(
        d.statusPengerjaan || d.status || ''
      ).trim();

      const allowed = [
        'Masuk',
        'Menunggu Sparepart',
        'Proses',
        'Selesai',
        'Sudah Diambil',
        'Dibatalkan'
      ];

      if (!sid) {
        throw Error('ID servis wajib diisi.');
      }

      if (!allowed.includes(status)) {
        throw Error('Status pengerjaan tidak valid.');
      }

      const lock = LockService.getScriptLock();
      lock.waitLock(15000);

      try {
        const f = findServis(sid);

        const current = String(
          f.row[f.col['Status Pengerjaan']] || 'Masuk'
        );

        const total = num(
          f.row[f.col['Total Tagihan']]
        );

        const paid = num(
          f.row[f.col['Dibayar']]
        );

        const oldFinal = num(
          f.row[f.col['Pendapatan Final']]
        );

        if (
          current === 'Sudah Diambil' &&
          status !== 'Sudah Diambil'
        ) {
          throw Error(
            'Servis yang sudah diambil tidak dapat diubah kembali.'
          );
        }

        if (
          current === 'Dibatalkan' &&
          status !== 'Dibatalkan'
        ) {
          throw Error(
            'Servis yang dibatalkan tidak dapat diaktifkan kembali.'
          );
        }

        if (
          status === 'Sudah Diambil' &&
          current !== 'Sudah Diambil'
        ) {
          const gap = Math.max(0, total - paid);

          /*
           * Jika ada selisih, catat sebagai pelunasan otomatis.
           * Keterangan ini menandakan pencatatan otomatis,
           * bukan bukti uang benar-benar diterima secara fisik.
           */
          if (gap > 0) {
            appendPaymentLog(
              sid,
              String(
                f.row[f.col['Nama Pelanggan']] || ''
              ),
              gap,
              'Pelunasan saat pengambilan',
              'Pelunasan otomatis saat status diubah menjadi Sudah Diambil'
            );
          }

          writeCell(f, 'Dibayar', total);
          writeCell(f, 'Sisa Piutang', 0);
          writeCell(f, 'Status', 'Lunas');
          writeCell(f, 'Tanggal Diambil', new Date());

          /*
           * Jika pendapatan final sudah ada, jangan tambahkan lagi.
           */
          writeCell(
            f,
            'Pendapatan Final',
            oldFinal > 0 ? oldFinal : total
          );
        }

        if (status === 'Dibatalkan') {
          writeCell(f, 'Pendapatan Final', 0);
          writeCell(f, 'Sisa Piutang', 0);

          /*
           * Status pembayaran tetap dipertahankan.
           * Pengembalian DP dicatat melalui aksi kembalikanDP.
           */
        }

        writeCell(f, 'Status Pengerjaan', status);

        writeCell(
          f,
          'Catatan Pengerjaan',
          String(d.catatanPengerjaan || d.catatan || '')
        );

        return json({
          ok: true,

          message:
            status === 'Sudah Diambil'
              ? 'Status Sudah Diambil. Pembayaran otomatis Lunas dan pendapatan final dicatat satu kali.'
              : status === 'Dibatalkan'
                ? 'Servis dibatalkan. Status pembayaran asli dipertahankan; catat pengembalian DP setelah uang benar-benar dikembalikan.'
                : 'Status pengerjaan berhasil diperbarui.',

          idServis: sid,
          statusPengerjaan: status,

          statusPembayaran:
            status === 'Sudah Diambil'
              ? 'Lunas'
              : String(f.row[f.col['Status']] || ''),

          totalTagihan: total,

          dibayar:
            status === 'Sudah Diambil'
              ? total
              : paid,

          sisaPiutang:
            status === 'Sudah Diambil' ||
            status === 'Dibatalkan'
              ? 0
              : Math.max(0, total - paid),

          pendapatanFinal:
            status === 'Sudah Diambil'
              ? (oldFinal > 0 ? oldFinal : total)
              : 0
        });

      } finally {
        lock.releaseLock();
      }
    }


    /* -----------------------------------------------
       PEMBAYARAN SERVIS MANUAL

       Aksi ini tetap tersedia untuk kompatibilitas backend.
       Halaman data.html terbaru tidak menyediakan tombol Bayar.
    ------------------------------------------------ */

    if (
      action === 'bayarservis' ||
      action === 'pembayaranservis'
    ) {
      const sid = String(d.idServis || '').trim();
      const amount = num(d.nominal);

      if (!sid || amount <= 0) {
        throw Error(
          'ID servis dan nominal pembayaran yang valid wajib diisi.'
        );
      }

      const lock = LockService.getScriptLock();
      lock.waitLock(15000);

      try {
        const f = findServis(sid);

        const work = String(
          f.row[f.col['Status Pengerjaan']] || 'Masuk'
        );

        if (
          work === 'Dibatalkan' ||
          work === 'Sudah Diambil'
        ) {
          throw Error(
            'Pembayaran manual tidak tersedia untuk servis yang dibatalkan atau sudah diambil.'
          );
        }

        const total = num(
          f.row[f.col['Total Tagihan']]
        );

        const paid = num(
          f.row[f.col['Dibayar']]
        );

        const remaining = Math.max(0, total - paid);

        if (amount > remaining) {
          throw Error(
            'Pembayaran melebihi sisa tagihan.'
          );
        }

        const newPaid = paid + amount;
        const newRemaining = Math.max(0, total - newPaid);

        appendPaymentLog(
          sid,
          String(f.row[f.col['Nama Pelanggan']] || ''),
          amount,
          d.metode || 'Tunai',
          d.catatan || 'Pembayaran servis'
        );

        writeCell(f, 'Dibayar', newPaid);
        writeCell(f, 'Sisa Piutang', newRemaining);

        writeCell(
          f,
          'Status',
          newRemaining === 0 ? 'Lunas' : 'Belum Lunas'
        );

        return json({
          ok: true,
          message: 'Pembayaran servis berhasil dicatat.',
          totalTagihan: total,
          dibayar: newPaid,
          sisaPiutang: newRemaining,
          status: newRemaining === 0 ? 'Lunas' : 'Belum Lunas'
        });

      } finally {
        lock.releaseLock();
      }
    }


    /* -----------------------------------------------
       PENGEMBALIAN DP

       Hanya berlaku untuk servis Dibatalkan.
       Nominal harus sama dengan total uang yang diterima.
       Status ditandai agar pengembalian tidak dicatat dua kali.
    ------------------------------------------------ */

    if (action === 'kembalikanDP') {
      const sid = String(d.idServis || '').trim();

      const amount = num(d.nominal);

      if (!sid) {
        throw Error('ID servis wajib diisi.');
      }

      const lock = LockService.getScriptLock();
      lock.waitLock(15000);

      try {
        const f = findServis(sid);

        const work = String(
          f.row[f.col['Status Pengerjaan']] || 'Masuk'
        );

        if (work !== 'Dibatalkan') {
          throw Error(
            'Pengembalian hanya bisa dilakukan untuk servis berstatus Dibatalkan.'
          );
        }

        const refundStatus = String(
          f.row[f.col['Status Pengembalian DP']] || ''
        );

        if (
          refundStatus === 'Sudah Dikembalikan'
        ) {
          throw Error(
            'Pengembalian DP sudah tercatat. Tidak boleh dicatat dua kali.'
          );
        }

        const paid = num(
          f.row[f.col['Dibayar']]
        );

        if (paid <= 0) {
          throw Error(
            'Tidak ada uang/DP yang tercatat diterima untuk servis ini.'
          );
        }

        if (amount !== paid) {
          throw Error(
            'Pengembalian harus 100% dari uang yang tercatat diterima: Rp' +
            paid.toLocaleString('id-ID')
          );
        }

        /*
         * Tandai lebih dahulu untuk mencegah klik ulang
         * mencatat pengembalian yang sama dua kali.
         */
        writeCell(
          f,
          'Status Pengembalian DP',
          'Sudah Dikembalikan'
        );

        appendRefundLog(
          sid,
          String(f.row[f.col['Nama Pelanggan']] || ''),
          amount,
          d.metode || 'Tunai',
          d.catatan ||
            'Pengembalian penuh DP servis dibatalkan'
        );

        append('Pengeluaran', {
          'ID Pengeluaran': id('EXP'),
          'Tanggal': new Date(),
          'Kategori': 'Pengembalian DP Servis',
          'Nominal': amount,
          'Metode': d.metode || 'Tunai',
          'Supplier': '',
          'Karyawan ID':
            f.row[f.col['Karyawan ID']] || '',
          'Keterangan':
            'Pengembalian penuh DP servis ' +
            sid +
            ' - ' +
            String(f.row[f.col['Nama Pelanggan']] || '') +
            (d.catatan ? ' | ' + d.catatan : '')
        });

        return json({
          ok: true,
          message:
            'Pengembalian DP 100% berhasil dicatat sebagai pengeluaran.',
          idServis: sid,
          nominal: amount,
          statusPengembalianDP: 'Sudah Dikembalikan'
        });

      } finally {
        lock.releaseLock();
      }
    }


    /* -----------------------------------------------
       PENJUALAN
    ------------------------------------------------ */

    if (action === 'createsale') {
      const lock = LockService.getScriptLock();
      lock.waitLock(15000);

      try {
        const items = Array.isArray(d.items)
          ? d.items
          : [];

        const products = rows('Produk');
        const demand = {};

        let total = 0;

        items.forEach(function(item) {
          const code = String(item.kode);
          const qty = num(item.qty);

          if (qty <= 0) {
            throw Error(
              'Jumlah produk harus lebih dari nol.'
            );
          }

          demand[code] = (demand[code] || 0) + qty;
        });

        Object.keys(demand).forEach(function(code) {
          const product = products.find(function(item) {
            return String(item['Kode Produk']) === code;
          });

          if (!product) {
            throw Error(
              'Produk tidak ditemukan: ' + code
            );
          }

          if (num(product.Stok) < demand[code]) {
            throw Error(
              'Stok tidak cukup: ' + product['Nama Produk']
            );
          }
        });

        items.forEach(function(item) {
          const product = products.find(function(p) {
            return String(p['Kode Produk']) ===
              String(item.kode);
          });

          total +=
            num(product['Harga Jual']) * num(item.qty);
        });

        total = Math.max(0, total - num(d.diskon));

        const paid = num(d.dibayar);

        if (paid < 0 || paid > total) {
          throw Error(
            'Nominal pembayaran penjualan tidak valid.'
          );
        }

        const sid = id('JUAL');

        append('Penjualan', {
          'ID Penjualan': sid,
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

        items.forEach(function(item) {
          updateStock(item.kode, -num(item.qty));
        });

        return json({
          ok: true,
          message: 'Penjualan tersimpan',
          idPenjualan: sid,
          total: total,
          sisaPiutang: Math.max(0, total - paid)
        });

      } finally {
        lock.releaseLock();
      }
    }


    return json({
      ok: false,
      message: 'Action POST tidak dikenal: ' + action
    });

  } catch (err) {
    return json({
      ok: false,
      message: err.message || String(err)
    });
  }
}


/* =====================================================
   STOK PRODUK
===================================================== */

function updateStock(code, delta) {
  const sh = sheet('Produk');
  const all = rows('Produk');

  const index = all.findIndex(function(item) {
    return String(item['Kode Produk']) === String(code);
  });

  if (index < 0) {
    throw Error('Produk tidak ditemukan: ' + code);
  }

  const headers = headersOf(sh);
  const stockColumn = headers.indexOf('Stok');

  if (stockColumn < 0) {
    throw Error('Kolom Stok tidak ditemukan.');
  }

  sh.getRange(index + 2, stockColumn + 1)
    .setValue(num(all[index].Stok) + delta);
}


/* =====================================================
   PENGHITUNGAN UPAH
===================================================== */

function calcPayroll(month) {
  const employees = rows('Karyawan');

  const attendance = rows('Absensi').filter(function(item) {
    return fmtDate(item.Tanggal).slice(0, 7) === month;
  });

  const kasbon = rows('Kasbon').filter(function(item) {
    return fmtDate(item.Tanggal).slice(0, 7) === month;
  });

  return employees.map(function(employee) {
    const hadir = attendance.filter(function(item) {
      return (
        String(item['Karyawan ID']) === String(employee.ID) &&
        item.Status === 'Hadir'
      );
    }).length;

    const gross = hadir * num(employee['Upah Harian']);

    const totalKasbon = kasbon
      .filter(function(item) {
        return (
          String(item['Karyawan ID']) === String(employee.ID) &&
          item.Jenis === 'Kasbon'
        );
      })
      .reduce(function(total, item) {
        return total + num(item.Nominal);
      }, 0);

    const paid = kasbon
      .filter(function(item) {
        return (
          String(item['Karyawan ID']) === String(employee.ID) &&
          item.Jenis === 'Pembayaran Upah'
        );
      })
      .reduce(function(total, item) {
        return total + num(item.Nominal);
      }, 0);

    return {
      ID: employee.ID,
      Nama: employee.Nama,
      HariHadir: hadir,
      UpahKotor: gross,
      Kasbon: totalKasbon,
      SudahDibayar: paid,
      SisaUpah: Math.max(0, gross - totalKasbon - paid)
    };
  });
}