/** FIXZY Apps Script Backend */
const SHEETS = {
  Produk: ['Kode Produk','Nama Produk','Kategori','Merek','Supplier','Stok','Stok Minimum','Modal Satuan','Harga Jual'],
  Servis: ['ID Servis','Tanggal','Nama Pelanggan','WhatsApp','Model HP','IMEI','Keluhan','Biaya Jasa','Teknisi','Karyawan ID','Catatan Internal','Sparepart JSON','Total Tagihan','Dibayar','Sisa Piutang','Status'],
  Penjualan: ['ID Penjualan','Tanggal','Nama Pelanggan','WhatsApp','Karyawan ID','Items JSON','Diskon','Total','Dibayar','Sisa Piutang','Metode'],
  Pengeluaran: ['ID Pengeluaran','Tanggal','Kategori','Nominal','Metode','Supplier','Karyawan ID','Keterangan'],
  Karyawan: ['ID','Nama','Telepon','Jabatan','Upah Harian','Status','Dibuat'],
  Absensi: ['ID','Tanggal','Karyawan ID','Nama','Status','Catatan','Jam Masuk','Jam Pulang','Jam Kerja','Rasio Upah'],
  Kasbon: ['ID','Tanggal','Karyawan ID','Nama','Jenis','Nominal','Keterangan'],
  Pengaturan: ['Kunci','Nilai']
};

const FIXZY_SPREADSHEET_ID = '1Dx5wRs39KLFaUStd9NynKJX7j80Zswj66P_dkD_N_Ao';

function ss() {
  return SpreadsheetApp.openById(FIXZY_SPREADSHEET_ID);
}

function setupSheets() {
  const book = ss();

  Object.keys(SHEETS).forEach(function(n) {
    let sh = book.getSheetByName(n);

    if (!sh) {
      sh = book.insertSheet(n);
    }

    if (sh.getLastRow() === 0) {
      sh.appendRow(SHEETS[n]);
    } else {
      const current = sh
        .getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
        .getValues()[0]
        .map(String);

      SHEETS[n].forEach(function(h) {
        if (!current.includes(h)) {
          sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
          current.push(h);
        }
      });
    }
  });

  return {
    ok: true,
    message: 'Sheet FIXZY siap dibuat/diperiksa.'
  };
}

function sheet(n) {
  let s = ss().getSheetByName(n);

  if (!s) {
    setupSheets();
    s = ss().getSheetByName(n);
  }

  return s;
}

function rows(n) {
  const s = sheet(n);
  const v = s.getDataRange().getValues();

  if (v.length < 2) return [];

  const h = v[0];

  return v.slice(1)
    .filter(function(r) {
      return r.some(function(x) {
        return x !== '' && x !== null;
      });
    })
    .map(function(r) {
      return Object.fromEntries(
        h.map(function(k, i) {
          return [k, r[i]];
        })
      );
    });
}

function append(n, obj) {
  const s = sheet(n);
  const h = SHEETS[n];

  s.appendRow(
    h.map(function(k) {
      return obj[k] === undefined ? '' : obj[k];
    })
  );
}

function id(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function json(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    setupSheets();

    const action = String((e.parameter.action || 'ping')).split('&')[0];
    const q = e.parameter;

    switch (action) {
      case 'ping':
        return json({ok:true, message:'FIXZY API siap'});

      case 'stok':
        return json({ok:true, data:rows('Produk')});

      case 'servis':
        return json({ok:true, data:rows('Servis')});

      case 'penjualan':
        return json({ok:true, data:rows('Penjualan')});

      case 'pengeluaran':
        return json({ok:true, data:rows('Pengeluaran')});

      case 'karyawan':
        return json({ok:true, data:rows('Karyawan')});

      case 'pengaturan': {
        const o = {};
        rows('Pengaturan').forEach(function(x) {
          o[x.Kunci] = x.Nilai;
        });
        return json({ok:true, data:o});
      }

      case 'absensi': {
        const bulan = q.bulan ||
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

        const data = rows('Absensi')
          .filter(function(x) {
            return fmtDate(x.Tanggal).slice(0, 7) === bulan;
          })
          .map(function(x) {
            return {...x, Tanggal:fmtDate(x.Tanggal)};
          });

        return json({ok:true, data:data});
      }

      case 'upah': {
        const bulan = q.bulan ||
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

        return json({
          ok:true,
          data:calcPayroll(bulan)
        });
      }

      case 'dashboard': {
        const sv = rows('Servis');
        const pe = rows('Penjualan');
        const pr = rows('Produk');

        return json({
          ok:true,
          ringkasan:{
            jumlahServis:sv.length,
            totalPiutangServis:sv.reduce(function(a,x) {
              return a + num(x['Sisa Piutang']);
            }, 0),
            totalPenjualan:pe.reduce(function(a,x) {
              return a + num(x.Total);
            }, 0),
            stokMenipis:pr.filter(function(x) {
              return num(x.Stok) <= num(x['Stok Minimum']);
            }).length
          }
        });
      }

      case 'pembayaran': {
        const data = rows('Servis').filter(function(x) {
          return num(x['Sisa Piutang']) > 0;
        });

        return json({ok:true, data:data});
      }

      default:
        return json({
          ok:false,
          message:'Action GET tidak dikenal: ' + action
        });
    }
  } catch (err) {
    return json({ok:false, message:err.message});
  }
}

function doPost(e) {
  try {
    setupSheets();

    const d = JSON.parse(e.postData.contents || '{}');
    const a = d.action || '';

    if (a === 'createproduk') {
      const code = d.kodeProduk || id('PRD');

      append('Produk', {
        'Kode Produk':code,
        'Nama Produk':d.namaProduk,
        'Kategori':d.kategori,
        'Merek':d.merek,
        'Supplier':d.supplier,
        'Stok':num(d.stokAwal),
        'Stok Minimum':num(d.stokMinimum),
        'Modal Satuan':num(d.modal),
        'Harga Jual':num(d.hargaJual)
      });

      return json({
        ok:true,
        message:'Produk tersimpan',
        kodeProduk:code
      });
    }

    if (a === 'savekaryawan') {
      const all = rows('Karyawan');
      const kid = d.id || id('KRY');

      const obj = {
        ID:kid,
        Nama:d.nama,
        Telepon:d.telepon,
        Jabatan:d.jabatan,
        'Upah Harian':num(d.upahHarian),
        Status:d.status || 'Aktif',
        Dibuat:new Date()
      };

      const found = all.findIndex(function(x) {
        return String(x.ID) === kid;
      });

      if (found >= 0) {
        const s = sheet('Karyawan');
        const h = SHEETS.Karyawan;

        s.getRange(found + 2, 1, 1, h.length).setValues([
          h.map(function(k) {
            return obj[k] === undefined ? '' : obj[k];
          })
        ]);
      } else {
        append('Karyawan', obj);
      }

      return json({
        ok:true,
        message:'Data karyawan tersimpan',
        id:kid
      });
    }

    // ABSENSI: menyimpan jam kerja dan rasio upah.
    if (a === 'saveabsensi') {
      const emp = rows('Karyawan').find(function(x) {
        return String(x.ID) === String(d.karyawanId);
      });

      if (!emp) throw Error('Karyawan tidak ditemukan');

      const sh = sheet('Absensi');
      const all = rows('Absensi');

      const idx = all.findIndex(function(x) {
        return fmtDate(x.Tanggal) === String(d.tanggal) &&
          String(x['Karyawan ID']) === String(d.karyawanId);
      });

      const jamMasuk = String(d.jamMasuk || '');
      const jamPulang = String(d.jamPulang || '');

      let jamKerja = num(d.jamKerja);

      if (jamMasuk && jamPulang && jamKerja <= 0) {
        const masuk = jamMasuk.split(':').map(Number);
        const pulang = jamPulang.split(':').map(Number);

        let menit =
          (pulang[0] * 60 + pulang[1]) -
          (masuk[0] * 60 + masuk[1]);

        if (menit < 0) menit += 1440;

        jamKerja = Math.max(0, menit / 60);
      }

      const rasio = d.rasioUpah !== undefined
        ? Math.max(0, Math.min(1, num(d.rasioUpah)))
        : Math.min(1, jamKerja / 9);

      const obj = {
        ID:idx >= 0 ? (all[idx].ID || id('ABS')) : id('ABS'),
        Tanggal:d.tanggal,
        'Karyawan ID':d.karyawanId,
        Nama:emp.Nama,
        Status:d.status,
        Catatan:d.catatan || '',
        'Jam Masuk':jamMasuk,
        'Jam Pulang':jamPulang,
        'Jam Kerja':jamKerja,
        'Rasio Upah':rasio
      };

      const headers = sh
        .getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map(String);

      const row = headers.map(function(k) {
        if (obj[k] !== undefined) return obj[k];
        return idx >= 0 ? (all[idx][k] || '') : '';
      });

      if (idx >= 0) {
        sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
      } else {
        sh.appendRow(row);
      }

      return json({ok:true, message:'Absensi tersimpan'});
    }

    if (a === 'savekasbon') {
      const emp = rows('Karyawan').find(function(x) {
        return String(x.ID) === String(d.karyawanId);
      });

      if (!emp) throw Error('Karyawan tidak ditemukan');

      append('Kasbon', {
        ID:id('KAS'),
        Tanggal:d.tanggal,
        'Karyawan ID':d.karyawanId,
        Nama:emp.Nama,
        Jenis:d.jenis,
        Nominal:num(d.nominal),
        Keterangan:d.keterangan
      });

      return json({
        ok:true,
        message:'Kasbon/pembayaran tersimpan'
      });
    }

    if (a === 'savepengaturan') {
      const sh = sheet('Pengaturan');

      sh.clearContents();
      sh.appendRow(['Kunci','Nilai']);

      Object.keys(d).filter(function(k) {
        return k !== 'action';
      }).forEach(function(k) {
        sh.appendRow([k, d[k]]);
      });

      return json({ok:true, message:'Pengaturan tersimpan'});
    }

    if (a === 'createexpense') {
      const eid = id('EXP');

      append('Pengeluaran', {
        'ID Pengeluaran':eid,
        Tanggal:new Date(),
        Kategori:d.kategori,
        Nominal:num(d.nominal),
        Metode:d.metode,
        Supplier:d.supplier,
        'Karyawan ID':d.karyawanId,
        Keterangan:d.keterangan
      });

      if (d.kategori === 'Kasbon Karyawan' && d.karyawanId) {
        const emp = rows('Karyawan').find(function(x) {
          return String(x.ID) === String(d.karyawanId);
        });

        append('Kasbon', {
          ID:id('KAS'),
          Tanggal:new Date(),
          'Karyawan ID':d.karyawanId,
          Nama:emp ? emp.Nama : '',
          Jenis:'Kasbon',
          Nominal:num(d.nominal),
          Keterangan:d.keterangan
        });
      }

      return json({
        ok:true,
        message:'Pengeluaran tersimpan',
        id:eid
      });
    }

    if (a === 'bayarservis' || a === 'pembayaranservis') {
      const idServis = String(d.idServis || '').trim();
      const nominal = num(d.nominal);

      if (!idServis) throw Error('ID servis wajib diisi.');
      if (nominal <= 0) throw Error('Nominal pembayaran harus lebih dari nol.');

      const lock = LockService.getScriptLock();
      lock.waitLock(15000);

      try {
        const sh = sheet('Servis');
        const values = sh.getDataRange().getValues();

        if (values.length < 2) throw Error('Data servis belum tersedia.');

        const headers = values[0].map(String);
        const col = {};

        headers.forEach(function(h, i) {
          col[h] = i;
        });

        ['ID Servis','Total Tagihan','Dibayar','Sisa Piutang','Status'].forEach(function(h) {
          if (col[h] === undefined) {
            throw Error('Kolom servis tidak ditemukan: ' + h);
          }
        });

        let rowIndex = -1;
        let service = null;

        for (let i = 1; i < values.length; i++) {
          if (String(values[i][col['ID Servis']]) === idServis) {
            rowIndex = i + 1;
            service = values[i];
            break;
          }
        }

        if (rowIndex < 0) throw Error('Data servis ' + idServis + ' tidak ditemukan.');

        const total = num(service[col['Total Tagihan']]);
        const dibayarLama = num(service[col['Dibayar']]);
        const sisaLama = Math.max(0, total - dibayarLama);

        if (sisaLama <= 0) throw Error('Tagihan servis ini sudah lunas.');

        if (nominal > sisaLama) {
          throw Error('Pembayaran melebihi sisa tagihan Rp' +
            sisaLama.toLocaleString('id-ID') + '.');
        }

        const dibayarBaru = dibayarLama + nominal;
        const sisaBaru = Math.max(0, total - dibayarBaru);
        const statusBaru = sisaBaru === 0 ? 'Lunas' : 'Belum Lunas';

        const book = ss();
        let log = book.getSheetByName('PembayaranServis');

        const logHeaders = [
          'ID Pembayaran','Tanggal','ID Servis',
          'Nama Pelanggan','Nominal','Metode','Catatan'
        ];

        if (!log) {
          log = book.insertSheet('PembayaranServis');
          log.appendRow(logHeaders);
        } else if (log.getLastRow() === 0) {
          log.appendRow(logHeaders);
        } else {
          const existing = log.getRange(1, 1, 1, log.getLastColumn())
            .getValues()[0].map(String);

          logHeaders.forEach(function(h) {
            if (!existing.includes(h)) {
              log.getRange(1, log.getLastColumn() + 1).setValue(h);
              existing.push(h);
            }
          });
        }

        const logMap = {};

        log.getRange(1, 1, 1, log.getLastColumn())
          .getValues()[0].forEach(function(h, i) {
            logMap[String(h)] = i;
          });

        const logRow = new Array(log.getLastColumn()).fill('');

        const logData = {
          'ID Pembayaran':id('BYR'),
          Tanggal:new Date(),
          'ID Servis':idServis,
          'Nama Pelanggan':col['Nama Pelanggan'] !== undefined
            ? service[col['Nama Pelanggan']] : '',
          Nominal:nominal,
          Metode:d.metode || 'Tunai',
          Catatan:d.catatan || ''
        };

        Object.keys(logData).forEach(function(k) {
          if (logMap[k] !== undefined) logRow[logMap[k]] = logData[k];
        });

        log.appendRow(logRow);

        sh.getRange(rowIndex, col['Dibayar'] + 1).setValue(dibayarBaru);
        sh.getRange(rowIndex, col['Sisa Piutang'] + 1).setValue(sisaBaru);
        sh.getRange(rowIndex, col['Status'] + 1).setValue(statusBaru);

        return json({
          ok:true,
          message:'Pembayaran servis berhasil dicatat.',
          idServis:idServis,
          totalTagihan:total,
          pembayaranTambahan:nominal,
          totalDibayar:dibayarBaru,
          sisaPiutang:sisaBaru,
          status:statusBaru
        });
      } finally {
        lock.releaseLock();
      }
    }

    if (a === 'createservis') {
      const parts = Array.isArray(d.spareparts) ? d.spareparts : [];
      const products = rows('Produk');
      let partTotal = 0;

      parts.forEach(function(p) {
        const ix = products.findIndex(function(x) {
          return String(x['Kode Produk']) === String(p.kode);
        });

        if (ix < 0) throw Error('Sparepart tidak ditemukan: ' + p.kode);

        const pr = products[ix];
        const qty = num(p.qty);

        if (num(pr.Stok) < qty) {
          throw Error('Stok tidak cukup untuk ' + pr['Nama Produk']);
        }

        partTotal += qty * num(pr['Harga Jual']);
        updateStock(p.kode, -qty);
      });

      const total = num(d.biayaJasa) + partTotal;
      const paid = num(d.dibayar);
      const sid = id('SRV');

      append('Servis', {
        'ID Servis':sid,
        Tanggal:new Date(),
        'Nama Pelanggan':d.namaPelanggan,
        WhatsApp:d.wa,
        'Model HP':d.model,
        IMEI:d.imei,
        Keluhan:d.keluhan,
        'Biaya Jasa':num(d.biayaJasa),
        Teknisi:d.teknisi,
        'Karyawan ID':d.karyawanId,
        'Catatan Internal':d.catatan,
        'Sparepart JSON':JSON.stringify(parts),
        'Total Tagihan':total,
        Dibayar:paid,
        'Sisa Piutang':Math.max(0, total - paid),
        Status:paid >= total ? 'Lunas' : 'Belum Lunas'
      });

      return json({
        ok:true,
        message:'Servis tersimpan',
        idServis:sid,
        total:total,
        sisaPiutang:Math.max(0, total - paid)
      });
    }

    if (a === 'createsale') {
      const items = Array.isArray(d.items) ? d.items : [];
      let total = 0;

      items.forEach(function(it) {
        const pr = rows('Produk').find(function(x) {
          return String(x['Kode Produk']) === String(it.kode);
        });

        if (!pr) throw Error('Produk tidak ditemukan');

        const qty = num(it.qty);

        if (num(pr.Stok) < qty) {
          throw Error('Stok tidak cukup: ' + pr['Nama Produk']);
        }

        total += num(pr['Harga Jual']) * qty;
        updateStock(it.kode, -qty);
      });

      total = Math.max(0, total - num(d.diskon));

      const paid = num(d.dibayar);
      const sid = id('JUAL');

      append('Penjualan', {
        'ID Penjualan':sid,
        Tanggal:new Date(),
        'Nama Pelanggan':d.namaPelanggan,
        WhatsApp:d.wa,
        'Karyawan ID':d.karyawanId,
        'Items JSON':JSON.stringify(items),
        Diskon:num(d.diskon),
        Total:total,
        Dibayar:paid,
        'Sisa Piutang':Math.max(0, total - paid),
        Metode:d.metode
      });

      return json({
        ok:true,
        message:'Penjualan tersimpan',
        idPenjualan:sid,
        total:total,
        sisaPiutang:Math.max(0, total - paid)
      });
    }

    return json({
      ok:false,
      message:'Action POST tidak dikenal: ' + a
    });

  } catch (err) {
    return json({ok:false, message:err.message});
  }
}

function updateStock(code, delta) {
  const s = sheet('Produk');
  const all = rows('Produk');

  const i = all.findIndex(function(x) {
    return String(x['Kode Produk']) === String(code);
  });

  if (i < 0) throw Error('Produk tidak ditemukan');

  s.getRange(i + 2, 6).setValue(num(all[i].Stok) + delta);
}

function num(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function fmtDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(
      v,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  return String(v || '').slice(0, 10);
}

/**
 * Perhitungan upah berdasarkan jam kerja, kasbon, pembayaran,
 * dan tanggal awal periode dari Pengaturan.
 */
function calcPayroll(month) {
  const emp = rows('Karyawan');

  const settings = {};
  rows('Pengaturan').forEach(function(x) {
    settings[String(x.Kunci)] = x.Nilai;
  });

  const startDay = Math.max(
    1,
    Math.min(28, Number(settings.tanggalMulai || 1))
  );

  const parts = String(month).split('-').map(Number);
  const year = parts[0];
  const mo = parts[1];

  if (!year || !mo || mo < 1 || mo > 12) {
    throw Error('Bulan penggajian tidak valid.');
  }

  let start;
  let end;

  if (startDay === 1) {
    start = new Date(year, mo - 1, 1);
    end = new Date(year, mo, 0);
  } else {
    start = new Date(year, mo - 2, startDay);
    end = new Date(year, mo - 1, startDay - 1);
  }

  const iso = function(d) {
    return Utilities.formatDate(
      d,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  };

  const startISO = iso(start);
  const endISO = iso(end);

  const inPeriod = function(x) {
    const d = fmtDate(x.Tanggal);
    return d >= startISO && d <= endISO;
  };

  const att = rows('Absensi').filter(inPeriod);
  const kas = rows('Kasbon').filter(inPeriod);

  return emp
    .filter(function(e) {
      return String(e.Status || 'Aktif').toLowerCase() !== 'nonaktif' &&
        String(e.Status || 'Aktif').toLowerCase() !== 'tidak aktif';
    })
    .map(function(e) {
      const hadir = att.filter(function(x) {
        return String(x['Karyawan ID']) === String(e.ID) &&
          String(x.Status).toLowerCase() === 'hadir';
      });

      const totalJam = hadir.reduce(function(sum, x) {
        return sum + num(x['Jam Kerja']);
      }, 0);

      const setaraHari = hadir.reduce(function(sum, x) {
        let ratio = x['Rasio Upah'];

        if (ratio === '' || ratio === null || ratio === undefined) {
          ratio = num(x['Jam Kerja']) / 9;
        } else {
          ratio = num(ratio);
        }

        return sum + Math.max(0, Math.min(1, ratio));
      }, 0);

      const upahHarian = num(e['Upah Harian']);
      const gross = Math.round(setaraHari * upahHarian);

      const totalKasbon = kas
        .filter(function(x) {
          return String(x['Karyawan ID']) === String(e.ID) &&
            String(x.Jenis).toLowerCase() === 'kasbon';
        })
        .reduce(function(sum, x) {
          return sum + num(x.Nominal);
        }, 0);

      const paid = kas
        .filter(function(x) {
          return String(x['Karyawan ID']) === String(e.ID) &&
            String(x.Jenis).toLowerCase() === 'pembayaran upah';
        })
        .reduce(function(sum, x) {
          return sum + num(x.Nominal);
        }, 0);

      return {
        ID:e.ID,
        Nama:e.Nama,
        UpahHarian:upahHarian,
        HariHadir:hadir.length,
        TotalJamKerja:Math.round(totalJam * 100) / 100,
        SetaraHariKerja:Math.round(setaraHari * 10000) / 10000,
        UpahKotor:gross,
        Kasbon:totalKasbon,
        SudahDibayar:paid,
        SisaUpah:Math.max(0, gross - totalKasbon - paid),
        PeriodeMulai:startISO,
        PeriodeSelesai:endISO
      };
    });
}