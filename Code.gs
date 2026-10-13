/**
 * FIXZY - Pembukuan Servis & Penjualan
 * Google Apps Script backend / Google Spreadsheet database
 * Kerangka awal: sesuaikan SPREADSHEET_ID sebelum deploy.
 */
const SPREADSHEET_ID = 'GANTI_DENGAN_ID_GOOGLE_SPREADSHEET_FIXZY';
const TZ = 'Asia/Jakarta';
const SHEETS = {
  SETTINGS: 'PENGATURAN', SERVIS: 'SERVIS', JASA: 'DETAIL_JASA', DETAIL_SERVIS: 'DETAIL_SPAREPART',
  PEMBAYARAN: 'PEMBAYARAN', PRODUK: 'STOK', PENJUALAN: 'PENJUALAN', DETAIL_JUAL: 'DETAIL_PENJUALAN',
  MUTASI: 'MUTASI_STOK', PENGELUARAN: 'PENGELUARAN', KARYAWAN: 'KARYAWAN', ABSENSI: 'ABSENSI',
  UPAH: 'UPAH', LOG: 'LOG_AKTIVITAS'
};
const HEADERS = {
  PENGATURAN: ['Kunci','Nilai'],
  SERVIS: ['ID Servis','Tanggal','Nama Pelanggan','WhatsApp','Merek/Model','IMEI/Serial','Keluhan','Catatan','Teknisi','Status Servis','Biaya Jasa','Total Sparepart','Total Tagihan','Total Dibayar','Sisa Piutang','Modal Sparepart','Laba Kotor Estimasi','Dibuat Oleh','Diperbarui'],
  DETAIL_JASA: ['ID Detail','ID Servis','Nama Jasa','Biaya'],
  DETAIL_SPAREPART: ['ID Detail','ID Servis','Kode Produk','Nama Produk','Qty','Modal Satuan','Harga Jual Satuan','Subtotal Modal','Subtotal Jual'],
  PEMBAYARAN: ['ID Pembayaran','Tanggal','Jenis Transaksi','ID Referensi','Nama Pelanggan','Metode','Nominal','Catatan','Petugas','Status'],
  STOK: ['Kode Produk','Nama Produk','Kategori','Merek','Supplier','Stok','Stok Minimum','Modal Satuan','Harga Jual','Aktif','Diperbarui'],
  PENJUALAN: ['ID Penjualan','Tanggal','Nama Pelanggan','WhatsApp','Subtotal','Diskon','Total','Total Dibayar','Sisa Piutang','Petugas','Status'],
  DETAIL_PENJUALAN: ['ID Detail','ID Penjualan','Kode Produk','Nama Produk','Qty','Modal Satuan','Harga Jual Satuan','Subtotal Modal','Subtotal Jual'],
  MUTASI_STOK: ['ID Mutasi','Tanggal','Kode Produk','Nama Produk','Jenis Mutasi','Qty Masuk','Qty Keluar','Referensi','Catatan','Petugas'],
  PENGELUARAN: ['ID Pengeluaran','Tanggal','Kategori','Nominal','Metode','Keterangan','Supplier/Penerima','Petugas','Jenis Akuntansi','Status'],
  KARYAWAN: ['ID Karyawan','Nama','WhatsApp','Peran','Upah Harian','Aktif'],
  ABSENSI: ['ID Absensi','Tanggal','ID Karyawan','Nama','Jam Masuk','Jam Pulang','Status','Catatan'],
  UPAH: ['ID Upah','Periode','ID Karyawan','Nama','Hari Hadir','Tarif','Bonus','Potongan','Total','Dibayar','Sisa','Status'],
  LOG_AKTIVITAS: ['Waktu','Pengguna','Aksi','Referensi','Keterangan']
};

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = String(p.action || 'ping').toLowerCase();
    if (action === 'ping') return json_({ok:true, app:'FIXZY', message:'API FIXZY aktif', time:new Date().toISOString()});
    if (action === 'setup') return json_({ok:true, message:'Gunakan fungsi setupFixzy() dari editor Apps Script untuk membuat sheet.'});
    if (action === 'dashboard') return json_(dashboard_());
    if (action === 'servis') return json_({ok:true,data:listObjects_(SHEETS.SERVIS, 300)});
    if (action === 'produk' || action === 'stok') return json_({ok:true,data:listObjects_(SHEETS.PRODUK, 1000)});
    if (action === 'penjualan') return json_({ok:true,data:listObjects_(SHEETS.PENJUALAN, 300)});
    if (action === 'pengeluaran') return json_({ok:true,data:listObjects_(SHEETS.PENGELUARAN, 500)});
    if (action === 'pembayaran') return json_({ok:true,data:listObjects_(SHEETS.PEMBAYARAN, 500)});
    return json_({ok:false,message:'Action GET tidak dikenal: '+action});
  } catch(err) { return json_({ok:false,message:err.message}); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const d = parseBody_(e);
    const a = String(d.action || '').toLowerCase();
    let result;
    switch(a) {
      case 'createservis': result = createServis_(d); break;
      case 'payment': result = addPayment_(d); break;
      case 'updatestatus': result = updateServisStatus_(d); break;
      case 'createproduk': result = createProduk_(d); break;
      case 'stockin': result = stockMove_(d, 'MASUK'); break;
      case 'createsale': result = createSale_(d); break;
      case 'createexpense': result = createExpense_(d); break;
      case 'checkin': result = attendance_(d, 'MASUK'); break;
      case 'checkout': result = attendance_(d, 'PULANG'); break;
      default: throw new Error('Action POST tidak dikenal: '+a);
    }
    return json_({ok:true,...result});
  } catch(err) { return json_({ok:false,message:err.message}); }
  finally { try { lock.releaseLock(); } catch(_) {} }
}

/** Jalankan sekali dari editor Apps Script setelah SPREADSHEET_ID diisi. */
function setupFixzy() {
  const ss = spreadsheet_();
  Object.keys(HEADERS).forEach(k => {
    const name = SHEETS[k]; let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = HEADERS[k];
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
    else {
      const existing = sh.getRange(1,1,1,headers.length).getDisplayValues()[0];
      if (headers.some((h,i) => existing[i] !== h)) throw new Error('Header sheet '+name+' berbeda. Periksa sebelum setup agar data tidak tertimpa.');
    }
    sh.setFrozenRows(1); sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  });
  const set = ss.getSheetByName(SHEETS.SETTINGS);
  if (set.getLastRow() < 2) set.getRange(2,1,5,2).setValues([
    ['NAMA_USAHA','FIXZY - Pembukuan Servis'],['MATA_UANG','IDR'],['TZ',TZ],['NOTA_PREFIX','SRV'],['VERSI','1.0.0-kerangka']
  ]);
  return 'Setup FIXZY selesai: '+ss.getUrl();
}

function createServis_(d) {
  const customer = clean_(d.namaPelanggan); const model = clean_(d.model); const keluhan = clean_(d.keluhan);
  if (!customer || !model || !keluhan) throw new Error('Nama pelanggan, model HP, dan keluhan wajib diisi.');
  const id = nextId_('SRV', SHEETS.SERVIS, 1);
  const jasa = Math.max(0, num_(d.biayaJasa));
  const parts = Array.isArray(d.spareparts) ? d.spareparts : [];
  let totalPart=0, modalPart=0;
  parts.forEach(p => {
    const qty = Math.max(0, num_(p.qty)); if (!qty) return;
    const prod = findProduct_(clean_(p.kode));
    if (!prod) throw new Error('Produk tidak ditemukan: '+p.kode);
    if (num_(prod.Stok) < qty) throw new Error('Stok tidak cukup untuk '+prod['Nama Produk']+'. Tersedia '+prod.Stok);
    const modal = num_(prod['Modal Satuan']); const jual = p.hargaJual !== undefined ? num_(p.hargaJual) : num_(prod['Harga Jual']);
    totalPart += qty*jual; modalPart += qty*modal;
    append_(SHEETS.DETAIL_SERVIS, [nextId_('DSP',SHEETS.DETAIL_SERVIS,0),id,prod['Kode Produk'],prod['Nama Produk'],qty,modal,jual,qty*modal,qty*jual]);
    stockMoveRow_(prod, 'PAKAI SERVIS', 0, qty, id, 'Pemakaian sparepart servis', actor_(d));
    updateProductStock_(prod['Kode Produk'], num_(prod.Stok)-qty);
  });
  const total = jasa + totalPart; const now = new Date();
  append_(SHEETS.SERVIS,[id,now,customer,clean_(d.wa),model,clean_(d.imei),keluhan,clean_(d.catatan),clean_(d.teknisi),'DITERIMA',jasa,totalPart,total,0,total,modalPart,total-modalPart,actor_(d),now]);
  if (jasa > 0) append_(SHEETS.DETAIL_JASA,[nextId_('DJS',SHEETS.JASA,0),id,'Jasa servis',jasa]);
  log_(actor_(d),'CREATE_SERVIS',id,'Tagihan '+total);
  return {idServis:id,totalTagihan:total,sisaPiutang:total,message:'Servis berhasil disimpan.'};
}

function addPayment_(d) {
  const ref = clean_(d.idReferensi); const nominal = num_(d.nominal);
  if (!ref || nominal <= 0) throw new Error('Referensi transaksi dan nominal pembayaran wajib diisi.');
  const kind = String(d.jenisTransaksi || 'SERVIS').toUpperCase();
  const sheet = kind === 'PENJUALAN' ? SHEETS.PENJUALAN : SHEETS.SERVIS;
  const obj = findById_(sheet, ref); if (!obj) throw new Error('Transaksi tidak ditemukan: '+ref);
  const total = num_(obj['Total Tagihan'] !== undefined ? obj['Total Tagihan'] : obj.Total);
  const paid = num_(obj['Total Dibayar']); if (paid + nominal > total) throw new Error('Pembayaran melebihi sisa tagihan.');
  const now = new Date(); const payId = nextId_('PAY',SHEETS.PEMBAYARAN,1);
  append_(SHEETS.PEMBAYARAN,[payId,now,kind,ref,obj['Nama Pelanggan']||'',clean_(d.metode||'Tunai'),nominal,clean_(d.catatan),actor_(d),'SAH']);
  const newPaid=paid+nominal, balance=total-newPaid;
  updateObject_(sheet,ref,{'Total Dibayar':newPaid,'Sisa Piutang':balance,'Status':balance===0?'LUNAS':(newPaid>0?'DP':'BELUM BAYAR'),'Diperbarui':now});
  if (sheet===SHEETS.SERVIS) updateObject_(sheet,ref,{'Total Dibayar':newPaid,'Sisa Piutang':balance,'Diperbarui':now});
  log_(actor_(d),'PAYMENT',ref,'Pembayaran '+nominal+'; sisa '+balance);
  return {idPembayaran:payId,totalDibayar:newPaid,sisaPiutang:balance,message:'Pembayaran berhasil dicatat.'};
}

function updateServisStatus_(d) {
  const id=clean_(d.idServis), status=clean_(d.status); const allowed=['DITERIMA','DIKERJAKAN','MENUNGGU_SPAREPART','SELESAI','DIAMBIL','DIBATALKAN'];
  if (!allowed.includes(status)) throw new Error('Status servis tidak valid.');
  const ok=updateObject_(SHEETS.SERVIS,id,{'Status Servis':status,'Diperbarui':new Date()}); if (!ok) throw new Error('Servis tidak ditemukan.');
  log_(actor_(d),'UPDATE_STATUS',id,status); return {message:'Status servis diperbarui.'};
}

function createProduk_(d) {
  const code=clean_(d.kodeProduk)||nextId_('PRD',SHEETS.PRODUK,1); const name=clean_(d.namaProduk);
  if (!name) throw new Error('Nama produk wajib diisi.'); if (findProduct_(code)) throw new Error('Kode produk sudah digunakan.');
  const row=[code,name,clean_(d.kategori),clean_(d.merek),clean_(d.supplier),Math.max(0,num_(d.stokAwal)),Math.max(0,num_(d.stokMinimum)),Math.max(0,num_(d.modal)),Math.max(0,num_(d.hargaJual)),true,new Date()];
  append_(SHEETS.PRODUK,row); if (num_(d.stokAwal)>0) stockMoveRow_({ 'Kode Produk':code,'Nama Produk':name },'STOK AWAL',num_(d.stokAwal),0,code,'Stok awal produk',actor_(d));
  log_(actor_(d),'CREATE_PRODUK',code,name); return {kodeProduk:code,message:'Produk berhasil ditambahkan.'};
}

function stockMove_(d, kind) {
  const code=clean_(d.kodeProduk), qty=num_(d.qty); if (!code || qty<=0) throw new Error('Produk dan jumlah wajib diisi.');
  const prod=findProduct_(code); if (!prod) throw new Error('Produk tidak ditemukan.');
  if (kind==='MASUK') {
    updateProductStock_(code,num_(prod.Stok)+qty); stockMoveRow_(prod,clean_(d.jenisMutasi||'PEMBELIAN'),qty,0,clean_(d.referensi),clean_(d.catatan),actor_(d));
    if (String(d.jenisMutasi||'PEMBELIAN').toUpperCase()==='PEMBELIAN' && num_(d.totalBiaya)>0) createExpense_({kategori:'Modal Belanja Stok',nominal:num_(d.totalBiaya),metode:d.metode,keterangan:d.catatan||('Pembelian '+prod['Nama Produk']),jenisAkuntansi:'PERSEDIAAN',petugas:actor_(d)});
  } else {
    if (num_(prod.Stok)<qty) throw new Error('Stok tidak cukup.'); updateProductStock_(code,num_(prod.Stok)-qty); stockMoveRow_(prod,'KELUAR',0,qty,clean_(d.referensi),clean_(d.catatan),actor_(d));
  }
  return {message:'Mutasi stok berhasil dicatat.'};
}

function createSale_(d) {
  const items=Array.isArray(d.items)?d.items:[]; if (!items.length) throw new Error('Tambahkan minimal satu barang.');
  const id=nextId_('JUAL',SHEETS.PENJUALAN,1); let subtotal=0, cost=0; const detail=[];
  items.forEach(it=>{const prod=findProduct_(clean_(it.kodeProduk)); const qty=num_(it.qty); if(!prod||qty<=0) throw new Error('Barang atau jumlah tidak valid.'); if(num_(prod.Stok)<qty) throw new Error('Stok tidak cukup: '+prod['Nama Produk']); const modal=num_(prod['Modal Satuan']), price=it.hargaJual===undefined?num_(prod['Harga Jual']):num_(it.hargaJual); subtotal+=qty*price; cost+=qty*modal; detail.push({prod,qty,modal,price});});
  const discount=Math.max(0,num_(d.diskon)); const total=Math.max(0,subtotal-discount); const paid=Math.min(total,Math.max(0,num_(d.dibayar))); const now=new Date();
  append_(SHEETS.PENJUALAN,[id,now,clean_(d.namaPelanggan),clean_(d.wa),subtotal,discount,total,paid,total-paid,actor_(d),(total-paid===0?'LUNAS':paid>0?'DP':'BELUM BAYAR')]);
  detail.forEach(x=>{append_(SHEETS.DETAIL_JUAL,[nextId_('DJL',SHEETS.DETAIL_JUAL,0),id,x.prod['Kode Produk'],x.prod['Nama Produk'],x.qty,x.modal,x.price,x.qty*x.modal,x.qty*x.price]); updateProductStock_(x.prod['Kode Produk'],num_(x.prod.Stok)-x.qty); stockMoveRow_(x.prod,'PENJUALAN',0,x.qty,id,'Penjualan barang',actor_(d));});
  if(paid>0) append_(SHEETS.PEMBAYARAN,[nextId_('PAY',SHEETS.PEMBAYARAN,1),now,'PENJUALAN',id,clean_(d.namaPelanggan),clean_(d.metode||'Tunai'),paid,'Pembayaran saat penjualan',actor_(d),'SAH']);
  log_(actor_(d),'CREATE_SALE',id,'Total '+total); return {idPenjualan:id,subtotal:subtotal,total:total,dibayar:paid,sisaPiutang:total-paid,labaKotorEstimasi:total-cost-discount,message:'Penjualan berhasil disimpan.'};
}

function createExpense_(d) {
  const amount=num_(d.nominal); const category=clean_(d.kategori); if(amount<=0||!category) throw new Error('Kategori dan nominal pengeluaran wajib diisi.');
  const id=nextId_('EXP',SHEETS.PENGELUARAN,1); const now=new Date();
  append_(SHEETS.PENGELUARAN,[id,now,category,amount,clean_(d.metode||'Tunai'),clean_(d.keterangan),clean_(d.supplier),actor_(d),clean_(d.jenisAkuntansi||expenseType_(category)),'SAH']);
  log_(actor_(d),'CREATE_EXPENSE',id,category+' '+amount); return {idPengeluaran:id,message:'Pengeluaran berhasil dicatat.'};
}

function attendance_(d, mode) {
  const id=clean_(d.idKaryawan); if(!id) throw new Error('Pilih karyawan.'); const emp=findById_(SHEETS.KARYAWAN,id); if(!emp) throw new Error('Karyawan tidak ditemukan.');
  const today=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd'); const sh=sheet_(SHEETS.ABSENSI); const rows=sh.getDataRange().getValues(); let rowNum=0;
  for(let i=1;i<rows.length;i++){if(String(rows[i][1])===today&&String(rows[i][2])===id){rowNum=i+1;break;}}
  if(mode==='MASUK'){if(rowNum) throw new Error('Karyawan sudah absen masuk hari ini.'); append_(SHEETS.ABSENSI,[nextId_('ABS',SHEETS.ABSENSI,1),today,id,emp.Nama,new Date(),'','HADIR','']);}
  else {if(!rowNum) throw new Error('Belum ada absen masuk hari ini.'); sh.getRange(rowNum,6).setValue(new Date());}
  log_(actor_(d),'ABSENSI_'+mode,id,emp.Nama); return {message:'Absensi '+mode.toLowerCase()+' berhasil.'};
}

function dashboard_() {
  const servis=listObjects_(SHEETS.SERVIS,10000), sales=listObjects_(SHEETS.PENJUALAN,10000), exp=listObjects_(SHEETS.PENGELUARAN,10000), products=listObjects_(SHEETS.PRODUK,10000);
  const sum=(arr,key)=>arr.reduce((s,o)=>s+num_(o[key]),0);
  return {ok:true,ringkasan:{jumlahServis:servis.length,servisBelumLunas:servis.filter(x=>num_(x['Sisa Piutang'])>0).length,totalPiutangServis:sum(servis,'Sisa Piutang'),totalPenjualan:sum(sales,'Total'),totalPengeluaran:sum(exp.filter(x=>x['Jenis Akuntansi']!=='PERSEDIAAN'),'Nominal'),jumlahProduk:products.length,stokMenipis:products.filter(x=>num_(x.Stok)<=num_(x['Stok Minimum'])).length}};
}

function spreadsheet_(){ if(!SPREADSHEET_ID||SPREADSHEET_ID.indexOf('GANTI_')===0) throw new Error('Isi SPREADSHEET_ID terlebih dahulu di Code.gs.'); return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sheet_(name){const ss=spreadsheet_();let sh=ss.getSheetByName(name);if(!sh)throw new Error('Sheet '+name+' belum ada. Jalankan setupFixzy().');return sh;}
function append_(name,row){sheet_(name).appendRow(row);}
function listObjects_(name,limit){const sh=sheet_(name), values=sh.getDataRange().getValues(); if(values.length<2)return [];const h=values[0];return values.slice(1).filter(r=>r.some(v=>v!==''&&v!==null)).slice(-limit).reverse().map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i] instanceof Date?Utilities.formatDate(r[i],TZ,'yyyy-MM-dd HH:mm:ss'):r[i]);return o;});}
function findById_(name,id){return listObjects_(name,10000).find(o=>String(Object.values(o)[0])===String(id))||null;}
function findProduct_(code){return listObjects_(SHEETS.PRODUK,10000).find(o=>String(o['Kode Produk'])===String(code))||null;}
function updateObject_(name,id,updates){const sh=sheet_(name), vals=sh.getDataRange().getValues();if(vals.length<2)return false;const headers=vals[0], keyIndex=0;for(let i=1;i<vals.length;i++){if(String(vals[i][keyIndex])===String(id)){Object.keys(updates).forEach(k=>{const c=headers.indexOf(k);if(c>=0)sh.getRange(i+1,c+1).setValue(updates[k]);});return true;}}return false;}
function updateProductStock_(code,stock){const sh=sheet_(SHEETS.PRODUK),vals=sh.getDataRange().getValues();for(let i=1;i<vals.length;i++){if(String(vals[i][0])===String(code)){sh.getRange(i+1,6).setValue(stock);sh.getRange(i+1,11).setValue(new Date());return;}}throw new Error('Produk tidak ditemukan: '+code);}
function stockMoveRow_(prod,kind,inQty,outQty,ref,note,actor){append_(SHEETS.MUTASI,[nextId_('MUT',SHEETS.MUTASI,1),new Date(),prod['Kode Produk'],prod['Nama Produk'],kind,inQty,outQty,ref,note,actor]);}
function nextId_(prefix,sheetName,width){const sh=sheet_(sheetName);const n=Math.max(0,sh.getLastRow()-1)+1;return prefix+'-'+Utilities.formatDate(new Date(),TZ,'yyMMdd')+'-'+String(n).padStart(width||4,'0');}
function actor_(d){return clean_(d.petugas||d.user||'OWNER');}
function expenseType_(category){const c=String(category).toLowerCase();return (c.indexOf('modal belanja')>=0||c.indexOf('stok')>=0||c.indexOf('persediaan')>=0)?'PERSEDIAAN':(c.indexOf('alat')>=0?'ASET':'OPERASIONAL');}
function log_(user,action,ref,detail){append_(SHEETS.LOG,[new Date(),user,action,ref,detail]);}
function parseBody_(e){if(!e||!e.postData||!e.postData.contents)throw new Error('Body permintaan kosong.');const raw=e.postData.contents;try{return JSON.parse(raw);}catch(_){return e.parameter||{};}}
function clean_(v){return String(v==null?'':v).trim();}
function num_(v){if(typeof v==='number')return isFinite(v)?v:0;const n=Number(String(v==null?'':v).replace(/[^0-9,.-]/g,'').replace(/\.(?=.*\.)/g,'').replace(',','.'));return isFinite(n)?n:0;}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
