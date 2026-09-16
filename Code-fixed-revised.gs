/*************************************************
 * HOLDTIGHT VOL.4 — SISTEM TIKET
 * Google Apps Script + Google Sheets
 * ------------------------------------------------
 * SETUP (10 menit, panduan lengkap di PANDUAN-SETUP.md):
 * 1. Buat Google Sheet baru (nama bebas, misal "HT4 Tiket")
 * 2. Extensions > Apps Script > hapus isi > paste file ini > Save
 * 3. Deploy > New deployment > type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy Web App URL -> tempel di website (TICKET_API)
 * ------------------------------------------------
 * URL penting setelah deploy (ganti URL dengan Web App URL, dan KEY dengan secret lo sendiri):
 * - Panel  : URL?page=panel&key=KEY_RAHASIA_LO
 * - Tiket  : URL?page=ticket&code=HT4-ABCD
 * (link lama ?page=admin & ?page=scan tetap jalan, otomatis ke tab yang bener)
 *
 * WAJIB SEBELUM DEPLOY — set SECRET (kunci panel), JANGAN ditulis di kode ini:
 * 1. Buka editor Apps Script > ikon gerigi "Project Settings" di kiri > scroll ke "Script Properties"
 * 2. Add script property > key: SECRET, value: string acak panjang (min 20 karakter, contoh: HT4-x7Qp9zL2mK4vR8wD)
 * 3. Save. Dari situ pakai value itu sebagai ?key=... buat akses panel.
 * (Alternatif cepat: jalankan fungsi setSecretSekali_() sekali dari editor lalu HAPUS isinya lagi
 *  biar secret gak nempel permanen di source code yang mungkin lo commit/share.)
 *
 * MODEL DATA:
 * - 1 orang = 1 baris = 1 kode = 1 QR (order 3 tiket -> 3 baris)
 * - Baris-baris satu order diikat kolom OrderID (kode tiket pertama)
 * - Kode tiket ACAK (misal HT4-7F3A) biar susah ditebak orang iseng
 *************************************************/

var SHEET_NAME = 'ORDERS';
var HARGA = 50000;          // presale 50K

// Opsional: kalau script lo standalone / bukan dibuat dari Extensions > Apps Script di Sheet,
// isi ID spreadsheet di sini. Kalau dikosongkan, script akan pakai spreadsheet aktif,
// atau bikin spreadsheet baru sekali lalu menyimpan ID-nya di Script Properties.
var SPREADSHEET_ID = '';

// PERINGATAN: ini bikin spreadsheet bisa diedit oleh siapa pun yang punya link.
// Default DIMATIKAN — kalau panitia butuh edit dari HP tanpa login akun lo, kasih akses
// manual per-akun (Share > invite email panitia) daripada buka ke siapa pun yang punya link.
var SHARE_SPREADSHEET_ANYONE_WITH_LINK_EDIT = false;

/* ---------- secret (disimpan di Script Properties, bukan di kode) ---------- */

function getSecret_(){
  var s = PropertiesService.getScriptProperties().getProperty('SECRET');
  if(!s){
    throw new Error('SECRET belum di-set. Buka Project Settings > Script Properties, tambah key "SECRET" dengan value string acak panjang.');
  }
  return s;
}

// Jalankan SEKALI dari editor Apps Script (pilih fungsi ini di dropdown run > klik Run),
// abis itu ganti value di bawah jadi string kosong lagi / hapus baris isinya.
function setSecretSekali_(){
  PropertiesService.getScriptProperties().setProperty('SECRET', 'GANTI_DENGAN_STRING_ACAK_PANJANG_MIN_20_KARAKTER');
}

/* ---------- util ---------- */

function getSpreadsheet_(){
  var props = PropertiesService.getScriptProperties();
  var id = String(SPREADSHEET_ID || props.getProperty('HT4_SPREADSHEET_ID') || '').trim();
  if(id){
    return SpreadsheetApp.openById(id);
  }

  // Kalau script ini dibuat dari menu Extensions > Apps Script di Google Sheet, ini akan ada.
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if(active){
    props.setProperty('HT4_SPREADSHEET_ID', active.getId());
    return active;
  }

  // Fallback untuk kasus script standalone: bikin Sheet baru sekali saja.
  var created = SpreadsheetApp.create('HT4 Tiket');
  props.setProperty('HT4_SPREADSHEET_ID', created.getId());
  return created;
}

function spreadsheetUrl_(){
  return getSpreadsheet_().getUrl();
}

function ensureSpreadsheetPublicEdit_(){
  if(!SHARE_SPREADSHEET_ANYONE_WITH_LINK_EDIT) return {ok:true, skipped:true};
  try{
    var ss = getSpreadsheet_();
    DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    return {ok:true};
  }catch(err){
    return {ok:false, error:String(err)};
  }
}

// Jalankan manual sekali dari editor Apps Script kalau mau paksa buka akses spreadsheet.
function bukaAksesSpreadsheet(){
  var r = ensureSpreadsheetPublicEdit_();
  if(!r.ok) throw new Error(r.error);
  return spreadsheetUrl_();
}

function sheet(){
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Waktu','Kode','Nama','WA','Qty','Status','Check-in','OrderID']);
    sh.setFrozenRows(1);
  } else if(String(sh.getRange(1,8).getValue()) !== 'OrderID'){
    sh.getRange(1,8).setValue('OrderID'); // migrasi header lama
  }
  return sh;
}

function json(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(callback, o){
  var cb = String(callback || 'callback').replace(/[^a-zA-Z0-9_.$]/g, '');
  if(!cb) cb = 'callback';
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(o) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function findRow(sh, code){
  var last = sh.getLastRow();
  if(last < 2) return -1;
  var codes = sh.getRange(2, 2, last-1, 1).getValues();
  for(var i=0; i<codes.length; i++){
    if(String(codes[i][0]).toUpperCase() === String(code).toUpperCase()) return i+2;
  }
  return -1;
}

// cari SEMUA baris satu order (cocokkan OrderID atau kode tiketnya)
function findRowsOrder(sh, key){
  var last = sh.getLastRow();
  if(last < 2) return [];
  var vals = sh.getRange(2, 2, last-1, 7).getValues(); // kolom B..H
  var target = String(key).toUpperCase();
  var out = [];
  for(var i=0;i<vals.length;i++){
    var code = String(vals[i][0]);
    var oid  = String(vals[i][6]||'') || code;
    if(oid.toUpperCase() === target || code.toUpperCase() === target) out.push(i+2);
  }
  return out;
}

// kode acak 4 karakter — tanpa 0/O/1/I biar gak ketuker pas diketik manual
function bikinKode(sh){
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for(var tries=0; tries<50; tries++){
    var s = '';
    for(var i=0;i<4;i++) s += chars.charAt(Math.floor(Math.random()*chars.length));
    var code = 'HT4-' + s;
    if(findRow(sh, code) < 0) return code;
  }
  return 'HT4-' + String(Date.now());
}

function normWa(wa){
  var w = String(wa||'').replace(/[^0-9]/g,'');
  if(w.indexOf('62')===0) return w;
  if(w.indexOf('0')===0) return '62' + w.slice(1);
  if(w.indexOf('8')===0) return '62' + w; // Sheets kadang makan 0 di depan
  return w;
}

function baseUrl(){ return ScriptApp.getService().getUrl(); }

/* ---------- POST/GET: order baru & aksi admin ---------- */

function createOrder_(data){
  // order baru dari form website — 1 baris per tiket, tiap orang dapet QR sendiri
  var nama = String(data.nama||'').trim();
  var wa = normWa(data.wa);
  var qty = Math.max(1, Math.min(10, Number(data.qty||1)));

  // validasi dasar — tolak submit kosong/asal biar sheet gak penuh sampah
  if(!nama) return {ok:false, error:'nama_kosong'};
  if(wa.length < 10) return {ok:false, error:'wa_tidak_valid'};

  // rate limit sederhana per nomor WA — cegah spam klik/bot ngirim order berkali-kali dalam hitungan detik
  var cache = CacheService.getScriptCache();
  var cacheKey = 'order_' + wa;
  if(cache.get(cacheKey)) return {ok:false, error:'coba_lagi_sebentar'};
  cache.put(cacheKey, '1', 30); // 1 nomor WA maks 1 order tiap 30 detik

  var sh = sheet();

  var codes = [];
  for(var k=0;k<qty;k++){
    var code = bikinKode(sh);
    var orderId = codes.length ? codes[0] : code;
    sh.appendRow([new Date(), code, nama, '', 1, 'MENUNGGU', '', orderId]);
    var r = sh.getLastRow();
    sh.getRange(r, 4).setNumberFormat('@').setValue(wa);
    codes.push(code);
  }
  return {ok:true, codes:codes, code:codes[0]};
}

function doPost(e){
  var data = {};
  try{ data = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch(err){ return json({ok:false, error:'bad json'}); }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // antri maks 20 dtk — cegah 2 request tabrakan di detik yang sama
  try{
    if(data.action) return adminAction(data);

    return json(createOrder_(data));
  } finally {
    lock.releaseLock();
  }
}

function adminAction(d){
  if(d.key !== getSecret_()) return json({ok:false, error:'forbidden'});
  var sh = sheet();

  // confirm & hapus kerja per-ORDER (semua tiket dalam order itu sekalian)
  if(d.action === 'confirm' || d.action === 'hapus'){
    var rows = findRowsOrder(sh, d.code);
    if(!rows.length) return json({ok:false, error:'not found'});
    if(d.action === 'confirm'){
      for(var i=0;i<rows.length;i++) sh.getRange(rows[i], 6).setValue('LUNAS');
      return json({ok:true, code:d.code});
    }
    for(var j=rows.length-1;j>=0;j--) sh.deleteRow(rows[j]); // dari bawah biar nomor baris gak geser
    return json({ok:true, code:d.code});
  }

  // check-in tetap per-TIKET (1 QR = 1 orang masuk)
  if(d.action === 'checkin'){
    var row = findRow(sh, d.code);
    if(row < 0) return json({ok:false, error:'not found'});
    var rowData = sh.getRange(row, 1, 1, 7).getValues()[0];
    var nama   = rowData[2];
    var status = rowData[5];
    var cin    = rowData[6];
    if(status !== 'LUNAS') return json({ok:false, status:'BELUM_LUNAS', code:d.code, nama:nama});
    if(cin) return json({ok:false, status:'SUDAH_DIPAKE', code:d.code, nama:nama, qty:1, waktu:String(cin)});
    sh.getRange(row, 7).setValue(new Date());
    return json({ok:true, status:'VALID', code:d.code, nama:nama, qty:1});
  }

  return json({ok:false, error:'unknown action'});
}

/* ---------- GET: halaman tiket / panel ---------- */

function doGet(e){
  var p = (e && e.parameter) || {};

  // Endpoint JSONP untuk website statis. Ini sengaja pakai GET supaya aman dari masalah CORS
  // saat website dipasang di hosting/domain berbeda dari script.google.com.
  if(p.action === 'order'){
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try{
      var result = createOrder_({nama:p.nama, wa:p.wa, qty:p.qty});
      return p.callback ? jsonp_(p.callback, result) : json(result);
    } finally {
      lock.releaseLock();
    }
  }

  if(p.page === 'ticket') return ticketPage(String(p.code||''));
  if(p.page === 'admin' || p.page === 'scan' || p.page === 'panel')
    return panelPage(String(p.key||''), p.page === 'scan' ? 'scan' : 'orders');
  return ContentService.createTextOutput('HOLDTIGHT VOL.4 — sistem tiket aktif.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function pageShell(title, body){
  return HtmlService.createHtmlOutput(
'<!DOCTYPE html><html lang="id"><head><meta charset="UTF-8">' +
'<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
'<title>' + title + '</title>' +
'<link href="https://fonts.googleapis.com/css2?family=Anton&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">' +
'<style>' +
':root{--bg:#0d0d0d;--bone:#E8E4DC;--dim:rgba(232,228,220,.6);--faint:rgba(232,228,220,.35);--line:rgba(232,228,220,.14);--green:#2e9e6b;--red:#c1121f}' +
'*{margin:0;padding:0;box-sizing:border-box}' +
'body{background:var(--bg);color:var(--bone);font-family:"Space Mono",monospace;font-size:14px;line-height:1.7;padding:28px 20px}' +
'.wrap{max-width:460px;margin:0 auto}' +
'h1{font-family:"Anton",sans-serif;font-weight:400;text-transform:uppercase;letter-spacing:.02em;line-height:.95}' +
'.lbl{font-size:10px;letter-spacing:.32em;text-transform:uppercase;color:var(--faint)}' +
'.card{border:1px solid var(--line);padding:26px 22px;margin-top:18px}' +
'.btn{display:inline-block;font-family:"Space Mono",monospace;font-weight:700;font-size:12px;letter-spacing:.22em;text-transform:uppercase;padding:14px 26px;text-decoration:none;border:1px solid var(--bone);background:var(--bone);color:#0d0d0d;cursor:pointer}' +
'.btn.green{background:var(--green);border-color:var(--green);color:#fff}' +
'.btn.ghost{background:transparent;color:var(--bone)}' +
'</style></head><body><div class="wrap">' + body + '</div></body></html>'
  ).setTitle(title);
}

/* --- halaman e-ticket (dibuka pembeli) --- */

function ticketPage(code){
  var sh = sheet();
  var last = sh.getLastRow();
  var row = -1, canon = String(code||''), pos = 1, total = 1;

  if(canon && last >= 2){
    var vals = sh.getRange(2, 2, last-1, 7).getValues(); // kolom B..H
    var target = canon.toUpperCase();
    var hit = -1;
    for(var i=0;i<vals.length;i++){
      if(String(vals[i][0]).toUpperCase() === target){ hit = i; break; }
    }
    if(hit >= 0){
      row = hit + 2;
      canon = String(vals[hit][0]);
      var oid = String(vals[hit][6]||'') || canon;
      total = 0; pos = 0;
      for(var j=0;j<vals.length;j++){
        var o2 = String(vals[j][6]||'') || String(vals[j][0]);
        if(o2 === oid){ total++; if(j <= hit) pos = total; }
      }
    }
  }

  var body;
  if(row < 0){
    body = '<div class="lbl">Holdtight Vol.4 — 26.09.2026</div>' +
      '<h1 style="font-size:42px;margin-top:12px">Tiket tidak<br>ditemukan.</h1>' +
      '<div class="card"><span class="lbl">kode</span><br>' + esc(canon) + '</div>';
    return pageShell('Tiket — HOLDTIGHT VOL.4', body);
  }

  var nama   = sh.getRange(row, 3).getValue();
  var status = sh.getRange(row, 6).getValue();
  var cin    = sh.getRange(row, 7).getValue();

  var statusHtml;
  if(cin){
    statusHtml = '<div class="card" style="border-color:var(--faint)"><span class="lbl">status</span>' +
      '<h1 style="font-size:34px;margin-top:8px;color:var(--faint)">Sudah<br>check-in.</h1></div>';
  } else if(status === 'LUNAS'){
    statusHtml = '<div class="card" style="text-align:center;border-color:var(--green)">' +
      '<div id="qr" style="display:inline-block;background:#fff;padding:14px;margin:6px auto 4px"></div>' +
      '<div class="lbl" style="margin-top:10px">tunjukkan qr ini di pintu — 1 qr untuk 1 orang</div></div>';
  } else {
    statusHtml = '<div class="card" style="border-color:var(--red)"><span class="lbl">status</span>' +
      '<h1 style="font-size:34px;margin-top:8px;color:var(--red)">Menunggu<br>konfirmasi.</h1>' +
      '<p style="color:var(--dim);margin-top:10px;font-size:12px">Pembayaran lo belum terkonfirmasi panitia. QR bakal muncul di halaman ini setelah LUNAS.</p></div>';
  }

  body = '<div class="lbl">Holdtight Vol.4 — Sabtu 26.09.2026 — Kaliwungu</div>' +
    '<h1 style="font-size:46px;margin:12px 0 4px">' + esc(canon) + '</h1>' +
    '<div class="lbl">' + esc(nama) + (total > 1 ? ' — tiket ' + pos + ' dari ' + total : '') + '</div>' +
    statusHtml +
    '<p class="lbl" style="margin-top:20px;text-align:center">tanpa korporasi — dari kita, untuk kita</p>' +
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>' +
    '<script>if(document.getElementById("qr") && typeof QRCode!=="undefined"){new QRCode(document.getElementById("qr"),{text:' + JSON.stringify(canon) + ',width:220,height:220,correctLevel:QRCode.CorrectLevel.M});}</script>';

  return pageShell('E-Tiket ' + canon + ' — HOLDTIGHT VOL.4', body);
}

/* --- panel panitia: pesanan + scan gate + spreadsheet (1 halaman) --- */

function panelPage(key, tab){
  if(key !== getSecret_()){
    return ContentService.createTextOutput('403 — butuh kunci.')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  var share = ensureSpreadsheetPublicEdit_();
  var sh = sheet();
  var last = sh.getLastRow();
  var rows = '';
  var msgs = {};
  var orderCount = 0, totalQty = 0, lunasQty = 0, cinQty = 0;

  if(last >= 2){
    var vals = sh.getRange(2, 1, last-1, 8).getValues();
    var orderMap = {}, orderKeys = [];
    for(var i=0;i<vals.length;i++){
      var v = vals[i];
      var code=String(v[1]), nama=String(v[2]), wa=String(v[3]), status=String(v[5]), cin=String(v[6]||''), oid=String(v[7]||'')||code;
      totalQty++;
      if(status==='LUNAS') lunasQty++;
      if(cin) cinQty++;
      if(!orderMap[oid]){ orderMap[oid]={nama:nama,wa:wa,tickets:[]}; orderKeys.push(oid); }
      var o=orderMap[oid];
      if(!o.nama&&nama) o.nama=nama;
      if(!o.wa&&wa) o.wa=wa;
      o.tickets.push({code:code,status:status,cin:cin});
    }
    orderCount = orderKeys.length;

    for(var k=orderKeys.length-1;k>=0;k--){
      var oid2 = orderKeys[k];
      var g = orderMap[oid2];
      var n = g.tickets.length;
      var allCin = true, allLunas = true;
      var list = '';
      for(var ti=0;ti<n;ti++){
        var t = g.tickets[ti];
        if(!t.cin) allCin=false;
        if(t.status!=='LUNAS') allLunas=false;
        var st = t.cin ? ('masuk ' + t.cin) : t.status;
        var wr = t.cin ? 'var(--faint)' : (t.status==='LUNAS' ? 'var(--green)' : 'var(--red)');
        list += '<div style="display:flex;justify-content:space-between;gap:10px;border-top:1px solid var(--line);padding:8px 0">' +
          '<b style="font-family:Anton;font-size:16px;letter-spacing:.04em">' + esc(t.code) + '</b>' +
          '<span class="lbl" style="color:' + wr + '">' + esc(st) + '</span></div>';
      }

      var aksi;
      if(allCin){
        aksi = '<span class="lbl">semua masuk ✓</span>';
      } else if(allLunas){
        aksi = '<button class="btn green" style="padding:10px 16px;font-size:10px" onclick="kirim(\'' + oid2 + '\')">Kirim Tiket →</button>';
      } else {
        aksi = '<button class="btn" style="padding:10px 16px;font-size:10px" onclick="lunas(\'' + oid2 + '\',this)">LUNAS ✓</button>';
      }

      var lines = '';
      for(var mi=0;mi<n;mi++){
        lines += '\nTiket ' + (mi+1) + ': ' + baseUrl() + '?page=ticket&code=' + encodeURIComponent(g.tickets[mi].code);
      }
      var waMsg = 'Halo ' + g.nama + '! ' + n + ' tiket HOLDTIGHT VOL.4 lo udah LUNAS ✓\n' + lines +
        '\n\nTiap orang pegang tiketnya masing-masing ya (QR-nya beda-beda). PENTING: buka link-nya pas ada sinyal, terus SCREENSHOT QR-nya — di venue tinggal tunjukkin screenshot-nya, gak butuh sinyal. Sabtu 26 Sept 2026, gate open 15.00. Sampai ketemu di pit!';
      msgs[oid2] = {wa: normWa(g.wa), msg: waMsg};

      rows += '<div class="card" style="margin-top:12px">' +
        '<div style="display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;align-items:flex-start">' +
        '<div><b style="font-family:Anton;font-size:20px;letter-spacing:.04em">' + esc(oid2) + '</b>' +
        '<div style="color:var(--dim);font-size:12px">' + esc(g.nama) + ' — ' + esc(g.wa) + ' — ' + n + ' tiket</div></div>' +
        '<div>' + aksi + '<div style="margin-top:8px;text-align:right"><a href="#" style="color:var(--faint);font-size:10px;letter-spacing:.2em" onclick="hapus(\'' + oid2 + '\',this);return false">HAPUS X</a></div></div></div>' +
        '<div style="margin-top:10px">' + list + '</div></div>';
    }
  } else {
    rows = '<div class="card">Belum ada order masuk.</div>';
  }

  var body = '<div class="lbl" style="margin-top:18px">Daftar pesanan</div>' +
    '<h1 style="font-size:38px;margin-top:10px">Orders.</h1>' +
    '<div class="card" style="display:flex;gap:20px;flex-wrap:wrap">' +
    '<div><div class="lbl">order</div><b style="font-family:Anton;font-size:28px">' + orderCount + '</b></div>' +
    '<div><div class="lbl">tiket</div><b style="font-family:Anton;font-size:28px">' + totalQty + '</b></div>' +
    '<div><div class="lbl">lunas</div><b style="font-family:Anton;font-size:28px;color:var(--green)">' + lunasQty + '</b></div>' +
    '<div><div class="lbl">masuk</div><b style="font-family:Anton;font-size:28px">' + cinQty + '</b></div></div>' +
    '<p class="lbl" style="margin-top:10px">1 orang = 1 qr. kode acak, susah ditebak orang iseng.</p>' +
    (share.ok ? '' : '<div class="card" style="border-color:var(--red);color:var(--red)">Spreadsheet belum bisa dibuka otomatis: ' + esc(share.error) + '</div>') +
    '<p style="margin-top:14px"><a class="btn ghost" style="padding:10px 16px;font-size:10px" href="" onclick="location.reload();return false">Refresh</a></p>' +
    rows +
    '<script>var MSGS=' + JSON.stringify(msgs) + ';' +
'function kirim(code){var d=MSGS[code];if(!d||!d.wa){alert("Nomor WA pembeli kosong — isi dulu di sheet");return}var t=encodeURIComponent(d.msg);var start=Date.now();location.href="whatsapp://send?phone="+d.wa+"&text="+t;setTimeout(function(){if(!document.hidden&&Date.now()-start<2500){location.href="https://wa.me/"+d.wa+"?text="+t}},1800)}' +
'function lunas(code,btn){btn.disabled=true;btn.textContent="...";fetch("' + baseUrl() + '",{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({action:"confirm",key:' + JSON.stringify(key) + ',code:code})}).then(function(){location.reload()}).catch(function(){btn.textContent="GAGAL — coba lagi";btn.disabled=false})}' +
'function hapus(code,el){if(!confirm("Hapus order "+code+"? semua tiket di order ini ilang dari spreadsheet."))return;el.textContent="MENGHAPUS...";fetch("' + baseUrl() + '",{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({action:"hapus",key:' + JSON.stringify(key) + ',code:code})}).then(function(){location.reload()}).catch(function(){el.textContent="GAGAL"})}</script>';

  var tabOrdersStyle = tab === 'scan' ? 'display:none' : '';
  var tabScanStyle = tab === 'scan' ? '' : 'display:none';
  var btnOrdersCls = 'btn' + (tab === 'scan' ? ' ghost' : ' green');
  var btnScanCls = 'btn' + (tab === 'scan' ? ' green' : ' ghost');

  var panel = '<div style="display:flex;justify-content:space-between;align-items:center;gap:10px">' +
    '<div class="lbl">Holdtight Vol.4 — Panel Panitia</div>' +
    '<a class="btn ghost" style="padding:8px 14px;font-size:9px" href="' + spreadsheetUrl_() + '" target="_blank">Spreadsheet ↗</a></div>' +
    '<div style="display:flex;gap:8px;margin-top:14px">' +
    '<button id="btnOrders" class="' + btnOrdersCls + '" style="flex:1;padding:12px 8px" onclick="pindah(\'orders\')">PESANAN</button>' +
    '<button id="btnScan" class="' + btnScanCls + '" style="flex:1;padding:12px 8px" onclick="pindah(\'scan\')">SCAN GATE</button></div>' +
    '<div id="tabOrders" style="' + tabOrdersStyle + '">' + body + '</div>' +
    '<div id="tabScan" style="' + tabScanStyle + '">' + scanBody(key) + '</div>' +
    '<script>function pindah(t){' +
    'document.getElementById("tabOrders").style.display=t==="orders"?"":"none";' +
    'document.getElementById("tabScan").style.display=t==="scan"?"":"none";' +
    'document.getElementById("btnOrders").className="btn"+(t==="orders"?" green":" ghost");' +
    'document.getElementById("btnScan").className="btn"+(t==="scan"?" green":" ghost");' +
    'if(t!=="scan" && window.__ht4StopScanCamera) window.__ht4StopScanCamera();' +
    '}</script>';

  return pageShell('Panel — HOLDTIGHT VOL.4', panel);
}

/* --- isi tab scan check-in (pintu hari-H) --- */

function scanBody(key){
  var body = '<div class="lbl">Holdtight Vol.4 — Check-in Gate</div>' +
    '<h1 style="font-size:38px;margin-top:10px">Scan.</h1>' +
    '<div class="card" style="padding:0;overflow:hidden;position:relative">' +
    '<video id="video" playsinline muted style="width:100%;display:block;background:#000"></video>' +
    '<div id="frame" style="position:absolute;inset:0;display:none;align-items:center;justify-content:center;pointer-events:none">' +
    '<div style="width:58%;aspect-ratio:1/1;border:3px solid var(--green);border-radius:14px;box-shadow:0 0 0 999px rgba(0,0,0,.4)"></div></div>' +
    '<canvas id="canvas" style="display:none"></canvas>' +
    '</div>' +
    '<div class="card" style="text-align:center">' +
    '<button id="btnCam" class="btn green" style="width:100%" onclick="mulaiKamera()">Aktifkan Kamera</button>' +
    '<p class="lbl" style="margin-top:12px" id="camStatus">arahkan kamera ke qr — otomatis kedetek, gak perlu jepret</p></div>' +
    '<div class="card"><label class="lbl">atau ketik kode manual</label>' +
    '<div style="display:flex;gap:8px;margin-top:10px">' +
    '<input id="manual" placeholder="HT4-7F3A" style="flex:1;background:transparent;border:1px solid var(--line);color:var(--bone);font-family:inherit;font-size:16px;padding:12px">' +
    '<button class="btn" onclick="cek(document.getElementById(\'manual\').value)">CEK</button></div></div>' +
    '<div id="hasil" style="margin-top:18px"></div>' +
    '<script src="https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js"></script>' +
    '<script>' +
    'var KEY=' + JSON.stringify(key) + ';var BASE=' + JSON.stringify(baseUrl()) + ';var sibuk=false;var camStream=null;var camRaf=null;' +
    'function e2(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}' +
    'function tampil(warna,pesan){document.getElementById("hasil").innerHTML="<div class=card style=\'border-color:"+warna+"\'><h1 style=\'font-size:28px;color:"+warna+";line-height:1.4\'>"+pesan+"</h1></div>"}' +
    'function cek(code){' +
    '  code=String(code||"").trim().toUpperCase(); if(!code||sibuk) return; sibuk=true;' +
    '  tampil("var(--faint)","MEMERIKSA…");' +
    '  fetch(BASE,{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({action:"checkin",key:KEY,code:code})})' +
    '  .then(function(r){return r.json()}).then(function(d){' +
    '    if(d.ok){tampil("var(--green)","VALID ✓ — SILAKAN MASUK<br>"+e2(d.nama)+"<br>"+e2(d.code))}' +
    '    else if(d.status==="SUDAH_DIPAKE"){tampil("var(--red)","SUDAH DIPAKE ✕<br>"+e2(d.nama)+"<br>"+e2(d.code)+"<br><small>check-in: "+e2(d.waktu)+"</small>")}' +
    '    else if(d.status==="BELUM_LUNAS"){tampil("var(--red)","BELUM LUNAS ✕<br>"+e2(d.nama)+"<br>"+e2(d.code))}' +
    '    else{tampil("var(--red)","TIDAK DITEMUKAN ✕<br>"+e2(code))}' +
    '    setTimeout(function(){sibuk=false},1800);' + // jeda dikit abis hasil muncul, terus scan lanjut otomatis buat orang berikutnya
    '  }).catch(function(){tampil("var(--red)","KONEKSI GAGAL — coba lagi");sibuk=false});' +
    '}' +
    'function mulaiKamera(){' +
    '  if(!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia){' +
    '    var st0=document.getElementById("camStatus"); st0.textContent="kamera live gak didukung browser ini — pakai input manual di bawah"; st0.style.color="var(--red)";' +
    '    return;' +
    '  }' +
    '  var btn=document.getElementById("btnCam"); btn.disabled=true; btn.textContent="MEMINTA IZIN KAMERA…";' +
    '  navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:"environment"}}, audio:false}).then(function(s){' +
    '    camStream=s; var v=document.getElementById("video"); v.srcObject=s; v.play();' +
    '    btn.style.display="none";' +
    '    document.getElementById("frame").style.display="flex";' +
    '    var st=document.getElementById("camStatus"); st.textContent="live — arahkan ke qr, otomatis kedetek"; st.style.color="";' +
    '    v.addEventListener("loadedmetadata",function(){ scanLoop(); }, {once:true});' +
    '  }).catch(function(){' +
    '    btn.disabled=false; btn.textContent="Aktifkan Kamera";' +
    '    var st=document.getElementById("camStatus"); st.textContent="izin kamera ditolak/gagal — pakai input manual di bawah"; st.style.color="var(--red)";' +
    '  });' +
    '}' +
    'function scanLoop(){' +
    '  var v=document.getElementById("video"); var c=document.getElementById("canvas"); var x=c.getContext("2d",{willReadFrequently:true});' +
    '  function tick(){' +
    '    camRaf=requestAnimationFrame(tick);' +
    '    if(sibuk || !v.videoWidth) return;' +
    '    var w=Math.min(640, v.videoWidth); var h=Math.round(w * v.videoHeight / v.videoWidth);' +
    '    c.width=w; c.height=h; x.drawImage(v,0,0,w,h);' +
    '    var d; try{ d=x.getImageData(0,0,w,h); }catch(e){ return; }' +
    '    var q=jsQR(d.data,w,h,{inversionAttempts:"attemptBoth"});' +
    '    if(q && q.data){ cek(q.data); }' +
    '  }' +
    '  tick();' +
    '}' +
    'function hentikanKamera(){' +
    '  if(camRaf) cancelAnimationFrame(camRaf); camRaf=null;' +
    '  if(camStream){ camStream.getTracks().forEach(function(t){t.stop()}); camStream=null; }' +
    '  var v=document.getElementById("video"); if(v) v.srcObject=null;' +
    '  var btn=document.getElementById("btnCam"); if(btn){ btn.style.display=""; btn.disabled=false; btn.textContent="Aktifkan Kamera"; }' +
    '  var fr=document.getElementById("frame"); if(fr) fr.style.display="none";' +
    '  var st=document.getElementById("camStatus"); if(st){ st.textContent="arahkan kamera ke qr — otomatis kedetek, gak perlu jepret"; st.style.color=""; }' +
    '}' +
    'window.__ht4StopScanCamera = hentikanKamera;' +
    '</script>';

  return body;
}

/* ---------- escape ---------- */
function esc(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
