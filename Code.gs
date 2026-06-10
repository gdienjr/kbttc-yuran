// =============================================
// KBTTC YURAN COLLECTION APP
// Google Apps Script - API Backend
// =============================================

var YURAN_BULANAN = 5;
var YURAN_DAFTAR = 5;
var YURAN_TAHUNAN = 50;
var BULAN = ['Jan','Feb','Mac','Apr','Mei','Jun','Jul','Ogos','Sep','Okt','Nov','Dis'];
var VIP_COLOR = '#b7dfbf';

// =============================================
// WEB APP ENTRY POINT - Handle CORS
// =============================================

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var callback = e && e.parameter && e.parameter.callback;
  
  var result = handleAction(action, e.parameter);
  var json = JSON.stringify(result);
  
  // JSONP support for CORS
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action;
  var result = handleAction(action, data);
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleAction(action, params) {
  try {
    if (action === 'getSenariAhli') return getSenariAhli();
    if (action === 'getStatusBayaran') return getStatusBayaran(params.noAhli);
    if (action === 'rekodBayaran') return rekodBayaran(params.noAhli, params.amaun, params.bulanSemasa);
    if (action === 'rekodYuranTahunan') return rekodYuranTahunan(params.noAhli);
    if (action === 'daftarAhliBaru') return daftarAhliBaru(params.nama, params.noPhone);
    if (action === 'semakStatusAhli') return semakStatusAhli(params.last4);
    return { success: false, msg: 'Action tidak dikenali.' };
  } catch(e) {
    return { success: false, msg: e.message };
  }
}

// =============================================
// SETUP SHEETS (run once)
// =============================================

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetAhli = ss.getSheetByName('Senarai Ahli');
  if (!sheetAhli) sheetAhli = ss.insertSheet('Senarai Ahli');
  sheetAhli.clearContents();
  sheetAhli.getRange(1,1,1,5).setValues([['No. Ahli','Nama','No. Phone','Tarikh Daftar','Status']]);
  sheetAhli.getRange(1,1,1,5).setFontWeight('bold');

  var ahliSedia = [
    ['0001','DINRC','','','Aktif'],['0002','PAWANG','','','Aktif'],['0003','CGHALIM','','','Aktif'],
    ['0004','DIN SUNGAI','','','Aktif'],['0005','MAT G','','','Aktif'],['0006','MATKONG','','','Aktif'],
    ['0007','AMIR KB','','','Aktif'],['0008','AYOHWAN','','','Aktif'],['0009','BAHRIN','','','Aktif'],
    ['0010','LAMIN','','','Aktif'],['0011','MADU','','','Aktif'],['0012','NUAR TM','','','Aktif'],
    ['0013','PARI','','','Aktif'],['0014','ACHIK CH','','','Aktif'],['0015','ANGOH CH','','','Aktif'],
    ['0016','APIS S','','','Aktif'],['0017','AZREN C','','','Aktif'],['0018','FARIS KB','','','Aktif'],
    ['0019','HADIF','','','Aktif'],['0020','MRILEK','','','Aktif'],['0021','MUSLIM','','','Aktif'],
    ['0022','AMAN','','','Aktif'],['0023','AMIR PAI','','','Aktif'],['0024','AYUB KEPONG','','','Aktif'],
    ['0025','CG ADDA','','','Aktif'],['0026','AYUB NB','','','Aktif'],['0027','CGMAD','','','Aktif'],
    ['0028','CHAD C','','','Aktif'],['0029','DIN HSNZ','','','Aktif'],['0030','MANAN','','','Aktif'],
    ['0031','MAT MRG','','','Aktif'],['0032','NUAR 1','','','Aktif'],['0033','PIAN','','','Aktif'],
    ['0034','PKAMIL','','','Aktif'],['0035','PMERAN','','','Aktif'],['0036','PZAM','','','Aktif'],
    ['0037','PSYED','','','Aktif'],['0038','RIZAL','','','Aktif'],['0039','ROZLIM','','','Aktif'],
    ['0040','SARIMAN','','','Aktif'],['0041','SEPU KS','','','Aktif'],['0042','YAT','','','Aktif'],
    ['0043','ALLY','','','Aktif'],['0044','AMIRUL PKD','','','Aktif'],['0045','HAFIZ C','','','Aktif'],
    ['0046','HAFIZ HSNZ MA','','','Aktif'],['0047','HAIRY JKR','','','Aktif'],['0048','IMAN PM','','','Aktif'],
    ['0049','IQBAL 1','','','Aktif'],['0050','KOLEY','','','Aktif'],['0051','LEMAN','','','Aktif'],
    ['0052','MATYI AT','','','Aktif'],['0053','MEZOH PB','','','Aktif'],['0054','PAJEY','','','Aktif'],
    ['0055','PEJO','','','Aktif'],['0056','PGHANI','','','Aktif'],['0057','TINI (SHAM)','','','Aktif'],
    ['0058','DIN CONWAY','','','Aktif'],['0059','EZRI','','','Aktif'],['0060','KAK_ANI','','','Aktif'],
    ['0061','KAK_SAL','','','Aktif'],['0062','MI LAUT','','','Aktif'],['0063','MIZIE','','','Aktif'],
    ['0064','ADAM','','','Aktif'],['0065','AIMAN KMMN','','','Aktif'],['0066','ALIF PB','','','Aktif'],
    ['0067','AYIE','','','Aktif'],['0068','DANIEL','','','Aktif'],['0069','DRLAN','','','Aktif'],
    ['0070','DUAN PB','','','Aktif'],['0071','ELEAS','','','Aktif'],['0072','FAKHRUL ST','','','Aktif'],
    ['0073','FARIS','','','Aktif'],['0074','FAZRUL','','','Aktif'],['0075','HAFIZ HSNZ','','','Aktif'],
    ['0076','HANAN','','','Aktif'],['0077','HAZIQ AP','','','Aktif'],['0078','IMAN BR','','','Aktif'],
    ['0079','LIE MEK','','','Aktif'],['0080','ROKIN','','','Aktif'],['0081','TAJUDIN DGN','','','Aktif']
  ];
  sheetAhli.getRange(2,1,ahliSedia.length,5).setValues(ahliSedia);

  var sheetLog = ss.getSheetByName('Log Transaksi');
  if (!sheetLog) sheetLog = ss.insertSheet('Log Transaksi');
  sheetLog.clearContents();
  sheetLog.getRange(1,1,1,6).setValues([['Tarikh','No. Ahli','Nama','Jenis','Bulan Dibayar','Amaun (RM)']]);
  sheetLog.getRange(1,1,1,6).setFontWeight('bold');

  var sheetMatriks = ss.getSheetByName('Matriks Yuran');
  if (!sheetMatriks) sheetMatriks = ss.insertSheet('Matriks Yuran');
  rebuildMatriks();

  SpreadsheetApp.getUi().alert('Setup selesai!');
}

function rebuildMatriks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetMatriks = ss.getSheetByName('Matriks Yuran');
  var sheetAhli = ss.getSheetByName('Senarai Ahli');

  sheetMatriks.clearContents();
  sheetMatriks.clearFormats();

  var header = ['No.','No. Ahli','Nama'].concat(BULAN).concat(['Jumlah (RM)']);
  sheetMatriks.getRange(1,1,1,header.length).setValues([header]);
  sheetMatriks.getRange(1,1,1,header.length).setFontWeight('bold');

  var ahliData = sheetAhli.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < ahliData.length; i++) {
    if (ahliData[i][0] && ahliData[i][1])
      rows.push([i, ahliData[i][0], ahliData[i][1]].concat(new Array(12).fill('')).concat([0]));
  }
  if (rows.length > 0) sheetMatriks.getRange(2,1,rows.length,rows[0].length).setValues(rows);

  for (var r = 2; r <= rows.length+1; r++)
    sheetMatriks.getRange(r,16).setFormula('=SUM(D'+r+':O'+r+')');

  var kutipanRow = rows.length+2;
  sheetMatriks.getRange(kutipanRow,3).setValue('Kutipan Bulanan').setFontWeight('bold');
  for (var c = 4; c <= 15; c++) {
    var col = columnToLetter(c);
    sheetMatriks.getRange(kutipanRow,c).setFormula('=SUM('+col+'2:'+col+(rows.length+1)+')');
  }
  sheetMatriks.getRange(kutipanRow,16).setFormula('=SUM(P2:P'+(rows.length+1)+')').setFontWeight('bold');
}

function columnToLetter(col) {
  var letter = '';
  while (col > 0) {
    var mod = (col-1)%26;
    letter = String.fromCharCode(65+mod)+letter;
    col = Math.floor((col-1)/26);
  }
  return letter;
}

function formatNoAhli(no) {
  return String(no).padStart(4,'0');
}

// =============================================
// API FUNCTIONS
// =============================================

function getSenariAhli() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Senarai Ahli');
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1])
      result.push({ noAhli: formatNoAhli(data[i][0]), nama: data[i][1], status: data[i][4] });
  }
  return result;
}

function getStatusBayaran(noAhli) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetMatriks = ss.getSheetByName('Matriks Yuran');
  var sheetAhli = ss.getSheetByName('Senarai Ahli');
  var matriksData = sheetMatriks.getDataRange().getValues();
  var ahliData = sheetAhli.getDataRange().getValues();

  var isVip = false;
  for (var i = 1; i < ahliData.length; i++) {
    if (formatNoAhli(ahliData[i][0]) == noAhli) { isVip = ahliData[i][4]==='VIP'; break; }
  }

  for (var i = 1; i < matriksData.length; i++) {
    if (formatNoAhli(matriksData[i][1]) == noAhli) {
      var status = {};
      for (var b = 0; b < 12; b++) status[BULAN[b]] = matriksData[i][b+3];
      return { nama: matriksData[i][2], status: status, isVip: isVip };
    }
  }
  return null;
}

function rekodBayaran(noAhli, amaun, bulanSemasa) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName('Log Transaksi');
  var sheetMatriks = ss.getSheetByName('Matriks Yuran');
  var sheetAhli = ss.getSheetByName('Senarai Ahli');

  var ahliData = sheetAhli.getDataRange().getValues();
  var nama = '';
  for (var i = 1; i < ahliData.length; i++) {
    if (formatNoAhli(ahliData[i][0]) == noAhli) { nama = ahliData[i][1]; break; }
  }
  if (!nama) return { success: false, msg: 'Ahli tidak dijumpai.' };

  amaun = parseFloat(amaun);
  if (isNaN(amaun) || amaun <= 0) return { success: false, msg: 'Amaun tidak sah.' };

  var matriksData = sheetMatriks.getDataRange().getValues();
  var matriksRow = -1;
  for (var i = 1; i < matriksData.length; i++) {
    if (formatNoAhli(matriksData[i][1]) == noAhli) { matriksRow = i+1; break; }
  }
  if (matriksRow === -1) return { success: false, msg: 'Ahli tidak ada dalam matriks.' };

  var bulanIdx = BULAN.indexOf(bulanSemasa);
  if (bulanIdx === -1) return { success: false, msg: 'Bulan tidak sah.' };

  var baki = amaun;
  var logEntries = [];
  var tarikh = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy');

  var colSemasa = bulanIdx+4;
  var currentVal = sheetMatriks.getRange(matriksRow, colSemasa).getValue();
  if (!currentVal || currentVal === '') {
    sheetMatriks.getRange(matriksRow, colSemasa).setValue(YURAN_BULANAN);
    logEntries.push([tarikh, noAhli, nama, 'Yuran Bulanan', bulanSemasa, YURAN_BULANAN]);
    baki -= YURAN_BULANAN;
  }

  for (var b = bulanIdx-1; b >= 0; b--) {
    if (baki < YURAN_BULANAN) break;
    var col = b+4;
    var val = sheetMatriks.getRange(matriksRow, col).getValue();
    if (!val || val === '') {
      sheetMatriks.getRange(matriksRow, col).setValue(YURAN_BULANAN);
      logEntries.push([tarikh, noAhli, nama, 'Yuran Bulanan', BULAN[b], YURAN_BULANAN]);
      baki -= YURAN_BULANAN;
    }
  }

  if (logEntries.length > 0)
    sheetLog.getRange(sheetLog.getLastRow()+1, 1, logEntries.length, 6).setValues(logEntries);

  return {
    success: true, nama: nama,
    jumlahDirekod: logEntries.length * YURAN_BULANAN,
    baki: Math.round(baki*100)/100,
    bulanDibayar: logEntries.map(function(e){ return e[4]; }),
    isVip: false
  };
}

function rekodYuranTahunan(noAhli) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName('Log Transaksi');
  var sheetMatriks = ss.getSheetByName('Matriks Yuran');
  var sheetAhli = ss.getSheetByName('Senarai Ahli');

  var ahliData = sheetAhli.getDataRange().getValues();
  var nama = '', ahliRow = -1;
  for (var i = 1; i < ahliData.length; i++) {
    if (formatNoAhli(ahliData[i][0]) == noAhli) { nama = ahliData[i][1]; ahliRow = i+1; break; }
  }
  if (!nama) return { success: false, msg: 'Ahli tidak dijumpai.' };

  var matriksData = sheetMatriks.getDataRange().getValues();
  var matriksRow = -1;
  for (var i = 1; i < matriksData.length; i++) {
    if (formatNoAhli(matriksData[i][1]) == noAhli) { matriksRow = i+1; break; }
  }
  if (matriksRow === -1) return { success: false, msg: 'Ahli tidak ada dalam matriks.' };

  var tarikh = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy');

  sheetMatriks.getRange(matriksRow, 15).setValue(YURAN_TAHUNAN).setBackground(VIP_COLOR);
  for (var c = 4; c <= 14; c++)
    sheetMatriks.getRange(matriksRow, c).setValue('').setBackground(VIP_COLOR);

  sheetAhli.getRange(ahliRow, 5).setValue('VIP');
  sheetLog.appendRow([tarikh, noAhli, nama, 'Yuran Tahunan (VIP)', 'Jan-Dis', YURAN_TAHUNAN]);

  return { success: true, nama: nama, noAhli: noAhli, isVip: true };
}

function daftarAhliBaru(nama, noPhone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAhli = ss.getSheetByName('Senarai Ahli');
  var sheetLog = ss.getSheetByName('Log Transaksi');
  var sheetMatriks = ss.getSheetByName('Matriks Yuran');

  if (!nama || nama.trim() === '') return { success: false, msg: 'Nama tidak boleh kosong.' };

  var ahliData = sheetAhli.getDataRange().getValues();
  var lastNo = 81;
  for (var i = 1; i < ahliData.length; i++) {
    var no = parseInt(ahliData[i][0]);
    if (!isNaN(no) && no > lastNo) lastNo = no;
  }
  var noAhliBaru = String(lastNo+1).padStart(4,'0');
  var tarikh = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy');

  sheetAhli.appendRow([noAhliBaru, nama.trim().toUpperCase(), noPhone||'', tarikh, 'Aktif']);

  var matriksData = sheetMatriks.getDataRange().getValues();
  var kutipanRow = -1;
  for (var i = 1; i < matriksData.length; i++) {
    if (matriksData[i][2] === 'Kutipan Bulanan') { kutipanRow = i+1; break; }
  }
  var insertAt = kutipanRow > 0 ? kutipanRow : sheetMatriks.getLastRow()+1;
  sheetMatriks.insertRowBefore(insertAt);
  var noUrut = insertAt-1;
  sheetMatriks.getRange(insertAt, 1, 1, 16).setValues([[noUrut, noAhliBaru, nama.trim().toUpperCase()].concat(new Array(12).fill('')).concat(['=SUM(D'+insertAt+':O'+insertAt+')'])]);

  sheetLog.appendRow([tarikh, noAhliBaru, nama.trim().toUpperCase(), 'Daftar Ahli', '-', YURAN_DAFTAR]);

  return { success: true, noAhli: noAhliBaru, nama: nama.trim().toUpperCase() };
}

function semakStatusAhli(last4) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAhli = ss.getSheetByName('Senarai Ahli');
  var ahliData = sheetAhli.getDataRange().getValues();

  var found = null;
  for (var i = 1; i < ahliData.length; i++) {
    var phone = String(ahliData[i][2]).replace(/\D/g,'');
    if (phone.length >= 4 && phone.slice(-4) === String(last4)) {
      found = { noAhli: formatNoAhli(ahliData[i][0]), nama: ahliData[i][1], status: ahliData[i][4] };
      break;
    }
  }
  if (!found) return { success: false, msg: 'Tiada ahli dijumpai dengan nombor tersebut.' };

  var statusData = getStatusBayaran(found.noAhli);
  return {
    success: true, noAhli: found.noAhli, nama: found.nama,
    status: found.status, bayaran: statusData ? statusData.status : {},
    isVip: statusData ? statusData.isVip : false
  };
}

function buatRingkasanKutipan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Ringkasan Kutipan');
  if (!sheet) sheet = ss.insertSheet('Ringkasan Kutipan');
  sheet.clearContents(); sheet.clearFormats();

  sheet.getRange(1,1).setValue('RINGKASAN KUTIPAN YURAN KBTTC 2026').setFontWeight('bold').setFontSize(13);
  sheet.getRange(2,1).setValue('Dikemaskini: '+Utilities.formatDate(new Date(),'Asia/Kuala_Lumpur','dd/MM/yyyy HH:mm')).setFontColor('#888888');
  sheet.getRange(4,1,1,3).setValues([['Bulan','Kutipan Bulanan (RM)','Kutipan VIP (RM)']]).setFontWeight('bold').setBackground('#1a73e8');
  sheet.getRange(4,1,1,3).setFontColor('#ffffff');

  var sheetLog = ss.getSheetByName('Log Transaksi');
  var logData = sheetLog.getDataRange().getValues();
  var kb = {}, kv = {};
  for (var b = 0; b < BULAN.length; b++) { kb[BULAN[b]] = 0; kv[BULAN[b]] = 0; }

  for (var i = 1; i < logData.length; i++) {
    var jenis = logData[i][3], bln = logData[i][4], amaun = parseFloat(logData[i][5])||0;
    if (jenis === 'Yuran Tahunan (VIP)') kv['Dis'] += amaun;
    else if (jenis === 'Yuran Bulanan' && kb.hasOwnProperty(bln)) kb[bln] += amaun;
  }

  var tb = 0, tv = 0;
  for (var b = 0; b < BULAN.length; b++) {
    var row = b+5, bln = BULAN[b];
    sheet.getRange(row,1).setValue(bln);
    sheet.getRange(row,2).setValue(kb[bln]);
    sheet.getRange(row,3).setValue(kv[bln]);
    tb += kb[bln]; tv += kv[bln];
    if (b%2===0) sheet.getRange(row,1,1,3).setBackground('#f8f9fa');
  }

  var totalRow = 5+BULAN.length;
  sheet.getRange(totalRow,1,1,4).setValues([['JUMLAH', tb, tv, tb+tv]]).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.autoResizeColumns(1,4);
  SpreadsheetApp.getUi().alert('Ringkasan dikemaskini!');
}
