/**
 * SILAP-BMN - Backend Web API (Code.gs)
 * Arsitektur: REST API Decoupled untuk Frontend Blogger
 */

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getDatabaseData';
    var forceRefresh = (e && e.parameter && e.parameter.nocache === '1');

    var result;
    if (action === 'getDatabaseData') {
      result = getDatabaseData(forceRefresh);
    } else {
      result = { success: false, error: 'Aksi tidak dikenal: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  return doGet(e);
}

/**
 * Membaca & Mengolah Data Spreadsheet dengan Cepat
 */
function getDatabaseData(forceRefresh) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'SILAP_BMN_DB_DATA_V2';

    if (forceRefresh) {
      cache.remove(cacheKey); // Paksa hapus cache GAS saat ada permintaan data baru
    } else {
      var cached = cache.get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var sheet = null;
    
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].getName().trim().toLowerCase() === 'sheet1') {
        sheet = sheets[s];
        break;
      }
    }
    
    if (!sheet) {
      sheet = sheets[0];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, items: [] };
    }

    var rawData = sheet.getRange(2, 1, lastRow - 1, 26).getValues();

    var itemsList = [];
    var currentItem = null;

    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];

      var no = row[0];
      var kodeSatker = row[1];
      var kodeBarang = row[2];
      var nup = row[3];
      var namaBarang = row[4];
      var tahun = row[5];
      var nilaiPerolehan = row[6];
      var merk = row[7];
      var tipe = row[8];
      var kondisi = row[9];
      var jenisBmn = row[10];
      var nopol = row[11];
      var foto = row[13];

      if (String(namaBarang).trim() !== "" || String(nopol).trim() !== "" || String(nup).trim() !== "") {
        var itemId = "ITEM_" + (nup || i) + "_" + String(nopol || i).replace(/\s+/g, '');
        
        var numericNilai = 0;
        if (typeof nilaiPerolehan === 'number') {
          numericNilai = nilaiPerolehan;
        } else if (typeof nilaiPerolehan === 'string') {
          numericNilai = parseFloat(nilaiPerolehan.replace(/[^0-9.-]+/g, "")) || 0;
        }

        currentItem = {
          id: itemId,
          no: no || (itemsList.length + 1),
          kodeSatker: String(kodeSatker || '-'),
          kodeBarang: String(kodeBarang || '-'),
          nup: String(nup || '-'),
          namaBarang: String(namaBarang || '-'),
          tahun: String(tahun || '-'),
          nilaiPerolehan: numericNilai,
          nilaiPerolehanFormatted: formatRupiah(numericNilai),
          merk: String(merk || '-'),
          tipe: String(tipe || '-'),
          kondisi: String(kondisi || 'Baik'),
          jenisBmn: String(jenisBmn || '-'),
          nopol: String(nopol || '-'),
          foto: fixDriveUrl(foto),
          history: []
        };

        itemsList.push(currentItem);
      }

      if (currentItem) {
        var tglKeluar = row[14] ? formatDate(row[14]) : '';
        var tglMasuk = row[20] ? formatDate(row[20]) : '';

        if (tglKeluar || row[15] || tglMasuk || row[21]) {
          currentItem.history.push({
            trxNo: currentItem.history.length + 1,
            keluarTgl: tglKeluar || '-',
            keluarPetugasSerah: String(row[15] || '-'),
            keluarPetugasTerima: String(row[16] || '-'),
            keluarBA: String(row[17] || '-'),
            keluarFoto: fixDriveUrl(row[18]),
            keluarLokasi: String(row[19] || '-'),
            masukTgl: tglMasuk || '-',
            masukPetugasTerima: String(row[21] || '-'),
            masukPetugasSerah: String(row[22] || '-'),
            masukBA: String(row[23] || '-'),
            masukFoto: fixDriveUrl(row[24]),
            masukLokasi: String(row[25] || '-')
          });
        }
      }
    }

    var responsePayload = {
      success: true,
      items: itemsList
    };

    try {
      cache.put(cacheKey, JSON.stringify(responsePayload), 300);
    } catch (cErr) {}

    return responsePayload;

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function formatRupiah(val) {
  if (!val || isNaN(val)) return 'Rp 0';
  return 'Rp ' + Math.round(val).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function formatDate(d) {
  if (!d) return '';
  if (d instanceof Date) {
    var day = ("0" + d.getDate()).slice(-2);
    var month = ("0" + (d.getMonth() + 1)).slice(-2);
    var year = d.getFullYear();
    return day + "/" + month + "/" + year;
  }
  return String(d);
}

function fixDriveUrl(url) {
  if (!url || typeof url !== 'string') return '';
  if (url.indexOf('drive.google.com') !== -1) {
    var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      return 'https://lh3.googleusercontent.com/d/' + match[1];
    }
  }
  return url;
}
