/**
 * ============================================================================
 * ZETTBOS SYSTEM ARCHITECTURE - SETUP & SAFE MIGRATE
 * File: setup.gs
 * Deskripsi: Inisialisasi Database 13 Kolom Presisi & Medium Bus (Bus 1 & Bus 2 - Total 66 Seat)
 * ============================================================================
 */

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Sheet Utama: Peserta (Struktur 13 Kolom Presisi - Sistem Tabungan)
  var sheetPesertaName = 'Peserta';
  var sheetPeserta = ss.getSheetByName(sheetPesertaName);
  var headerPeserta = [
    'ID Transaksi', 
    'Tanggal', 
    'Nama Peserta', 
    'No. WhatsApp',
    'PIN Akses',
    'Alamat (Fix)', 
    'No. Rumah', 
    'Catatan Lokasi', 
    'Metode Pembayaran', 
    'Jumlah Tabungan (Rp)', 
    'SponsorID', 
    'SponsorNama', 
    'RelationToSponsor'
  ];

  if (!sheetPeserta) {
    sheetPeserta = ss.insertSheet(sheetPesertaName);
    sheetPeserta.getRange(1, 1, 1, headerPeserta.length).setValues([headerPeserta]);
    formatHeaderRow_(sheetPeserta, headerPeserta.length, '#EA6A9C');
  } else {
    // Safe Migrate: Update Header Baris 1 ke 13 Kolom Presisi tanpa menghapus data
    sheetPeserta.getRange(1, 1, 1, headerPeserta.length).setValues([headerPeserta]);
    formatHeaderRow_(sheetPeserta, headerPeserta.length, '#EA6A9C');
  }

  // 2. Sheet Log Setoran Tabungan
  var sheetLogName = 'Log_Tabungan';
  var sheetLog = ss.getSheetByName(sheetLogName);
  var headerLog = [
    'ID Log', 
    'Tanggal & Waktu', 
    'ID Transaksi', 
    'Nama Peserta', 
    'Nominal Setoran (Rp)', 
    'Total Tabungan (Rp)', 
    'Admin Penyetor'
  ];

  if (!sheetLog) {
    sheetLog = ss.insertSheet(sheetLogName);
    sheetLog.getRange(1, 1, 1, headerLog.length).setValues([headerLog]);
    formatHeaderRow_(sheetLog, headerLog.length, '#265768');
  } else {
    sheetLog.getRange(1, 1, 1, headerLog.length).setValues([headerLog]);
    formatHeaderRow_(sheetLog, headerLog.length, '#265768');
  }

  // 3. Sheet Kursi Bus (Medium Bus: Bus 1 & Bus 2 @ 33 Seats = Total 66 Seats)
  var sheetBusName = 'Kursi_Bus';
  var sheetBus = ss.getSheetByName(sheetBusName);
  var headerBus = ['ID Bus', 'No Kursi', 'Baris', 'Tipe Kursi', 'ID Peserta', 'Nama Terisi'];

  if (!sheetBus) {
    sheetBus = ss.insertSheet(sheetBusName);
    sheetBus.getRange(1, 1, 1, headerBus.length).setValues([headerBus]);
    formatHeaderRow_(sheetBus, headerBus.length, '#5A56EC');
    
    var seatsBus1 = generateBus33MediumSeatsStructure_('Bus 1');
    var seatsBus2 = generateBus33MediumSeatsStructure_('Bus 2');
    var allSeats = seatsBus1.concat(seatsBus2);
    sheetBus.getRange(2, 1, allSeats.length, headerBus.length).setValues(allSeats);
  } else {
    sheetBus.getRange(1, 1, 1, headerBus.length).setValues([headerBus]);
    formatHeaderRow_(sheetBus, headerBus.length, '#5A56EC');

    var lastRowBus = sheetBus.getLastRow();
    // Safe Migrate: Jika total kursi belum 66 kursi medium (33 x 2), perbarui struktur kursi
    if (lastRowBus !== 67) {
      // Simpan data pendaftar kursi yang sudah terisi sebelumnya
      var existingAssignments = {};
      if (lastRowBus > 1) {
        var oldData = sheetBus.getRange(2, 1, lastRowBus - 1, 6).getValues();
        for (var i = 0; i < oldData.length; i++) {
          var pId = oldData[i][4];
          var pName = oldData[i][5];
          if (pId && pName) {
            existingAssignments[pId] = pName;
          }
        }
      }

      sheetBus.getRange(2, 1, Math.max(lastRowBus, 120), 6).clearContent();
      
      var seatsBus1 = generateBus33MediumSeatsStructure_('Bus 1');
      var seatsBus2 = generateBus33MediumSeatsStructure_('Bus 2');
      var newAllSeats = seatsBus1.concat(seatsBus2);

      // Re-assign pendaftar ke kursi baru secara sekuensial
      var assignedKeys = Object.keys(existingAssignments);
      for (var k = 0; k < assignedKeys.length && k < newAllSeats.length; k++) {
        var curId = assignedKeys[k];
        newAllSeats[k][4] = curId;
        newAllSeats[k][5] = existingAssignments[curId];
      }

      sheetBus.getRange(2, 1, newAllSeats.length, 6).setValues(newAllSeats);
    }
  }

  // 4. Sheet Konfigurasi / Pengaturan Sistem
  var sheetConfigName = 'Konfigurasi';
  var sheetConfig = ss.getSheetByName(sheetConfigName);
  var headerConfig = ['Key', 'Value'];

  if (!sheetConfig) {
    sheetConfig = ss.insertSheet(sheetConfigName);
    sheetConfig.getRange(1, 1, 1, headerConfig.length).setValues([headerConfig]);
    formatHeaderRow_(sheetConfig, headerConfig.length, '#1E2238');
    
    var defaultConfigs = [
      ['ALAMAT_FIX', 'Kamp baru I Jl. Marga Mulya'],
      ['ADMIN_PIN', '1234'],
      ['SUPERADMIN_PIN', '9999'],
      ['NAMA_KEGIATAN', 'Halal Bihalal Warga RT']
    ];
    sheetConfig.getRange(2, 1, defaultConfigs.length, 2).setValues(defaultConfigs);
  }

  // 5. Auto-Clean Migration
  fixScrambledData();

  SpreadsheetApp.flush();
  Logger.log('Safe Migrate Selesai: Database Medium Bus (33 Seat/Bus) Siap Digunakan.');
}

/**
 * ZettBOT Feature: Safe Data Cleanup Script
 */
function fixScrambledData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPeserta = ss.getSheetByName('Peserta');
  if (!sheetPeserta || sheetPeserta.getLastRow() <= 1) return;

  var lastRow = sheetPeserta.getLastRow();
  var range = sheetPeserta.getRange(2, 1, lastRow - 1, 13);
  var values = range.getDisplayValues();

  var cleanedValues = [];
  var isModified = false;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var id = row[0];
    if (!id) continue;

    var colD = String(row[3] || '').trim();

    if (/^kamp/i.test(colD) || colD.indexOf('Jl.') !== -1) {
      isModified = true;
      var generatedPin = String(Math.floor(1000 + Math.random() * 9000));
      
      var cleanedRow = [
        row[0],                  // ID Transaksi
        row[1],                  // Tanggal
        row[2],                  // Nama Peserta
        '081234567890',          // No. WhatsApp
        generatedPin,            // PIN Akses
        row[3],                  // Alamat (Fix)
        row[4],                  // No. Rumah
        row[5],                  // Catatan Lokasi
        row[6] || 'Tunai',       // Metode Pembayaran
        0,                       // Jumlah Tabungan Default 0
        row[8] || '',            // SponsorID
        row[9] || '',            // SponsorNama
        row[10] || ''            // RelationToSponsor
      ];
      cleanedValues.push(cleanedRow);
    } else {
      var tabVal = parseInt(String(row[9]).replace(/\D/g, ''), 10);
      cleanedValues.push([
        row[0], row[1], row[2], row[3], row[4], row[5], 
        row[6], row[7], row[8], isNaN(tabVal) ? 0 : tabVal, 
        row[10], row[11], row[12]
      ]);
    }
  }

  if (isModified && cleanedValues.length > 0) {
    sheetPeserta.getRange(2, 1, cleanedValues.length, 13).setValues(cleanedValues);
    SpreadsheetApp.flush();
  }
}

/**
 * Helper: Membentuk struktur 33 kursi Armada Medium Bus (Konfigurasi 2-2 & 5 Belakang)
 */
function generateBus33MediumSeatsStructure_(busId) {
  var bId = busId || 'Bus 1';
  var rows = [];

  // Baris 1 s/d 7: 4 Kursi per Baris (Konfigurasi 2-2)
  for (var r = 1; r <= 7; r++) {
    var startNum = (r - 1) * 4 + 1;
    var s1 = ('0' + startNum).slice(-2);
    var s2 = ('0' + (startNum + 1)).slice(-2);
    var s3 = ('0' + (startNum + 2)).slice(-2);
    var s4 = ('0' + (startNum + 3)).slice(-2);

    rows.push([bId, s1, r, 'Window (Kiri)', '', '']);
    rows.push([bId, s2, r, 'Aisle (Kiri)', '', '']);
    rows.push([bId, s3, r, 'Aisle (Kanan)', '', '']);
    rows.push([bId, s4, r, 'Window (Kanan)', '', '']);
  }

  // Baris 8 (Belakang Penuh): 5 Kursi (29, 30, 31, 32, 33)
  rows.push([bId, '29', 8, 'Window (Kiri Belakang)', '', '']);
  rows.push([bId, '30', 8, 'Middle (Kiri Belakang)', '', '']);
  rows.push([bId, '31', 8, 'Middle (Tengah Belakang)', '', '']);
  rows.push([bId, '32', 8, 'Middle (Kanan Belakang)', '', '']);
  rows.push([bId, '33', 8, 'Window (Kanan Belakang)', '', '']);

  return rows;
}

function formatHeaderRow_(sheet, colCount, hexColor) {
  var headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setBackground(hexColor)
             .setFontColor('#FFFFFF')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}
