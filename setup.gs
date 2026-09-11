/**
 * ============================================================================
 * ZETTBOS SYSTEM ARCHITECTURE - SETUP & SAFE MIGRATE
 * File: setup.gs
 * Deskripsi: Inisialisasi Database 13 Kolom Presisi & Multi-Bus (Bus 1 & Bus 2 - 108 Seat)
 * ============================================================================
 */

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Sheet Utama: Peserta (Struktur 13 Kolom Presisi)
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
    'Status Pembayaran', 
    'SponsorID', 
    'SponsorNama', 
    'RelationToSponsor'
  ];

  if (!sheetPeserta) {
    sheetPeserta = ss.insertSheet(sheetPesertaName);
    sheetPeserta.getRange(1, 1, 1, headerPeserta.length).setValues([headerPeserta]);
    formatHeaderRow_(sheetPeserta, headerPeserta.length, '#EA6A9C');
  } else {
    // Safe Migrate: Paksakan update Header Baris 1 ke 13 Kolom Presisi tanpa menghapus data
    sheetPeserta.getRange(1, 1, 1, headerPeserta.length).setValues([headerPeserta]);
    formatHeaderRow_(sheetPeserta, headerPeserta.length, '#EA6A9C');
  }

  // 2. Sheet Kursi Bus (Multi-Bus: Bus 1 & Bus 2 @ 54 Seats)
  var sheetBusName = 'Kursi_Bus';
  var sheetBus = ss.getSheetByName(sheetBusName);
  var headerBus = ['ID Bus', 'No Kursi', 'Baris', 'Tipe Kursi', 'ID Peserta', 'Nama Terisi'];

  if (!sheetBus) {
    sheetBus = ss.insertSheet(sheetBusName);
    sheetBus.getRange(1, 1, 1, headerBus.length).setValues([headerBus]);
    formatHeaderRow_(sheetBus, headerBus.length, '#5A56EC');
  } else {
    // Safe Migrate header baris 1
    sheetBus.getRange(1, 1, 1, headerBus.length).setValues([headerBus]);
    formatHeaderRow_(sheetBus, headerBus.length, '#5A56EC');
  }

  // Safe Migrate kursi: Jika data kursi masih format lama (5 kolom) atau kurang dari 108 kursi, perbarui/tambahkan
  var lastRowBus = sheetBus.getLastRow();
  if (lastRowBus <= 1) {
    var seatsDataBus1 = generateBus54SeatsStructure_('Bus 1');
    var seatsDataBus2 = generateBus54SeatsStructure_('Bus 2');
    var allSeats = seatsDataBus1.concat(seatsDataBus2);
    sheetBus.getRange(2, 1, allSeats.length, headerBus.length).setValues(allSeats);
  } else {
    // Cek apakah data bus lama hanya 5 kolom (belum ada ID Bus di kolom 1)
    var firstDataCol = String(sheetBus.getRange(2, 1, 1, 1).getDisplayValue() || '');
    if (firstDataCol.indexOf('Bus') === -1) {
      // Format lama: sisipkan kolom ID Bus = 'Bus 1'
      var oldBusData = sheetBus.getRange(2, 1, lastRowBus - 1, 5).getValues();
      var migratedBusData = [];
      for (var b = 0; b < oldBusData.length; b++) {
        migratedBusData.push([
          'Bus 1',
          oldBusData[b][0],
          oldBusData[b][1],
          oldBusData[b][2],
          oldBusData[b][3],
          oldBusData[b][4]
        ]);
      }
      sheetBus.getRange(2, 1, migratedBusData.length, 6).setValues(migratedBusData);
      
      // Jika belum ada Bus 2, tambahkan 54 seat Bus 2 di bawahnya
      if (migratedBusData.length < 108) {
        var bus2Seats = generateBus54SeatsStructure_('Bus 2');
        sheetBus.getRange(migratedBusData.length + 2, 1, bus2Seats.length, 6).setValues(bus2Seats);
      }
    } else {
      // Sudah ada ID Bus, pastikan Bus 2 juga terinisialisasi jika total kursi < 108
      if (lastRowBus - 1 < 108) {
        var bus2Data = generateBus54SeatsStructure_('Bus 2');
        sheetBus.getRange(56, 1, bus2Data.length, 6).setValues(bus2Data);
      }
    }
  }

  // 3. Sheet Konfigurasi / Pengaturan Sistem
  var sheetConfigName = 'Konfigurasi';
  var sheetConfig = ss.getSheetByName(sheetConfigName);
  var headerConfig = ['Key', 'Value'];

  if (!sheetConfig) {
    sheetConfig = ss.insertSheet(sheetConfigName);
    sheetConfig.getRange(1, 1, 1, headerConfig.length).setValues([headerConfig]);
    formatHeaderRow_(sheetConfig, headerConfig.length, '#265768');
    
    var defaultConfigs = [
      ['ALAMAT_FIX', 'Kamp baru I Jl. Marga Mulya'],
      ['ADMIN_PIN', '1234'],
      ['SUPERADMIN_PIN', '9999'],
      ['NAMA_KEGIATAN', 'Halal Bihalal Warga RT']
    ];
    sheetConfig.getRange(2, 1, defaultConfigs.length, 2).setValues(defaultConfigs);
  }

  // 4. Jalankan Auto-Clean Migration untuk merapikan data teracak/tergeser
  fixScrambledData();

  SpreadsheetApp.flush();
  Logger.log('Safe Migrate & Auto-Clean Multi-Bus Selesai: Database Siap Digunakan.');
}

/**
 * ZettBOT Feature: Auto-Clean / Fix Migration Script
 * Memeriksa dan merapikan data format lama (11 kolom) menjadi 13 kolom presisi.
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

    // Jika Kolom D berisi Alamat "Kamp...", artinya ini baris data format lama (11 kolom)
    if (/^kamp/i.test(colD) || colD.indexOf('Jl.') !== -1) {
      isModified = true;
      var generatedPin = String(Math.floor(1000 + Math.random() * 9000));
      
      // Susun ulang posisi data secara presisi ke 13 kolom
      var cleanedRow = [
        row[0],                  // ID Transaksi
        row[1],                  // Tanggal
        row[2],                  // Nama Peserta
        '081234567890',          // No. WhatsApp (Default Placeholder)
        generatedPin,            // PIN Akses (Auto-Generate 4 Digit)
        row[3],                  // Alamat (Fix) -> "Kamp baru I Jl..."
        row[4],                  // No. Rumah
        row[5],                  // Catatan Lokasi
        row[6] || 'Tunai',       // Metode Pembayaran
        row[7] || 'Lunas',       // Status Pembayaran
        row[8] || '',            // SponsorID
        row[9] || '',            // SponsorNama
        row[10] || ''            // RelationToSponsor
      ];
      cleanedValues.push(cleanedRow);
    } else {
      // Data sudah dalam format 13 kolom yang benar
      cleanedValues.push([
        row[0], row[1], row[2], row[3], row[4], row[5], 
        row[6], row[7], row[8], row[9], row[10], row[11], row[12]
      ]);
    }
  }

  if (isModified && cleanedValues.length > 0) {
    sheetPeserta.getRange(2, 1, cleanedValues.length, 13).setValues(cleanedValues);
    SpreadsheetApp.flush();
    Logger.log('Auto-Clean Selesai: Data lama berhasil disesuaikan dengan header baru!');
  }
}

/**
 * Helper: Membentuk struktur 54 kursi bus pariwisata (Konfigurasi 2-3) untuk ID Bus tertentu
 */
function generateBus54SeatsStructure_(busId) {
  var bId = busId || 'Bus 1';
  var rows = [];
  for (var r = 1; r <= 10; r++) {
    rows.push([bId, r + 'A', r, 'Window (Kiri)', '', '']);
    rows.push([bId, r + 'B', r, 'Aisle (Kiri)', '', '']);
    rows.push([bId, r + 'C', r, 'Aisle (Kanan)', '', '']);
    rows.push([bId, r + 'D', r, 'Middle (Kanan)', '', '']);
    rows.push([bId, r + 'E', r, 'Window (Kanan)', '', '']);
  }
  rows.push([bId, '11A', 11, 'Window (Kiri Belakang)', '', '']);
  rows.push([bId, '11B', 11, 'Middle (Kiri Belakang)', '', '']);
  rows.push([bId, '11C', 11, 'Middle (Kanan Belakang)', '', '']);
  rows.push([bId, '11D', 11, 'Window (Kanan Belakang)', '', '']);

  return rows;
}

/**
 * Format baris header tabel dengan warna identitas Zettbos
 */
function formatHeaderRow_(sheet, colCount, hexColor) {
  var headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setBackground(hexColor)
             .setFontColor('#FFFFFF')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}
