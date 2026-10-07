/**
 * ============================================================================
 * ZETTBOS SYSTEM ARCHITECTURE - SETUP & SAFE MIGRATE
 * File: setup.gs
 * Deskripsi: Inisialisasi Database 13 Kolom Presisi & Medium Bus (Bus 1, 2, 3 - 99 Seats)
 *            Fitur Auto-Sync 67 Peserta Lengkap & Pengelompokan Ulang Kursi Keluarga
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

  // 3. Sheet Kursi Bus (Medium Bus: 3 Armada @ 33 Kursi = Total 99 Kursi)
  var sheetBusName = 'Kursi_Bus';
  var sheetBus = ss.getSheetByName(sheetBusName);
  var headerBus = ['ID Bus', 'No Kursi', 'Baris', 'Tipe Kursi', 'ID Peserta', 'Nama Terisi'];

  if (!sheetBus) {
    sheetBus = ss.insertSheet(sheetBusName);
    sheetBus.getRange(1, 1, 1, headerBus.length).setValues([headerBus]);
    formatHeaderRow_(sheetBus, headerBus.length, '#5A56EC');
    
    var seatsBus1 = generateBus33MediumSeatsStructure_('Bus 1');
    var seatsBus2 = generateBus33MediumSeatsStructure_('Bus 2');
    var seatsBus3 = generateBus33MediumSeatsStructure_('Bus 3');
    var allSeats = seatsBus1.concat(seatsBus2).concat(seatsBus3);
    sheetBus.getRange(2, 1, allSeats.length, headerBus.length).setValues(allSeats);
  } else {
    sheetBus.getRange(1, 1, 1, headerBus.length).setValues([headerBus]);
    formatHeaderRow_(sheetBus, headerBus.length, '#5A56EC');
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

  // 5. Auto-Clean Migration Data
  fixScrambledData();

  // 6. SINKRONISASI TOTAL: Pastikan seluruh 67 peserta dari sheet Peserta teralokasikan ke 99 kursi bus
  //    sekaligus merapikan posisi keluarga agar berdampingan (contoh: 4A Suwito & 4B Bu RT)
  syncAllPesertaToSeats_();

  SpreadsheetApp.flush();
  Logger.log('Setup & Sinkronisasi Selesai: Seluruh peserta teralokasi di Bus 1, Bus 2, & Bus 3.');
}

/**
 * REKONSILIASI PENUH: Membaca sheet Peserta (Single Source of Truth) dan memetakan ke Kursi_Bus
 * Menghilangkan masalah 23 peserta hilang serta menyatukan anggota keluarga yang terpisah.
 */
function syncAllPesertaToSeats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPeserta = ss.getSheetByName('Peserta');
  var sheetBus = ss.getSheetByName('Kursi_Bus');

  if (!sheetPeserta || !sheetBus) return;

  var lastRowPeserta = sheetPeserta.getLastRow();
  if (lastRowPeserta <= 1) return;

  // 1. Ambil seluruh data peserta dari sheet Peserta
  var rawPeserta = sheetPeserta.getRange(2, 1, lastRowPeserta - 1, 13).getValues();
  var pesertaMap = {};
  var allPesertaList = [];

  for (var p = 0; p < rawPeserta.length; p++) {
    var pId = String(rawPeserta[p][0] || '').trim();
    var pName = String(rawPeserta[p][2] || '').trim();
    var pSponsorId = String(rawPeserta[p][10] || '').trim();
    var pSponsorNama = String(rawPeserta[p][11] || '').trim();
    var pRel = String(rawPeserta[p][12] || '').trim();

    if (!pId || !pName) continue;

    var pObj = {
      id: pId,
      name: pName,
      sponsorId: pSponsorId,
      sponsorNama: pSponsorNama,
      hubungan: pRel
    };

    pesertaMap[pId] = pObj;
    allPesertaList.push(pObj);
  }

  // 2. Siapkan kerangka 99 Kursi Standar (33 Bus 1, 33 Bus 2, 33 Bus 3)
  var seatsBus1 = generateBus33MediumSeatsStructure_('Bus 1');
  var seatsBus2 = generateBus33MediumSeatsStructure_('Bus 2');
  var seatsBus3 = generateBus33MediumSeatsStructure_('Bus 3');
  var newAllSeats = seatsBus1.concat(seatsBus2).concat(seatsBus3);

  // 3. Baca data kursi yang ada saat ini
  var lastRowBus = sheetBus.getLastRow();
  var currentSeatAssignments = {};

  if (lastRowBus > 1) {
    var currentBusData = sheetBus.getRange(2, 1, lastRowBus - 1, 6).getValues();
    for (var b = 0; b < currentBusData.length; b++) {
      var bId = String(currentBusData[b][0] || 'Bus 1').trim();
      var bSeat = convertSeatToLetterFormat_(String(currentBusData[b][1] || '').trim().toUpperCase());
      var seatedId = String(currentBusData[b][4] || '').trim();
      var seatedName = String(currentBusData[b][5] || '').trim();

      if (seatedName && isSeatValid33_(bSeat)) {
        // Cocokkan ID jika ada di sheet Peserta, atau cari ID berdasarkan nama
        var matchedId = seatedId;
        if (!matchedId || !pesertaMap[matchedId]) {
          for (var idKey in pesertaMap) {
            if (pesertaMap[idKey].name.toLowerCase() === seatedName.toLowerCase()) {
              matchedId = idKey;
              break;
            }
          }
        }

        var key = bId + '_' + bSeat;
        if (matchedId && pesertaMap[matchedId]) {
          currentSeatAssignments[key] = {
            id: matchedId,
            name: pesertaMap[matchedId].name
          };
        }
      }
    }
  }

  // 4. Petakan kursi yang sudah ada ke struktur baru
  var assignedParticipants = {};
  for (var s = 0; s < newAllSeats.length; s++) {
    var seatKey = newAllSeats[s][0] + '_' + newAllSeats[s][1];
    if (currentSeatAssignments[seatKey]) {
      var curAssign = currentSeatAssignments[seatKey];
      // Pastikan seorang peserta tidak mendapat kursi ganda
      if (!assignedParticipants[curAssign.id]) {
        newAllSeats[s][4] = curAssign.id;
        newAllSeats[s][5] = curAssign.name;
        assignedParticipants[curAssign.id] = {
          busId: newAllSeats[s][0],
          noKursi: newAllSeats[s][1],
          seatIndex: s
        };
      }
    }
  }

  // 5. SOLUSI KHUSUS KELUARGA TERPISAH:
  //    Periksa apakah ada anggota keluarga yang sudah duduk tapi tidak berdampingan dengan sponsornya
  //    Contoh: Suwito di 4A dan Bu RT di 3C -> tukar 4B dengan 3C jika 4B adalah peserta solo
  for (var f = 0; f < allPesertaList.length; f++) {
    var member = allPesertaList[f];
    if (member.sponsorId && assignedParticipants[member.sponsorId] && assignedParticipants[member.id]) {
      var sponsorSeatInfo = assignedParticipants[member.sponsorId];
      var memberSeatInfo = assignedParticipants[member.id];

      // Jika keduanya berada di armada yang sama
      if (sponsorSeatInfo.busId === memberSeatInfo.busId) {
        var idealPartnerSeat = getPrimaryPartnerSeatLetter_(sponsorSeatInfo.noKursi);

        // Jika anggota keluarga saat ini BELUM duduk di kursi pasangannya
        if (memberSeatInfo.noKursi !== idealPartnerSeat) {
          // Cari index kursi ideal di armada tersebut
          var idealSeatIdx = findSeatIndexInArray_(newAllSeats, sponsorSeatInfo.busId, idealPartnerSeat);
          var memberSeatIdx = memberSeatInfo.seatIndex;

          if (idealSeatIdx !== -1) {
            var currentOccupantId = newAllSeats[idealSeatIdx][4];
            var currentOccupantName = newAllSeats[idealSeatIdx][5];

            // Jika kursi ideal kosong, langsung pindahkan anggota keluarga ke sana
            if (!currentOccupantName) {
              newAllSeats[idealSeatIdx][4] = member.id;
              newAllSeats[idealSeatIdx][5] = member.name;
              newAllSeats[memberSeatIdx][4] = '';
              newAllSeats[memberSeatIdx][5] = '';
              assignedParticipants[member.id] = { busId: sponsorSeatInfo.busId, noKursi: idealPartnerSeat, seatIndex: idealSeatIdx };
            } 
            // Jika kursi ideal diduduki peserta mandiri (solo), lakukan pertukaran (swap)
            else if (isSoloFromMap_(pesertaMap, currentOccupantId)) {
              newAllSeats[idealSeatIdx][4] = member.id;
              newAllSeats[idealSeatIdx][5] = member.name;

              newAllSeats[memberSeatIdx][4] = currentOccupantId;
              newAllSeats[memberSeatIdx][5] = currentOccupantName;

              assignedParticipants[member.id] = { busId: sponsorSeatInfo.busId, noKursi: idealPartnerSeat, seatIndex: idealSeatIdx };
              assignedParticipants[currentOccupantId] = { busId: memberSeatInfo.busId, noKursi: memberSeatInfo.noKursi, seatIndex: memberSeatIdx };
            }
          }
        }
      }
    }
  }

  // 6. ALOKASIKAN SELURUH PESERTA YANG BELUM DAPAT KURSI (23 PESERTA YANG HILANG)
  //    Prioritas 1: Jika anggota keluarga, coba tempatkan di samping sponsor
  //    Prioritas 2: Jika tidak, tempatkan di kursi kosong pertama di Bus 1, Bus 2, lalu Bus 3
  for (var u = 0; u < allPesertaList.length; u++) {
    var pItem = allPesertaList[u];
    if (assignedParticipants[pItem.id]) continue; // Sudah punya kursi

    var placed = false;

    // Jika memiliki sponsor dan sponsor sudah punya kursi
    if (pItem.sponsorId && assignedParticipants[pItem.sponsorId]) {
      var spInfo = assignedParticipants[pItem.sponsorId];
      var candidateSeats = getAdjacentSeatsLetter_(spInfo.noKursi);

      // Cari kursi tetangga yang kosong di bus sponsor
      for (var cs = 0; cs < candidateSeats.length; cs++) {
        var candCode = candidateSeats[cs];
        var candIdx = findSeatIndexInArray_(newAllSeats, spInfo.busId, candCode);
        if (candIdx !== -1 && !newAllSeats[candIdx][4]) {
          newAllSeats[candIdx][4] = pItem.id;
          newAllSeats[candIdx][5] = pItem.name;
          assignedParticipants[pItem.id] = { busId: spInfo.busId, noKursi: candCode, seatIndex: candIdx };
          placed = true;
          break;
        }
      }
    }

    // Jika belum dapat tempat duduk berdampingan, ambil kursi kosong pertama yang tersedia
    if (!placed) {
      for (var em = 0; em < newAllSeats.length; em++) {
        if (!newAllSeats[em][4]) {
          newAllSeats[em][4] = pItem.id;
          newAllSeats[em][5] = pItem.name;
          assignedParticipants[pItem.id] = { busId: newAllSeats[em][0], noKursi: newAllSeats[em][1], seatIndex: em };
          placed = true;
          break;
        }
      }
    }
  }

  // 7. Bersihkan dan tulis ulang 99 baris ke sheet Kursi_Bus
  sheetBus.getRange(2, 1, Math.max(lastRowBus, 120), 6).clearContent();
  sheetBus.getRange(2, 1, newAllSeats.length, 6).setValues(newAllSeats);
}

/**
 * Helper: Mencari index kursi dalam array newAllSeats
 */
function findSeatIndexInArray_(seatsArray, busId, seatNo) {
  var sClean = String(seatNo || '').trim().toUpperCase();
  for (var i = 0; i < seatsArray.length; i++) {
    if (seatsArray[i][0] === busId && String(seatsArray[i][1]).trim().toUpperCase() === sClean) {
      return i;
    }
  }
  return -1;
}

/**
 * Helper: Menemukan pasangan deret sebangku primer
 * 1A berpasangan dengan 1B, 1C berpasangan dengan 1D
 */
function getPrimaryPartnerSeatLetter_(seatCode) {
  var s = String(seatCode || '').trim().toUpperCase();
  var match = s.match(/^([1-7])([A-D])$/);
  if (match) {
    var row = match[1];
    var col = match[2];
    if (col === 'A') return row + 'B';
    if (col === 'B') return row + 'A';
    if (col === 'C') return row + 'D';
    if (col === 'D') return row + 'C';
  } else if (/^8[A-E]$/.test(s)) {
    if (s === '8A') return '8B';
    if (s === '8B') return '8A';
    if (s === '8C') return '8D';
    if (s === '8D') return '8C';
    if (s === '8E') return '8D';
  }
  return s;
}

/**
 * Helper: Mengetahui apakah seseorang adalah penumpang mandiri (solo) dari objek pesertaMap
 */
function isSoloFromMap_(pesertaMap, pId) {
  if (!pId || !pesertaMap[pId]) return true;
  var p = pesertaMap[pId];
  if (p.sponsorId && p.sponsorId !== '') return false;
  for (var key in pesertaMap) {
    if (pesertaMap[key].sponsorId === pId) return false;
  }
  return true;
}

/**
 * Helper: Membentuk struktur 33 kursi Armada Medium Bus dengan kode huruf (1A-7D & 8A-8E)
 */
function generateBus33MediumSeatsStructure_(busId) {
  var bId = busId || 'Bus 1';
  var rows = [];

  for (var r = 1; r <= 7; r++) {
    rows.push([bId, r + 'A', r, 'Window (Kiri)', '', '']);
    rows.push([bId, r + 'B', r, 'Aisle (Kiri)', '', '']);
    rows.push([bId, r + 'C', r, 'Aisle (Kanan)', '', '']);
    rows.push([bId, r + 'D', r, 'Window (Kanan)', '', '']);
  }

  rows.push([bId, '8A', 8, 'Window (Kiri Belakang)', '', '']);
  rows.push([bId, '8B', 8, 'Middle (Kiri Belakang)', '', '']);
  rows.push([bId, '8C', 8, 'Middle (Tengah Belakang)', '', '']);
  rows.push([bId, '8D', 8, 'Middle (Kanan Belakang)', '', '']);
  rows.push([bId, '8E', 8, 'Window (Kanan Belakang)', '', '']);

  return rows;
}

/**
 * Helper: Konversi nomor kursi lama (numerik) ke format huruf resmi denah
 */
function convertSeatToLetterFormat_(seatStr) {
  var s = String(seatStr || '').trim().toUpperCase();
  if (/^([1-7][A-D]|8[A-E])$/.test(s)) {
    return s;
  }
  var num = parseInt(s, 10);
  if (isNaN(num)) return s;

  if (num >= 1 && num <= 28) {
    var r = Math.floor((num - 1) / 4) + 1;
    var colIdx = (num - 1) % 4;
    var colLetters = ['A', 'B', 'C', 'D'];
    return r + colLetters[colIdx];
  } else if (num >= 29 && num <= 33) {
    var row8Letters = ['8A', '8B', '8C', '8D', '8E'];
    return row8Letters[num - 29];
  }
  return s;
}

function isSeatValid33_(seatStr) {
  return /^([1-7][A-D]|8[A-E])$/.test(String(seatStr || '').trim().toUpperCase());
}

function formatHeaderRow_(sheet, colCount, hexColor) {
  var headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setBackground(hexColor)
             .setFontColor('#FFFFFF')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

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
        row[0],
        row[1],
        row[2],
        '081234567890',
        generatedPin,
        row[3],
        row[4],
        row[5],
        row[6] || 'Tunai',
        0,
        row[8] || '',
        row[9] || '',
        row[10] || ''
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
