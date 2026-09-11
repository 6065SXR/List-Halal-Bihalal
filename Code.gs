/**
 * ============================================================================
 * ZETTBOS BACKEND SERVICE & BUSINESS LOGIC (MODULAR ARCHITECTURE)
 * File: code.gs
 * Deskripsi: Serves Modular Web App templates, API CRUD, Duplicate Checks,
 *            Multi-Bus Allocation, Role-Based Access & Auto-Cleanup
 * ============================================================================
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Registrasi Halal Bihalal & Seat Bus RT')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Scriptlet Helper: Memasukkan sub-file HTML (css.html, js.html) ke dalam index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function generateSequentialId_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Peserta');
  if (!sheet || sheet.getLastRow() <= 1) {
    return 'HAL-0001';
  }

  var lastRow = sheet.getLastRow();
  var idRange = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  
  var maxNum = 0;
  for (var i = 0; i < idRange.length; i++) {
    var str = idRange[i][0] || '';
    if (str.indexOf('HAL-') === 0) {
      var numPart = parseInt(str.replace('HAL-', ''), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  }
  
  var nextNum = maxNum + 1;
  var padded = ('0000' + nextNum).slice(-4);
  return 'HAL-' + padded;
}

/**
 * Formatting Nomor Rumah: Mendukung Angka murni (12) maupun Alfanumerik (32 A, 32a)
 */
function formatNoRumah_(val) {
  var str = String(val || '').trim();
  if (!str) return 'No. -';
  var clean = str.replace(/^no\.?\s*/i, '').trim();
  return 'No. ' + clean;
}

/**
 * Logika Pembersihan Otomatis:
 * Menghapus permanen peserta yang belum Lunas jika sudah melewati batas waktu Desember 2026.
 */
function autoCleanUnpaidAfterDeadline_() {
  try {
    // Batas waktu akhir: 31 Desember 2026 pukul 23:59:59 WIB
    var deadline = new Date(2026, 11, 31, 23, 59, 59);
    var now = new Date();

    // Jika belum melewati batas waktu, lewati pembersihan
    if (now <= deadline) return;

    var ss = getSpreadsheet_();
    var sheetPeserta = ss.getSheetByName('Peserta');
    var sheetBus = ss.getSheetByName('Kursi_Bus');

    if (!sheetPeserta || !sheetBus || sheetPeserta.getLastRow() <= 1) return;

    var lastRow = sheetPeserta.getLastRow();
    var data = sheetPeserta.getRange(2, 1, lastRow - 1, 13).getDisplayValues();

    var rowsToDelete = [];
    var affectedIds = [];

    for (var i = 0; i < data.length; i++) {
      var status = String(data[i][9] || '').trim(); // Kolom Status Pembayaran (Kolom J / Indeks 9)
      if (status !== 'Lunas') {
        rowsToDelete.push(i + 2); // Simpan nomor baris sheet
        affectedIds.push(data[i][0]); // Simpan ID Transaksi
      }
    }

    if (rowsToDelete.length === 0) return;

    // Hapus baris dari yang paling bawah agar indeks tidak bergeser
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var r = 0; r < rowsToDelete.length; r++) {
      sheetPeserta.deleteRow(rowsToDelete[r]);
    }

    // Kosongkan alokasi kursi di sheet Kursi_Bus untuk ID yang terhapus
    var busLastRow = sheetBus.getLastRow();
    if (busLastRow > 1) {
      var busRange = sheetBus.getRange(2, 1, busLastRow - 1, 6);
      var busData = busRange.getValues();
      var changed = false;

      for (var b = 0; b < busData.length; b++) {
        var seatedId = String(busData[b][4] || '').trim();
        if (affectedIds.indexOf(seatedId) !== -1) {
          busData[b][4] = ''; // Reset ID Peserta
          busData[b][5] = ''; // Reset Nama Terisi
          changed = true;
        }
      }

      if (changed) {
        busRange.setValues(busData);
      }
    }

    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log('Error auto-cleanup: ' + err.message);
  }
}

function getAppData(searchQuery, statusFilter) {
  try {
    // Jalankan auto-cleanup terlebih dahulu jika sudah melewai deadline 2026
    autoCleanUnpaidAfterDeadline_();

    var ss = getSpreadsheet_();
    var sheetPeserta = ss.getSheetByName('Peserta');
    var sheetBus = ss.getSheetByName('Kursi_Bus');

    if (!sheetPeserta || !sheetBus) {
      return { success: false, message: 'Database belum diinisialisasi. Silakan jalankan setupDatabase() dahulu.' };
    }

    var pesertaLastRow = sheetPeserta.getLastRow();
    var pesertaList = [];
    var totalKepalaKeluarga = 0;
    var totalWarga = 0;
    var totalLunas = 0;
    var totalMenunggu = 0;

    if (pesertaLastRow > 1) {
      var rawData = sheetPeserta.getRange(2, 1, pesertaLastRow - 1, 13).getDisplayValues();
      var query = (searchQuery || '').toLowerCase().trim();
      var filter = (statusFilter || 'SEMUA').toUpperCase().trim();

      var allRows = [];

      for (var i = 0; i < rawData.length; i++) {
        var row = rawData[i];
        var item = {
          id: row[0],
          tanggal: row[1],
          nama: row[2],
          noWhatsapp: row[3],
          pinAkses: row[4],
          alamat: row[5],
          noRumah: row[6],
          catatanLokasi: row[7],
          metodePembayaran: row[8],
          statusPembayaran: row[9],
          sponsorId: row[10],
          sponsorNama: row[11],
          hubungan: row[12],
          familyMembers: []
        };

        totalWarga++;
        if (item.statusPembayaran === 'Lunas') {
          totalLunas++;
        } else {
          totalMenunggu++;
        }

        allRows.push(item);
      }

      var primaries = [];
      var familyMap = {};

      for (var j = 0; j < allRows.length; j++) {
        var current = allRows[j];
        if (!current.sponsorId) {
          totalKepalaKeluarga++;
          primaries.push(current);
          familyMap[current.id] = current;
        }
      }

      for (var k = 0; k < allRows.length; k++) {
        var member = allRows[k];
        if (member.sponsorId && familyMap[member.sponsorId]) {
          familyMap[member.sponsorId].familyMembers.push(member);
        }
      }

      for (var m = 0; m < primaries.length; m++) {
        var p = primaries[m];
        var matchStatus = (filter === 'SEMUA' || p.statusPembayaran.toUpperCase() === filter);
        
        var matchQuery = false;
        if (!query) {
          matchQuery = true;
        } else {
          var combinedSearch = (p.nama + ' ' + p.id + ' ' + p.noRumah + ' ' + p.noWhatsapp + ' ' + p.catatanLokasi).toLowerCase();
          if (combinedSearch.indexOf(query) !== -1) {
            matchQuery = true;
          } else {
            for (var f = 0; f < p.familyMembers.length; f++) {
              if (p.familyMembers[f].nama.toLowerCase().indexOf(query) !== -1) {
                matchQuery = true;
                break;
              }
            }
          }
        }

        if (matchStatus && matchQuery) {
          pesertaList.push(p);
        }
      }
    }

    var busLastRow = sheetBus.getLastRow();
    var seatsList = [];
    var occupiedBus1 = 0;
    var occupiedBus2 = 0;

    if (busLastRow > 1) {
      var rawBus = sheetBus.getRange(2, 1, busLastRow - 1, 6).getDisplayValues();
      for (var b = 0; b < rawBus.length; b++) {
        var seatRow = rawBus[b];
        var busId = seatRow[0] || 'Bus 1';
        var isOccupied = (seatRow[5] && seatRow[5].trim() !== '');
        
        if (isOccupied) {
          if (busId === 'Bus 1') occupiedBus1++;
          if (busId === 'Bus 2') occupiedBus2++;
        }

        seatsList.push({
          idBus: busId,
          noKursi: seatRow[1],
          baris: parseInt(seatRow[2], 10) || 1,
          tipeKursi: seatRow[3],
          idPeserta: seatRow[4],
          namaTerisi: seatRow[5]
        });
      }
    }

    return {
      success: true,
      data: {
        peserta: pesertaList,
        stats: {
          totalWarga: totalWarga,
          totalKepalaKeluarga: totalKepalaKeluarga,
          totalLunas: totalLunas,
          totalMenunggu: totalMenunggu,
          totalKursiBus1: 54,
          kursiTerisiBus1: occupiedBus1,
          kursiTersisaBus1: Math.max(0, 54 - occupiedBus1),
          totalKursiBus2: 54,
          kursiTerisiBus2: occupiedBus2,
          kursiTersisaBus2: Math.max(0, 54 - occupiedBus2),
          totalKursiAll: 108,
          kursiTerisiAll: occupiedBus1 + occupiedBus2
        },
        seats: seatsList
      }
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat data: ' + err.message };
  }
}

function registerParticipant(payload) {
  try {
    var ss = getSpreadsheet_();
    var sheetPeserta = ss.getSheetByName('Peserta');
    var sheetBus = ss.getSheetByName('Kursi_Bus');

    if (!sheetPeserta || !sheetBus) {
      return { success: false, message: 'Sheet database tidak ditemukan.' };
    }

    var inputWaRaw = String(payload.noWhatsapp || '').trim();
    var inputWaClean = inputWaRaw.replace(/\D/g, ''); // Ambil digit murni

    var formattedNoRumah = formatNoRumah_(payload.noRumah);
    var inputRumahClean = formattedNoRumah.toLowerCase().replace(/\s+/g, '');

    // Check duplicate Phone/WhatsApp & Duplicate House Number
    var lastRow = sheetPeserta.getLastRow();
    if (lastRow > 1) {
      var existingData = sheetPeserta.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
      for (var i = 0; i < existingData.length; i++) {
        var existWaClean = String(existingData[i][3] || '').replace(/\D/g, '');
        var existRumahClean = String(existingData[i][6] || '').toLowerCase().replace(/\s+/g, '');

        if (inputWaClean && existWaClean && inputWaClean === existWaClean) {
          return {
            success: false,
            message: 'nomor telepon anda sudah terdaftar silahkan hubungi Pak Agus atau Tommy'
          };
        }

        if (inputRumahClean && existRumahClean && inputRumahClean === existRumahClean) {
          return {
            success: false,
            message: 'nomor rumah anda sudah terdaftar silahkan hubungi Pak Agus atau Tommy'
          };
        }
      }
    }

    var idTransaksi = generateSequentialId_();
    var nowFormatted = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var alamatFix = 'Kamp baru I Jl. Marga Mulya';
    var statusDefault = payload.statusPembayaran || 'Menunggu Konfirmasi';
    
    var generatedPin = String(Math.floor(1000 + Math.random() * 9000));

    var newRow = [
      idTransaksi,
      nowFormatted,
      String(payload.nama || '').trim(),
      inputWaRaw,
      generatedPin,
      alamatFix,
      formattedNoRumah,
      String(payload.catatanLokasi || '').trim(),
      'Tunai',
      statusDefault,
      '',
      '',
      ''
    ];

    sheetPeserta.appendRow(newRow);

    var assignedInfo = autoAssignFirstAvailableSeatMultiBus_(sheetBus, idTransaksi, newRow[2]);

    SpreadsheetApp.flush();

    return {
      success: true,
      idTransaksi: idTransaksi,
      nama: newRow[2],
      pinAkses: generatedPin,
      noWhatsapp: newRow[3],
      assignedBus: assignedInfo ? assignedInfo.busId : 'Bus 1',
      assignedSeat: assignedInfo ? assignedInfo.seatNo : '1A',
      message: 'Pendaftaran berhasil! PIN Akses Anda adalah: ' + generatedPin
    };
  } catch (err) {
    return { success: false, message: 'Terjadi kesalahan saat pendaftaran: ' + err.message };
  }
}

function addFamilyMember(payload) {
  try {
    var ss = getSpreadsheet_();
    var sheetPeserta = ss.getSheetByName('Peserta');
    var sheetBus = ss.getSheetByName('Kursi_Bus');

    var sponsorId = String(payload.sponsorId || '').trim();
    var sponsorNama = String(payload.sponsorNama || '').trim();
    var inputPin = String(payload.pinAkses || '').trim();
    var namaAnggota = String(payload.namaAnggota || '').trim();
    var hubungan = String(payload.hubungan || 'Istri').trim();

    if (!sponsorId || !namaAnggota) {
      return { success: false, message: 'ID Sponsor dan Nama Anggota wajib diisi.' };
    }

    if (!inputPin) {
      return { success: false, message: 'PIN Akses 4-digit wajib diisi untuk keamanan.' };
    }

    var lastRow = sheetPeserta.getLastRow();
    var rawData = sheetPeserta.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var sponsorData = null;

    for (var i = 0; i < rawData.length; i++) {
      if (rawData[i][0] === sponsorId) {
        sponsorData = rawData[i];
        break;
      }
    }

    if (!sponsorData) {
      return { success: false, message: 'Data Kepala Keluarga tidak ditemukan di database.' };
    }

    var correctPin = String(sponsorData[4] || '').trim();
    if (inputPin !== correctPin) {
      return { success: false, message: 'PIN Akses Warga salah! Silakan minta PIN 4-digit kepada Kepala Keluarga.' };
    }

    var noWhatsapp = sponsorData[3];
    var noRumah = sponsorData[6];
    var catatanLokasi = sponsorData[7];
    var statusPembayaran = sponsorData[9];
    var alamatFix = 'Kamp baru I Jl. Marga Mulya';

    var idTransaksi = generateSequentialId_();
    var nowFormatted = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');

    var newMemberRow = [
      idTransaksi,
      nowFormatted,
      namaAnggota,
      noWhatsapp,
      correctPin,
      alamatFix,
      noRumah,
      catatanLokasi,
      'Tunai',
      statusPembayaran,
      sponsorId,
      sponsorNama,
      hubungan
    ];

    sheetPeserta.appendRow(newMemberRow);

    var assignedInfo = autoAssignFirstAvailableSeatMultiBus_(sheetBus, idTransaksi, namaAnggota);

    SpreadsheetApp.flush();

    return {
      success: true,
      idTransaksi: idTransaksi,
      namaAnggota: namaAnggota,
      assignedBus: assignedInfo ? assignedInfo.busId : 'Bus 1',
      assignedSeat: assignedInfo ? assignedInfo.seatNo : '1A',
      message: 'Anggota keluarga (' + hubungan + ') berhasil ditambahkan.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal menambah anggota keluarga: ' + err.message };
  }
}

function autoAssignFirstAvailableSeatMultiBus_(sheetBus, idPeserta, namaPeserta) {
  var lastRow = sheetBus.getLastRow();
  if (lastRow <= 1) return null;

  var busData = sheetBus.getRange(2, 1, lastRow - 1, 6).getValues();
  for (var i = 0; i < busData.length; i++) {
    var namaTerisi = busData[i][5];
    if (!namaTerisi || String(namaTerisi).trim() === '') {
      var rowTarget = i + 2;
      sheetBus.getRange(rowTarget, 5, 1, 2).setValues([[idPeserta, namaPeserta]]);
      return {
        busId: busData[i][0],
        seatNo: busData[i][1]
      };
    }
  }
  return null;
}

function adminLogin(pin) {
  try {
    var inputPin = String(pin || '').trim();
    var ss = getSpreadsheet_();
    var sheetConfig = ss.getSheetByName('Konfigurasi');
    
    var adminPin = '1234';
    var superAdminPin = '9999';

    if (sheetConfig && sheetConfig.getLastRow() > 1) {
      var configs = sheetConfig.getRange(2, 1, sheetConfig.getLastRow() - 1, 2).getDisplayValues();
      for (var i = 0; i < configs.length; i++) {
        if (configs[i][0] === 'ADMIN_PIN') adminPin = configs[i][1];
        if (configs[i][0] === 'SUPERADMIN_PIN') superAdminPin = configs[i][1];
      }
    }

    if (inputPin === superAdminPin) {
      return { success: true, role: 'Super Admin', token: 'SUPER_' + new Date().getTime() };
    } else if (inputPin === adminPin) {
      return { success: true, role: 'Admin', token: 'ADMIN_' + new Date().getTime() };
    } else {
      return { success: false, message: 'PIN Admin tidak sesuai. Silakan coba kembali.' };
    }
  } catch (err) {
    return { success: false, message: 'Error otentikasi: ' + err.message };
  }
}

/**
 * Update Status Pembayaran
 * Aturan Role-Based Access Control:
 * - Admin biasa: Hanya dapat mengubah dari Menunggu ke LUNAS. Tidak bisa mengubah dari Lunas ke Menunggu.
 * - Super Admin: Dapat mengubah status bolak-balik (Lunas <-> Menunggu) tanpa batasan.
 */
function updatePaymentStatus(idTransaksi, newStatus, currentRole) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('Peserta');
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Data tidak ditemukan.' };

    var role = String(currentRole || '').trim();
    var targetStatus = String(newStatus || '').trim();

    // Validasi Akses Peran Admin vs Super Admin
    if (targetStatus !== 'Lunas' && role !== 'Super Admin') {
      return {
        success: false,
        message: 'Akses ditolak: Hanya Super Admin yang berhak mengembalikan status ke Menunggu Konfirmasi.'
      };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var updatedCount = 0;

    for (var i = 0; i < data.length; i++) {
      var currentId = data[i][0];
      var sponsorId = data[i][10];

      if (currentId === idTransaksi || sponsorId === idTransaksi) {
        sheet.getRange(i + 2, 10).setValue(targetStatus);
        updatedCount++;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, updatedCount: updatedCount, message: 'Status pembayaran berhasil diperbarui menjadi ' + targetStatus + '.' };
  } catch (err) {
    return { success: false, message: 'Gagal memperbarui status: ' + err.message };
  }
}

function deleteParticipant(idTransaksi) {
  try {
    var ss = getSpreadsheet_();
    var sheetPeserta = ss.getSheetByName('Peserta');
    var sheetBus = ss.getSheetByName('Kursi_Bus');

    var lastRow = sheetPeserta.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Data kosong.' };

    var data = sheetPeserta.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var rowsToDelete = [];
    var affectedIds = [idTransaksi];

    for (var i = 0; i < data.length; i++) {
      var curId = data[i][0];
      var sponsorId = data[i][10];
      if (curId === idTransaksi || sponsorId === idTransaksi) {
        rowsToDelete.push(i + 2);
        if (curId !== idTransaksi) affectedIds.push(curId);
      }
    }

    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var r = 0; r < rowsToDelete.length; r++) {
      sheetPeserta.deleteRow(rowsToDelete[r]);
    }

    var busLastRow = sheetBus.getLastRow();
    if (busLastRow > 1) {
      var busRange = sheetBus.getRange(2, 1, busLastRow - 1, 6);
      var busData = busRange.getValues();
      var changed = false;

      for (var b = 0; b < busData.length; b++) {
        var seatedId = String(busData[b][4] || '').trim();
        if (affectedIds.indexOf(seatedId) !== -1) {
          busData[b][4] = '';
          busData[b][5] = '';
          changed = true;
        }
      }

      if (changed) {
        busRange.setValues(busData);
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Data peserta dan alokasi kursi berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus data: ' + err.message };
  }
}

function swapBusSeats(busIdA, seatNoA, busIdB, seatNoB) {
  try {
    var ss = getSpreadsheet_();
    var sheetBus = ss.getSheetByName('Kursi_Bus');
    var lastRow = sheetBus.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Data kursi tidak tersedia.' };

    var range = sheetBus.getRange(2, 1, lastRow - 1, 6);
    var data = range.getValues();

    var indexA = -1;
    var indexB = -1;

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(busIdA) && String(data[i][1]) === String(seatNoA)) indexA = i;
      if (String(data[i][0]) === String(busIdB) && String(data[i][1]) === String(seatNoB)) indexB = i;
    }

    if (indexA === -1 || indexB === -1) {
      return { success: false, message: 'Salah satu posisi kursi tidak ditemukan.' };
    }

    var tempId = data[indexA][4];
    var tempName = data[indexA][5];

    data[indexA][4] = data[indexB][4];
    data[indexA][5] = data[indexB][5];

    data[indexB][4] = tempId;
    data[indexB][5] = tempName;

    range.setValues(data);
    SpreadsheetApp.flush();

    return { 
      success: true, 
      message: 'Kursi (' + busIdA + ' - ' + seatNoA + ') dan (' + busIdB + ' - ' + seatNoB + ') berhasil ditukar!' 
    };
  } catch (err) {
    return { success: false, message: 'Gagal menukar kursi: ' + err.message };
  }
}

function getExportData() {
  try {
    var ss = getSpreadsheet_();
    var sheetPeserta = ss.getSheetByName('Peserta');
    var sheetBus = ss.getSheetByName('Kursi_Bus');

    if (!sheetPeserta || !sheetBus) {
      return { success: false, message: 'Data tidak ditemukan.' };
    }

    var pLastRow = sheetPeserta.getLastRow();
    var bLastRow = sheetBus.getLastRow();

    var seatMapping = {};
    if (bLastRow > 1) {
      var busValues = sheetBus.getRange(2, 1, bLastRow - 1, 6).getDisplayValues();
      for (var b = 0; b < busValues.length; b++) {
        var pId = busValues[b][4];
        if (pId && pId.trim() !== '') {
          seatMapping[pId] = {
            idBus: busValues[b][0],
            noKursi: busValues[b][1]
          };
        }
      }
    }

    var exportRows = [];
    if (pLastRow > 1) {
      var pValues = sheetPeserta.getRange(2, 1, pLastRow - 1, 13).getDisplayValues();
      for (var p = 0; p < pValues.length; p++) {
        var row = pValues[p];
        var id = row[0];
        var seatInfo = seatMapping[id] || { idBus: 'Belum Ada', noKursi: '-' };

        exportRows.push({
          idTransaksi: row[0],
          tanggal: row[1],
          namaPeserta: row[2],
          noWhatsapp: row[3],
          pinAkses: row[4],
          alamat: row[5],
          noRumah: row[6],
          catatanLokasi: row[7],
          metodePembayaran: row[8],
          statusPembayaran: row[9],
          sponsorId: row[10] || '-',
          sponsorNama: row[11] || '-',
          hubungan: row[12] || 'Kepala Keluarga',
          idBus: seatInfo.idBus,
          noKursi: seatInfo.noKursi
        });
      }
    }

    return {
      success: true,
      exportRows: exportRows,
      timestamp: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (err) {
    return { success: false, message: 'Gagal menyiapkan data ekspor: ' + err.message };
  }
}
