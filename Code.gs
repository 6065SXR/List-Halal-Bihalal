/**
 * ============================================================================
 * ZETTBOS BACKEND SERVICE & BUSINESS LOGIC (MODULAR ARCHITECTURE)
 * File: Code.gs
 * Deskripsi: API CRUD, Multi-Bus Allocation, A-Z Sorting, System Tabungan,
 *            WA Kwitansi Digital, Rundown Info & Backup/Restore Database 1-Klik
 * ============================================================================
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Registrasi Halal Bihalal & Seat Bus RT')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

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

function generateSequentialLogId_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Log_Tabungan');
  if (!sheet || sheet.getLastRow() <= 1) {
    return 'LOG-0001';
  }

  var lastRow = sheet.getLastRow();
  var idRange = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  
  var maxNum = 0;
  for (var i = 0; i < idRange.length; i++) {
    var str = idRange[i][0] || '';
    if (str.indexOf('LOG-') === 0) {
      var numPart = parseInt(str.replace('LOG-', ''), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  }
  
  var nextNum = maxNum + 1;
  var padded = ('0000' + nextNum).slice(-4);
  return 'LOG-' + padded;
}

function formatNoRumah_(val) {
  var str = String(val || '').trim();
  if (!str) return 'No. -';
  var clean = str.replace(/^no\.?\s*/i, '').trim();
  return 'No. ' + clean;
}

function getAppData(searchQuery, statusFilter) {
  try {
    var ss = getSpreadsheet_();
    var sheetPeserta = ss.getSheetByName('Peserta');
    var sheetBus = ss.getSheetByName('Kursi_Bus');

    if (!sheetPeserta || !sheetBus) {
      return { success: false, message: 'Database belum diinisialisasi.' };
    }

    var pesertaLastRow = sheetPeserta.getLastRow();
    var pesertaList = [];
    var totalKepalaKeluarga = 0;
    var totalWarga = 0;
    var totalTabunganSeluruhnya = 0;

    if (pesertaLastRow > 1) {
      var rawData = sheetPeserta.getRange(2, 1, pesertaLastRow - 1, 13).getDisplayValues();
      var query = (searchQuery || '').toLowerCase().trim();
      var filter = (statusFilter || 'SEMUA').toUpperCase().trim();

      var allRows = [];

      for (var i = 0; i < rawData.length; i++) {
        var row = rawData[i];
        var tabVal = parseInt(String(row[9]).replace(/\D/g, ''), 10);
        if (isNaN(tabVal)) tabVal = 0;

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
          jumlahTabungan: tabVal,
          sponsorId: row[10],
          sponsorNama: row[11],
          hubungan: row[12],
          familyMembers: []
        };

        totalWarga++;
        totalTabunganSeluruhnya += tabVal;

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

      // Urutkan Kepala Keluarga Berdasarkan Abjad Nama (A ke Z)
      primaries.sort(function(a, b) {
        return (a.nama || '').localeCompare(b.nama || '', 'id', { sensitivity: 'base' });
      });

      for (var m = 0; m < primaries.length; m++) {
        var p = primaries[m];
        
        var totalFamilyTab = p.jumlahTabungan || 0;
        for (var fm = 0; fm < p.familyMembers.length; fm++) {
          totalFamilyTab += (p.familyMembers[fm].jumlahTabungan || 0);
        }

        var matchStatus = true;
        if (filter === 'ADA TABUNGAN') matchStatus = (totalFamilyTab > 0);
        if (filter === 'BELUM TABUNGAN') matchStatus = (totalFamilyTab === 0);

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

    var infoData = getInfoAcaraData_();

    return {
      success: true,
      data: {
        peserta: pesertaList,
        stats: {
          totalWarga: totalWarga,
          totalKepalaKeluarga: totalKepalaKeluarga,
          totalTabungan: totalTabunganSeluruhnya,
          totalKursiBus1: 33,
          kursiTerisiBus1: occupiedBus1,
          kursiTersisaBus1: Math.max(0, 33 - occupiedBus1),
          totalKursiBus2: 33,
          kursiTerisiBus2: occupiedBus2,
          kursiTersisaBus2: Math.max(0, 33 - occupiedBus2),
          totalKursiAll: 66,
          kursiTerisiAll: occupiedBus1 + occupiedBus2
        },
        seats: seatsList,
        info: infoData
      }
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat data: ' + err.message };
  }
}

function getInfoAcaraData_() {
  var ss = getSpreadsheet_();
  var sheetRundown = ss.getSheetByName('Rundown_Acara');
  var sheetKontak = ss.getSheetByName('Kontak_Panitia');

  var rundownList = [];
  var kontakList = [];

  var defaultRundown = [
    { id: 'RD-001', hari: 'Hari 1', jam: '06:00 - 06:30', judul: 'Kumpul & Absensi Peserta Bus', keterangan: 'Titik Kumpul: Kp. Baru I Jl. Marga Mulya (Depan Pos Satpam RT 03)' },
    { id: 'RD-002', hari: 'Hari 1', jam: '06:45 WIB', judul: 'Keberangkatan Bus Medium 1 & Bus 2', keterangan: 'Seluruh armada berangkat bersama. Mohon hadir tepat waktu!' },
    { id: 'RD-003', hari: 'Hari 1', jam: '09:30 - 11:30', judul: 'Pembukaan, Tausiyah & Halal Bihalal', keterangan: 'Sambutan Ketua RT, Mushafahah/Salaman Warga, dan Tausiyah Agama' },
    { id: 'RD-004', hari: 'Hari 1', jam: '11:30 - 13:30', judul: 'Makan Siang & Ramah Tamah', keterangan: 'Santap Catering bersama & Pembagian Doorprize Digital' },
    { id: 'RD-005', hari: 'Hari 1', jam: '15:30 WIB', judul: 'Persiapan Pulang ke Kamp Baru I', keterangan: 'Pengecekan ulang seluruh anggota keluarga di Bus 1 & Bus 2' },
    { id: 'RD-006', hari: 'Hari 2', jam: '07:30 - 09:00', judul: 'Senam Pagi Warga & Sarapan Bersama', keterangan: 'Area Lapangan Villa / Resort Halal Bihalal' },
    { id: 'RD-007', hari: 'Hari 2', jam: '09:30 - 12:00', judul: 'Fun Games Warga & Pembagian Hadiah', keterangan: 'Lomba keakraban antar RT & Door Prize Utama' },
    { id: 'RD-008', hari: 'Hari 2', jam: '13:00 WIB', judul: 'Check Out & Perjalanan Pulang', keterangan: 'Seluruh armada bus kembali ke Jakarta' }
  ];

  var defaultKontak = [
    { id: 'KT-001', jabatan: 'Koordinator Bus 1', nama: 'Pak Agus', noWhatsapp: '081234567890' },
    { id: 'KT-002', jabatan: 'Koordinator Bus 2', nama: 'Tommy H', noWhatsapp: '081234567890' },
    { id: 'KT-003', jabatan: 'Ketua RT 03', nama: 'Pak Sugeng', noWhatsapp: '081398765432' }
  ];

  if (!sheetRundown || sheetRundown.getLastRow() <= 1) {
    rundownList = defaultRundown;
  } else {
    var rawRD = sheetRundown.getRange(2, 1, sheetRundown.getLastRow() - 1, 5).getDisplayValues();
    for (var r = 0; r < rawRD.length; r++) {
      rundownList.push({
        id: rawRD[r][0],
        hari: rawRD[r][1] || 'Hari 1',
        jam: rawRD[r][2],
        judul: rawRD[r][3],
        keterangan: rawRD[r][4]
      });
    }
  }

  if (!sheetKontak || sheetKontak.getLastRow() <= 1) {
    kontakList = defaultKontak;
  } else {
    var rawKT = sheetKontak.getRange(2, 1, sheetKontak.getLastRow() - 1, 4).getDisplayValues();
    for (var k = 0; k < rawKT.length; k++) {
      kontakList.push({
        id: rawKT[k][0],
        jabatan: rawKT[k][1],
        nama: rawKT[k][2],
        noWhatsapp: rawKT[k][3]
      });
    }
  }

  return { rundown: rundownList, kontak: kontakList };
}

function saveRundownItem(item, adminRole) {
  try {
    if (adminRole !== 'Super Admin' && adminRole !== 'Admin') {
      return { success: false, message: 'Akses ditolak. Membutuhkan hak akses Admin / Super Admin.' };
    }

    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('Rundown_Acara');
    if (!sheet) {
      sheet = ss.insertSheet('Rundown_Acara');
      sheet.appendRow(['ID', 'Hari', 'Jam', 'Judul', 'Keterangan']);
      formatHeaderRow_(sheet, 5, '#5A56EC');
    }

    var id = item.id || ('RD-' + ('000' + (sheet.getLastRow())).slice(-3));
    var isEdit = false;

    if (sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getDisplayValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === id) {
          sheet.getRange(i + 2, 1, 1, 5).setValues([[id, item.hari, item.jam, item.judul, item.keterangan]]);
          isEdit = true;
          break;
        }
      }
    }

    if (!isEdit) {
      sheet.appendRow([id, item.hari, item.jam, item.judul, item.keterangan]);
    }

    SpreadsheetApp.flush();
    return { success: true, message: isEdit ? 'Agenda berhasil diperbarui!' : 'Agenda baru berhasil ditambahkan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan agenda: ' + err.message };
  }
}

function deleteRundownItem(id, adminRole) {
  try {
    if (adminRole !== 'Super Admin' && adminRole !== 'Admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('Rundown_Acara');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Data tidak ditemukan.' };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { success: true, message: 'Agenda berhasil dihapus!' };
      }
    }
    return { success: false, message: 'ID Agenda tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus agenda: ' + err.message };
  }
}

function savePanitiaContact(contact, adminRole) {
  try {
    if (adminRole !== 'Super Admin' && adminRole !== 'Admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('Kontak_Panitia');
    if (!sheet) {
      sheet = ss.insertSheet('Kontak_Panitia');
      sheet.appendRow(['ID', 'Jabatan', 'Nama', 'No. WhatsApp']);
      formatHeaderRow_(sheet, 4, '#265768');
    }

    var id = contact.id || ('KT-' + ('000' + (sheet.getLastRow())).slice(-3));
    var isEdit = false;

    if (sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getDisplayValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === id) {
          sheet.getRange(i + 2, 1, 1, 4).setValues([[id, contact.jabatan, contact.nama, contact.noWhatsapp]]);
          isEdit = true;
          break;
        }
      }
    }

    if (!isEdit) {
      sheet.appendRow([id, contact.jabatan, contact.nama, contact.noWhatsapp]);
    }

    SpreadsheetApp.flush();
    return { success: true, message: isEdit ? 'Kontak panitia diperbarui!' : 'Kontak panitia ditambahkan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan kontak: ' + err.message };
  }
}

function deletePanitiaContact(id, adminRole) {
  try {
    if (adminRole !== 'Super Admin' && adminRole !== 'Admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName('Kontak_Panitia');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Data tidak ditemukan.' };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { success: true, message: 'Kontak panitia berhasil dihapus!' };
      }
    }
    return { success: false, message: 'ID Kontak tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus kontak: ' + err.message };
  }
}

function registerParticipant(payload) {
  try {
    var ss = getSpreadsheet_();
    var sheetPeserta = ss.getSheetByName('Peserta');
    var sheetBus = ss.getSheetByName('Kursi_Bus');
    var sheetLog = ss.getSheetByName('Log_Tabungan');

    if (!sheetPeserta || !sheetBus) {
      return { success: false, message: 'Sheet database tidak ditemukan.' };
    }

    var inputWaRaw = String(payload.noWhatsapp || '').trim();
    var inputWaClean = inputWaRaw.replace(/\D/g, '');
    var formattedNoRumah = formatNoRumah_(payload.noRumah);
    var inputRumahClean = formattedNoRumah.toLowerCase().replace(/\s+/g, '');

    var lastRow = sheetPeserta.getLastRow();
    if (lastRow > 1) {
      var existingData = sheetPeserta.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
      for (var i = 0; i < existingData.length; i++) {
        var existWaClean = String(existingData[i][3] || '').replace(/\D/g, '');
        var existRumahClean = String(existingData[i][6] || '').toLowerCase().replace(/\s+/g, '');

        if (inputWaClean && existWaClean && inputWaClean === existWaClean) {
          return { success: false, message: 'nomor telepon anda sudah terdaftar silahkan hubungi Pak Agus atau Tommy' };
        }
        if (inputRumahClean && existRumahClean && inputRumahClean === existRumahClean) {
          return { success: false, message: 'nomor rumah anda sudah terdaftar silahkan hubungi Pak Agus atau Tommy' };
        }
      }
    }

    var idTransaksi = generateSequentialId_();
    var nowFormatted = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var alamatFix = 'Kamp baru I Jl. Marga Mulya';
    
    // Proteksi Keuangan: Hanya izinkan setoran awal jika pendaftaran dilakukan oleh Admin / Super Admin
    var setoranAwal = 0;
    if (payload.adminRole === 'Admin' || payload.adminRole === 'Super Admin') {
      setoranAwal = parseInt(payload.setoranAwal || 0, 10);
      if (isNaN(setoranAwal)) setoranAwal = 0;
    }

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
      setoranAwal,
      '',
      '',
      ''
    ];

    sheetPeserta.appendRow(newRow);

    if (setoranAwal > 0 && sheetLog) {
      var logId = generateSequentialLogId_();
      sheetLog.appendRow([
        logId,
        nowFormatted,
        idTransaksi,
        newRow[2],
        setoranAwal,
        setoranAwal,
        payload.adminRole || 'Admin'
      ]);
    }

    var assignedInfo = autoAssignFirstAvailableSeatMultiBus_(sheetBus, idTransaksi, newRow[2]);
    SpreadsheetApp.flush();

    return {
      success: true,
      idTransaksi: idTransaksi,
      nama: newRow[2],
      pinAkses: generatedPin,
      noWhatsapp: newRow[3],
      assignedBus: assignedInfo ? assignedInfo.busId : 'Bus 1',
      assignedSeat: assignedInfo ? assignedInfo.seatNo : '01',
      setoranAwal: setoranAwal,
      message: 'Pendaftaran berhasil! PIN Akses Anda: ' + generatedPin
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
    var sheetLog = ss.getSheetByName('Log_Tabungan');

    var sponsorId = String(payload.sponsorId || '').trim();
    var sponsorNama = String(payload.sponsorNama || '').trim();
    var inputPin = String(payload.pinAkses || '').trim();
    var namaAnggota = String(payload.namaAnggota || '').trim();
    var hubungan = String(payload.hubungan || 'Istri').trim();

    // Proteksi Keuangan: Hanya izinkan setoran awal jika ditambahkan oleh Admin / Super Admin
    var setoranAwal = 0;
    if (payload.adminRole === 'Admin' || payload.adminRole === 'Super Admin') {
      setoranAwal = parseInt(payload.setoranAwal || 0, 10);
      if (isNaN(setoranAwal)) setoranAwal = 0;
    }

    if (!sponsorId || !namaAnggota) {
      return { success: false, message: 'ID Sponsor dan Nama Anggota wajib diisi.' };
    }
    if (!inputPin) {
      return { success: false, message: 'PIN Akses 4-digit wajib diisi.' };
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
      return { success: false, message: 'Data Kepala Keluarga tidak ditemukan.' };
    }

    var correctPin = String(sponsorData[4] || '').trim();
    if (inputPin !== correctPin) {
      return { success: false, message: 'PIN Akses Warga salah!' };
    }

    var idTransaksi = generateSequentialId_();
    var nowFormatted = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');

    var newMemberRow = [
      idTransaksi,
      nowFormatted,
      namaAnggota,
      sponsorData[3],
      correctPin,
      'Kamp baru I Jl. Marga Mulya',
      sponsorData[6],
      sponsorData[7],
      'Tunai',
      setoranAwal,
      sponsorId,
      sponsorNama,
      hubungan
    ];

    sheetPeserta.appendRow(newMemberRow);

    if (setoranAwal > 0 && sheetLog) {
      var logId = generateSequentialLogId_();
      var currentHeadTab = parseInt(String(sponsorData[9]).replace(/\D/g, ''), 10) || 0;
      var newTotal = currentHeadTab + setoranAwal;

      sheetLog.appendRow([
        logId,
        nowFormatted,
        idTransaksi,
        namaAnggota,
        setoranAwal,
        newTotal,
        payload.adminRole || 'Admin'
      ]);
    }

    var assignedInfo = autoAssignFirstAvailableSeatMultiBus_(sheetBus, idTransaksi, namaAnggota);
    SpreadsheetApp.flush();

    return {
      success: true,
      idTransaksi: idTransaksi,
      namaAnggota: namaAnggota,
      assignedBus: assignedInfo ? assignedInfo.busId : 'Bus 1',
      assignedSeat: assignedInfo ? assignedInfo.seatNo : '01',
      message: 'Anggota keluarga (' + hubungan + ') berhasil ditambahkan.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal menambah anggota keluarga: ' + err.message };
  }
}

function addTabunganDeposit(idTransaksi, nominalSetoran, adminRole) {
  try {
    var ss = getSpreadsheet_();
    var sheetPeserta = ss.getSheetByName('Peserta');
    var sheetLog = ss.getSheetByName('Log_Tabungan');

    var nominal = parseInt(nominalSetoran, 10);
    if (isNaN(nominal) || nominal <= 0) {
      return { success: false, message: 'Nominal setoran harus berupa angka positif.' };
    }

    var lastRow = sheetPeserta.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Data peserta tidak ditemukan.' };

    var data = sheetPeserta.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var targetRowIndex = -1;
    var currentTabungan = 0;
    var namaPeserta = '';
    var noWa = '';

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === idTransaksi) {
        targetRowIndex = i + 2;
        currentTabungan = parseInt(String(data[i][9]).replace(/\D/g, ''), 10) || 0;
        namaPeserta = data[i][2];
        noWa = data[i][3];
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { success: false, message: 'ID Peserta tidak ditemukan.' };
    }

    var newTotalTabungan = currentTabungan + nominal;
    sheetPeserta.getRange(targetRowIndex, 10).setValue(newTotalTabungan);

    var nowFormatted = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var logId = generateSequentialLogId_();
    
    if (sheetLog) {
      sheetLog.appendRow([
        logId,
        nowFormatted,
        idTransaksi,
        namaPeserta,
        nominal,
        newTotalTabungan,
        adminRole || 'Admin'
      ]);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      idTransaksi: idTransaksi,
      namaPeserta: namaPeserta,
      noWhatsapp: noWa,
      nominalSetor: nominal,
      totalTabungan: newTotalTabungan,
      logId: logId,
      tanggal: nowFormatted,
      message: 'Setoran Rp ' + nominal.toLocaleString('id-ID') + ' berhasil ditambahkan!'
    };
  } catch (err) {
    return { success: false, message: 'Gagal menambah tabungan: ' + err.message };
  }
}

function getTabunganLog(idTransaksi) {
  try {
    var ss = getSpreadsheet_();
    var sheetLog = ss.getSheetByName('Log_Tabungan');
    var sheetPeserta = ss.getSheetByName('Peserta');
    if (!sheetLog || sheetLog.getLastRow() <= 1) {
      return { success: true, logs: [] };
    }

    var familyIds = [idTransaksi];
    if (sheetPeserta && sheetPeserta.getLastRow() > 1) {
      var pData = sheetPeserta.getRange(2, 1, sheetPeserta.getLastRow() - 1, 13).getDisplayValues();
      for (var p = 0; p < pData.length; p++) {
        if (pData[p][10] === idTransaksi) {
          familyIds.push(pData[p][0]);
        }
      }
    }

    var lastRow = sheetLog.getLastRow();
    var logsData = sheetLog.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
    var userLogs = [];

    for (var i = 0; i < logsData.length; i++) {
      if (familyIds.indexOf(logsData[i][2]) !== -1) {
        userLogs.push({
          idLog: logsData[i][0],
          tanggal: logsData[i][1],
          idTransaksi: logsData[i][2],
          namaPeserta: logsData[i][3],
          nominalSetoran: parseInt(String(logsData[i][4]).replace(/\D/g, ''), 10) || 0,
          totalTabungan: parseInt(String(logsData[i][5]).replace(/\D/g, ''), 10) || 0,
          adminPenyetor: logsData[i][6]
        });
      }
    }

    userLogs.reverse();
    return { success: true, logs: userLogs };
  } catch (err) {
    return { success: false, message: 'Gagal mengambil log tabungan: ' + err.message };
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
      return { success: false, message: 'PIN Admin salah.' };
    }
  } catch (err) {
    return { success: false, message: 'Error otentikasi: ' + err.message };
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

    var primaries = [];
    var familyMembersMap = {};

    if (pLastRow > 1) {
      var pValues = sheetPeserta.getRange(2, 1, pLastRow - 1, 13).getDisplayValues();
      for (var p = 0; p < pValues.length; p++) {
        var row = pValues[p];
        var id = row[0];
        var sponsorId = row[10];
        var seatInfo = seatMapping[id] || { idBus: 'Belum Ada', noKursi: '-' };

        var item = {
          idTransaksi: row[0],
          tanggal: row[1],
          namaPeserta: row[2],
          noWhatsapp: row[3],
          pinAkses: row[4],
          alamat: row[5],
          noRumah: row[6],
          catatanLokasi: row[7],
          metodePembayaran: row[8],
          jumlahTabungan: parseInt(String(row[9]).replace(/\D/g, ''), 10) || 0,
          sponsorId: row[10] || '',
          sponsorNama: row[11] || '',
          hubungan: row[12] || 'Kepala Keluarga',
          idBus: seatInfo.idBus,
          noKursi: seatInfo.noKursi
        };

        if (!sponsorId) {
          primaries.push(item);
          familyMembersMap[id] = [];
        } else {
          if (!familyMembersMap[sponsorId]) {
            familyMembersMap[sponsorId] = [];
          }
          familyMembersMap[sponsorId].push(item);
        }
      }
    }

    primaries.sort(function(a, b) {
      return (a.namaPeserta || '').localeCompare(b.namaPeserta || '', 'id', { sensitivity: 'base' });
    });

    var exportRows = [];

    for (var i = 0; i < primaries.length; i++) {
      var head = primaries[i];
      var members = familyMembersMap[head.idTransaksi] || [];
      var totalFamilyCount = 1 + members.length;

      exportRows.push({
        idTransaksi: head.idTransaksi,
        tanggal: head.tanggal,
        keluarga: head.namaPeserta,
        noWhatsapp: head.noWhatsapp,
        namaPeserta: head.namaPeserta,
        statusKeluarga: 'Kepala Keluarga',
        pinAkses: head.pinAkses,
        alamat: head.alamat,
        noRumah: head.noRumah,
        catatanLokasi: head.catatanLokasi,
        metodePembayaran: head.metodePembayaran,
        jumlahTabungan: head.jumlahTabungan,
        sponsorId: '-',
        sponsorNama: '-',
        idBus: head.idBus,
        noKursi: head.noKursi,
        familyCount: totalFamilyCount,
        isHead: true
      });

      for (var m = 0; m < members.length; m++) {
        var sub = members[m];
        exportRows.push({
          idTransaksi: sub.idTransaksi,
          tanggal: sub.tanggal,
          keluarga: head.namaPeserta,
          noWhatsapp: head.noWhatsapp,
          namaPeserta: sub.namaPeserta,
          statusKeluarga: sub.hubungan || 'Anggota Keluarga',
          pinAkses: sub.pinAkses,
          alamat: sub.alamat,
          noRumah: sub.noRumah,
          catatanLokasi: sub.catatanLokasi,
          metodePembayaran: sub.metodePembayaran,
          jumlahTabungan: sub.jumlahTabungan,
          sponsorId: head.idTransaksi,
          sponsorNama: head.namaPeserta,
          idBus: sub.idBus,
          noKursi: sub.noKursi,
          familyCount: totalFamilyCount,
          isHead: false
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

function backupDatabase() {
  try {
    var ss = getSpreadsheet_();
    var sheetsToBackup = ['Peserta', 'Kursi_Bus', 'Log_Tabungan', 'Konfigurasi', 'Rundown_Acara', 'Kontak_Panitia'];
    var backupObj = {
      appName: 'Web App Halal Bihalal RT',
      backupTime: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss'),
      data: {}
    };

    for (var i = 0; i < sheetsToBackup.length; i++) {
      var sheetName = sheetsToBackup[i];
      var sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() >= 1) {
        backupObj.data[sheetName] = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
      } else {
        backupObj.data[sheetName] = [];
      }
    }

    return {
      success: true,
      backupJson: JSON.stringify(backupObj),
      timestamp: backupObj.backupTime,
      message: 'Backup database berhasil dibuat!'
    };
  } catch (err) {
    return { success: false, message: 'Gagal membuat backup database: ' + err.message };
  }
}

function clearAllDatabase(adminRole) {
  try {
    if (adminRole !== 'Super Admin') {
      return { success: false, message: 'Akses ditolak. Fitur reset database ini khusus Super Admin!' };
    }

    var ss = getSpreadsheet_();

    // 1. Kosongkan Data Peserta (mulai baris 2)
    var sheetPeserta = ss.getSheetByName('Peserta');
    if (sheetPeserta && sheetPeserta.getLastRow() > 1) {
      sheetPeserta.deleteRows(2, sheetPeserta.getLastRow() - 1);
    }

    // 2. Kosongkan Log Tabungan (mulai baris 2)
    var sheetLog = ss.getSheetByName('Log_Tabungan');
    if (sheetLog && sheetLog.getLastRow() > 1) {
      sheetLog.deleteRows(2, sheetLog.getLastRow() - 1);
    }

    // 3. Reset Alokasi Kursi Bus (kosongkan ID Peserta & Nama Terisi)
    var sheetBus = ss.getSheetByName('Kursi_Bus');
    if (sheetBus && sheetBus.getLastRow() > 1) {
      var numRows = sheetBus.getLastRow() - 1;
      var emptyValues = [];
      for (var b = 0; b < numRows; b++) {
        emptyValues.push(['', '']);
      }
      sheetBus.getRange(2, 5, numRows, 2).setValues(emptyValues);
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: '💥 Seluruh database peserta, log tabungan, dan alokasi kursi berhasil dikosongkan! Aplikasi siap untuk Launching.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengosongkan database: ' + err.message };
  }
}

function restoreDatabase(backupJsonString) {
  try {
    if (!backupJsonString) {
      return { success: false, message: 'Data backup tidak boleh kosong.' };
    }

    var backupObj = JSON.parse(backupJsonString);
    if (!backupObj || !backupObj.data) {
      return { success: false, message: 'Format data backup JSON tidak valid.' };
    }

    var ss = getSpreadsheet_();
    var sheetNames = ['Peserta', 'Kursi_Bus', 'Log_Tabungan', 'Konfigurasi', 'Rundown_Acara', 'Kontak_Panitia'];

    for (var s = 0; s < sheetNames.length; s++) {
      var name = sheetNames[s];
      var rawMatrix = backupObj.data[name];
      if (!rawMatrix || rawMatrix.length === 0) continue;

      var sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
      }

      sheet.clearContents();
      sheet.getRange(1, 1, rawMatrix.length, rawMatrix[0].length).setValues(rawMatrix);
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      timestamp: backupObj.backupTime || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss'),
      message: 'Restore database berhasil dipulihkan!'
    };
  } catch (err) {
    return { success: false, message: 'Gagal memulihkan database: ' + err.message };
  }
}
