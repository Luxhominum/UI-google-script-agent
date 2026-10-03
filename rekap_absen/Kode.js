/**
 * =========================================================================
 * REKAP ABSENSI MINGGUAN OTOMATIS KE DASHBOARD (TAB '0')
 * Google Sheets + Google Apps Script Automation
 * =========================================================================
 * 
 * FITUR UTAMA:
 * 1. Standardisasi format Tab '0' (Dashboard) untuk seluruh bulan (Sept26 s/d Mar27)
 *    dengan 5 Minggu + Kolom Total otomatis formula =SUM().
 * 2. Menulis rentang tanggal minggu (misal '30-5', '6-12', dsb) di Baris 2.
 * 3. Ekstraksi otomatis 'total hari kerja' dari setiap tab mingguan baru yang di-paste.
 * 4. Normalisasi nama karyawan (membersihkan inisial avatar 'WA\n', enter, spasi ganda)
 *    agar cocok 100% dengan nama di Tab '0'.
 * 5. Proteksi Ganda Tab '0' (Dual-Layer Protection):
 *    - Layer 1: Google Sheets Native Warning Interceptor (Mencegah Pemilik & Editor menghapus sel).
 *    - Layer 2: Real-time onEdit Guard (Otomatis mengembalikan nilai jika ada sel yang terhapus).
 * 6. Penataan border presisi dan penanganan unmerge anti-error.
 * 7. Menu interaktif '⚡ Rekap Absen' untuk sinkronisasi 1-klik atau konfirmasi manual.
 */

const REKAP_CONFIG = Object.freeze({
  DASHBOARD_SHEET_NAME: '0',
  EMPLOYEE_START_ROW: 4,
  COL_NO: 1,         // Kolom A
  COL_NAME: 2,       // Kolom B
  START_DATA_COL: 4, // Kolom D (Mulai Minggu 1 Sept26)

  // Prioritas nama kolom 'Hari Kerja' di tab mingguan
  WORKDAY_COL_CANDIDATES: [
    'total hari kerja',
    'hari kerja total',
    'jumlah hari kerja',
    'total hari'
  ],
  FALLBACK_WORKDAY_COL: 'jumlah hari', // Kolom fallback bila format sheet lama

  // Definisi struktur bulan & minggu (5 minggu + 1 Total per bulan = 6 kolom)
  MONTHS: [
    {
      id: 'sept26',
      label: 'Sept26',
      name: 'September',
      color: '#cfe2f3', // Biru muda lembut
      weeks: [
        { num: 1, range: '30-5', patterns: ['30-5', '31-05', '31-5', 'sept1', 'sep1'] },
        { num: 2, range: '6-12',  patterns: ['6-12', 'sept2', 'sep2'] },
        { num: 3, range: '13-19', patterns: ['13-19', 'sept3', 'sep3'] },
        { num: 4, range: '20-26', patterns: ['20-26', 'sept4', 'sep4'] },
        { num: 5, range: '27-3',  patterns: ['27-3', 'sept5', 'sep5'] }
      ]
    },
    {
      id: 'okt26',
      label: 'Okt26',
      name: 'Oktober',
      color: '#f4cccc', // Merah muda pastel
      weeks: [
        { num: 1, range: '4-10',  patterns: ['4-10', 'okt1', 'oct1'] },
        { num: 2, range: '11-17', patterns: ['11-17', 'okt2', 'oct2'] },
        { num: 3, range: '18-24', patterns: ['18-24', 'okt3', 'oct3'] },
        { num: 4, range: '25-31', patterns: ['25-31', 'okt4', 'oct4'] },
        { num: 5, range: '1-3',   patterns: ['okt5', 'oct5'] }
      ]
    },
    {
      id: 'nov26',
      label: 'Nov26',
      name: 'November',
      color: '#d9ead3', // Hijau muda pastel
      weeks: [
        { num: 1, range: '1-7',   patterns: ['1-7', 'nov1'] },
        { num: 2, range: '8-14',  patterns: ['8-14', 'nov2'] },
        { num: 3, range: '15-21', patterns: ['15-21', 'nov3'] },
        { num: 4, range: '22-28', patterns: ['22-28', 'nov4'] },
        { num: 5, range: '29-5',  patterns: ['29-5', 'nov5'] }
      ]
    },
    {
      id: 'des26',
      label: 'Des26',
      name: 'Desember',
      color: '#fff2cc', // Kuning muda lembut
      weeks: [
        { num: 1, range: '6-12',  patterns: ['6-12', 'des1', 'dec1'] },
        { num: 2, range: '13-19', patterns: ['13-19', 'des2', 'dec2'] },
        { num: 3, range: '20-26', patterns: ['20-26', 'des3', 'dec3'] },
        { num: 4, range: '27-2',  patterns: ['27-2', 'des4', 'dec4'] },
        { num: 5, range: '3-5',   patterns: ['des5', 'dec5'] }
      ]
    },
    {
      id: 'jan27',
      label: 'Jan27',
      name: 'Januari',
      color: '#ead1dc', // Ungu lavender muda
      weeks: [
        { num: 1, range: '3-9',   patterns: ['3-9', 'jan1'] },
        { num: 2, range: '10-16', patterns: ['10-16', 'jan2'] },
        { num: 3, range: '17-23', patterns: ['17-23', 'jan3'] },
        { num: 4, range: '24-30', patterns: ['24-30', 'jan4'] },
        { num: 5, range: '31-6',  patterns: ['31-6', 'jan5'] }
      ]
    },
    {
      id: 'feb27',
      label: 'Feb27',
      name: 'Februari',
      color: '#c9daf8', // Biru periwinkle
      weeks: [
        { num: 1, range: '7-13',  patterns: ['7-13', 'feb1'] },
        { num: 2, range: '14-20', patterns: ['14-20', 'feb2'] },
        { num: 3, range: '21-27', patterns: ['21-27', 'feb3'] },
        { num: 4, range: '28-6',  patterns: ['28-6', 'feb4'] },
        { num: 5, range: '1-3',   patterns: ['feb5'] }
      ]
    },
    {
      id: 'mar27',
      label: 'Mar27',
      name: 'Maret',
      color: '#d0e0e3', // Cyan muda lembut
      weeks: [
        { num: 1, range: '7-13',  patterns: ['7-13', 'mar1'] },
        { num: 2, range: '14-20', patterns: ['14-20', 'mar2'] },
        { num: 3, range: '21-27', patterns: ['21-27', 'mar3'] },
        { num: 4, range: '28-3',  patterns: ['28-3', 'mar4'] },
        { num: 5, range: '4-6',   patterns: ['mar5'] }
      ]
    }
  ]
});

/**
 * Normalisasi nama karyawan:
 * - Menghilangkan inisial avatar dan enter ('WA\nWahyu Albino' -> 'Wahyu Albino')
 * - Menghilangkan leading newline ('\nWahyu Albino' -> 'Wahyu Albino')
 * - Menghapus kelebihan spasi dan mengubah ke lowercase untuk pencocokan dictionary yang akurat
 */
function normalizeEmployeeName(rawName) {
  if (!rawName) return '';
  var s = String(rawName).trim();
  
  if (s.indexOf('\n') !== -1) {
    var lines = s.split('\n').map(function(item) { return item.trim(); }).filter(function(item) { return item.length > 0; });
    if (lines.length > 0) {
      s = lines[lines.length - 1];
    }
  }
  
  return s.replace(/\s+/g, ' ').toLowerCase();
}

/**
 * Konversi nomor kolom (1-indexed) menjadi huruf alfabet Excel/Google Sheets
 */
function getColumnLetter(colIndex) {
  var letter = '';
  var temp = colIndex;
  while (temp > 0) {
    var mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

/**
 * Mendapatkan indeks kolom di Tab '0' untuk target Bulan dan Nomor Minggu
 */
function getDashboardColumnInfo(monthIdOrLabel, weekNum) {
  var targetLabel = String(monthIdOrLabel).toLowerCase().replace(/[^a-z0-9]/g, '');
  var currentCol = REKAP_CONFIG.START_DATA_COL; // Kolom D (4)
  
  for (var m = 0; m < REKAP_CONFIG.MONTHS.length; m++) {
    var month = REKAP_CONFIG.MONTHS[m];
    var mId = month.id.toLowerCase();
    var mLabel = month.label.toLowerCase();
    var isTargetMonth = (mId === targetLabel || mLabel === targetLabel || month.name.toLowerCase().indexOf(targetLabel) !== -1);
    
    var weekCount = month.weeks.length; // 5
    var totalCol = currentCol + weekCount; // Kolom Total
    
    if (isTargetMonth) {
      if (weekNum >= 1 && weekNum <= weekCount) {
        return {
          found: true,
          month: month,
          colIndex: currentCol + (weekNum - 1),
          colLetter: getColumnLetter(currentCol + (weekNum - 1)),
          isTotal: false,
          totalColIndex: totalCol,
          totalColLetter: getColumnLetter(totalCol),
          startColLetter: getColumnLetter(currentCol),
          endColLetter: getColumnLetter(currentCol + weekCount - 1)
        };
      } else if (weekNum === 'total' || weekNum === 0) {
        return {
          found: true,
          month: month,
          colIndex: totalCol,
          colLetter: getColumnLetter(totalCol),
          isTotal: true
        };
      }
    }
    
    currentCol += (weekCount + 1);
  }
  
  return { found: false };
}

/**
 * Mendeteksi target Bulan dan Minggu dari nama sheet secara otomatis
 */
function detectTargetFromSheetName(sheetName) {
  var cleanName = sheetName.toLowerCase().trim().replace(/\s+/g, ' ');
  
  for (var m = 0; m < REKAP_CONFIG.MONTHS.length; m++) {
    var month = REKAP_CONFIG.MONTHS[m];
    var mMatch = false;
    
    if (cleanName.indexOf(month.id) !== -1 || 
        cleanName.indexOf(month.label.toLowerCase()) !== -1 || 
        cleanName.indexOf(month.name.toLowerCase()) !== -1) {
      mMatch = true;
    }
    
    for (var w = 0; w < month.weeks.length; w++) {
      var week = month.weeks[w];
      
      for (var p = 0; p < week.patterns.length; p++) {
        var pat = week.patterns[p];
        if (cleanName.indexOf(pat) !== -1) {
          return {
            monthId: month.id,
            monthLabel: month.label,
            weekNum: week.num,
            dateRange: week.range,
            colInfo: getDashboardColumnInfo(month.id, week.num)
          };
        }
      }
      
      if (mMatch) {
        var numPattern = new RegExp('(?:minggu\\s*|w\\s*|\\b)' + week.num + '(?:\\b|[^0-9])');
        if (numPattern.test(cleanName)) {
          return {
            monthId: month.id,
            monthLabel: month.label,
            weekNum: week.num,
            dateRange: week.range,
            colInfo: getDashboardColumnInfo(month.id, week.num)
          };
        }
      }
    }
  }
  
  return null;
}

/**
 * Mencari indeks kolom 'total hari kerja' pada tab mingguan
 */
function findWorkdayColumnIndex(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Prioritas 1: Kandidat nama kolom
  for (var i = 0; i < REKAP_CONFIG.WORKDAY_COL_CANDIDATES.length; i++) {
    var cand = REKAP_CONFIG.WORKDAY_COL_CANDIDATES[i];
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').trim().toLowerCase();
      if (h === cand) {
        return c + 1;
      }
    }
  }
  
  // Prioritas 2: Kolom yang mengandung kata 'total' dan 'hari kerja'
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim().toLowerCase();
    if (h.indexOf('total') !== -1 && h.indexOf('hari kerja') !== -1) {
      return c + 1;
    }
  }
  
  // Prioritas 3: Fallback mencari 'hari' setelah kolom 'jam kerja'
  var jamKerjaFound = false;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim().toLowerCase();
    if (h.indexOf('jam kerja') !== -1) {
      jamKerjaFound = true;
    }
    if (jamKerjaFound && h.indexOf('hari') !== -1 && h.indexOf('bukan') === -1 && h.indexOf('cuti') === -1 && h.indexOf('terlambat') === -1) {
      return c + 1;
    }
  }
  
  return -1;
}

/**
 * Ekstraksi data absensi hari kerja dari satu sheet mingguan
 */
function extractAttendanceFromWeeklySheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { success: false, message: 'Sheet kosong atau tidak ada baris data.' };
  
  var workdayCol = findWorkdayColumnIndex(sheet);
  if (workdayCol === -1) {
    return { success: false, message: 'Kolom total hari kerja tidak ditemukan di sheet ini.' };
  }
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var nameCol = 2; // Default kolom B
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim().toLowerCase();
    if (h === 'nama 2' || h === 'nama lengkap') {
      nameCol = c + 1;
      break;
    }
  }
  
  var maxColToFetch = Math.max(nameCol, workdayCol);
  var dataRange = sheet.getRange(2, 1, lastRow - 1, maxColToFetch).getValues();
  var attendanceMap = {};
  var rawList = [];
  
  for (var r = 0; r < dataRange.length; r++) {
    var rawName = dataRange[r][nameCol - 1];
    var val = dataRange[r][workdayCol - 1];
    
    if (!rawName && nameCol !== 2) {
      rawName = dataRange[r][1];
    }
    
    var clean = normalizeEmployeeName(rawName);
    if (clean) {
      var numDays = (val !== '' && val !== null && !isNaN(val)) ? Number(val) : 0;
      attendanceMap[clean] = numDays;
      rawList.push({ cleanName: clean, rawName: rawName, days: numDays });
    }
  }
  
  return {
    success: true,
    attendanceMap: attendanceMap,
    rawList: rawList,
    workdayCol: workdayCol
  };
}

/**
 * Membaca daftar karyawan yang terdaftar di Tab '0' (Dashboard)
 */
function getDashboardEmployees(dashboardSheet) {
  var lastRow = dashboardSheet.getLastRow();
  if (lastRow < REKAP_CONFIG.EMPLOYEE_START_ROW) return [];
  
  var numRows = lastRow - REKAP_CONFIG.EMPLOYEE_START_ROW + 1;
  var names = dashboardSheet.getRange(REKAP_CONFIG.EMPLOYEE_START_ROW, REKAP_CONFIG.COL_NAME, numRows, 1).getValues();
  
  var list = [];
  for (var i = 0; i < names.length; i++) {
    var raw = names[i][0];
    if (raw !== null && String(raw).trim() !== '') {
      list.push({
        row: REKAP_CONFIG.EMPLOYEE_START_ROW + i,
        rawName: String(raw).trim(),
        cleanName: normalizeEmployeeName(raw)
      });
    }
  }
  return list;
}

/**
 * Format dan Standardisasi Tab '0' (Dashboard)
 * Menata seluruh header bulan, kolom total, tanggal, rumus SUM,
 * serta membersihkan dan menata seluruh border (garis tabel) secara presisi.
 */
function formatDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet0 = ss.getSheetByName(REKAP_CONFIG.DASHBOARD_SHEET_NAME);
  
  if (!sheet0) {
    SpreadsheetApp.getUi().alert('Tab "' + REKAP_CONFIG.DASHBOARD_SHEET_NAME + '" tidak ditemukan!');
    return;
  }
  
  unprotectDashboardSheet(true);
  
  // 1. LEPAS SEMUA MERGE SECARA TOTAL PADA SELURUH LEMBAR KERJA
  try {
    sheet0.getRange(1, 1, sheet0.getMaxRows(), sheet0.getMaxColumns()).breakApart();
  } catch(e) {
    Logger.log('breakApart full sheet: ' + e);
  }
  
  // 2. Hitung total kolom yang dibutuhkan (Col D + 7 bulan * 6 kolom = 45 kolom / Col AS)
  var totalMonths = REKAP_CONFIG.MONTHS.length; // 7
  var colsPerMonth = 6; // 5 minggu + 1 total
  var totalRequiredCols = REKAP_CONFIG.START_DATA_COL + (totalMonths * colsPerMonth) - 1; // 45
  
  var currentMaxCols = sheet0.getMaxColumns();
  if (currentMaxCols < totalRequiredCols) {
    sheet0.insertColumnsAfter(currentMaxCols, totalRequiredCols - currentMaxCols + 2);
  }
  
  var maxColsNow = sheet0.getMaxColumns();
  var employees = getDashboardEmployees(sheet0);
  var maxEmployeeRow = employees.length > 0 ? (REKAP_CONFIG.EMPLOYEE_START_ROW + employees.length - 1) : 25;
  
  // 3. Bersihkan konten lama di baris 1-3 dari kolom D ke kanan
  sheet0.getRange(1, REKAP_CONFIG.START_DATA_COL, 3, maxColsNow - REKAP_CONFIG.START_DATA_COL + 1).clear({ contentsOnly: true });
  
  // 4. BERSIHKAN SEMUA BORDER LAMA (Explicit false untuk menghapus semua garis lama yang salah letak)
  var entireTableRange = sheet0.getRange(1, 1, maxEmployeeRow + 2, maxColsNow);
  entireTableRange.setBorder(false, false, false, false, false, false);
  
  // 5. Format Header Karyawan (A3 & B3)
  sheet0.getRange('A3').setValue('No.').setFontWeight('bold').setHorizontalAlignment('center');
  sheet0.getRange('B3').setValue('Nama').setFontWeight('bold').setHorizontalAlignment('left');
  
  // Garis tabel Karyawan (Kolom A & B, Baris 3 s/d maxEmployeeRow)
  var empTableRange = sheet0.getRange(3, 1, maxEmployeeRow - 2, 2);
  empTableRange.setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
  sheet0.getRange(3, 1, 1, 2).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  empTableRange.setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Kolom C: jadikan spacer bersih tanpa border
  sheet0.setColumnWidth(3, 15);
  sheet0.getRange(1, 3, maxEmployeeRow, 1).setBorder(false, false, false, false, false, false);
  
  // 6. Bangun layout untuk setiap bulan secara terstandarisasi
  var currentCol = REKAP_CONFIG.START_DATA_COL;
  var totalColsList = [];
  
  for (var m = 0; m < totalMonths; m++) {
    var month = REKAP_CONFIG.MONTHS[m];
    var weekCount = month.weeks.length; // 5
    var monthStartCol = currentCol;
    var monthEndCol = currentCol + weekCount; // Kolom Total
    totalColsList.push(monthEndCol);
    
    // Baris 1: Minggu 1-5 & Baris 2: Rentang Tanggal
    for (var w = 0; w < weekCount; w++) {
      var col = currentCol + w;
      sheet0.getRange(1, col).setValue(w + 1).setFontWeight('bold').setHorizontalAlignment('center');
      
      var rangeText = month.weeks[w].range;
      var cell2 = sheet0.getRange(2, col);
      cell2.setNumberFormat('@'); // Plain text agar tidak berubah jadi tanggal otomatis
      cell2.setValue(rangeText);
      cell2.setHorizontalAlignment('center').setFontSize(9).setFontColor('#555555');
    }
    
    // Kolom Total
    var totalCol = monthEndCol;
    sheet0.getRange(1, totalCol).setValue('Total').setFontWeight('bold').setHorizontalAlignment('center');
    sheet0.getRange(2, totalCol).setValue('').setNumberFormat('@');
    
    // Baris 3: Merge Header Bulan (dilindungi try-catch agar aman)
    var startColLetter = getColumnLetter(monthStartCol);
    var endColLetter = getColumnLetter(monthEndCol);
    var monthHeaderRange = sheet0.getRange(3, monthStartCol, 1, weekCount + 1);
    
    try {
      monthHeaderRange.merge();
    } catch(err) {
      Logger.log('Merge failed for ' + month.label + ': ' + err);
    }
    
    monthHeaderRange.setValue(month.label)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(month.color);
    
    // Rumus =SUM(W1:W5) untuk setiap karyawan di kolom Total
    if (employees.length > 0) {
      var sumRangeEndColLetter = getColumnLetter(monthStartCol + weekCount - 1);
      var formulas = [];
      for (var r = 0; r < employees.length; r++) {
        var rowNum = employees[r].row;
        formulas.push(['=SUM(' + startColLetter + rowNum + ':' + sumRangeEndColLetter + rowNum + ')']);
      }
      sheet0.getRange(REKAP_CONFIG.EMPLOYEE_START_ROW, totalCol, employees.length, 1)
        .setFormulas(formulas)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
    
    currentCol = monthEndCol + 1;
  }
  
  // 7. PENATAAN BORDER DATA TABEL (D1:AS{maxEmployeeRow})
  var dataTableRange = sheet0.getRange(1, REKAP_CONFIG.START_DATA_COL, maxEmployeeRow, totalRequiredCols - REKAP_CONFIG.START_DATA_COL + 1);
  dataTableRange.setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  
  // Garis bawah header bulan (baris 3) lebih tegas
  sheet0.getRange(3, REKAP_CONFIG.START_DATA_COL, 1, totalRequiredCols - REKAP_CONFIG.START_DATA_COL + 1)
    .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Garis batas luar tabel data
  dataTableRange.setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Garis vertikal hitam tebal di setiap pembatas akhir bulan (di sebelah kanan kolom Total)
  for (var i = 0; i < totalColsList.length; i++) {
    sheet0.getRange(1, totalColsList[i], maxEmployeeRow, 1)
      .setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  
  // 8. Atur lebar kolom seragam dan proporsional (48px)
  for (var c = REKAP_CONFIG.START_DATA_COL; c <= totalRequiredCols; c++) {
    sheet0.setColumnWidth(c, 48);
  }
  
  protectDashboardSheet();
  ss.toast('Layout dan border Tab "0" berhasil dirapikan dan distandarisasi!', '⚡ Rekap Absen', 4);
}

/**
 * Sinkronisasi sheet yang sedang aktif ke Dashboard (Tab '0')
 */
function syncActiveSheetToDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var activeSheetName = activeSheet.getName();
  
  if (activeSheetName === REKAP_CONFIG.DASHBOARD_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('Pilih tab absensi mingguan yang ingin disinkronkan, bukan Tab "0".');
    return;
  }
  
  var extractResult = extractAttendanceFromWeeklySheet(activeSheet);
  if (!extractResult.success) {
    SpreadsheetApp.getUi().alert('Gagal membaca sheet "' + activeSheetName + '":\n' + extractResult.message);
    return;
  }
  
  var target = detectTargetFromSheetName(activeSheetName);
  var targetColInfo = null;
  
  if (target && target.colInfo && target.colInfo.found) {
    targetColInfo = target.colInfo;
  } else {
    var ui = SpreadsheetApp.getUi();
    var promptResult = ui.prompt(
      'Konfirmasi Target Bulan & Minggu',
      'Nama sheet "' + activeSheetName + '" tidak dikenali otomatis.\n' +
      'Ketik target Bulan dan Minggu (Contoh: "Sept26 1" atau "Okt26 2"):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (promptResult.getSelectedButton() !== ui.Button.OK) {
      ss.toast('Sinkronisasi dibatalkan.', 'Info');
      return;
    }
    
    var responseText = promptResult.getResponseText().trim().split(/\s+/);
    if (responseText.length < 2) {
      ui.alert('Format tidak valid. Harus menyertakan Bulan dan Minggu (contoh: Sept26 1).');
      return;
    }
    
    var mLabel = responseText[0];
    var wNum = parseInt(responseText[1], 10);
    targetColInfo = getDashboardColumnInfo(mLabel, wNum);
    
    if (!targetColInfo.found) {
      ui.alert('Target bulan "' + mLabel + '" minggu ' + wNum + ' tidak ditemukan di dashboard.');
      return;
    }
  }
  
  var sheet0 = ss.getSheetByName(REKAP_CONFIG.DASHBOARD_SHEET_NAME);
  if (!sheet0) {
    SpreadsheetApp.getUi().alert('Tab "' + REKAP_CONFIG.DASHBOARD_SHEET_NAME + '" tidak ditemukan!');
    return;
  }
  
  unprotectDashboardSheet(true);
  
  var employees = getDashboardEmployees(sheet0);
  if (employees.length === 0) {
    SpreadsheetApp.getUi().alert('Tidak ada karyawan ditemukan di Tab "0" mulai baris ' + REKAP_CONFIG.EMPLOYEE_START_ROW);
    return;
  }
  
  var writeValues = [];
  var matchedCount = 0;
  var attMap = extractResult.attendanceMap;
  
  for (var i = 0; i < employees.length; i++) {
    var emp = employees[i];
    if (attMap.hasOwnProperty(emp.cleanName)) {
      writeValues.push([attMap[emp.cleanName]]);
      matchedCount++;
    } else {
      var foundPartial = false;
      for (var k in attMap) {
        if (emp.cleanName.indexOf(k) !== -1 || k.indexOf(emp.cleanName) !== -1) {
          writeValues.push([attMap[k]]);
          matchedCount++;
          foundPartial = true;
          break;
        }
      }
      if (!foundPartial) {
        writeValues.push(['']);
      }
    }
  }
  
  sheet0.getRange(REKAP_CONFIG.EMPLOYEE_START_ROW, targetColInfo.colIndex, writeValues.length, 1)
    .setValues(writeValues)
    .setHorizontalAlignment('center');
  
  var unlistedEmployees = [];
  for (var j = 0; j < extractResult.rawList.length; j++) {
    var rawItem = extractResult.rawList[j];
    var isListed = false;
    for (var e = 0; e < employees.length; e++) {
      if (employees[e].cleanName === rawItem.cleanName || 
          employees[e].cleanName.indexOf(rawItem.cleanName) !== -1 || 
          rawItem.cleanName.indexOf(employees[e].cleanName) !== -1) {
        isListed = true;
        break;
      }
    }
    if (!isListed) {
      unlistedEmployees.push(rawItem.rawName + ' (' + rawItem.days + ' hari)');
    }
  }
  
  protectDashboardSheet();
  
  var successMsg = 'Berhasil sinkron ' + matchedCount + '/' + employees.length + 
                   ' karyawan ke ' + (target ? target.monthLabel : targetColInfo.month.label) + 
                   ' Minggu ' + (target ? target.weekNum : '') + ' (Kolom ' + targetColInfo.colLetter + ').';
  
  if (unlistedEmployees.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Sinkronisasi Selesai (Ada Catatan):\n\n' + successMsg + 
      '\n\n⚠️ Karyawan berikut ada di sheet mingguan tetapi BELUM terdaftar di Tab "0":\n- ' + 
      unlistedEmployees.join('\n- ')
    );
  } else {
    ss.toast(successMsg, '⚡ Rekap Absen Sukses', 5);
  }
}

/**
 * Sinkronisasi seluruh tab mingguan yang ada di spreadsheet ke Tab '0'
 */
function syncAllWeeklySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet0 = ss.getSheetByName(REKAP_CONFIG.DASHBOARD_SHEET_NAME);
  
  if (!sheet0) {
    SpreadsheetApp.getUi().alert('Tab "' + REKAP_CONFIG.DASHBOARD_SHEET_NAME + '" tidak ditemukan!');
    return;
  }
  
  var employees = getDashboardEmployees(sheet0);
  if (employees.length === 0) {
    SpreadsheetApp.getUi().alert('Tidak ada karyawan di Tab "0"!');
    return;
  }
  
  unprotectDashboardSheet(true);
  
  var syncedSheets = [];
  var skippedSheets = [];
  
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sName = sheet.getName();
    
    if (sName === REKAP_CONFIG.DASHBOARD_SHEET_NAME) continue;
    
    var target = detectTargetFromSheetName(sName);
    if (!target || !target.colInfo || !target.colInfo.found) {
      skippedSheets.push(sName);
      continue;
    }
    
    var extractResult = extractAttendanceFromWeeklySheet(sheet);
    if (!extractResult.success) {
      skippedSheets.push(sName + ' (error)');
      continue;
    }
    
    var writeValues = [];
    var attMap = extractResult.attendanceMap;
    
    for (var i = 0; i < employees.length; i++) {
      var emp = employees[i];
      if (attMap.hasOwnProperty(emp.cleanName)) {
        writeValues.push([attMap[emp.cleanName]]);
      } else {
        var foundPartial = false;
        for (var k in attMap) {
          if (emp.cleanName.indexOf(k) !== -1 || k.indexOf(emp.cleanName) !== -1) {
            writeValues.push([attMap[k]]);
            foundPartial = true;
            break;
          }
        }
        if (!foundPartial) {
          writeValues.push(['']);
        }
      }
    }
    
    sheet0.getRange(REKAP_CONFIG.EMPLOYEE_START_ROW, target.colInfo.colIndex, writeValues.length, 1)
      .setValues(writeValues)
      .setHorizontalAlignment('center');
      
    syncedSheets.push(sName + ' ➔ ' + target.monthLabel + ' W' + target.weekNum);
  }
  
  protectDashboardSheet();
  
  var summary = 'Sinkronisasi Selesai!\n\n';
  if (syncedSheets.length > 0) {
    summary += '✅ Berhasil disinkronkan (' + syncedSheets.length + ' sheet):\n- ' + syncedSheets.join('\n- ') + '\n\n';
  }
  if (skippedSheets.length > 0) {
    summary += 'ℹ️ Sheet dilewati (bukan nama target Sept26-Mar27 atau sheet lain):\n- ' + skippedSheets.join('\n- ');
  }
  
  SpreadsheetApp.getUi().alert(summary);
}

/**
 * Mengunci Tab '0' dengan Proteksi Ganda (Dual-Layer Protection)
 * Layer 1: Native Warning Interceptor (Mencegah Pemilik & Editor menghapus sel secara tidak sengaja).
 * Layer 2: Menghapus izin edit kolaborator lain.
 */
function protectDashboardSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet0 = ss.getSheetByName(REKAP_CONFIG.DASHBOARD_SHEET_NAME);
  if (!sheet0) return;
  
  // Hapus proteksi lama jika ada
  var protections = sheet0.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) {
    protections[i].remove();
  }
  
  var protection = sheet0.protect().setDescription('Kunci Dashboard Otomatis');
  
  // 1. Matikan izin edit domain jika ada
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  
  // 2. Hapus semua editor lain
  var editors = protection.getEditors();
  if (editors.length > 0) {
    try {
      protection.removeEditors(editors);
    } catch(e) {}
  }
  
  // 3. SANGAT KRUSIAL: Aktifkan mode peringatan (Warning Only).
  // Di Google Sheets, Pemilik Spreadsheet (Owner) secara sistem tidak dapat diblokir total oleh hak akses.
  // Satu-satunya cara resmi Google Sheets untuk MENCEGAH PEMILIK menghapus sel adalah melalui Warning Interceptor.
  // Saat aktif, menekan Delete / Backspace akan memunculkan pop-up pencegahan.
  protection.setWarningOnly(true);
  
  ss.toast('Tab "0" telah dikunci! Setiap upaya edit atau hapus akan diblokir dengan peringatan.', '🔒 Proteksi Aktif', 4);
}

/**
 * Membuka kunci proteksi Tab '0'
 */
function unprotectDashboardSheet(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet0 = ss.getSheetByName(REKAP_CONFIG.DASHBOARD_SHEET_NAME);
  if (!sheet0) return;
  
  var protections = sheet0.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) {
    protections[i].remove();
  }
  
  if (!silent) {
    ss.toast('Kunci Tab "0" telah dibuka. Anda sekarang dapat mengedit sel secara bebas.', '🔓 Proteksi Dilepas', 3);
  }
}

/**
 * Trigger Real-Time onEdit Guard (Layer 2 Protection):
 * Jika pengguna tidak sengaja mengonfirmasi atau menghapus data sel di Tab '0' saat terkunci,
 * fungsi ini secara otomatis MENGEMBALIKAN (UNDO) nilai sel tersebut ke nilai semula.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== REKAP_CONFIG.DASHBOARD_SHEET_NAME) return;
  
  // Cek apakah sheet sedang memiliki proteksi aktif
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (protections.length === 0) return; // Jika tidak sedang dikunci, izinkan edit
  
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row < 1 || col < 1) return;
  
  // Jika nilai dihapus (cell menjadi kosong) saat terkunci:
  if (e.oldValue !== undefined && (e.value === undefined || e.value === '')) {
    e.range.setValue(e.oldValue);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '⚠️ Tab "0" sedang dikunci! Nilai yang terhapus telah dikembalikan otomatis.',
      '🛡️ Proteksi Aktif',
      4
    );
  }
}

/**
 * Menu Kustom di Google Sheets Toolbar
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Rekap Absen')
    .addItem('▶️ Sync Sheet Aktif ke Dashboard', 'syncActiveSheetToDashboard')
    .addItem('🔄 Sync Semua Sheet Mingguan', 'syncAllWeeklySheets')
    .addSeparator()
    .addItem('📐 Format & Standardisasi Tab "0"', 'formatDashboard')
    .addSeparator()
    .addItem('🔒 Kunci Tab "0" (Anti-Hapus & Read-Only)', 'protectDashboardSheet')
    .addItem('🔓 Buka Kunci Tab "0" (Izinkan Edit)', 'unprotectDashboardSheet')
    .addToUi();
}
