/**
 * =========================================================================
 * PURCHASING - DYNAMIC WEEKLY DROPDOWN & STOCK QUANTITY LOOKUP
 * Google Sheets + Google Apps Script Automation
 * =========================================================================
 * 
 * FITUR KEAMANAN GANTI NAMA SHEET (ANTI-BREAK):
 * Menggunakan GID (Sheet ID Unik Permanen) dari Google Sheets, sehingga jika
 * nama tab diubah/direname, script dan formula tetap bekerja 100% normal.
 */

const PURCHASING_CONFIG = Object.freeze({
  // GID Permanen (Tidak akan berubah meskipun nama tab diganti)
  SOURCE_SHEET_GID: 2088019415,      // Tab sumber Qty mingguan (Sheet10)
  MIN_STOCK_SHEET_GID: 1499872074,   // Tab sumber Batas Minimum Stok (Copy of Data)
  TARGET_SHEET_GID: 1711808705,      // Tab target (Step 3a)

  // Nama Fallback (jika GID tidak cocok)
  FALLBACK_SOURCE_NAME: 'Sheet10',
  FALLBACK_MIN_STOCK_NAME: 'Copy of Data',
  FALLBACK_TARGET_NAME: 'Step 3a',
  
  // Posisi di Sheet Sumber Mingguan
  SOURCE_HEADER_ROW: 3,      // Baris nama periode/minggu (K3, L3, ...)
  SOURCE_START_COL: 11,      // Kolom K (11)
  SOURCE_DATA_START_ROW: 5,  // Baris mulai data stok barang
  SOURCE_NAME_COL: 1,        // Kolom A (Nama Barang)
  
  // Posisi di Tab Master Batas Stok
  MIN_NAME_COL_LETTER: 'C',  // Kolom C (Nama Barang)
  MIN_STOCK_COL_LETTER: 'J', // Kolom J (Batas Minimum Stok)
  
  // Posisi di Tab Target
  DROPDOWN_CELL: 'C3',       // Sel dropdown minggu
  TARGET_DATA_START_ROW: 5,  // Baris mulai data
  TARGET_NAME_COL: 1,        // Kolom A (Nama Barang)
  TARGET_MIN_COL: 2,         // Kolom B (Batas Minimal)
  TARGET_QTY_COL: 3          // Kolom C (Qty)
});

/**
 * Helper untuk mencari Sheet berdasarkan GID Permanen atau Nama Fallback
 */
function getTargetSheet_(ss, gid, fallbackName) {
  if (gid !== null && gid !== undefined) {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === gid) {
        return sheets[i];
      }
    }
  }
  return ss.getSheetByName(fallbackName);
}

/**
 * Menu Khusus saat Spreadsheet dibuka
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Purchasing')
    .addItem('🔄 Perbarui Dropdown Minggu (C3)', 'updateDropdownWeekList')
    .addItem('📐 Pasang Formula Batas Minimal & Qty Otomatis', 'applyFormulasStep3a')
    .addSeparator()
    .addItem('🚀 Sinkronisasi Penuh (Dropdown + Formula)', 'syncWeeklyDropdownAndFormulas')
    .addToUi();

  // Otomatis sinkronisasi penuh saat spreadsheet dibuka
  try {
    syncWeeklyDropdownAndFormulas();
  } catch (err) {
    console.warn('Gagal sinkronisasi otomatis saat onOpen: ' + err.message);
  }
}

/**
 * Trigger onEdit:
 * Otomatis sinkronisasi jika ada perubahan pada sheet sumber mingguan atau sel dropdown target
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetId = sheet.getSheetId();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheetId === PURCHASING_CONFIG.SOURCE_SHEET_GID || sheet.getName() === PURCHASING_CONFIG.FALLBACK_SOURCE_NAME) {
    if (row === PURCHASING_CONFIG.SOURCE_HEADER_ROW && col >= PURCHASING_CONFIG.SOURCE_START_COL) {
      updateDropdownWeekList();
    }
  } else if (sheetId === PURCHASING_CONFIG.TARGET_SHEET_GID || sheet.getName() === PURCHASING_CONFIG.FALLBACK_TARGET_NAME) {
    if (e.range.getA1Notation() === PURCHASING_CONFIG.DROPDOWN_CELL) {
      applyFormulasStep3a();
    }
  }
}

/**
 * 1. Mengatur Data Validation di Step 3a!C3 sebagai "Dropdown dari rentang"
 */
function updateDropdownWeekList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = getTargetSheet_(ss, PURCHASING_CONFIG.SOURCE_SHEET_GID, PURCHASING_CONFIG.FALLBACK_SOURCE_NAME);
  const targetSheet = getTargetSheet_(ss, PURCHASING_CONFIG.TARGET_SHEET_GID, PURCHASING_CONFIG.FALLBACK_TARGET_NAME);

  if (!sourceSheet) {
    throw new Error('Tab sumber mingguan tidak ditemukan.');
  }
  if (!targetSheet) {
    throw new Error('Tab target (Step 3a) tidak ditemukan.');
  }

  const sourceSheetName = sourceSheet.getName();
  const maxCol = Math.max(sourceSheet.getLastColumn(), sourceSheet.getMaxColumns());
  const maxColLetter = getColumnLetter_(maxCol);

  // Range rentang minggu: misal 'Sheet10'!K3:Z3
  const rangeRule = sourceSheet.getRange(`K3:${maxColLetter}3`);

  // Buat aturan Data Validation: Dropdown dari Rentang (requireValueInRange)
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeRule, true)
    .setAllowInvalid(true)
    .build();

  const dropdownCell = targetSheet.getRange(PURCHASING_CONFIG.DROPDOWN_CELL);
  dropdownCell.setDataValidation(rule);

  // Jika nilai saat ini kosong, set otomatis ke minggu terakhir yang terisi
  const lastCol = sourceSheet.getLastColumn();
  if (lastCol >= PURCHASING_CONFIG.SOURCE_START_COL) {
    const lastWeekVal = sourceSheet.getRange(PURCHASING_CONFIG.SOURCE_HEADER_ROW, lastCol).getValue();
    const currentVal = dropdownCell.getValue();
    if (!currentVal && lastWeekVal) {
      dropdownCell.setValue(lastWeekVal);
    }
  }

  console.log(`Dropdown C3 berhasil diatur dinamis dari rentang '${sourceSheetName}'!K3:${maxColLetter}3`);
}

/**
 * Konversi nomor kolom (1-based) menjadi huruf kolom (contoh: 1 -> A, 26 -> Z, 27 -> AA)
 */
function getColumnLetter_(colIndex) {
  let temp, letter = '';
  let col = colIndex;
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = Math.floor((col - temp - 1) / 26);
  }
  return letter || 'Z';
}

/**
 * 2. Memasang formula pencarian dinamis di kolom Batas Minimal (B5:B) dan Qty (C5:C)
 */
function applyFormulasStep3a() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = getTargetSheet_(ss, PURCHASING_CONFIG.SOURCE_SHEET_GID, PURCHASING_CONFIG.FALLBACK_SOURCE_NAME);
  const minSheet = getTargetSheet_(ss, PURCHASING_CONFIG.MIN_STOCK_SHEET_GID, PURCHASING_CONFIG.FALLBACK_MIN_STOCK_NAME) || ss.getSheetByName('Master Data') || ss.getSheetByName('Batas Minimal Stok');
  const targetSheet = getTargetSheet_(ss, PURCHASING_CONFIG.TARGET_SHEET_GID, PURCHASING_CONFIG.FALLBACK_TARGET_NAME);

  if (!sourceSheet || !targetSheet) {
    throw new Error('Sheet sumber atau target tidak ditemukan.');
  }

  const sourceName = sourceSheet.getName();
  const minName = minSheet ? minSheet.getName() : PURCHASING_CONFIG.FALLBACK_MIN_STOCK_NAME;

  const lastRow = targetSheet.getLastRow();
  if (lastRow < PURCHASING_CONFIG.TARGET_DATA_START_ROW) {
    console.warn('Tidak ada baris data di sheet target.');
    return;
  }

  const maxCol = Math.max(sourceSheet.getLastColumn(), sourceSheet.getMaxColumns());
  const maxColLetter = getColumnLetter_(maxCol);

  const numRows = lastRow - PURCHASING_CONFIG.TARGET_DATA_START_ROW + 1;
  const nameValues = targetSheet.getRange(
    PURCHASING_CONFIG.TARGET_DATA_START_ROW,
    PURCHASING_CONFIG.TARGET_NAME_COL,
    numRows,
    1
  ).getValues();

  const batasMinimalFormulas = [];
  const qtyFormulas = [];

  for (let i = 0; i < numRows; i++) {
    const rowNum = PURCHASING_CONFIG.TARGET_DATA_START_ROW + i;
    const name = String(nameValues[i][0]).trim();
    if (name === '') {
      batasMinimalFormulas.push(['']);
      qtyFormulas.push(['']);
    } else {
      // 1. Formula Batas Minimal (dinamis ke nama tab master saat ini)
      const batasFormula = `=IF(ISBLANK(A${rowNum}); ""; IFERROR(INDEX('${minName}'!$${PURCHASING_CONFIG.MIN_STOCK_COL_LETTER}:$${PURCHASING_CONFIG.MIN_STOCK_COL_LETTER}; MATCH(A${rowNum}; '${minName}'!$${PURCHASING_CONFIG.MIN_NAME_COL_LETTER}:$${PURCHASING_CONFIG.MIN_NAME_COL_LETTER}; 0)); 0))`;
      batasMinimalFormulas.push([batasFormula]);

      // 2. Formula Qty Mingguan (dinamis ke nama tab mingguan saat ini)
      const qtyFormula = `=IF(ISBLANK(A${rowNum}); ""; IFERROR(INDEX('${sourceName}'!$A:$${maxColLetter}; MATCH(A${rowNum}; '${sourceName}'!$A:$A; 0); MATCH($C$3; '${sourceName}'!$3:$3; 0)); 0))`;
      qtyFormulas.push([qtyFormula]);
    }
  }

  // Pasang formula Kolom B (Batas Minimal)
  const batasRange = targetSheet.getRange(
    PURCHASING_CONFIG.TARGET_DATA_START_ROW,
    PURCHASING_CONFIG.TARGET_MIN_COL,
    numRows,
    1
  );
  try {
    batasRange.setFormulasLocal(batasMinimalFormulas);
  } catch (err) {
    batasRange.setFormulas(batasMinimalFormulas);
  }

  // Pasang formula Kolom C (Qty)
  const qtyRange = targetSheet.getRange(
    PURCHASING_CONFIG.TARGET_DATA_START_ROW,
    PURCHASING_CONFIG.TARGET_QTY_COL,
    numRows,
    1
  );
  try {
    qtyRange.setFormulasLocal(qtyFormulas);
  } catch (err) {
    qtyRange.setFormulas(qtyFormulas);
  }

  console.log(`Formula berhasil dipasang merujuk ke tab '${minName}' & '${sourceName}'.`);
}

// Alias untuk kompatibilitas
function applyQtyFormulasStep3a() {
  applyFormulasStep3a();
}

/**
 * 3. Sinkronisasi penuh (Update Dropdown + Terapkan Formula)
 */
function syncWeeklyDropdownAndFormulas() {
  updateDropdownWeekList();
  applyFormulasStep3a();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Dropdown C3, Batas Minimal & Qty berhasil disinkronkan!', '⚡ Purchasing', 4);
  } catch (e) {}
}
