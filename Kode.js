// Updated: 1789958589.454242
/**
 * PGE Order Tracking & Alert System
 * Google Sheets + Google Apps Script operational monitoring for unfinished production orders
 * PRD Revision: 1.0
 */

const ORDER_TRACKER = Object.freeze({
  START_ROW: 5,

  // Summary & Button Anchors
  UNFINISHED_SUMMARY_CELL: 'K1',
  BUTTON_ANCHOR_CELL: 'L1',
  BUTTON_ALT_TITLE: 'PGE_OPEN_UNFINISHED_ORDERS',
  BUTTON_HANDLER: 'showUnfinishedOrdersSidebar',
  BUTTON_WIDTH: 180,
  BUTTON_HEIGHT: 38,
  HOVER_NOTE_MAX_ITEMS: 20,

  // Operational Column Indexes (1-indexed)
  ORDER_COLUMN: 3,          // Column C — Order Detail / Deskripsi Pesanan
  DATE_COLUMN: 4,           // Column D — Order Date / Tanggal Pesanan
  PROCESS_START_COLUMN: 6,  // Column F — Process Start / Selesai Print (Timer Start)
  COMPLETION_COLUMN: 9,     // Column I — Completion Timestamp / Waktu Selesai
  DURATION_COLUMN: 10,      // Column J — Duration / Durasi Pengerjaan
  ORDER_TYPE_COLUMN: 11,    // Column K — Order Type / Urgent Marker

  // Visual Highlight Ranges
  NORMAL_END_COLUMN: 10,    // Color range A:J for standard orders
  URGENT_END_COLUMN: 11,    // Color range A:K for urgent orders

  // Thresholds
  ONE_HOUR_MS: 60 * 60 * 1000,
  TWO_HOURS_MS: 2 * 60 * 60 * 1000,
  THREE_HOURS_MS: 3 * 60 * 60 * 1000,

  // Status Alert Colors
  COLORS: {
    YELLOW: '#fff2cc', // Warning: >= 1 hour
    ORANGE: '#f9cb9c', // High: >= 2 hours
    RED: '#f4cccc',    // Critical: >= 3 hours / Urgent
    PURPLE: '#ede9fe', // Keperluan Video / Internal Content
    WHITE: '#ffffff'   // Finished / Normal baseline
  },

  // Indonesian Uppercase Month Names
  MONTH_NAMES_ID: [
    'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI',
    'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
  ],

  BUTTON_PNG_BASE64: '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',
  STATE_PREFIX: 'PGE_D_DATE_TRACKER_V1_',
  LEGACY_STATE_PREFIX: 'PGE_SIMPLE_TRACKER_'
});


/* =========================================================
 * OPEN AND MENU
 * ========================================================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pelacak PGE')
    .addItem('Upload Foto Desain', 'showImageUploadDialog')
    .addItem('Buka Tracker', 'showUnfinishedOrdersSidebar')
    .addItem('Ke Tanggal Hari Ini', 'jumpToTodayDate')
    .addSeparator()
    .addItem('Pasang Tombol Tracker (L1)', 'drawButtonL1')
    .addItem('Cek Baris Terakhir', 'showLastRowAlert')
    .addItem('Perbarui Warna', 'updateOrderTrackingColors')
    .addSeparator()
    .addItem('Pasang Pelacak Otomatis', 'installOrderTrackingTrigger')
    .addItem('Reset Warna', 'resetOrderTrackingColorsOnce')
    .addToUi();

  try {
    ensureTrackedSheetsConfigured_();
  } catch (error) {
    console.warn('Gagal cek tombol otomatis: ' + error.message);
  }

  try {
    jumpToTodayDate();
  } catch (error) {
    console.warn('Tidak dapat otomatis menuju tanggal hari ini: ' + error.message);
  }
}


/**
 * Navigates to the current monthly sheet and activates the first row for today.
 */
function jumpToTodayDate() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Skrip harus terhubung ke Google Sheet.');
  }

  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const now = new Date();
  const currentMonthName = getIndonesianMonthName_(now.getMonth());
  const sheet = spreadsheet.getSheetByName(currentMonthName);

  if (!sheet) {
    throw new Error('Sheet bulan ini (' + currentMonthName + ') tidak ditemukan.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < ORDER_TRACKER.START_ROW) {
    return {
      found: false,
      reason: 'Sheet ' + currentMonthName + ' belum memiliki data pesanan.'
    };
  }

  const todayKey = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');
  const rowCount = lastRow - ORDER_TRACKER.START_ROW + 1;
  const dateValues = sheet
    .getRange(ORDER_TRACKER.START_ROW, ORDER_TRACKER.DATE_COLUMN, rowCount, 1)
    .getValues();

  let targetRow = null;

  for (let index = 0; index < dateValues.length; index += 1) {
    const dateValue = getValidDate_(dateValues[index][0]);
    if (!dateValue) {
      continue;
    }

    const dateKey = formatDayKey_(dateValue, timeZone);
    if (dateKey === todayKey) {
      targetRow = ORDER_TRACKER.START_ROW + index;
      break;
    }
  }

  if (!targetRow) {
    spreadsheet.setActiveSheet(sheet);
    return {
      found: false,
      today: todayKey,
      sheetName: currentMonthName,
      reason: 'Tanggal hari ini (' + todayKey + ') tidak ditemukan di kolom D sheet ' + currentMonthName + '.'
    };
  }

  spreadsheet.setActiveSheet(sheet);
  const targetCell = sheet.getRange(targetRow, ORDER_TRACKER.DATE_COLUMN);
  spreadsheet.setActiveRange(targetCell);
  targetCell.activate();
  SpreadsheetApp.flush();

  return {
    found: true,
    sheetName: sheet.getName(),
    rowNumber: targetRow,
    today: todayKey
  };
}


/**
 * Opens a specific order row on the specified sheet from the sidebar.
 */
function jumpToOrderRow(sheetName, rowNumber) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Skrip harus terhubung ke Google Sheet.');
  }

  const requestedName = String(sheetName || '').trim();
  let sheet = requestedName ? spreadsheet.getSheetByName(requestedName) : null;

  if (!sheet) {
    const defaultSheets = getTrackedSheets_(spreadsheet);
    sheet = defaultSheets.length ? defaultSheets[0] : null;
  }

  if (!sheet) {
    throw new Error('Sheet pesanan (' + requestedName + ') tidak ditemukan.');
  }

  const row = Number(rowNumber);
  if (!Number.isInteger(row) || row < ORDER_TRACKER.START_ROW || row > sheet.getMaxRows()) {
    throw new Error('Nomor baris pesanan tidak valid: ' + rowNumber);
  }

  spreadsheet.setActiveSheet(sheet);
  const targetRange = sheet.getRange(
    row,
    1,
    1,
    Math.min(ORDER_TRACKER.URGENT_END_COLUMN, sheet.getMaxColumns())
  );

  spreadsheet.setActiveRange(targetRange);
  SpreadsheetApp.flush();

  return {
    sheetName: sheet.getName(),
    rowNumber: row
  };
}


/* =========================================================
 * EDIT AND TRIGGERS
 * ========================================================= */


/**
 * Automatically opens the sidebar whenever a user clicks on cell K1 or L1 (BUKA TRACKER area).
 * This guarantees the tracker opens on click even if an image button script is unassigned.
 */
function onSelectionChange(e) {
  if (!e || !e.range) return;
  try {
    const sheet = e.range.getSheet();
    if (!isTrackedSheet_(sheet)) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Column 11 (K1) or Column 12 (L1) or Column 13 (M1)
    if (row === 1 && (col === 11 || col === 12 || col === 13)) {
      showUnfinishedOrdersSidebar();
    }
  } catch (error) {
    console.warn('onSelectionChange error: ' + error.message);
  }
}


function onEdit(event) {
  if (!event || !event.range) {
    return;
  }

  const sheet = event.range.getSheet();
  if (!isTrackedSheet_(sheet)) {
    return;
  }

  const range = event.range;
  const startRow = Math.max(ORDER_TRACKER.START_ROW, range.getRow());
  const endRow = range.getLastRow();
  const startCol = range.getColumn();
  const endCol = range.getLastColumn();

  // Auto-fill Column A (TGL) with today's date when Column C (Detail) is populated and Column A is empty
  if (startCol <= ORDER_TRACKER.ORDER_COLUMN && endCol >= ORDER_TRACKER.ORDER_COLUMN && endRow >= ORDER_TRACKER.START_ROW) {
    autoFillDateColumnA_(sheet, startRow, endRow);
  }

  const touchesTrackedColumns =
    endCol >= ORDER_TRACKER.ORDER_COLUMN &&
    startCol <= ORDER_TRACKER.ORDER_TYPE_COLUMN;

  const touchesOrderRows = endRow >= ORDER_TRACKER.START_ROW;

  if (!touchesTrackedColumns || !touchesOrderRows) {
    return;
  }

  SpreadsheetApp.flush();
  updateOneTrackingSheet_(sheet);
}


/**
 * Automatically populates Column A with today's date ONLY for the first order created on that day.
 * Subsequent orders on the same day will keep Column A blank.
 */
function autoFillDateColumnA_(sheet, startRow, endRow) {
  const rowCount = endRow - startRow + 1;
  if (rowCount < 1) return;

  const timeZone = sheet.getParent().getSpreadsheetTimeZone();
  const today = new Date();
  const todayKey = Utilities.formatDate(today, timeZone, 'yyyy-MM-dd');
  const lastRow = sheet.getLastRow();

  // 1. Check if today's date already exists anywhere in Column A (from START_ROW down to lastRow)
  let todayAlreadyExists = false;
  if (lastRow >= ORDER_TRACKER.START_ROW) {
    const totalDataRows = lastRow - ORDER_TRACKER.START_ROW + 1;
    const colAValues = sheet.getRange(ORDER_TRACKER.START_ROW, 1, totalDataRows, 1).getValues();

    for (let i = 0; i < colAValues.length; i++) {
      const dateVal = getValidDate_(colAValues[i][0]);
      if (dateVal) {
        const key = formatDayKey_(dateVal, timeZone);
        if (key === todayKey) {
          todayAlreadyExists = true;
          break;
        }
      }
    }
  }

  // If today's date is already recorded in Column A on an earlier row, do not write it again
  if (todayAlreadyExists) {
    return;
  }

  // 2. Otherwise, write today's date on the very first row where Column C is populated
  const rangeAC = sheet.getRange(startRow, 1, rowCount, ORDER_TRACKER.ORDER_COLUMN);
  const values = rangeAC.getValues();

  for (let i = 0; i < rowCount; i++) {
    const colAValue = values[i][0];
    const colCValue = values[i][ORDER_TRACKER.ORDER_COLUMN - 1];

    if (!isBlank_(colCValue) && isBlank_(colAValue)) {
      sheet.getRange(startRow + i, 1).setValue(today);
      SpreadsheetApp.flush();
      break; // Only write date for the first row of today
    }
  }
}


/**
 * Installs or validates the 1-minute tracking trigger.
 */
function installOrderTrackingTrigger() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Skrip harus terhubung ke Google Sheet.');
  }

  const trackedSheets = getTrackedSheets_(spreadsheet);
  if (!trackedSheets.length) {
    const expectedMonth = getIndonesianMonthName_(new Date().getMonth());
    throw new Error('Sheet bulan ini (' + expectedMonth + ') tidak ditemukan.');
  }

  const properties = PropertiesService.getDocumentProperties();
  trackedSheets.forEach(function (sheet) {
    restoreLegacyTrackingState_(sheet, properties);
    configureSummaryCell_(sheet);
    installOpenListButton_(sheet);
  });

  removeOrderTrackingColorTriggers_();

  ScriptApp.newTrigger('updateOrderTrackingColors')
    .timeBased()
    .everyMinutes(1)
    .create();

  updateOrderTrackingColors();

  spreadsheet.toast('Pelacak pesanan otomatis aktif (1 menit).', 'Pelacak PGE', 6);
}


/**
 * Primary trigger handler: updates colors across current and prior monthly tabs.
 */
function updateOrderTrackingColors() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Skrip harus terhubung ke Google Sheet.');
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return; // Another run is currently active
  }

  try {
    const nowMs = Date.now();
    const sheets = getTrackedSheets_(spreadsheet);
    sheets.forEach(function (sheet) {
      updateTrackingSheet_(sheet, nowMs);
    });
    SpreadsheetApp.flush();

    // Auto-ensure tracked monthly sheets have the L1 button and K1 formula installed
    try {
      ensureTrackedSheetsConfigured_(spreadsheet);
    } catch (e) {
      // Ignore background button install error
    }
  } finally {
    lock.releaseLock();
  }
}



/* =========================================================
 * K1 SUMMARY, L1 BUTTON SETUP & HOVER NOTES
 * ========================================================= */

function configureSummaryCell_(sheet) {
  const startRow = ORDER_TRACKER.START_ROW;
  const colLetterD = '$D$' + startRow + ':$D';
  const colLetterI = '$I$' + startRow + ':$I';
  
  const summaryCell = sheet.getRange(ORDER_TRACKER.UNFINISHED_SUMMARY_CELL);
  try {
    summaryCell.setFormula('=COUNTIFS(' + colLetterD + '; ">="&TODAY()-1; ' + colLetterD + '; "<"&TODAY()+1; ' + colLetterI + '; "")');
  } catch (e) {
    summaryCell.setFormula('=COUNTIFS(' + colLetterD + ', ">="&TODAY()-1, ' + colLetterD + ', "<"&TODAY()+1, ' + colLetterI + ', "")');
  }

  sheet.getRange(ORDER_TRACKER.BUTTON_ANCHOR_CELL).clearContent();
}


function installOpenListButton_(sheet) {
  sheet.getImages().forEach(function(img) {
    const title = img.getAltTextTitle();
    if (title === ORDER_TRACKER.BUTTON_ALT_TITLE || title === 'buka-tracker') {
      try {
        img.remove();
      } catch (e) {}
    }
  });

  const anchor = sheet.getRange(ORDER_TRACKER.BUTTON_ANCHOR_CELL);
  const anchorColumn = anchor.getColumn();
  const anchorRow = anchor.getRow();

  // Read actual template dimensions so we never break header row height or column width
  const cellWidth = Math.max(80, sheet.getColumnWidth(anchorColumn));
  const cellHeight = Math.max(24, sheet.getRowHeight(anchorRow));

  // Seamlessly style the underlying cell with matching navy background
  anchor.clearContent();
  anchor.setBackground('#1f4e79');

  const bytes = Utilities.base64Decode(ORDER_TRACKER.BUTTON_PNG_BASE64);
  const blob = Utilities.newBlob(bytes, 'image/png', 'buka-tracker.png');

    // Calculate natural button dimensions preserving aspect ratio and centering inside cell L1
  const targetWidth = Math.min(190, Math.max(60, cellWidth - 10));
  const targetHeight = Math.min(48, Math.max(26, cellHeight - 12));
  const offsetX = Math.max(2, Math.floor((cellWidth - targetWidth) / 2));
  const offsetY = Math.max(2, Math.floor((cellHeight - targetHeight) / 2));

  sheet.insertImage(blob, anchorColumn, anchorRow, offsetX, offsetY)
    .setAltTextTitle(ORDER_TRACKER.BUTTON_ALT_TITLE)
    .setAltTextDescription('Buka daftar pesanan yang belum selesai')
    .setWidth(targetWidth)
    .setHeight(targetHeight)
    .assignScript(ORDER_TRACKER.BUTTON_HANDLER);
}


function updateUnfinishedOrdersNote_(sheet, nowMs, timeZone) {
  const summaryCell = sheet.getRange(ORDER_TRACKER.UNFINISHED_SUMMARY_CELL);
  const orders = collectUnfinishedOrders_(sheet, nowMs, timeZone);

  if (!orders.length) {
    summaryCell.setNote('Tidak ada pesanan kemarin atau hari ini yang belum selesai.');
    return;
  }

  const visibleOrders = orders.slice(0, ORDER_TRACKER.HOVER_NOTE_MAX_ITEMS);
  const lines = [
    '=== DAFTAR PESANAN BELUM SELESAI (' + sheet.getName() + ') ==='
  ];

  visibleOrders.forEach(function(order) {
    const desc = order.detail ? order.detail : '(Tanpa detail)';
    const timer = order.elapsedText && order.elapsedText !== '—' ? ' (' + order.elapsedText + ')' : '';
    lines.push('• Baris ' + order.rowNumber + ': ' + desc + ' — ' + order.statusLabel + timer);
  });

  if (orders.length > visibleOrders.length) {
    lines.push('...dan ' + (orders.length - visibleOrders.length) + ' pesanan lainnya. Klik BUKA TRACKER untuk melihat semuanya.');
  }

  summaryCell.setNote(lines.join('\n'));
}



/**
 * Dedicated function to draw and anchor the clickable BUKA TRACKER button image in cell L1.
 * Automatically cleans any old button, sizes column L and row 1, and assigns 'showUnfinishedOrdersSidebar'.
 */
function drawButtonL1() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Skrip harus terhubung ke Google Sheet.');
  }

  const sheet = spreadsheet.getActiveSheet();
  if (!sheet) {
    throw new Error('Sheet aktif tidak ditemukan.');
  }

  installOpenListButton_(sheet);
  configureSummaryCell_(sheet);

  spreadsheet.toast('Tombol BUKA TRACKER berhasil dipasang pas di sel L1!', 'Pelacak PGE', 6);
}


function reinstallButtonOnCurrentSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  if (!sheet) return;

  configureSummaryCell_(sheet);
  installOpenListButton_(sheet);
  updateOrderTrackingColors();

  spreadsheet.toast('Tombol BUKA TRACKER dan catatan K1 berhasil dipasang pada sheet ini.', 'Pelacak PGE', 5);
}


function updateOneTrackingSheet_(sheet) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return;
  }

  try {
    updateTrackingSheet_(sheet, Date.now());
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}


/* =========================================================
 * ROW COLORING LOGIC
 * ========================================================= */

function updateTrackingSheet_(sheet, nowMs) {
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();
  const currentState = buildCurrentTrackingState_(sheet, nowMs, timeZone);
  const properties = PropertiesService.getDocumentProperties();
  const previousState = loadStoredTrackingState_(sheet, properties);

  const lastRow = sheet.getLastRow();
  const finishedRows = new Set();

  // Identify finished orders from Column I (completion timestamp) or Column J (duration)
  if (lastRow >= ORDER_TRACKER.START_ROW) {
    const rowCount = lastRow - ORDER_TRACKER.START_ROW + 1;
    const completionRange = sheet.getRange(
      ORDER_TRACKER.START_ROW,
      ORDER_TRACKER.COMPLETION_COLUMN,
      rowCount,
      2 // Read Col I and Col J
    );
    const completionValues = completionRange.getValues();

    completionValues.forEach(function (rowVals, index) {
      const colI = rowVals[0];
      const colJ = rowVals[1];
      if (!isBlank_(colI) || !isBlank_(colJ)) {
        finishedRows.add(ORDER_TRACKER.START_ROW + index);
      }
    });
  }

  // Set finished rows back to white across configured range
  setRowsBackground_(
    sheet,
    Array.from(finishedRows),
    ORDER_TRACKER.COLORS.WHITE,
    ORDER_TRACKER.URGENT_END_COLUMN
  );

  // Restore rows that left active tracking
  Object.keys(previousState).forEach(function (rowKey) {
    const rowNumber = Number(rowKey);
    if (!currentState.has(rowNumber)) {
      if (finishedRows.has(rowNumber)) {
        sheet.getRange(rowNumber, 1, 1, ORDER_TRACKER.URGENT_END_COLUMN)
          .setBackground(ORDER_TRACKER.COLORS.WHITE);
      } else {
        restoreOriginalRowColor_(sheet, rowNumber, previousState[rowKey]);
      }
      deleteStoredRowState_(sheet, rowNumber, properties);
    }
  });

  const colorGroups = {
    yellow: [],
    orange: [],
    red: [],
    urgent: [],
    video: []
  };

  currentState.forEach(function (status, rowNumber) {
    let savedState = previousState[String(rowNumber)] || null;

    if (!savedState) {
      savedState = {
        status: status,
        backgrounds: captureOriginalRowColor_(sheet, rowNumber)
      };
      saveStoredRowState_(sheet, rowNumber, savedState, properties);
    } else if (savedState.status !== status) {
      savedState.status = status;
      saveStoredRowState_(sheet, rowNumber, savedState, properties);
    }

    if (colorGroups[status]) {
      colorGroups[status].push(rowNumber);
    }
  });

  setRowsBackground_(
    sheet,
    colorGroups.yellow,
    ORDER_TRACKER.COLORS.YELLOW,
    ORDER_TRACKER.NORMAL_END_COLUMN
  );

  setRowsBackground_(
    sheet,
    colorGroups.orange,
    ORDER_TRACKER.COLORS.ORANGE,
    ORDER_TRACKER.NORMAL_END_COLUMN
  );

  setRowsBackground_(
    sheet,
    colorGroups.red,
    ORDER_TRACKER.COLORS.RED,
    ORDER_TRACKER.NORMAL_END_COLUMN
  );

  setRowsBackground_(
    sheet,
    colorGroups.urgent,
    ORDER_TRACKER.COLORS.RED,
    ORDER_TRACKER.URGENT_END_COLUMN
  );

  setRowsBackground_(
    sheet,
    colorGroups.video,
    ORDER_TRACKER.COLORS.PURPLE,
    ORDER_TRACKER.URGENT_END_COLUMN
  );
}


function buildCurrentTrackingState_(sheet, nowMs, timeZone) {
  const state = new Map();
  const lastRow = sheet.getLastRow();

  if (lastRow < ORDER_TRACKER.START_ROW) {
    return state;
  }

  const rowCount = lastRow - ORDER_TRACKER.START_ROW + 1;
  const values = sheet
    .getRange(
      ORDER_TRACKER.START_ROW,
      1,
      rowCount,
      ORDER_TRACKER.ORDER_TYPE_COLUMN
    )
    .getValues();

  const allowedDayKeys = getAllowedDayKeys_(nowMs, timeZone);

  values.forEach(function (row, index) {
    const order = row[ORDER_TRACKER.ORDER_COLUMN - 1];
    const orderDate = getValidDate_(row[ORDER_TRACKER.DATE_COLUMN - 1]);
    const completedAt = row[ORDER_TRACKER.COMPLETION_COLUMN - 1];
    const duration = row[ORDER_TRACKER.DURATION_COLUMN - 1];

    if (isBlank_(order) || !orderDate || !isBlank_(completedAt) || !isBlank_(duration)) {
      return;
    }

    const dayKey = formatDayKey_(orderDate, timeZone);
    if (!allowedDayKeys[dayKey]) {
      return;
    }

    const rowNumber = ORDER_TRACKER.START_ROW + index;
    const orderType = normalizeText_(row[ORDER_TRACKER.ORDER_TYPE_COLUMN - 1]);
    const orderDetail = normalizeText_(row[ORDER_TRACKER.ORDER_COLUMN - 1]);

    // Keperluan Video: special pastel purple, immune to overdue timer alerts
    if (orderType.indexOf('video') !== -1 || orderDetail.indexOf('keperluan video') !== -1) {
      state.set(rowNumber, 'video');
      return;
    }

    // Urgent orders are critical red immediately
    if (orderType === 'urgent') {
      state.set(rowNumber, 'urgent');
      return;
    }

    const processStart = getValidDate_(row[ORDER_TRACKER.PROCESS_START_COLUMN - 1]);
    if (!processStart) {
      return;
    }

    const elapsedMs = nowMs - processStart.getTime();

    if (elapsedMs >= ORDER_TRACKER.THREE_HOURS_MS) {
      state.set(rowNumber, 'red');
    } else if (elapsedMs >= ORDER_TRACKER.TWO_HOURS_MS) {
      state.set(rowNumber, 'orange');
    } else if (elapsedMs >= ORDER_TRACKER.ONE_HOUR_MS) {
      state.set(rowNumber, 'yellow');
    }
  });

  return state;
}


/* =========================================================
 * SIDEBAR & DATA RETRIEVAL
 * ========================================================= */

function showUnfinishedOrdersSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('UnfinishedOrdersSidebar')
    .setTitle('Pesanan Belum Selesai');

  SpreadsheetApp.getUi().showSidebar(html);
}


function openTrackerFromImageButton() {
  showUnfinishedOrdersSidebar();
}


/**
 * Fetches unfinished orders from both current and previous monthly sheets.
 */
function getUnfinishedOrdersForSidebar() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Skrip harus terhubung ke Google Sheet.');
  }

  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const nowMs = Date.now();
  const resolution = resolveTrackedSheetNames_(new Date(nowMs), timeZone);

  const sheets = [];
  const warnings = [];

  const currentSheet = spreadsheet.getSheetByName(resolution.currentMonth);
  if (currentSheet) {
    sheets.push(currentSheet);
  } else {
    throw new Error('Sheet bulan ini (' + resolution.currentMonth + ') tidak ditemukan di workbook.');
  }

  const prevSheet = spreadsheet.getSheetByName(resolution.previousMonth);
  if (prevSheet) {
    sheets.push(prevSheet);
  } else {
    warnings.push('Sheet bulan sebelumnya (' + resolution.previousMonth + ') tidak ditemukan; memantau ' + resolution.currentMonth + '.');
  }

  let allOrders = [];
  sheets.forEach(function (sheet) {
    const sheetOrders = collectUnfinishedOrders_(sheet, nowMs, timeZone);
    allOrders = allOrders.concat(sheetOrders);
  });

  const statusRank = {
    overdue: 0,
    active: 1,
    'invalid-start': 2,
    'not-started': 3
  };

  allOrders.sort(function (a, b) {
    const urgentDifference =
      Number(normalizeText_(b.orderType) === 'urgent') -
      Number(normalizeText_(a.orderType) === 'urgent');

    if (urgentDifference !== 0) {
      return urgentDifference;
    }

    const rankDifference = statusRank[a.statusKind] - statusRank[b.statusKind];
    if (rankDifference !== 0) {
      return rankDifference;
    }

    if (a.elapsedMs !== null && b.elapsedMs !== null) {
      return (b.elapsedMs - a.elapsedMs) || (a.rowNumber - b.rowNumber);
    }

    return a.rowNumber - b.rowNumber;
  });

  return {
    currentMonth: resolution.currentMonth,
    previousMonth: resolution.previousMonth,
    count: allOrders.length,
    orders: allOrders,
    warnings: warnings,
    scopeLabel: 'Menampilkan pesanan kemarin & hari ini dari ' + sheets.map(function(s){ return s.getName(); }).join(', ') + '.',
    updatedAt: Utilities.formatDate(new Date(nowMs), timeZone, 'dd/MM/yyyy HH:mm:ss')
  };
}


function collectUnfinishedOrders_(sheet, nowMs, timeZone) {
  const lastRow = sheet.getLastRow();
  if (lastRow < ORDER_TRACKER.START_ROW) {
    return [];
  }

  const rowCount = lastRow - ORDER_TRACKER.START_ROW + 1;
  const range = sheet.getRange(
    ORDER_TRACKER.START_ROW,
    1,
    rowCount,
    ORDER_TRACKER.ORDER_TYPE_COLUMN
  );

  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const allowedDayKeys = getAllowedDayKeys_(nowMs, timeZone);
  const orders = [];

  values.forEach(function (row, index) {
    const orderValue = row[ORDER_TRACKER.ORDER_COLUMN - 1];
    const orderDate = getValidDate_(row[ORDER_TRACKER.DATE_COLUMN - 1]);
    const completedAt = row[ORDER_TRACKER.COMPLETION_COLUMN - 1];
    const duration = row[ORDER_TRACKER.DURATION_COLUMN - 1];

    if (isBlank_(orderValue) || !orderDate || !isBlank_(completedAt) || !isBlank_(duration)) {
      return;
    }

    const dayKey = formatDayKey_(orderDate, timeZone);
    if (!allowedDayKeys[dayKey]) {
      return;
    }

    const rowNumber = ORDER_TRACKER.START_ROW + index;
    const detail = String(displayValues[index][ORDER_TRACKER.ORDER_COLUMN - 1] || '').trim();
    const orderType = String(displayValues[index][ORDER_TRACKER.ORDER_TYPE_COLUMN - 1] || '').trim();
    const normType = normalizeText_(orderType);
    const isVideo = normType.indexOf('video') !== -1;
    const isUrgent = normType === 'urgent';

    const rawProcessStart = row[ORDER_TRACKER.PROCESS_START_COLUMN - 1];
    const processStart = getValidDate_(rawProcessStart);

    const order = {
      sheetName: sheet.getName(),
      rowNumber: rowNumber,
      detail: detail,
      orderType: orderType,
      dateFormatted: Utilities.formatDate(orderDate, timeZone, 'dd/MM'),
      elapsedMs: null,
      elapsedText: '—',
      scheduleText: '',
      statusKind: '',
      statusLabel: '',
      alertLevel: 'normal'
    };

    if (!processStart) {
      if (!isBlank_(rawProcessStart)) {
        order.statusKind = 'invalid-start';
        order.statusLabel = 'KOLOM F TIDAK VALID';
        order.scheduleText = 'Kolom F harus berisi tanggal dan waktu lengkap.';
      } else {
        if (isVideo) {
          order.statusKind = 'video';
          order.statusLabel = 'KEPERLUAN VIDEO';
          order.scheduleText = 'Pesanan konten internal — keperluan video';
          order.alertLevel = 'video';
        } else {
          order.statusKind = 'not-started';
          order.statusLabel = isUrgent ? 'URGENT — BELUM PRODUKSI' : 'BELUM MASUK PRODUKSI';
          order.scheduleText = isUrgent ? 'Pesanan urgent; kolom F masih kosong.' : 'Print belum selesai; kolom F masih kosong.';
          if (isUrgent) {
            order.alertLevel = 'critical';
          }
        }
      }

      orders.push(order);
      return;
    }

    const elapsedMs = Math.max(0, nowMs - processStart.getTime());
    order.elapsedMs = elapsedMs;
    order.elapsedText = formatElapsedTime_(elapsedMs);
    order.scheduleText = isUrgent ? 'Pesanan urgent — dihitung sejak kolom F.' : 'Dihitung sejak selesai print di kolom F.';

    if (isVideo) {
      order.statusKind = 'video';
      order.statusLabel = 'KEPERLUAN VIDEO (' + order.elapsedText + ')';
      order.scheduleText = 'Pesanan konten internal — keperluan video';
      order.alertLevel = 'video';
    } else if (elapsedMs >= ORDER_TRACKER.THREE_HOURS_MS) {
      order.statusKind = 'overdue';
      order.statusLabel = 'LEBIH 3 JAM';
      order.alertLevel = 'critical';
    } else if (elapsedMs >= ORDER_TRACKER.TWO_HOURS_MS) {
      order.statusKind = 'overdue';
      order.statusLabel = 'LEBIH 2 JAM';
      order.alertLevel = 'high';
    } else if (elapsedMs >= ORDER_TRACKER.ONE_HOUR_MS) {
      order.statusKind = 'overdue';
      order.statusLabel = 'LEBIH 1 JAM';
      order.alertLevel = 'warning';
    } else {
      order.statusKind = 'active';
      order.statusLabel = isUrgent ? 'URGENT — BERJALAN' : 'PRODUKSI BERJALAN';
      if (isUrgent) {
        order.alertLevel = 'critical';
      }
    }

    orders.push(order);
  });

  return orders;
}


/* =========================================================
 * ADMIN RECOVERY & STATE MANAGEMENT
 * ========================================================= */

function resetOrderTrackingColorsOnce() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Skrip harus terhubung ke Google Sheet.');
  }

  const sheets = getTrackedSheets_(spreadsheet);
  if (!sheets.length) {
    throw new Error('Sheet pelacakan bulanan tidak ditemukan.');
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    throw new Error('Sistem sedang memperbarui warna. Coba jalankan kembali dalam beberapa detik.');
  }

  let resetError = null;

  try {
    removeOrderTrackingColorTriggers_();
    const properties = PropertiesService.getDocumentProperties();

    sheets.forEach(function (sheet) {
      clearStoredTrackingStateForSheet_(sheet, properties);
      const lastRow = sheet.getLastRow();
      if (lastRow < ORDER_TRACKER.START_ROW) {
        return;
      }

      const rowCount = lastRow - ORDER_TRACKER.START_ROW + 1;
      sheet.getRange(
        ORDER_TRACKER.START_ROW,
        1,
        rowCount,
        ORDER_TRACKER.URGENT_END_COLUMN
      ).setBackground(ORDER_TRACKER.COLORS.WHITE);
    });

    const nowMs = Date.now();
    sheets.forEach(function (sheet) {
      updateTrackingSheet_(sheet, nowMs);
    });

    SpreadsheetApp.flush();
  } catch (error) {
    resetError = error;
  } finally {
    lock.releaseLock();
    ensureOrderTrackingColorTrigger_();
  }

  if (resetError) {
    throw resetError;
  }

  spreadsheet.toast(
    'Warna dasar telah dipulihkan ke putih dan status aktif dihitung ulang.',
    'Pelacak PGE',
    6
  );
}


function clearStoredTrackingStateForSheet_(sheet, properties) {
  const currentPrefix = ORDER_TRACKER.STATE_PREFIX + sheet.getSheetId() + '_';
  const legacyPrefix = ORDER_TRACKER.LEGACY_STATE_PREFIX + sheet.getSheetId() + '_';
  const allProperties = properties.getProperties();

  Object.keys(allProperties).forEach(function (key) {
    if (key.indexOf(currentPrefix) === 0 || key.indexOf(legacyPrefix) === 0) {
      properties.deleteProperty(key);
    }
  });
}


function captureOriginalRowColor_(sheet, rowNumber) {
  return sheet
    .getRange(rowNumber, 1, 1, ORDER_TRACKER.URGENT_END_COLUMN)
    .getBackgrounds()[0];
}


function restoreOriginalRowColor_(sheet, rowNumber, savedState) {
  if (
    !savedState ||
    !Array.isArray(savedState.backgrounds) ||
    savedState.backgrounds.length < ORDER_TRACKER.URGENT_END_COLUMN ||
    rowNumber < ORDER_TRACKER.START_ROW ||
    rowNumber > sheet.getMaxRows()
  ) {
    return;
  }

  sheet
    .getRange(rowNumber, 1, 1, ORDER_TRACKER.URGENT_END_COLUMN)
    .setBackgrounds([
      savedState.backgrounds.slice(0, ORDER_TRACKER.URGENT_END_COLUMN)
    ]);
}


function setRowsBackground_(sheet, rows, color, endColumn) {
  const uniqueRows = Array.from(new Set(rows)).filter(function (row) {
    return row >= ORDER_TRACKER.START_ROW && row <= sheet.getMaxRows();
  });

  if (!uniqueRows.length) {
    return;
  }

  const endLetter = columnNumberToLetter_(endColumn);
  const ranges = uniqueRows.map(function (row) {
    return 'A' + row + ':' + endLetter + row;
  });

  sheet.getRangeList(ranges).setBackground(color);
}


function loadStoredTrackingState_(sheet, properties) {
  const prefix = getSheetStatePrefix_(sheet);
  const allProperties = properties.getProperties();
  const state = {};

  Object.keys(allProperties).forEach(function (key) {
    if (key.indexOf(prefix) !== 0) {
      return;
    }

    const rowNumber = Number(key.slice(prefix.length));
    if (!Number.isInteger(rowNumber)) {
      return;
    }

    try {
      const parsed = JSON.parse(allProperties[key]);
      if (parsed && typeof parsed === 'object') {
        state[String(rowNumber)] = parsed;
      }
    } catch (error) {
      properties.deleteProperty(key);
    }
  });

  return state;
}


function saveStoredRowState_(sheet, rowNumber, state, properties) {
  properties.setProperty(
    getRowStateKey_(sheet, rowNumber),
    JSON.stringify(state)
  );
}


function deleteStoredRowState_(sheet, rowNumber, properties) {
  properties.deleteProperty(getRowStateKey_(sheet, rowNumber));
}


function getSheetStatePrefix_(sheet) {
  return ORDER_TRACKER.STATE_PREFIX + sheet.getSheetId() + '_';
}


function getRowStateKey_(sheet, rowNumber) {
  return getSheetStatePrefix_(sheet) + rowNumber;
}


function restoreLegacyTrackingState_(sheet, properties) {
  const prefix = ORDER_TRACKER.LEGACY_STATE_PREFIX + sheet.getSheetId() + '_';
  const allProperties = properties.getProperties();

  Object.keys(allProperties).forEach(function (key) {
    if (key.indexOf(prefix) !== 0) {
      return;
    }

    const rowNumber = Number(key.slice(prefix.length));
    try {
      const savedState = JSON.parse(allProperties[key]);
      if (
        Number.isInteger(rowNumber) &&
        rowNumber >= ORDER_TRACKER.START_ROW &&
        rowNumber <= sheet.getMaxRows() &&
        savedState &&
        Array.isArray(savedState.backgrounds) &&
        savedState.backgrounds.length >= ORDER_TRACKER.URGENT_END_COLUMN
      ) {
        sheet
          .getRange(rowNumber, 1, 1, ORDER_TRACKER.URGENT_END_COLUMN)
          .setBackgrounds([
            savedState.backgrounds.slice(0, ORDER_TRACKER.URGENT_END_COLUMN)
          ]);
      }
    } catch (error) {
      console.warn('Status warna lama pada baris ' + rowNumber + ' tidak dapat dibaca.');
    }

    properties.deleteProperty(key);
  });
}


function removeOrderTrackingColorTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'updateOrderTrackingColors') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}


function ensureOrderTrackingColorTrigger_() {
  removeOrderTrackingColorTriggers_();
  ScriptApp.newTrigger('updateOrderTrackingColors')
    .timeBased()
    .everyMinutes(1)
    .create();
}


/* =========================================================
 * DATE, MONTH AND RESOLUTION HELPERS
 * ========================================================= */

function getIndonesianMonthName_(monthIndexZeroBased) {
  const safeIndex = (monthIndexZeroBased % 12 + 12) % 12;
  return ORDER_TRACKER.MONTH_NAMES_ID[safeIndex];
}


/**
 * Resolves current and previous month uppercase tab names based on reference date and timezone.
 */
function resolveTrackedSheetNames_(dateObj, timeZone) {
  const currentMonthStr = Utilities.formatDate(dateObj, timeZone, 'MM');
  const currentMonthIndex = parseInt(currentMonthStr, 10) - 1;
  const prevMonthIndex = (currentMonthIndex - 1 + 12) % 12;

  return {
    currentMonth: getIndonesianMonthName_(currentMonthIndex),
    previousMonth: getIndonesianMonthName_(prevMonthIndex)
  };
}


function getTrackedSheets_(spreadsheet) {
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const names = resolveTrackedSheetNames_(new Date(), timeZone);
  const sheets = [];

  const curSheet = spreadsheet.getSheetByName(names.currentMonth);
  if (curSheet) {
    sheets.push(curSheet);
  }

  const prevSheet = spreadsheet.getSheetByName(names.previousMonth);
  if (prevSheet) {
    sheets.push(prevSheet);
  }

  return sheets;
}


function isTrackedSheet_(sheet) {
  if (!sheet) {
    return false;
  }
  const spreadsheet = sheet.getParent();
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const names = resolveTrackedSheetNames_(new Date(), timeZone);
  const sheetName = sheet.getName();

  return sheetName === names.currentMonth || sheetName === names.previousMonth;
}


function getAllowedDayKeys_(nowMs, timeZone) {
  const todayKey = Utilities.formatDate(new Date(nowMs), timeZone, 'yyyy-MM-dd');
  const yesterdayKey = shiftIsoDate_(todayKey, -1);

  const result = {};
  result[todayKey] = true;
  result[yesterdayKey] = true;
  return result;
}


function shiftIsoDate_(isoDate, dayOffset) {
  const parts = isoDate.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + dayOffset));

  return [
    date.getUTCFullYear(),
    String(date.getUTCMonth() + 1).padStart(2, '0'),
    String(date.getUTCDate()).padStart(2, '0')
  ].join('-');
}


function formatDayKey_(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
}


function getValidDate_(value) {
  if (!(value instanceof Date)) {
    return null;
  }
  const timestamp = value.getTime();
  if (Number.isNaN(timestamp) || value.getFullYear() < 2000) {
    return null;
  }
  return value;
}


function formatElapsedTime_(elapsedMs) {
  const totalSeconds = Math.floor(elapsedMs / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  return (
    hours +
    ':' +
    String(minutes).padStart(2, '0') +
    ':' +
    String(seconds).padStart(2, '0')
  );
}


function columnNumberToLetter_(columnNumber) {
  let value = columnNumber;
  let result = '';

  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }

  return result;
}


function normalizeText_(value) {
  return String(value || '').trim().toLowerCase();
}


function isBlank_(value) {
  return (
    value === null ||
    value === undefined ||
    value === '' ||
    (typeof value === 'string' && value.trim() === '')
  );
}


/* =========================================================
 * ON-SHEET BUTTON HANDLERS / ALIASES
 * Use any of these names when assigning a script to the
 * 'BUKA TRACKER' drawing/button on your Google Sheet.
 * ========================================================= */

function openTrackerFromImageButton() {
  showUnfinishedOrdersSidebar();
}

function bukaTracker() {
  showUnfinishedOrdersSidebar();
}

function openTracker() {
  showUnfinishedOrdersSidebar();
}

function showSidebar() {
  showUnfinishedOrdersSidebar();
}

function bukaSidebar() {
  showUnfinishedOrdersSidebar();
}

function bukaPesananBelumSelesai() {
  showUnfinishedOrdersSidebar();
}

function tampilkanPesananBelumSelesai() {
  showUnfinishedOrdersSidebar();
}

function openUnfinishedOrdersSidebar() {
  showUnfinishedOrdersSidebar();
}


/**
 * Opens the tracker in a floating modeless window with comfortable width (420px).
 */
function showUnfinishedOrdersDialog() {
  const html = HtmlService.createHtmlOutputFromFile('UnfinishedOrdersSidebar')
    .setWidth(420)
    .setHeight(680)
    .setTitle('Pesanan Belum Selesai - PGE');

  SpreadsheetApp.getUi().showModelessDialog(html, 'Pesanan Belum Selesai');
}



/* =========================================================
 * RAPID IMAGE UPLOADER (COLUMN C MODAL)
 * ========================================================= */

/**
 * Opens the rapid modeless Image Upload dialog.
 */
function showImageUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ImageUploadDialog')
    .setWidth(820)
    .setHeight(720)
    .setTitle('Upload Foto Desain');

  SpreadsheetApp.getUi().showModelessDialog(html, 'Upload Foto Desain');
}


/**
 * Returns current active sheet and selected row number for the modal.
 */
function getActiveRowContext() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) return null;
  const sheet = spreadsheet.getActiveSheet();
  if (!sheet) return null;

  const totalLastRow = sheet.getLastRow();
  let lastOrderRow = 0;

  // Scan Column D (Waktu Order) downwards to find the true last order row
  if (totalLastRow >= ORDER_TRACKER.START_ROW) {
    const rowCount = totalLastRow - ORDER_TRACKER.START_ROW + 1;
    const colDValues = sheet.getRange(ORDER_TRACKER.START_ROW, ORDER_TRACKER.DATE_COLUMN, rowCount, 1).getValues();

    for (let i = colDValues.length - 1; i >= 0; i--) {
      const val = colDValues[i][0];
      if (!isBlank_(val)) {
        lastOrderRow = ORDER_TRACKER.START_ROW + i;
        break;
      }
    }
  }

  const nextEmptyRow = lastOrderRow > 0 ? (lastOrderRow + 1) : ORDER_TRACKER.START_ROW;

  return {
    sheetName: sheet.getName(),
    lastRow: lastOrderRow > 0 ? lastOrderRow : totalLastRow,
    nextRow: nextEmptyRow,
    activeRow: nextEmptyRow // Automatically target the next empty row!
  };
}


/**
 * Uploads base64 image data to Google Drive, embeds it into Column C,
 * and automatically triggers Column D (Order Date) and Column A (Tgl).
 */
function uploadOrderImageToCell(sheetName, rowNumber, base64Data, mimeType, detailText) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getActiveSheet();
    if (!sheet) {
      return { success: false, error: 'Sheet "' + sheetName + '" tidak ditemukan.' };
    }

    const row = parseInt(rowNumber, 10);
    if (isNaN(row) || row < ORDER_TRACKER.START_ROW) {
      return { success: false, error: 'Nomor baris tidak valid (minimal baris 5).' };
    }

    // 1. Get or create dedicated Google Drive folder
    const folderName = 'PGE_Foto_Pesanan';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    // 2. Decode and save file to Drive
    const bytes = Utilities.base64Decode(base64Data);
    const fileName = 'Foto_Baris_' + row + '_' + Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmmss') + '.jpg';
    const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName);
    const file = folder.createFile(blob);

    // Make file viewable so =IMAGE() formula can render
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // Ignore if domain policy limits sharing
    }

    const fileId = file.getId();
    const imageUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;

    // 3. Set Column C (Col 3) strictly inside the cell (as CellImage or formula)
    const cellC = sheet.getRange(row, ORDER_TRACKER.ORDER_COLUMN);
    try {
      const cellImage = SpreadsheetApp.newCellImage()
        .setSourceUrl(imageUrl)
        .setAltTextTitle('Foto Desain Baris ' + row)
        .build();
      cellC.setValue(cellImage);
    } catch (e) {
      cellC.setFormula('=IMAGE("' + imageUrl + '"; 1)');
    }

    // Adjust Column C width (175px) and row height (95px) for balanced, crisp proportion
    if (sheet.getColumnWidth(ORDER_TRACKER.ORDER_COLUMN) < 175) {
      sheet.setColumnWidth(ORDER_TRACKER.ORDER_COLUMN, 175);
    }
    if (sheet.getRowHeight(row) < 95) {
      sheet.setRowHeight(row, 95);
    }

    // 4. Set Keterangan directly into Column L (Column 12)
    if (detailText && detailText.trim()) {
      sheet.getRange(row, 12).setValue(detailText.trim());
    }

    // 5. Column D is managed by sheet formulas; preserve formula completely
    // (Do not write anything into Column D)

    // 6. Trigger Column A (TGL) if first order of the day
    autoFillDateColumnA_(sheet, row, row);

    // 6. Update tracking background colors
    SpreadsheetApp.flush();
    updateOneTrackingSheet_(sheet);

    return {
      success: true,
      nextRow: row + 1,
      sheetName: sheet.getName()
    };
  } catch (error) {
    return {
      success: false,
      error: error.message || String(error)
    };
  }
}



/**
 * Automatically ensures that the current and previous month sheets have:
 * 1. Clickable BUKA TRACKER image button in cell L1 (with script assigned)
 * 2. Unfinished orders counter formula in cell K1
 */
function ensureTrackedSheetsConfigured_(optSpreadsheet) {
  const spreadsheet = optSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) return;

  const sheets = getTrackedSheets_(spreadsheet);
  sheets.forEach(function (sheet) {
    try {
      const images = sheet.getImages();
      let hasButton = false;
      for (let i = 0; i < images.length; i++) {
        const title = images[i].getAltTextTitle();
        if (title === ORDER_TRACKER.BUTTON_ALT_TITLE || title === 'buka-tracker') {
          hasButton = true;
          break;
        }
      }

      if (!hasButton) {
        configureSummaryCell_(sheet);
        installOpenListButton_(sheet);
      }
    } catch (e) {
      console.warn('Gagal cek/pasang tombol di sheet ' + sheet.getName() + ': ' + e.message);
    }
  });
}


function getSeptemberLastRowInfo() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const currentMonthName = getIndonesianMonthName_(new Date().getMonth());
  const sheet = spreadsheet.getSheetByName(currentMonthName) || spreadsheet.getSheetByName('SEPTEMBER') || spreadsheet.getActiveSheet();
  if (!sheet) {
    return JSON.stringify({ error: 'Sheet tidak ditemukan.' });
  }
  const lastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  
  let lastDataRow = 0;
  if (lastRow >= ORDER_TRACKER.START_ROW) {
    const rowCount = lastRow - ORDER_TRACKER.START_ROW + 1;
    const values = sheet.getRange(ORDER_TRACKER.START_ROW, 1, rowCount, 12).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      const row = values[i];
      const hasContent = row.some(function(cell) { return cell !== '' && cell !== null && cell !== undefined; });
      if (hasContent) {
        lastDataRow = ORDER_TRACKER.START_ROW + i;
        break;
      }
    }
  }

  const result = {
    sheetName: sheet.getName(),
    sheetGetLastRow: lastRow,
    lastDataRow: lastDataRow,
    nextNewRow: Math.max(ORDER_TRACKER.START_ROW, lastRow + 1),
    maxRows: maxRows
  };
  Logger.log(JSON.stringify(result));
  return JSON.stringify(result);
}



/**
 * User-invocable menu item to check and display the exact last row of the current active sheet.
 */
function showLastRowAlert() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const currentMonthName = getIndonesianMonthName_(new Date().getMonth());
  const sheet = spreadsheet.getActiveSheet() || spreadsheet.getSheetByName(currentMonthName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet tidak ditemukan.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  let lastDataRow = 0;

  if (lastRow >= ORDER_TRACKER.START_ROW) {
    const rowCount = lastRow - ORDER_TRACKER.START_ROW + 1;
    const values = sheet.getRange(ORDER_TRACKER.START_ROW, 1, rowCount, 12).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      const row = values[i];
      const hasContent = row.some(function(cell) { return cell !== '' && cell !== null && cell !== undefined; });
      if (hasContent) {
        lastDataRow = ORDER_TRACKER.START_ROW + i;
        break;
      }
    }
  }

  const nextEmpty = Math.max(ORDER_TRACKER.START_ROW, lastRow + 1);
  const msg = '📊 STATUS SHEET: ' + sheet.getName() + '\n\n' +
              '• Baris Terakhir Terisi (Data): Baris ' + lastDataRow + '\n' +
              '• Baris Terakhir Terbaca (getLastRow): Baris ' + lastRow + '\n' +
              '• Baris Kosong Baru Berikutnya: Baris ' + nextEmpty + '\n' +
              '• Total Kapasitas Baris Sheet: ' + maxRows + ' Baris';

  SpreadsheetApp.getUi().alert('Info Baris Sheet', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// Backward-compatibility alias
function showSeptemberLastRowAlert() {
  showLastRowAlert();
}
