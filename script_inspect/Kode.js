/**
 * PGE — Barang Baru Automation with Manual Verification
 *
 * SOURCE WORKBOOK
 *   Sheet: Barang Baru
 *   PIC row: 2
 *   Header row: 3
 *   First product row: 4
 *   Warna Produk is inserted directly after Nama Barang.
 *
 * DESTINATION MAPPING
 *   Design1 → A: NEW LINDA, B: product name, starting row 3
 *   DAdmin  → B: NEW LINDA, C: product name, starting row 3
 *   Video   → B: automatic sent date, C: product name, starting row 4
 *   NEW_LINDA → A: product name, B: color, C: HPP, starting row 2
 *
 * FINAL FLOW
 *   1. A manual edit completes Upload.
 *   2. Status becomes "Menunggu Konfirmasi".
 *   3. User selects the product row and opens:
 *      Barang Baru → Verifikasi & Kirim
 *   4. The script checks the product name live in all destination tabs.
 *   5. The dialog shows which destinations already contain the product and
 *      lets the user select only destinations that are missing/incomplete.
 *   6. Only selected destinations are processed. Existing complete destinations
 *      are disabled and never duplicated.
 *   7. Status becomes "Selesai" only when all four destinations are complete.
 *   8. Closing the dialog makes no changes.
 */

const BARANG_BARU_CONFIG = Object.freeze({
  sourceSheetName: "Barang Baru",
  picRow: 2,
  headerRow: 3,
  firstDataRow: 4,

  processHeader: "Proses Saat Ini",
  productColorHeader: "Warna Produk",
  logSheetName: "Log Distribusi",
  verifiedMarkerPrefix: "PGE_VERIFIED_PRODUCT::",
  verifiedSnapshotPrefix: "PGE_VERIFIED_SNAPSHOT::",

  productColorOptions: Object.freeze([
    "Hitam",
    "Putih",
    "Merah",
    "Biru",
    "Hijau",
    "Kuning",
    "Cokelat",
    "Abu-abu",
    "Krem",
    "Navy",
    "Pink",
    "Ungu",
    "Oranye",
    "Gold",
    "Silver",
    "Custom",
  ]),

  destinations: Object.freeze([
    Object.freeze({
      key: "DESIGN1",
      spreadsheetId: "1V8iCpaABrVLB0iMEAe_58xFEd8puCbx2jwDLQujiGHI",
      sheetName: "Design1",
      productColumn: 2,
      labelColumn: 1,
      labelValue: "NEW LINDA",
      colorColumn: null,
      hppColumn: null,
      firstDataRow: 3,
      appendAfterLastProduct: true,
    }),
    Object.freeze({
      key: "DADMIN",
      spreadsheetId: "1V8iCpaABrVLB0iMEAe_58xFEd8puCbx2jwDLQujiGHI",
      sheetName: "DAdmin",
      productColumn: 3,
      labelColumn: 2,
      labelValue: "NEW LINDA",
      colorColumn: null,
      hppColumn: null,
      firstDataRow: 3,
      appendAfterLastProduct: true,
    }),
    Object.freeze({
      key: "VIDEO",
      spreadsheetId: "1V8iCpaABrVLB0iMEAe_58xFEd8puCbx2jwDLQujiGHI",
      sheetName: "Video",
      dateColumn: 2,
      productColumn: 3,
      labelColumn: null,
      labelValue: "",
      colorColumn: null,
      hppColumn: null,
      headerRow: 3,
      firstDataRow: 4,
      appendAfterLastProduct: true,
      clearCopiedBackground: true,
      expectedHeaders: Object.freeze([
        Object.freeze({
          column: 2,
          names: Object.freeze(["Yang Request Tanggal"]),
        }),
        Object.freeze({
          column: 3,
          names: Object.freeze(["Nama Barang"]),
        }),
      ]),
    }),
    Object.freeze({
      key: "NEW_LINDA",
      spreadsheetId: "12yEx2Abl05tiLWlxGwp2McidQ6L5nIUpwczQvU9G8HQ",
      sheetName: "NEW_LINDA",
      productColumn: 1,
      labelColumn: null,
      labelValue: "",
      colorColumn: 2,
      hppColumn: 3,
      firstDataRow: 2,
      expectedHeaders: Object.freeze([
        Object.freeze({ column: 1, names: Object.freeze(["Nama"]) }),
        Object.freeze({ column: 2, names: Object.freeze(["Warna"]) }),
        Object.freeze({ column: 3, names: Object.freeze(["HPP"]) }),
      ]),
    }),
  ]),

  logHeaders: Object.freeze([
    "Waktu",
    "Nama Produk",
    "Tab Tujuan",
    "Hasil",
    "Baris Tujuan",
    "Pemicu",
    "Pesan",
    "Baris Sumber",
  ]),

  colors: Object.freeze({
    EMPTY: "#FFFFFF",
    RESEARCH: "#D9EAF7",
    PRODUCT_COLOR: "#E2F0D9",
    PATTERN: "#CFE2F3",
    MATERIAL_SAMPLE: "#B6D7E8",
    PATTERN_LAYOUT: "#FFF2CC",
    WAITING_APPROVAL: "#EADCF8",
    REVISION: "#F4CCCC",
    BLUEPRINT: "#D9D2E9",
    HPP: "#FCE5CD",
    ODOO: "#F9CB9C",
    SELLING_PRICE: "#D9EAD3",
    UPLOAD: "#D0EAE5",
    WAITING_CONFIRMATION: "#FCE8B2",
    COMPLETE: "#B7E1CD",
    CANCELLED: "#B7B7B7",
  }),
});

/* =========================================================
 * MENU
 * ======================================================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Barang Baru")
    .addItem("Setup sistem & trigger", "setupBarangBaruSystem")
    .addItem("Periksa trigger & koneksi", "checkBarangBaruSystem")
    .addItem("Refresh progress", "refreshBarangBaruProgress")
    .addSeparator()
    .addItem("Verifikasi & Kirim", "verifyAndSendActiveProduct")
    .addItem("Buka log distribusi", "openDistributionLog")
    .addSeparator()
    .addItem("Petunjuk", "showBarangBaruInstructions")
    .addToUi();
}

/* =========================================================
 * SETUP AND AUTOMATIC STATUS UPDATE
 * ======================================================= */

function setupBarangBaruSystem() {
  const sheet = getBarangBaruSheet_();

  ensureProcessColumn_(sheet);
  ensureProductColorColumn_(sheet);
  styleProcessColumn_(sheet);
  styleProductColorColumn_(sheet);
  ensureProductColorDropdown_(sheet);
  ensureDistributionLogSheet_(sheet.getParent());

  // Preserve rows successfully distributed by the previous script version.
  migrateLegacyDistributionNotes_(sheet);

  installEditTrigger_();
  refreshBarangBaruProgress();

  const connection = checkDestinationConnection_();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Trigger aktif dan " +
      connection.checks.length +
      " tab tujuan berhasil ditemukan.",
    "Setup Berhasil",
    8,
  );
}

/**
 * Installable edit-trigger handler.
 *
 * This trigger updates only the source-row status and color.
 * It never sends products to the destination workbook.
 *
 * Do not run this function manually.
 */
function handleBarangBaruEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() !== BARANG_BARU_CONFIG.sourceSheetName) {
    return;
  }

  if (!hasProcessColumn_(sheet)) return;

  const endRow = e.range.getLastRow();

  if (endRow < BARANG_BARU_CONFIG.firstDataRow) {
    return;
  }

  const startRow = Math.max(e.range.getRow(), BARANG_BARU_CONFIG.firstDataRow);

  try {
    const result = updateRows_(sheet, startRow, endRow);

    if (
      result.statuses.some(function (status) {
        return status === "Menunggu Konfirmasi";
      })
    ) {
      e.source.toast(
        "Produk belum dikirim. Pilih barisnya lalu gunakan " +
          "Barang Baru → Verifikasi & Kirim.",
        "Menunggu Konfirmasi",
        8,
      );
    }
  } catch (error) {
    recordTriggerFailure_(sheet, startRow, endRow, error);

    e.source.toast(
      "Status gagal diperbarui: " +
        cleanText_(error.message) +
        ". Periksa Log Distribusi.",
      "Sistem Barang Baru",
      10,
    );

    throw error;
  }
}

/* =========================================================
 * USER COMMANDS
 * ======================================================= */

function refreshBarangBaruProgress() {
  const sheet = getBarangBaruSheet_();

  if (!hasProcessColumn_(sheet)) {
    throw new Error(
      'Kolom "Proses Saat Ini" belum tersedia. ' +
        "Jalankan Barang Baru → Setup sistem & trigger.",
    );
  }

  const lastRow = Math.max(sheet.getLastRow(), BARANG_BARU_CONFIG.firstDataRow);

  updateRows_(sheet, BARANG_BARU_CONFIG.firstDataRow, lastRow);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Progress diperbarui. Tidak ada produk yang dikirim.",
    "Refresh Progress",
    5,
  );
}

function checkBarangBaruSystem() {
  const triggers = ScriptApp.getProjectTriggers().filter(function (trigger) {
    return trigger.getHandlerFunction() === "handleBarangBaruEdit";
  });

  let message =
    "Trigger edit: " + (triggers.length > 0 ? "AKTIF" : "TIDAK DITEMUKAN");

  try {
    const connection = checkDestinationConnection_();

    connection.checks.forEach(function (check) {
      message +=
        "\n" +
        check.fileName +
        " / " +
        check.sheetName +
        " — " +
        check.mapping +
        "; baris kosong berikutnya: " +
        check.nextRow;
    });

    message +=
      "\n\nKoneksi baca berhasil.\n" +
      "Pengiriman hanya dilakukan melalui " +
      '"Verifikasi & Kirim".';
  } catch (error) {
    message += "\n\nKoneksi tujuan GAGAL:\n" + error.message;

    SpreadsheetApp.getUi().alert(
      "Pemeriksaan Sistem",
      message,
      SpreadsheetApp.getUi().ButtonSet.OK,
    );

    throw error;
  }

  SpreadsheetApp.getUi().alert(
    "Pemeriksaan Sistem",
    message,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function verifyAndSendActiveProduct() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const activeRange = SpreadsheetApp.getActiveRange();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== BARANG_BARU_CONFIG.sourceSheetName) {
    ui.alert("Pilih satu sel pada tab Barang Baru terlebih dahulu.");
    return;
  }

  if (!activeRange || activeRange.getRow() < BARANG_BARU_CONFIG.firstDataRow) {
    ui.alert("Pilih satu sel pada baris produk terlebih dahulu.");
    return;
  }

  const sourceRow = activeRange.getRow();
  const columns = getColumns_(sheet);

  // Recalculate before deciding whether the row is eligible.
  const statusResult = updateRows_(sheet, sourceRow, sourceRow);
  let currentStatus = statusResult.statuses[0];
  const product = getSourceProduct_(sheet, sourceRow, columns);

  if (!product.productName) {
    ui.alert("Nama Barang pada baris " + sourceRow + " masih kosong.");
    return;
  }

  if (!product.productColor) {
    ui.alert(
      "Warna Produk Belum Diisi",
      'Pilih minimal satu warna untuk "' +
        product.productName +
        '" pada kolom Warna Produk.',
      ui.ButtonSet.OK,
    );
    return;
  }

  if (currentStatus !== "Menunggu Konfirmasi" && currentStatus !== "Selesai") {
    ui.alert(
      "Produk Belum Siap",
      '"' +
        product.productName +
        '" masih berada pada proses:\n\n' +
        currentStatus +
        "\n\nLengkapi proses tersebut terlebih dahulu.",
      ui.ButtonSet.OK,
    );
    return;
  }

  const destinationStates = inspectDestinationStates_(product);

  const dialogData = buildDestinationDialogData_(
    product,
    destinationStates,
    currentStatus,
  );

  const template = HtmlService.createTemplateFromFile("DestinationSelector");
  template.initialData = JSON.stringify(dialogData).replace(/</g, "\\u003c");

  const html = template
    .evaluate()
    .setWidth(880)
    .setHeight(620);

  ui.showModalDialog(html, "Verifikasi & Kirim");
}

function getSourceProduct_(sheet, sourceRow, columns) {
  const hppCell = sheet.getRange(sourceRow, columns.HPP);

  return {
    productName: cleanText_(
      sheet.getRange(sourceRow, columns.PRODUCT_NAME).getDisplayValue(),
    ),
    productColor: cleanText_(
      sheet.getRange(sourceRow, columns.PRODUCT_COLOR).getDisplayValue(),
    ),
    hpp: hppCell.getValue(),
    hppDisplay: cleanText_(hppCell.getDisplayValue()),
    hppNumberFormat: hppCell.getNumberFormat(),
    sourceRow: sourceRow,
  };
}

function buildDestinationDialogData_(product, destinationStates, currentStatus) {
  return {
    sourceRow: product.sourceRow,
    productName: product.productName,
    productColor: product.productColor,
    hppDisplay: product.hppDisplay,
    currentStatus: currentStatus,
    destinations: destinationStates,
    allComplete: allDestinationsComplete_(destinationStates),
    canCancelUpload: canCancelUploadFromStates_(destinationStates),
  };
}

/**
 * Called from DestinationSelector.html.
 * Re-reads the source row and destination sheets before writing anything.
 */
function sendSelectedDestinationsFromDialog(request) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getBarangBaruSheet_();
  const sourceRow = Number(request && request.sourceRow);
  const requestedProductName = cleanText_(request && request.productName);
  const selectedKeys = uniqueStrings_(
    request && Array.isArray(request.selectedKeys) ? request.selectedKeys : [],
  );

  if (
    !sourceRow ||
    sourceRow < BARANG_BARU_CONFIG.firstDataRow ||
    sourceRow > sheet.getMaxRows()
  ) {
    throw new Error("Baris produk tidak valid. Tutup dialog lalu coba lagi.");
  }

  const columns = getColumns_(sheet);
  let currentStatus = updateRows_(sheet, sourceRow, sourceRow).statuses[0];
  const product = getSourceProduct_(sheet, sourceRow, columns);

  if (
    requestedProductName &&
    normalizedProductName_(requestedProductName) !==
      normalizedProductName_(product.productName)
  ) {
    throw new Error(
      "Nama produk pada baris sumber sudah berubah. Tutup dialog lalu buka Verifikasi & Kirim lagi.",
    );
  }

  if (!product.productName || !product.productColor) {
    throw new Error(
      "Nama Barang atau Warna Produk sudah berubah/kosong. Tutup dialog lalu periksa baris sumber.",
    );
  }

  if (currentStatus !== "Menunggu Konfirmasi" && currentStatus !== "Selesai") {
    throw new Error(
      'Produk sekarang berada pada proses "' +
        currentStatus +
        '". Lengkapi proses tersebut terlebih dahulu.',
    );
  }

  const beforeStates = inspectDestinationStates_(product);
  const stateByKey = {};
  beforeStates.forEach(function (state) {
    stateByKey[state.key] = state;
  });

  const validDestinationKeys = {};
  BARANG_BARU_CONFIG.destinations.forEach(function (destination) {
    validDestinationKeys[destination.key] = true;
  });

  selectedKeys.forEach(function (key) {
    if (!validDestinationKeys[key]) {
      throw new Error('Tujuan "' + key + '" tidak dikenal oleh sistem.');
    }
  });

  const sendableKeys = selectedKeys.filter(function (key) {
    const state = stateByKey[key];
    return state && state.status !== "ERROR";
  });

  const blockedSelected = selectedKeys.filter(function (key) {
    const state = stateByKey[key];
    return state && state.status === "ERROR";
  });

  if (blockedSelected.length > 0) {
    throw new Error(
      "Ada tujuan yang tidak dapat diakses atau error: " +
        blockedSelected
          .map(function (key) {
            return stateByKey[key].sheetName;
          })
          .join(", ") +
        ". Tutup dialog, perbaiki koneksi tujuan, lalu coba lagi.",
    );
  }

  if (sendableKeys.length === 0) {
    throw new Error(
      "Pilih minimal satu tujuan yang ingin dikirim atau diperbarui.",
    );
  }

  const result = distributeVerifiedProduct_(
    product,
    spreadsheet,
    "VERIFIKASI MENU",
    sendableKeys,
  );

  // Re-check all destinations after sending. This live check determines
  // whether the source row may become Selesai.
  const afterStates = inspectDestinationStates_(product);
  const allComplete = allDestinationsComplete_(afterStates);

  writeSourceDistributionNote_(
    sheet,
    columns.PROCESS,
    product,
    result,
    allComplete,
  );

  const finalStatus = updateRows_(sheet, sourceRow, sourceRow).statuses[0];

  let message;

  if (allComplete) {
    message =
      "Pengiriman selesai. Produk sekarang tersedia di seluruh tujuan dan status menjadi Selesai.";
  } else if (result.failed > 0) {
    message =
      "Sebagian pengiriman gagal. Produk tetap Menunggu Konfirmasi. Periksa tujuan yang bermasalah.";
  } else {
    message =
      "Tujuan yang dipilih sudah diproses. Produk tetap Menunggu Konfirmasi karena masih ada tujuan yang belum lengkap.";
  }

  return buildDialogResponse_(product, afterStates, result, message, finalStatus);
}

/**
 * Preserves the old NO behavior, but only while the product has not appeared
 * in any destination. Once a destination already contains the product, clearing
 * Upload would not undo that send, so cancellation is intentionally blocked.
 */
function cancelUploadFromDestinationDialog(request) {
  const sheet = getBarangBaruSheet_();
  const sourceRow = Number(request && request.sourceRow);
  const requestedProductName = cleanText_(request && request.productName);

  if (!sourceRow || sourceRow < BARANG_BARU_CONFIG.firstDataRow) {
    throw new Error("Baris produk tidak valid.");
  }

  const columns = getColumns_(sheet);
  const product = getSourceProduct_(sheet, sourceRow, columns);

  if (
    requestedProductName &&
    normalizedProductName_(requestedProductName) !==
      normalizedProductName_(product.productName)
  ) {
    throw new Error(
      "Nama produk pada baris sumber sudah berubah. Tutup dialog lalu coba lagi.",
    );
  }

  const destinationStates = inspectDestinationStates_(product);

  if (!canCancelUploadFromStates_(destinationStates)) {
    throw new Error(
      "Status Odoo tidak dibatalkan karena produk sudah ditemukan di salah satu tujuan atau ada tujuan yang tidak dapat diperiksa. Pengiriman yang sudah terjadi tidak boleh dianggap terbatalkan.",
    );
  }

  cancelFinalUploadStep_(sheet, sourceRow, columns);

  return {
    cancelled: true,
    status: "Update Odoo — Oliv",
    message: "Status Odoo dibatalkan. Tidak ada data tujuan yang diubah.",
  };
}

function buildDialogResponse_(
  product,
  destinationStates,
  result,
  message,
  finalStatus,
) {
  return {
    sourceRow: product.sourceRow,
    productName: product.productName,
    productColor: product.productColor,
    hppDisplay: product.hppDisplay,
    destinations: destinationStates,
    allComplete: allDestinationsComplete_(destinationStates),
    canCancelUpload: canCancelUploadFromStates_(destinationStates),
    finalStatus: finalStatus || "Menunggu Konfirmasi",
    distribution: {
      inserted: result.inserted,
      skipped: result.skipped,
      failed: result.failed,
      details: result.details,
      logWarning: result.logWarning || "",
    },
    message: message || "",
  };
}

function uniqueStrings_(values) {
  const seen = {};
  const result = [];

  values.forEach(function (value) {
    const key = cleanText_(value);

    if (key && !seen[key]) {
      seen[key] = true;
      result.push(key);
    }
  });

  return result;
}

function cancelFinalUploadStep_(sheet, sourceRow, columns) {
  sheet.getRange(sourceRow, columns.ODOO).setValue(false);

  updateRows_(sheet, sourceRow, sourceRow);
}

/* =========================================================
 * PROGRESS CALCULATION
 * ======================================================= */

function updateRows_(sheet, startRow, endRow) {
  if (endRow < startRow) {
    return {
      statuses: [],
    };
  }

  const columns = getColumns_(sheet);

  const rowCount = endRow - startRow + 1;

  const lastWorkflowColumn = Math.max(columns.UPLOAD, columns.PROCESS);

  const values = sheet
    .getRange(startRow, 1, rowCount, lastWorkflowColumn)
    .getValues();

  const processNotes = sheet
    .getRange(startRow, columns.PROCESS, rowCount, 1)
    .getNotes();

  const processValues = [];
  const rowBackgrounds = [];
  const statuses = [];

  values.forEach(function (row, index) {
    const status = determineCurrentProcess_(
      row,
      columns,
      processNotes[index][0],
    );

    processValues.push([status.label]);

    statuses.push(status.label);

    rowBackgrounds.push(new Array(lastWorkflowColumn).fill(status.color));
  });

  sheet
    .getRange(startRow, columns.PROCESS, rowCount, 1)
    .setValues(processValues);

  // Formatting failures must not block status calculation.
  try {
    sheet
      .getRange(startRow, 1, rowCount, lastWorkflowColumn)
      .setBackgrounds(rowBackgrounds);
  } catch (formatError) {
    console.warn(
      "Status diperbarui, tetapi warna gagal: " + formatError.message,
    );
  }

  return {
    statuses: statuses,
  };
}

function determineCurrentProcess_(row, columns, processNote) {
  const productName = cleanText_(row[columns.PRODUCT_NAME - 1]);

  if (!productName) {
    return status_("", "EMPTY");
  }

  const productColor = cleanText_(row[columns.PRODUCT_COLOR - 1]);
  const hppValue = row[columns.HPP - 1];
  const approval = normalizedText_(row[columns.APPROVAL - 1]);

  if (isCancelled_(approval)) {
    return status_("Dibatalkan", "CANCELLED");
  }

  if (isRevision_(approval)) {
    return status_("Revisi — Linda", "REVISION");
  }

  if (
    isVerifiedProductNote_(
      processNote,
      productName,
      productColor,
      hppValue,
    )
  ) {
    return status_("Selesai", "COMPLETE");
  }

  if (!productColor) {
    return status_("Warna Produk — Linda", "PRODUCT_COLOR");
  }

  if (!hasValue_(row[columns.RESEARCH_DATE - 1])) {
    return status_("Riset — Linda", "RESEARCH");
  }

  if (!isChecked_(row[columns.PATTERN - 1])) {
    return status_("Pola — Linda", "PATTERN");
  }

  if (!isChecked_(row[columns.MATERIAL_SAMPLE - 1])) {
    return status_("Sample Bahan — Linda", "MATERIAL_SAMPLE");
  }

  if (!isChecked_(row[columns.PATTERN_LAYOUT - 1])) {
    return status_("Layout Pola — Ilham", "PATTERN_LAYOUT");
  }

  if (!isApproved_(approval)) {
    return status_("Menunggu ACC — Linda", "WAITING_APPROVAL");
  }

  if (!isChecked_(row[columns.BLUEPRINT - 1])) {
    return status_("Blueprint — Linda & Ilham", "BLUEPRINT");
  }

  const hpp = normalizedText_(hppValue);

  if (!hpp || hpp === "revisi" || hpp === "revision") {
    return status_("HPP — Oliv", "HPP");
  }

  if (!isChecked_(row[columns.ODOO - 1])) {
    return status_("Update Odoo — Oliv", "ODOO");
  }

  return status_("Menunggu Konfirmasi", "WAITING_CONFIRMATION");
}

/* =========================================================
 * CROSS-WORKBOOK DISTRIBUTION
 * ======================================================= */

function distributeVerifiedProduct_(
  product,
  sourceSpreadsheet,
  reason,
  selectedDestinationKeys,
) {
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const result = emptyDistributionResult_();
    const destinationFiles = {};
    const selected = selectedDestinationKeys
      ? uniqueStrings_(selectedDestinationKeys)
      : BARANG_BARU_CONFIG.destinations.map(function (destination) {
          return destination.key;
        });
    const selectedMap = {};

    selected.forEach(function (key) {
      selectedMap[key] = true;
    });

    const destinationsToProcess = BARANG_BARU_CONFIG.destinations.filter(
      function (destination) {
        return Boolean(selectedMap[destination.key]);
      },
    );

    if (destinationsToProcess.length !== selected.length) {
      throw new Error("Ada tujuan pengiriman yang tidak dikenal oleh sistem.");
    }

    destinationsToProcess.forEach(function (destination) {
      if (!destinationFiles[destination.spreadsheetId]) {
        try {
          destinationFiles[destination.spreadsheetId] = {
            file: SpreadsheetApp.openById(destination.spreadsheetId),
            error: null,
          };
        } catch (openError) {
          destinationFiles[destination.spreadsheetId] = {
            file: null,
            error: openError,
          };
        }
      }

      const fileResult = destinationFiles[destination.spreadsheetId];

      if (fileResult.error) {
        addDistributionError_(
          result,
          destination,
          product,
          "Tidak dapat membuka file tujuan: " + fileResult.error.message,
        );
        return;
      }

      const destinationFile = fileResult.file;
      const targetSheet = destinationFile.getSheetByName(destination.sheetName);

      if (!targetSheet) {
        addDistributionError_(
          result,
          destination,
          product,
          "Tab tujuan tidak ditemukan.",
        );
        return;
      }

      try {
        validateDestinationHeaders_(targetSheet, destination);
        writeProductToDestination_(targetSheet, destination, product, result);
      } catch (destinationError) {
        addDistributionError_(
          result,
          destination,
          product,
          destinationError.message,
        );
      }
    });

    SpreadsheetApp.flush();

    safeWriteDistributionLog_(sourceSpreadsheet, result, reason);

    return result;
  } finally {
    lock.releaseLock();
  }
}

const DESIGN1_COLOR_MAP = Object.freeze({
  "hitam": "#1E293B",
  "black": "#1E293B",
  "putih": "#64748B",
  "white": "#64748B",
  "merah": "#DC2626",
  "red": "#DC2626",
  "maroon": "#881337",
  "maron": "#881337",
  "biru": "#2563EB",
  "blue": "#2563EB",
  "bca": "#1D4ED8",
  "hijau": "#16A34A",
  "green": "#16A34A",
  "sage": "#4D7C0F",
  "mint": "#0D9488",
  "army": "#3F6212",
  "olive": "#4D7C0F",
  "kuning": "#D97706",
  "yellow": "#D97706",
  "mustard": "#B45309",
  "cokelat": "#78350F",
  "coklat": "#78350F",
  "brown": "#78350F",
  "tan": "#92400E",
  "moka": "#854D0E",
  "mocha": "#854D0E",
  "khaki": "#A16207",
  "krem": "#B45309",
  "cream": "#B45309",
  "beige": "#A16207",
  "abu-abu": "#64748B",
  "abu": "#64748B",
  "grey": "#64748B",
  "gray": "#64748B",
  "navy": "#1E3A8A",
  "dongker": "#1E3A8A",
  "pink": "#DB2777",
  "fanta": "#BE185D",
  "magenta": "#BE185D",
  "ungu": "#9333EA",
  "purple": "#9333EA",
  "lilac": "#7C3AED",
  "lavender": "#8B5CF6",
  "oranye": "#EA580C",
  "orange": "#EA580C",
  "terracotta": "#C2410C",
  "bata": "#C2410C",
  "gold": "#B45309",
  "silver": "#64748B",
  "tosca": "#0F766E",
  "toska": "#0F766E",
  "turkis": "#0284C7",
  "cyan": "#0284C7",
  "salem": "#E11D48",
  "peach": "#EA580C",
  "custom": "#4F46E5",
});

function getDesign1HexColor_(colorName) {
  if (!colorName) return "#2563EB";
  const normalized = cleanText_(colorName).toLowerCase();
  for (const key in DESIGN1_COLOR_MAP) {
    if (normalized.indexOf(key) !== -1) {
      return DESIGN1_COLOR_MAP[key];
    }
  }
  return "#2563EB";
}

function buildDesign1RichText_(productName, productColor) {
  const name = cleanText_(productName);
  const colorStr = cleanText_(productColor);
  if (!colorStr) {
    return SpreadsheetApp.newRichTextValue().setText(name).build();
  }

  const prefix = name + " (";
  const fullText = prefix + colorStr + ")";
  const builder = SpreadsheetApp.newRichTextValue().setText(fullText);

  // Parse color parts by comma / slash / semicolon
  const rawParts = colorStr.split(/([,;/]+)/);
  let runningOffset = prefix.length;

  rawParts.forEach(function (part) {
    const trimmed = part.trim();
    if (trimmed && !/^[,;/]+$/.test(trimmed)) {
      const colorStartIndex = fullText.indexOf(trimmed, runningOffset);
      if (colorStartIndex !== -1) {
        const colorEndIndex = colorStartIndex + trimmed.length;
        const hex = getDesign1HexColor_(trimmed);
        const style = SpreadsheetApp.newTextStyle()
          .setForegroundColor(hex)
          .setBold(true)
          .build();
        builder.setTextStyle(colorStartIndex, colorEndIndex, style);
        runningOffset = colorEndIndex;
      }
    } else {
      runningOffset += part.length;
    }
  });

  return builder.build();
}

function buildColorListRichText_(colorStr) {
  const text = cleanText_(colorStr);
  if (!text) {
    return SpreadsheetApp.newRichTextValue().setText("").build();
  }

  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  const rawParts = text.split(/([,;/]+)/);
  let runningOffset = 0;

  rawParts.forEach(function (part) {
    const trimmed = part.trim();
    if (trimmed && !/^[,;/]+$/.test(trimmed)) {
      const colorStartIndex = text.indexOf(trimmed, runningOffset);
      if (colorStartIndex !== -1) {
        const colorEndIndex = colorStartIndex + trimmed.length;
        const hex = getDesign1HexColor_(trimmed);
        const style = SpreadsheetApp.newTextStyle()
          .setForegroundColor(hex)
          .setBold(true)
          .build();
        builder.setTextStyle(colorStartIndex, colorEndIndex, style);
        runningOffset = colorEndIndex;
      }
    } else {
      runningOffset += part.length;
    }
  });

  return builder.build();
}

function extractBaseProductName_(value) {
  const text = cleanText_(value);
  const match = text.match(/^(.*?)\s*\([^)]*\)$/);
  if (match && match[1]) {
    return match[1].trim();
  }
  return text;
}

function writeProductToDestination_(targetSheet, destination, product, result) {
  const existingRow = findProductRow_(
    targetSheet,
    destination.productColumn,
    destination.firstDataRow,
    product.productName,
  );

  if (existingRow) {
    const updatedFields = ensureDestinationData_(targetSheet, destination, existingRow, product);

    if (updatedFields.length > 0) {
      result.updated = (result.updated || 0) + 1;
      result.details.push({
        sheetName: destination.sheetName,
        productName: product.productName,
        sourceRow: product.sourceRow,
        row: existingRow,
        action: "UPDATED",
        message: "Data produk diperbarui (" + updatedFields.join(", ") + ").",
      });
    } else {
      result.skipped += 1;
      result.details.push({
        sheetName: destination.sheetName,
        productName: product.productName,
        sourceRow: product.sourceRow,
        row: existingRow,
        action: "SKIPPED",
        message: "Nama produk sudah ada; data tujuan sudah sesuai.",
      });
    }

    return;
  }

  const targetRow = findNextDestinationRow_(targetSheet, destination);

  copyPreviousRowFormat_(
    targetSheet,
    targetRow,
    destination.firstDataRow,
    Boolean(destination.clearCopiedBackground),
  );

  if (destination.dateColumn) {
    targetSheet
      .getRange(targetRow, destination.dateColumn)
      .setValue(new Date())
      .setNumberFormat("dd/MM/yyyy");
  }

  if (destination.labelColumn) {
    targetSheet
      .getRange(targetRow, destination.labelColumn)
      .setValue(destination.labelValue);
  }

  if (destination.key === "DESIGN1") {
    const richText = buildDesign1RichText_(product.productName, product.productColor);
    targetSheet
      .getRange(targetRow, destination.productColumn)
      .setRichTextValue(richText);
  } else {
    targetSheet
      .getRange(targetRow, destination.productColumn)
      .setValue(product.productName);
  }

  if (destination.colorColumn) {
    const colorRichText = buildColorListRichText_(product.productColor);
    targetSheet
      .getRange(targetRow, destination.colorColumn)
      .setRichTextValue(colorRichText);
  }

  if (destination.hppColumn) {
    const hppCell = targetSheet.getRange(targetRow, destination.hppColumn);

    hppCell.setValue(product.hpp);

    if (product.hppNumberFormat) {
      hppCell.setNumberFormat(product.hppNumberFormat);
    }
  }

  result.inserted += 1;

  result.details.push({
    sheetName: destination.sheetName,
    productName: product.productName,
    sourceRow: product.sourceRow,
    row: targetRow,
    action: "INSERTED",
    message: "Data produk ditambahkan.",
  });
}

function ensureDestinationData_(sheet, destination, row, product) {
  const updatedFields = [];

  if (destination.key === "DESIGN1") {
    const currentCell = sheet.getRange(row, destination.productColumn);
    const richText = buildDesign1RichText_(product.productName, product.productColor);
    currentCell.setRichTextValue(richText);
    updatedFields.push("format warna Design1");
  }

  if (destination.dateColumn) {
    const dateCell = sheet.getRange(row, destination.dateColumn);

    if (!hasValue_(dateCell.getValue())) {
      dateCell.setValue(new Date()).setNumberFormat("dd/MM/yyyy");
      updatedFields.push("tanggal");
    }
  }

  if (destination.labelColumn) {
    const labelCell = sheet.getRange(row, destination.labelColumn);
    if (!valuesEquivalent_(labelCell.getValue(), destination.labelValue)) {
      labelCell.setValue(destination.labelValue);
      updatedFields.push("label");
    }
  }

  if (destination.colorColumn) {
    const colorCell = sheet.getRange(row, destination.colorColumn);
    const colorRichText = buildColorListRichText_(product.productColor);
    colorCell.setRichTextValue(colorRichText);
    updatedFields.push("warna");
  }

  if (destination.hppColumn) {
    const hppCell = sheet.getRange(row, destination.hppColumn);
    if (!valuesEquivalent_(hppCell.getValue(), product.hpp)) {
      hppCell.setValue(product.hpp);
      if (product.hppNumberFormat) {
        hppCell.setNumberFormat(product.hppNumberFormat);
      }
      updatedFields.push("HPP");
    }
  }

  return updatedFields;
}

function inspectDestinationStates_(product) {
  const destinationFiles = {};

  return BARANG_BARU_CONFIG.destinations.map(function (destination) {
    const base = {
      key: destination.key,
      sheetName: destination.sheetName,
      mapping: buildDestinationMapping_(destination),
      status: "ERROR",
      statusLabel: "Tidak dapat diperiksa",
      row: null,
      selectable: false,
      selectedByDefault: false,
      message: "",
    };

    try {
      if (!destinationFiles[destination.spreadsheetId]) {
        destinationFiles[destination.spreadsheetId] = SpreadsheetApp.openById(
          destination.spreadsheetId,
        );
      }

      const destinationFile = destinationFiles[destination.spreadsheetId];
      const sheet = destinationFile.getSheetByName(destination.sheetName);

      if (!sheet) {
        throw new Error('Tab tujuan "' + destination.sheetName + '" tidak ditemukan.');
      }

      validateDestinationHeaders_(sheet, destination);

      const existingRow = findProductRow_(
        sheet,
        destination.productColumn,
        destination.firstDataRow,
        product.productName,
      );

      if (!existingRow) {
        base.status = "MISSING";
        base.statusLabel = "Belum dikirim";
        base.selectable = true;
        base.selectedByDefault = true;
        base.message = "Nama produk belum ditemukan di tab tujuan.";
        return base;
      }

      const assessment = assessExistingDestinationData_(
        sheet,
        destination,
        existingRow,
        product,
      );

      base.row = existingRow;
      base.status = assessment.status;
      base.statusLabel = destinationStatusLabel_(assessment.status);
      base.selectable = true;
      base.selectedByDefault = assessment.status === "MISSING" || assessment.status === "INCOMPLETE" || assessment.status === "CONFLICT";
      base.message = assessment.message;

      return base;
    } catch (error) {
      base.status = "ERROR";
      base.statusLabel = "Tidak dapat diperiksa";
      base.selectable = false;
      base.selectedByDefault = false;
      base.message = cleanText_(error.message);
      return base;
    }
  });
}

function assessExistingDestinationData_(sheet, destination, row, product) {
  const missing = [];
  const conflicts = [];

  if (destination.key === "DESIGN1") {
    const nameCell = sheet.getRange(row, destination.productColumn);
    const expectedDesign1Name = cleanText_(product.productName) + " (" + cleanText_(product.productColor) + ")";
    const currentVal = cleanText_(nameCell.getDisplayValue());
    if (currentVal !== expectedDesign1Name) {
      conflicts.push('warna/format Design1 ("' + currentVal + '")');
    }
  }

  if (destination.dateColumn) {
    const dateCell = sheet.getRange(row, destination.dateColumn);

    if (!hasValue_(dateCell.getValue())) {
      missing.push("tanggal request");
    }
  }

  if (destination.labelColumn) {
    inspectExpectedDestinationValue_(
      sheet,
      row,
      destination.labelColumn,
      "label",
      destination.labelValue,
      valuesEquivalent_,
      missing,
      conflicts,
    );
  }

  if (destination.colorColumn) {
    inspectExpectedDestinationValue_(
      sheet,
      row,
      destination.colorColumn,
      "warna",
      product.productColor,
      colorListsEquivalent_,
      missing,
      conflicts,
    );
  }

  if (destination.hppColumn) {
    inspectExpectedDestinationValue_(
      sheet,
      row,
      destination.hppColumn,
      "HPP",
      product.hpp,
      valuesEquivalent_,
      missing,
      conflicts,
    );
  }

  if (conflicts.length > 0) {
    return {
      status: "CONFLICT",
      message:
        "Nama produk sudah ada di baris " +
        row +
        ", tetapi " +
        conflicts.join("; ") +
        ". Pilih tujuan ini untuk memperbarui datanya.",
    };
  }

  if (missing.length > 0) {
    return {
      status: "INCOMPLETE",
      message:
        "Nama produk sudah ada di baris " +
        row +
        ", tetapi " +
        missing.join(", ") +
        " masih kosong. Pilih tujuan ini untuk melengkapinya.",
    };
  }

  return {
    status: "COMPLETE",
    message: "Nama produk sudah ada dan data yang diperiksa sudah sesuai.",
  };
}

function inspectExpectedDestinationValue_(
  sheet,
  row,
  column,
  fieldName,
  expectedValue,
  comparator,
  missing,
  conflicts,
) {
  const cell = sheet.getRange(row, column);
  const existingValue = cell.getValue();

  if (!hasValue_(existingValue)) {
    missing.push(fieldName);
    return;
  }

  if (!comparator(existingValue, expectedValue)) {
    conflicts.push(
      fieldName +
        ' berbeda (tujuan: "' +
        cleanText_(cell.getDisplayValue()) +
        '")',
    );
  }
}

function destinationStatusLabel_(status) {
  const labels = {
    COMPLETE: "Sudah ada",
    MISSING: "Belum dikirim",
    INCOMPLETE: "Belum lengkap",
    CONFLICT: "Data berbeda",
    ERROR: "Tidak dapat diperiksa",
  };

  return labels[status] || status;
}

function allDestinationsComplete_(states) {
  return (
    states.length === BARANG_BARU_CONFIG.destinations.length &&
    states.every(function (state) {
      return state.status === "COMPLETE";
    })
  );
}

function canCancelUploadFromStates_(states) {
  return (
    states.length === BARANG_BARU_CONFIG.destinations.length &&
    states.every(function (state) {
      return state.status === "MISSING";
    })
  );
}

function checkDestinationConnection_() {
  const destinationFiles = {};
  const checks = BARANG_BARU_CONFIG.destinations.map(function (destination) {
    if (!destinationFiles[destination.spreadsheetId]) {
      destinationFiles[destination.spreadsheetId] = SpreadsheetApp.openById(
        destination.spreadsheetId,
      );
    }

    const destinationFile = destinationFiles[destination.spreadsheetId];
    const sheet = destinationFile.getSheetByName(destination.sheetName);

    if (!sheet) {
      throw new Error(
        'Tab tujuan "' + destination.sheetName + '" tidak ditemukan.',
      );
    }

    validateDestinationHeaders_(sheet, destination);

    return {
      fileName: destinationFile.getName(),
      sheetName: destination.sheetName,
      productColumn: destination.productColumn,
      labelColumn: destination.labelColumn,
      labelValue: destination.labelValue,
      mapping: buildDestinationMapping_(destination),
      nextRow: findNextDestinationRow_(sheet, destination),
    };
  });

  return {
    checks: checks,
  };
}

function validateDestinationHeaders_(sheet, destination) {
  if (!destination.expectedHeaders) return;

  const headerRow = destination.headerRow || 1;

  destination.expectedHeaders.forEach(function (expected) {
    const actual = cleanText_(
      sheet.getRange(headerRow, expected.column).getDisplayValue(),
    );
    const matches = expected.names.some(function (name) {
      return normalizeHeader_(actual) === normalizeHeader_(name);
    });

    if (!matches) {
      throw new Error(
        'Header tab "' +
          destination.sheetName +
          '" kolom ' +
          columnNumberToLetter_(expected.column) +
          " baris " +
          headerRow +
          ' harus "' +
          expected.names[0] +
          '", tetapi saat ini "' +
          actual +
          '".',
      );
    }
  });
}

function buildDestinationMapping_(destination) {
  const parts = [
    columnNumberToLetter_(destination.productColumn) + ": nama produk",
  ];

  if (destination.dateColumn) {
    parts.push(
      columnNumberToLetter_(destination.dateColumn) + ": tanggal request otomatis",
    );
  }

  if (destination.labelColumn) {
    parts.push(
      columnNumberToLetter_(destination.labelColumn) +
        ': "' +
        destination.labelValue +
        '"',
    );
  }

  if (destination.colorColumn) {
    parts.push(columnNumberToLetter_(destination.colorColumn) + ": warna");
  }

  if (destination.hppColumn) {
    parts.push(columnNumberToLetter_(destination.hppColumn) + ": HPP");
  }

  return parts.join(", ");
}

function findNextDestinationRow_(sheet, destination) {
  if (destination.appendAfterLastProduct) {
    return findRowAfterLastProduct_(
      sheet,
      destination.productColumn,
      destination.firstDataRow,
    );
  }

  return findFirstEmptyProductRow_(
    sheet,
    destination.productColumn,
    destination.firstDataRow,
  );
}

/**
 * Used only by destinations configured with appendAfterLastProduct.
 * It ignores historical gaps and appends immediately after the last
 * non-empty product name in the destination product column.
 */
function findRowAfterLastProduct_(sheet, productColumn, firstDataRow) {
  const lastSheetRow = Math.max(sheet.getLastRow(), firstDataRow);
  const rowCount = lastSheetRow - firstDataRow + 1;

  const values = sheet
    .getRange(firstDataRow, productColumn, rowCount, 1)
    .getDisplayValues();

  for (let index = values.length - 1; index >= 0; index -= 1) {
    if (cleanText_(values[index][0])) {
      return firstDataRow + index + 1;
    }
  }

  return firstDataRow;
}

function findFirstEmptyProductRow_(sheet, productColumn, firstDataRow) {
  const lastRow = Math.max(sheet.getLastRow(), firstDataRow);

  const values = sheet
    .getRange(firstDataRow, productColumn, lastRow - firstDataRow + 1, 1)
    .getDisplayValues();

  for (let index = 0; index < values.length; index += 1) {
    if (!cleanText_(values[index][0])) {
      return firstDataRow + index;
    }
  }

  return lastRow + 1;
}

function findProductRow_(sheet, productColumn, firstDataRow, productName) {
  const lastRow = sheet.getLastRow();

  if (lastRow < firstDataRow) {
    return null;
  }

  const targetName = normalizedProductName_(productName);

  const values = sheet
    .getRange(firstDataRow, productColumn, lastRow - firstDataRow + 1, 1)
    .getDisplayValues();

  for (let index = 0; index < values.length; index += 1) {
    const rawVal = values[index][0];
    const directName = normalizedProductName_(rawVal);
    const baseName = normalizedProductName_(extractBaseProductName_(rawVal));

    if (directName === targetName || baseName === targetName) {
      return firstDataRow + index;
    }
  }

  return null;
}

/**
 * Copies formatting and validation from the previous row.
 * Values and formulas are not copied.
 */
function copyPreviousRowFormat_(
  sheet,
  targetRow,
  firstDataRow,
  clearCopiedBackground,
) {
  if (targetRow <= firstDataRow) return;

  const previousRow = targetRow - 1;

  const lastColumn = Math.max(sheet.getLastColumn(), 1);

  const source = sheet.getRange(previousRow, 1, 1, lastColumn);

  const target = sheet.getRange(targetRow, 1, 1, lastColumn);

  source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  source.copyTo(
    target,
    SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
    false,
  );

  // Video should inherit borders/layout/validation, but not a status color
  // from the previous historical row. null restores the default background.
  if (clearCopiedBackground) {
    target.setBackground(null);
  }

  sheet.setRowHeight(targetRow, sheet.getRowHeight(previousRow));
}

/* =========================================================
 * DISTRIBUTION RESULT HELPERS
 * ======================================================= */

function emptyDistributionResult_() {
  return {
    inserted: 0,
    updated: 0,
    skipped: 0,
    failed: 0,
    details: [],
    logWarning: "",
  };
}

function addDistributionError_(result, destination, product, message) {
  result.failed += 1;

  result.details.push({
    sheetName: destination.sheetName,
    productName: product.productName,
    sourceRow: product.sourceRow,
    action: "ERROR",
    message: cleanText_(message),
  });
}

function buildDistributionMessage_(result) {
  const parts = [];
  if (result.inserted > 0) parts.push(result.inserted + " tujuan ditambahkan");
  if (result.updated > 0) parts.push(result.updated + " tujuan diperbarui");
  if (result.skipped > 0) parts.push(result.skipped + " tujuan sudah sesuai");
  if (result.failed > 0) parts.push(result.failed + " gagal");
  return parts.length > 0 ? parts.join("; ") + "." : "Tidak ada perubahan pada data tujuan.";
}

function buildDistributionDetails_(result) {
  if (!result.details.length) {
    return "Tidak ada detail distribusi.";
  }

  return result.details
    .map(function (detail) {
      let resultText;

      if (detail.action === "INSERTED") {
        resultText = "ditambahkan ke baris " + detail.row;
      } else if (detail.action === "UPDATED") {
        resultText = "diperbarui di baris " + detail.row + " (" + detail.message + ")";
      } else if (detail.action === "SKIPPED") {
        resultText = "sudah ada di baris " + detail.row;
      } else {
        resultText = "GAGAL — " + detail.message;
      }

      return detail.sheetName + ": " + resultText;
    })
    .join("\n");
}

/* =========================================================
 * VERIFIED STATE AND SOURCE NOTES
 * ======================================================= */

function writeSourceDistributionNote_(
  sheet,
  processColumn,
  product,
  result,
  verified,
) {
  const timezone = sheet.getParent().getSpreadsheetTimeZone();

  const note = [];

  if (verified) {
    note.push(buildVerifiedSnapshotMarker_(product));
  }

  note.push(
    "Distribusi terakhir: " +
      Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd HH:mm:ss"),
  );

  result.details.forEach(function (detail) {
    if (detail.action === "INSERTED") {
      note.push(detail.sheetName + ": ditambahkan ke baris " + detail.row);
    } else if (detail.action === "UPDATED") {
      note.push(detail.sheetName + ": diperbarui di baris " + detail.row);
    } else if (detail.action === "SKIPPED") {
      note.push(detail.sheetName + ": sudah ada di baris " + detail.row);
    } else {
      note.push(detail.sheetName + ": GAGAL — " + detail.message);
    }
  });

  if (result.logWarning) {
    note.push("Peringatan log: " + result.logWarning);
  }

  sheet.getRange(product.sourceRow, processColumn).setNote(note.join("\n"));
}

function clearVerifiedStateForRow_(sheet, sourceRow, processColumn) {
  const cell = sheet.getRange(sourceRow, processColumn);
  const note = cell.getNote();

  if (!note) return;

  const keptLines = String(note)
    .split("\n")
    .filter(function (line) {
      return (
        line.indexOf(BARANG_BARU_CONFIG.verifiedMarkerPrefix) !== 0 &&
        line.indexOf(BARANG_BARU_CONFIG.verifiedSnapshotPrefix) !== 0
      );
    });

  cell.setNote(keptLines.join("\n"));
}

function buildLegacyVerifiedMarker_(productName) {
  return (
    BARANG_BARU_CONFIG.verifiedMarkerPrefix +
    encodeURIComponent(normalizedProductName_(productName))
  );
}

function buildVerifiedSnapshotMarker_(product) {
  const snapshot = {
    productName: normalizedProductName_(product.productName),
    productColor: normalizeColorList_(product.productColor),
    hpp: normalizeComparableValue_(product.hpp),
  };

  return (
    BARANG_BARU_CONFIG.verifiedSnapshotPrefix +
    encodeURIComponent(JSON.stringify(snapshot))
  );
}

function isVerifiedProductNote_(
  note,
  productName,
  productColor,
  hpp,
) {
  if (!note) return false;

  const expectedSnapshot = buildVerifiedSnapshotMarker_({
    productName: productName,
    productColor: productColor,
    hpp: hpp,
  });
  const expectedLegacyMarker = buildLegacyVerifiedMarker_(productName);
  const lines = String(note)
    .split("\n")
    .map(cleanText_);

  if (lines.indexOf(expectedSnapshot) !== -1) {
    return true;
  }

  return !productColor && lines.indexOf(expectedLegacyMarker) !== -1;
}

function migrateLegacyDistributionNotes_(sheet) {
  if (!hasProcessColumn_(sheet)) {
    return;
  }

  const columns = getColumns_(sheet);

  const lastRow = sheet.getLastRow();

  if (lastRow < BARANG_BARU_CONFIG.firstDataRow) {
    return;
  }

  const rowCount = lastRow - BARANG_BARU_CONFIG.firstDataRow + 1;

  const productNames = sheet
    .getRange(
      BARANG_BARU_CONFIG.firstDataRow,
      columns.PRODUCT_NAME,
      rowCount,
      1,
    )
    .getDisplayValues();

  const processValues = sheet
    .getRange(BARANG_BARU_CONFIG.firstDataRow, columns.PROCESS, rowCount, 1)
    .getDisplayValues();

  const processNotes = sheet
    .getRange(BARANG_BARU_CONFIG.firstDataRow, columns.PROCESS, rowCount, 1)
    .getNotes();

  const updatedNotes = processNotes.map(function (row, index) {
    const productName = cleanText_(productNames[index][0]);

    const processValue = normalizedText_(processValues[index][0]);

    const note = cleanText_(row[0]);

    if (
      !productName ||
      processValue !== "selesai" ||
      isVerifiedProductNote_(note, productName, "", "") ||
      !isSuccessfulLegacyNote_(note)
    ) {
      return [row[0]];
    }

    return [[buildLegacyVerifiedMarker_(productName), note].join("\n")];
  });

  sheet
    .getRange(BARANG_BARU_CONFIG.firstDataRow, columns.PROCESS, rowCount, 1)
    .setNotes(updatedNotes);
}

function isSuccessfulLegacyNote_(note) {
  const normalized = normalizedText_(note);

  return (
    normalized.indexOf("distribusi terakhir:") !== -1 &&
    normalized.indexOf("design1:") !== -1 &&
    normalized.indexOf("dadmin:") !== -1 &&
    normalized.indexOf("vadmin:") !== -1 &&
    normalized.indexOf("gagal") === -1
  );
}

/* =========================================================
 * DISTRIBUTION LOG
 * ======================================================= */

function ensureDistributionLogSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(BARANG_BARU_CONFIG.logSheetName);

  let isNew = false;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(BARANG_BARU_CONFIG.logSheetName);

    isNew = true;
  }

  const headers = BARANG_BARU_CONFIG.logHeaders;

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground("#15395B");

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 320);
  sheet.setColumnWidth(8, 110);

  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");

  if (isNew && spreadsheet.getSheets().length > 1) {
    sheet.hideSheet();
  }

  return sheet;
}

function safeWriteDistributionLog_(spreadsheet, result, reason) {
  try {
    writeDistributionLog_(spreadsheet, result.details, reason);
  } catch (error) {
    result.logWarning = cleanText_(error.message);

    console.warn(
      "Distribusi selesai, tetapi log gagal ditulis: " + error.message,
    );
  }
}

function writeDistributionLog_(spreadsheet, details, reason) {
  if (!details || details.length === 0) {
    return;
  }

  const sheet = ensureDistributionLogSheet_(spreadsheet);

  const rows = details.map(function (detail) {
    return [
      new Date(),
      detail.productName || "",
      detail.sheetName || "",
      detail.action === "INSERTED"
        ? "DITAMBAHKAN"
        : detail.action === "UPDATED"
          ? "DIPERBARUI"
          : detail.action === "SKIPPED"
            ? "SUDAH ADA"
            : "GAGAL",
      detail.row || "",
      reason || "",
      detail.message || "",
      detail.sourceRow || "",
    ];
  });

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      BARANG_BARU_CONFIG.logHeaders.length,
    )
    .setValues(rows);
}

function recordTriggerFailure_(sheet, startRow, endRow, error) {
  try {
    const columns = getColumns_(sheet);

    const rowCount = endRow - startRow + 1;

    const productValues = sheet
      .getRange(startRow, columns.PRODUCT_NAME, rowCount, 1)
      .getDisplayValues();

    const details = productValues.map(function (row, index) {
      return {
        sheetName: "SISTEM",
        productName: cleanText_(row[0]) || "(tanpa nama)",
        sourceRow: startRow + index,
        action: "ERROR",
        message: cleanText_(error.message),
      };
    });

    writeDistributionLog_(sheet.getParent(), details, "UPDATE STATUS");
  } catch (logError) {
    console.error("Gagal mencatat error: " + logError.message);
  }
}

function openDistributionLog() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ensureDistributionLogSheet_(spreadsheet);

  sheet.showSheet();
  spreadsheet.setActiveSheet(sheet);
}

/* =========================================================
 * HEADER AND COLUMN DETECTION
 * ======================================================= */

function getColumns_(sheet) {
  const definitions = {
    NUMBER: ["No"],
    PROCESS: ["Proses Saat Ini"],
    PRODUCT_NAME: ["Nama Barang"],
    PRODUCT_COLOR: ["Warna Produk", "Warna"],
    RESEARCH_DATE: ["Tanggal Riset", "Tgl Riset"],
    PATTERN: ["Pola"],
    MATERIAL_SAMPLE: ["Sample Bahan"],
    PATTERN_LAYOUT: ["Layout Pola"],
    APPROVAL: ["ACC/REVISI", "ACC Revisi"],
    BLUEPRINT: ["Blueprint"],
    HPP: ["HPP"],
    ODOO: ["Update Odoo"],
    SELLING_PRICE: ["Harga Jual MP", "Harga Jual Marketplace"],
    UPLOAD: ["Upload"],
  };

  const headerMap = getHeaderMap_(sheet);

  const columns = {};
  const missing = [];

  Object.keys(definitions).forEach(function (key) {
    const column = findHeaderInMap_(headerMap, definitions[key]);

    if (!column) {
      missing.push(definitions[key][0]);
    } else {
      columns[key] = column;
    }
  });

  if (missing.length > 0) {
    throw new Error(
      "Header tidak ditemukan: " +
        missing.join(", ") +
        ". Periksa baris " +
        BARANG_BARU_CONFIG.headerRow +
        ".",
    );
  }

  return Object.freeze(columns);
}

function ensureProcessColumn_(sheet) {
  if (hasProcessColumn_(sheet)) {
    return;
  }

  const headerMap = getHeaderMap_(sheet);

  const numberColumn = findHeaderInMap_(headerMap, ["No"]);

  if (!numberColumn) {
    throw new Error(
      'Header "No" tidak ditemukan pada baris ' +
        BARANG_BARU_CONFIG.headerRow +
        ".",
    );
  }

  const processColumn = numberColumn + 1;

  sheet.insertColumnAfter(numberColumn);

  sheet
    .getRange(BARANG_BARU_CONFIG.headerRow, processColumn)
    .setValue(BARANG_BARU_CONFIG.processHeader);

  sheet.getRange(BARANG_BARU_CONFIG.picRow, processColumn).setValue("Sistem");
}

function ensureProductColorColumn_(sheet) {
  const headerMap = getHeaderMap_(sheet);
  const existingColumn = findHeaderInMap_(headerMap, [
    BARANG_BARU_CONFIG.productColorHeader,
    "Warna",
  ]);

  if (existingColumn) {
    return false;
  }

  const productColumn = findHeaderInMap_(headerMap, ["Nama Barang"]);

  if (!productColumn) {
    throw new Error(
      'Header "Nama Barang" tidak ditemukan pada baris ' +
        BARANG_BARU_CONFIG.headerRow +
        ".",
    );
  }

  sheet.insertColumnAfter(productColumn);

  const colorColumn = productColumn + 1;
  const sourceHeader = sheet.getRange(
    BARANG_BARU_CONFIG.headerRow,
    productColumn,
  );
  const targetHeader = sheet.getRange(
    BARANG_BARU_CONFIG.headerRow,
    colorColumn,
  );

  sourceHeader.copyTo(
    targetHeader,
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false,
  );
  targetHeader.setValue(BARANG_BARU_CONFIG.productColorHeader);

  const sourcePic = sheet.getRange(BARANG_BARU_CONFIG.picRow, productColumn);
  const targetPic = sheet.getRange(BARANG_BARU_CONFIG.picRow, colorColumn);

  sourcePic.copyTo(
    targetPic,
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false,
  );
  targetPic.setValue("Linda");

  return true;
}

function styleProductColorColumn_(sheet) {
  const colorColumn = findHeaderInMap_(getHeaderMap_(sheet), [
    BARANG_BARU_CONFIG.productColorHeader,
    "Warna",
  ]);

  if (!colorColumn) {
    throw new Error('Kolom "Warna Produk" tidak ditemukan.');
  }

  sheet
    .getRange(BARANG_BARU_CONFIG.headerRow, colorColumn)
    .setValue(BARANG_BARU_CONFIG.productColorHeader)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setNote(
      "Pilih satu atau beberapa warna dari dropdown. " +
        "Kolom ini wajib diisi sebelum produk dikirim.",
    );

  sheet.setColumnWidth(colorColumn, 180);
}

function ensureProductColorDropdown_(sheet) {
  const colorColumn = findHeaderInMap_(getHeaderMap_(sheet), [
    BARANG_BARU_CONFIG.productColorHeader,
  ]);
  const firstCell = sheet.getRange(
    BARANG_BARU_CONFIG.firstDataRow,
    colorColumn,
  );

  // Keep manually configured chip colors and multi-select settings.
  if (firstCell.getDataValidation()) {
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(BARANG_BARU_CONFIG.productColorOptions, true)
    .setAllowInvalid(false)
    .setHelpText("Pilih warna produk dari daftar; jangan diketik manual.")
    .build();

  sheet
    .getRange(
      BARANG_BARU_CONFIG.firstDataRow,
      colorColumn,
      sheet.getMaxRows() - BARANG_BARU_CONFIG.firstDataRow + 1,
      1,
    )
    .setDataValidation(rule);
}

function styleProcessColumn_(sheet) {
  const headerMap = getHeaderMap_(sheet);

  const processColumn = findHeaderInMap_(headerMap, [
    BARANG_BARU_CONFIG.processHeader,
  ]);

  const productColumn = findHeaderInMap_(headerMap, ["Nama Barang"]);

  if (!processColumn || !productColumn) {
    throw new Error("Kolom Proses Saat Ini atau Nama Barang tidak ditemukan.");
  }

  const processHeaderCell = sheet.getRange(
    BARANG_BARU_CONFIG.headerRow,
    processColumn,
  );

  sheet
    .getRange(BARANG_BARU_CONFIG.headerRow, productColumn)
    .copyTo(
      processHeaderCell,
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false,
    );

  processHeaderCell
    .setValue(BARANG_BARU_CONFIG.processHeader)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setNote(
      "Status dihitung otomatis. Setelah Upload lengkap, " +
        "status menjadi Menunggu Konfirmasi. " +
        "Gunakan menu Verifikasi & Kirim untuk mengirim produk.",
    );

  sheet.setColumnWidth(processColumn, 190);
}

function installEditTrigger_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === "handleBarangBaruEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("handleBarangBaruEdit")
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
}

function hasProcessColumn_(sheet) {
  return Boolean(
    findHeaderInMap_(getHeaderMap_(sheet), [BARANG_BARU_CONFIG.processHeader]),
  );
}

function getHeaderMap_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);

  const headers = sheet
    .getRange(BARANG_BARU_CONFIG.headerRow, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const map = {};

  headers.forEach(function (header, index) {
    const normalized = normalizeHeader_(header);

    if (normalized && !map[normalized]) {
      map[normalized] = index + 1;
    }
  });

  return map;
}

function findHeaderInMap_(headerMap, names) {
  for (let index = 0; index < names.length; index += 1) {
    const normalized = normalizeHeader_(names[index]);

    if (headerMap[normalized]) {
      return headerMap[normalized];
    }
  }

  return null;
}

function getBarangBaruSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    BARANG_BARU_CONFIG.sourceSheetName,
  );

  if (!sheet) {
    throw new Error('Sheet "Barang Baru" tidak ditemukan.');
  }

  return sheet;
}

/* =========================================================
 * STATUS AND VALUE HELPERS
 * ======================================================= */

function status_(label, key) {
  return {
    label: label,
    key: key,
    color: BARANG_BARU_CONFIG.colors[key],
  };
}

function isApproved_(value) {
  return value === "acc" || value === "approved" || value === "disetujui";
}

function isRevision_(value) {
  return value === "revisi" || value === "revision";
}

function isCancelled_(value) {
  return (
    value === "cancel" ||
    value === "cancelled" ||
    value === "canceled" ||
    value === "batal" ||
    value === "dibatalkan"
  );
}

function isChecked_(value) {
  if (value === true) {
    return true;
  }

  const text = normalizedText_(value);

  return text === "true" || text === "ya" || text === "yes" || text === "✓";
}

function hasValue_(value) {
  if (value === null || value === undefined || value === false) {
    return false;
  }

  if (value instanceof Date) {
    return !isNaN(value.getTime());
  }

  return cleanText_(value) !== "";
}

function normalizedText_(value) {
  return cleanText_(value).toLowerCase();
}

function normalizedProductName_(value) {
  return normalizedText_(value).replace(/\s+/g, " ");
}

function normalizeColorList_(value) {
  return cleanText_(value)
    .split(",")
    .map(function (color) {
      return normalizedText_(color);
    })
    .filter(Boolean)
    .sort()
    .join("|");
}

function colorListsEquivalent_(left, right) {
  return normalizeColorList_(left) === normalizeColorList_(right);
}

function normalizeComparableValue_(value) {
  if (typeof value === "number" && isFinite(value)) {
    return String(value);
  }

  return normalizedText_(value).replace(/\s+/g, " ");
}

function valuesEquivalent_(left, right) {
  return normalizeComparableValue_(left) === normalizeComparableValue_(right);
}

function normalizeHeader_(value) {
  return normalizedText_(value)
    .replace(/\s*\/\s*/g, "/")
    .replace(/\s+/g, " ");
}

function cleanText_(value) {
  return String(value == null ? "" : value).trim();
}

function columnNumberToLetter_(column) {
  let number = column;
  let letters = "";

  while (number > 0) {
    const remainder = (number - 1) % 26;

    letters = String.fromCharCode(65 + remainder) + letters;

    number = Math.floor((number - 1) / 26);
  }

  return letters;
}

/* =========================================================
 * INSTRUCTIONS
 * ======================================================= */

function showBarangBaruInstructions() {
  SpreadsheetApp.getUi().alert(
    "Petunjuk Barang Baru",
    "Alur:\n" +
      "Riset → Pola → Sample Bahan → Layout Pola → " +
      "ACC → Blueprint → HPP → Update Odoo → " +
      "Harga Jual → Upload → Menunggu Konfirmasi.\n\n" +
      "Untuk mengirim:\n" +
      "1. Pilih satu sel pada baris produk.\n" +
      "2. Buka Barang Baru → Verifikasi & Kirim.\n" +
      "3. Sistem memeriksa nama produk langsung di Design1, DAdmin, Video, dan NEW_LINDA.\n" +
      "4. Tujuan yang sudah lengkap akan ditandai Sudah ada dan tidak dapat dipilih.\n" +
      "5. Pilih tujuan yang berstatus Belum dikirim atau Belum lengkap, lalu klik Kirim.\n\n" +
      "Tujuan:\n" +
      "• Design1: A = NEW LINDA, B = nama produk\n" +
      "• DAdmin: B = NEW LINDA, C = nama produk\n" +
      "• Video: B = tanggal request otomatis, C = nama produk\n" +
      "• NEW_LINDA: A = nama, B = warna, C = HPP\n\n" +
      "Status Selesai hanya diberikan jika keempat tujuan sudah berisi data yang sesuai. " +
      "Jika hanya sebagian tujuan dikirim, status tetap Menunggu Konfirmasi.\n\n" +
      "Jika nama produk sudah ada tetapi data tujuan berbeda, sistem tidak menimpa data lama. " +
      "Perbaiki tujuan tersebut secara manual lalu periksa kembali.\n\n" +
      "Batalkan Upload hanya tersedia jika produk belum ditemukan di semua tujuan. " +
      "Menutup dialog tidak membuat perubahan.",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}
