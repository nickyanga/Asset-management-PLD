// ============================================================
// Asset Management PLD — Google Apps Script Backend
// ============================================================

// TODO: Replace with your Google Spreadsheet ID after creating the sheet
var SPREADSHEET_ID = "14O9Sx0k3xTvmJsGL8tV6xelhK8DETKe7Xym00Ga3AVM";

// ============================================================
// Entry Point
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Asset Management PLD")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// Private Sheet Helpers
// ============================================================

function getAssetsSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Assets");
  if (!sheet)
    throw new Error("Assets sheet not found. Please create it with headers.");
  return sheet;
}

function getTransactionsSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Transactions");
  if (!sheet)
    throw new Error(
      "Transactions sheet not found. Please create it with headers.",
    );
  return sheet;
}

/**
 * Returns all data rows (excluding header row 1) as array of arrays.
 */
function getSheetData_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}

/**
 * Generates a unique transaction ID: TXN-YYYYMMDD-HHMMSS-XXXX
 */
function generateTxnId_() {
  var now = new Date();
  var pad = function (n) {
    return String(n).padStart(2, "0");
  };
  var date = now.getFullYear() + pad(now.getMonth() + 1) + pad(now.getDate());
  var time =
    pad(now.getHours()) + pad(now.getMinutes()) + pad(now.getSeconds());
  var rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return "TXN-" + date + "-" + time + "-" + rand;
}

// ============================================================
// Email Helpers
// ============================================================

function sendLoanReceipt_(data) {
  var subject = "PLD Loan Receipt \u2013 " + data.serialNumber;
  var body = [
    "PLD Loan Receipt",
    "================",
    "",
    "Dear " + data.borrowerName + ",",
    "",
    "This confirms that you have borrowed the following device:",
    "",
    "Transaction ID : " + data.transactionId,
    "Loan Date/Time : " + data.loanDate,
    "",
    "Device Details",
    "--------------",
    "Serial Number  : " + data.serialNumber,
    "Model          : " + data.model,
    "Condition      : " + data.condition,
    "",
    "Please return the device in the same condition.",
    "If you have any questions, contact the asset manager.",
    "",
    "This is an automated receipt.",
  ].join("\n");

  MailApp.sendEmail({
    to: data.borrowerEmail,
    subject: subject,
    body: body,
  });
}

function sendReturnReceipt_(data) {
  var subject = "PLD Return Receipt \u2013 " + data.serialNumber;

  // Calculate loan duration
  var loanMs = new Date(data.returnDate) - new Date(data.loanDate);
  var days = Math.floor(loanMs / (1000 * 60 * 60 * 24));
  var hours = Math.floor((loanMs % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
  var duration = days + " day(s), " + hours + " hour(s)";

  var body = [
    "PLD Return Receipt",
    "==================",
    "",
    "Dear " + data.borrowerName + ",",
    "",
    "This confirms that you have returned the following device:",
    "",
    "Transaction ID : " + data.transactionId,
    "Loan Date/Time : " + data.loanDate,
    "Return Date/Time: " + data.returnDate,
    "Loan Duration  : " + duration,
    "",
    "Device Details",
    "--------------",
    "Serial Number  : " + data.serialNumber,
    "Model          : " + data.model,
    "Condition      : " + data.condition,
    "",
    "Thank you for returning the device.",
    "",
    "This is an automated receipt.",
  ].join("\n");

  MailApp.sendEmail({
    to: data.borrowerEmail,
    subject: subject,
    body: body,
  });
}

// ============================================================
// Asset Reads
// ============================================================

/**
 * Returns array of asset objects where Status === 'Available'.
 */
function getAvailableAssets() {
  var sheet = getAssetsSheet_();
  var rows = getSheetData_(sheet);
  var assets = [];
  rows.forEach(function (row) {
    var status = String(row[3] || "").trim();
    if (status === "Available") {
      assets.push({
        serialNumber: String(row[0] || "").trim(),
        model: String(row[1] || "").trim(),
        condition: String(row[2] || "").trim(),
        status: status,
        notes: String(row[4] || "").trim(),
      });
    }
  });
  return assets;
}

/**
 * Returns a single asset object by SerialNumber, or null if not found.
 */
function getAssetBySerial(serialNumber) {
  var sn = String(serialNumber || "")
    .trim()
    .toUpperCase();
  var sheet = getAssetsSheet_();
  var rows = getSheetData_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (
      String(rows[i][0] || "")
        .trim()
        .toUpperCase() === sn
    ) {
      return {
        serialNumber: String(rows[i][0] || "").trim(),
        model: String(rows[i][1] || "").trim(),
        condition: String(rows[i][2] || "").trim(),
        status: String(rows[i][3] || "").trim(),
        notes: String(rows[i][4] || "").trim(),
        dateAdded: rows[i][5]
          ? Utilities.formatDate(
              new Date(rows[i][5]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
        lastModified: rows[i][6]
          ? Utilities.formatDate(
              new Date(rows[i][6]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
      };
    }
  }
  return null;
}

/**
 * Returns all assets (for Records admin grid).
 */
function getAllAssets() {
  var sheet = getAssetsSheet_();
  var rows = getSheetData_(sheet);
  return rows
    .map(function (row) {
      return {
        serialNumber: String(row[0] || "").trim(),
        model: String(row[1] || "").trim(),
        condition: String(row[2] || "").trim(),
        status: String(row[3] || "").trim(),
        notes: String(row[4] || "").trim(),
        dateAdded: row[5]
          ? Utilities.formatDate(
              new Date(row[5]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd",
            )
          : "",
        lastModified: row[6]
          ? Utilities.formatDate(
              new Date(row[6]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
      };
    })
    .filter(function (a) {
      return a.serialNumber !== "";
    });
}

/**
 * Returns all transactions sorted newest-first (for Records history).
 */
function getAllTransactions() {
  var sheet = getTransactionsSheet_();
  var rows = getSheetData_(sheet);
  var txns = rows
    .map(function (row) {
      return {
        transactionId: String(row[0] || "").trim(),
        serialNumber: String(row[1] || "").trim(),
        condition: String(row[2] || "").trim(),
        borrowerName: String(row[3] || "").trim(),
        borrowerEmail: String(row[4] || "").trim(),
        loanDate: row[5]
          ? Utilities.formatDate(
              new Date(row[5]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
        returnDate: row[6]
          ? Utilities.formatDate(
              new Date(row[6]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
        type: String(row[7] || "").trim(),
        adminNotes: String(row[8] || "").trim(),
      };
    })
    .filter(function (t) {
      return t.transactionId !== "";
    });

  // Sort newest first by loanDate
  txns.sort(function (a, b) {
    return new Date(b.loanDate) - new Date(a.loanDate);
  });
  return txns;
}

// ============================================================
// Asset Writes
// ============================================================

/**
 * Appends a new asset row to the Assets sheet.
 * assetData: { serialNumber, model, condition, notes }
 */
function addAsset(assetData) {
  var sn = String(assetData.serialNumber || "").trim();
  if (!sn) throw new Error("Serial Number is required.");

  // Check for duplicate
  var existing = getAssetBySerial(sn);
  if (existing)
    throw new Error('Serial number "' + sn + '" already exists.');

  var now = new Date();
  var sheet = getAssetsSheet_();
  sheet.appendRow([
    sn,
    String(assetData.model || "").trim(),
    String(assetData.condition || "Good").trim(),
    "Available",
    String(assetData.notes || "").trim(),
    now,
    now,
  ]);
  return { success: true, serialNumber: sn };
}

/**
 * Updates specified fields for an existing asset.
 * Called from frontend as updateAsset({ serialNumber, updatedFields }).
 * params: { serialNumber, updatedFields: { model?, condition?, status?, notes? } }
 */
function updateAsset(params) {
  var serialNumber = params.serialNumber;
  var updatedFields = params.updatedFields;
  var sn = String(serialNumber || "")
    .trim()
    .toUpperCase();
  var sheet = getAssetsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("No assets found.");

  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (
      String(data[i][0] || "")
        .trim()
        .toUpperCase() === sn
    ) {
      var rowNum = i + 2;
      var row = data[i];

      // Map field names to column indices (0-based)
      var fieldMap = {
        model: 1,
        condition: 2,
        status: 3,
        notes: 4,
      };

      Object.keys(fieldMap).forEach(function (field) {
        if (updatedFields.hasOwnProperty(field)) {
          row[fieldMap[field]] = updatedFields[field];
        }
      });
      row[6] = new Date(); // LastModified

      sheet.getRange(rowNum, 1, 1, 7).setValues([row]);
      return { success: true };
    }
  }
  throw new Error('Serial number "' + serialNumber + '" not found.');
}

/**
 * Parses a CSV string and imports assets.
 * Returns { imported, skipped, errors[] }
 */
function importAssetsFromCSV(csvString) {
  var lines = csvString.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  if (lines.length < 2)
    return { imported: 0, skipped: 0, errors: ["CSV has no data rows."] };

  // Parse header
  var headers = lines[0].split(",").map(function (h) {
    return h.trim().toLowerCase().replace(/"/g, "");
  });
  var colIndex = {
    serialNumber: headers.indexOf("serialnumber"),
    model: headers.indexOf("model"),
    condition: headers.indexOf("condition"),
    notes: headers.indexOf("notes"),
  };

  if (colIndex.serialNumber === -1)
    return {
      imported: 0,
      skipped: 0,
      errors: ['CSV missing required "SerialNumber" column.'],
    };

  var imported = 0;
  var skipped = 0;
  var errors = [];

  for (var i = 1; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

    // Simple CSV parse (handles quoted fields)
    var cols = parseCSVLine_(line);

    var sn =
      colIndex.serialNumber >= 0
        ? String(cols[colIndex.serialNumber] || "").trim()
        : "";
    if (!sn) {
      errors.push("Row " + (i + 1) + ": SerialNumber is blank, skipped.");
      continue;
    }

    // Check duplicate
    var existing = getAssetBySerial(sn);
    if (existing) {
      skipped++;
      continue;
    }

    try {
      addAsset({
        serialNumber: sn,
        model: colIndex.model >= 0 ? cols[colIndex.model] || "" : "",
        condition:
          colIndex.condition >= 0 ? cols[colIndex.condition] || "Good" : "Good",
        notes: colIndex.notes >= 0 ? cols[colIndex.notes] || "" : "",
      });
      imported++;
    } catch (e) {
      errors.push("Row " + (i + 1) + " (" + sn + "): " + e.message);
    }
  }

  return { imported: imported, skipped: skipped, errors: errors };
}

/**
 * Simple CSV line parser that handles double-quoted fields.
 */
function parseCSVLine_(line) {
  var result = [];
  var current = "";
  var inQuotes = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === "," && !inQuotes) {
      result.push(current.trim());
      current = "";
    } else {
      current += ch;
    }
  }
  result.push(current.trim());
  return result;
}

// ============================================================
// Loan / Return (LockService-guarded)
// ============================================================

/**
 * Submits a loan transaction.
 * loanData: { borrowerName, borrowerEmail, serialNumber, condition }
 */
function submitLoan(loanData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // wait up to 15 seconds

    var sn = String(loanData.serialNumber || "").trim();
    if (!sn) throw new Error("Serial number is required.");
    if (!loanData.borrowerName) throw new Error("Borrower name is required.");
    if (!loanData.borrowerEmail) throw new Error("Borrower email is required.");

    // Verify asset is Available
    var asset = getAssetBySerial(sn);
    if (!asset) throw new Error('Asset "' + sn + '" not found.');
    if (asset.status !== "Available")
      throw new Error(
        'Asset "' +
          sn +
          '" is not available (current status: ' +
          asset.status +
          ").",
      );

    var now = new Date();
    var txnId = generateTxnId_();

    // Use condition from loan form (admin may have changed it)
    var loanCondition = String(loanData.condition || asset.condition).trim();

    // Append transaction
    var txnSheet = getTransactionsSheet_();
    txnSheet.appendRow([
      txnId,
      asset.serialNumber,
      loanCondition,
      loanData.borrowerName,
      loanData.borrowerEmail,
      now,
      "", // ReturnDate blank
      "Loan",
      "", // AdminNotes
    ]);

    // Update asset status (and condition if changed)
    var updates = { status: "On Loan" };
    if (loanCondition !== asset.condition) {
      updates.condition = loanCondition;
    }
    updateAsset({ serialNumber: sn, updatedFields: updates });

    // Send email receipt
    sendLoanReceipt_({
      transactionId: txnId,
      serialNumber: asset.serialNumber,
      model: asset.model,
      condition: loanCondition,
      borrowerName: loanData.borrowerName,
      borrowerEmail: loanData.borrowerEmail,
      loanDate: Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm:ss",
      ),
    });

    return { success: true, transactionId: txnId };
  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Finds open loans matching a borrower name or email query.
 * Returns array of open loan transaction objects.
 */
function lookupLoansByBorrower(query) {
  var q = String(query || "")
    .trim()
    .toLowerCase();
  if (!q) return [];

  var sheet = getTransactionsSheet_();
  var rows = getSheetData_(sheet);
  var results = [];

  rows.forEach(function (row, idx) {
    var type = String(row[7] || "").trim();
    var returnDate = String(row[6] || "").trim();
    if (type !== "Loan" || returnDate !== "") return; // only open loans

    var name = String(row[3] || "").toLowerCase();
    var email = String(row[4] || "").toLowerCase();
    if (name.indexOf(q) !== -1 || email.indexOf(q) !== -1) {
      results.push({
        transactionId: String(row[0] || "").trim(),
        serialNumber: String(row[1] || "").trim(),
        condition: String(row[2] || "").trim(),
        borrowerName: String(row[3] || "").trim(),
        borrowerEmail: String(row[4] || "").trim(),
        loanDate: row[5]
          ? Utilities.formatDate(
              new Date(row[5]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
        rowIndex: idx + 2, // 1-based sheet row
      });
    }
  });
  return results;
}

/**
 * Finds the open loan for a specific serial number.
 * Returns a single loan object or null.
 */
function lookupLoanBySerial(serialNumber) {
  var sn = String(serialNumber || "")
    .trim()
    .toUpperCase();
  if (!sn) return null;

  var sheet = getTransactionsSheet_();
  var rows = getSheetData_(sheet);

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var type = String(row[7] || "").trim();
    var returnDate = String(row[6] || "").trim();
    var rowSn = String(row[1] || "")
      .trim()
      .toUpperCase();
    if (type === "Loan" && returnDate === "" && rowSn === sn) {
      return {
        transactionId: String(row[0] || "").trim(),
        serialNumber: String(row[1] || "").trim(),
        condition: String(row[2] || "").trim(),
        borrowerName: String(row[3] || "").trim(),
        borrowerEmail: String(row[4] || "").trim(),
        loanDate: row[5]
          ? Utilities.formatDate(
              new Date(row[5]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
        rowIndex: i + 2,
      };
    }
  }
  return null;
}

/**
 * Processes a device return.
 * returnData: { transactionId, serialNumber, borrowerEmail, borrowerName, adminNotes? }
 */
function submitReturn(returnData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var txnId = String(returnData.transactionId || "").trim();
    var sn = String(returnData.serialNumber || "").trim();
    if (!txnId) throw new Error("Transaction ID is required.");
    if (!sn) throw new Error("Serial number is required.");

    var sheet = getTransactionsSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("No transactions found.");

    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var targetRow = -1;
    var loanDate = "";
    var txnData = null;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (
        String(row[0] || "").trim() === txnId &&
        String(row[7] || "").trim() === "Loan" &&
        String(row[6] || "").trim() === ""
      ) {
        targetRow = i + 2;
        loanDate = row[5]
          ? Utilities.formatDate(
              new Date(row[5]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "";
        txnData = row;
        break;
      }
    }

    if (targetRow === -1)
      throw new Error('Open loan transaction "' + txnId + '" not found.');

    var now = new Date();
    var nowStr = Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss",
    );

    // Write ReturnDate and AdminNotes to existing Loan row
    sheet.getRange(targetRow, 7).setValue(now); // col G = ReturnDate
    if (returnData.adminNotes) {
      sheet.getRange(targetRow, 9).setValue(returnData.adminNotes); // col I
    }

    // Append a Return transaction row for full history
    var returnTxnId = generateTxnId_();
    sheet.appendRow([
      returnTxnId,
      String(txnData[1] || ""),
      String(txnData[2] || ""),
      String(txnData[3] || ""),
      String(txnData[4] || ""),
      txnData[5], // original LoanDate
      now, // ReturnDate
      "Return",
      returnData.adminNotes || "",
    ]);

    // Update asset status back to Available
    updateAsset({ serialNumber: sn, updatedFields: { status: "Available" } });

    // Get full asset details for email
    var asset = getAssetBySerial(sn);

    // Send email receipt
    sendReturnReceipt_({
      transactionId: txnId,
      serialNumber: asset ? asset.serialNumber : sn,
      model: asset ? asset.model : "",
      condition: asset ? asset.condition : String(txnData[2] || ""),
      borrowerName: returnData.borrowerName || String(txnData[3] || ""),
      borrowerEmail: returnData.borrowerEmail || String(txnData[4] || ""),
      loanDate: loanDate,
      returnDate: nowStr,
    });

    return { success: true, returnDate: nowStr };
  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}
