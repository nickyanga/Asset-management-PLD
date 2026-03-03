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
  var subject =
    "PLD Loan Receipt \u2013 " + data.assetTag + " " + data.deviceName;
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
    "Asset Tag      : " + data.assetTag,
    "Serial Number  : " + data.serialNumber,
    "Device Name    : " + data.deviceName,
    "Model          : " + data.model,
    "Category       : " + data.category,
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
  var subject =
    "PLD Return Receipt \u2013 " + data.assetTag + " " + data.deviceName;

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
    "Asset Tag      : " + data.assetTag,
    "Serial Number  : " + data.serialNumber,
    "Device Name    : " + data.deviceName,
    "Model          : " + data.model,
    "Category       : " + data.category,
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
    var status = String(row[6] || "").trim();
    if (status === "Available") {
      assets.push({
        assetTag: String(row[0] || "").trim(),
        serialNumber: String(row[1] || "").trim(),
        deviceName: String(row[2] || "").trim(),
        model: String(row[3] || "").trim(),
        category: String(row[4] || "").trim(),
        condition: String(row[5] || "").trim(),
        status: status,
        notes: String(row[7] || "").trim(),
      });
    }
  });
  return assets;
}

/**
 * Returns a single asset object by AssetTag, or null if not found.
 */
function getAssetByTag(assetTag) {
  var tag = String(assetTag || "")
    .trim()
    .toUpperCase();
  var sheet = getAssetsSheet_();
  var rows = getSheetData_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (
      String(rows[i][0] || "")
        .trim()
        .toUpperCase() === tag
    ) {
      return {
        assetTag: String(rows[i][0] || "").trim(),
        serialNumber: String(rows[i][1] || "").trim(),
        deviceName: String(rows[i][2] || "").trim(),
        model: String(rows[i][3] || "").trim(),
        category: String(rows[i][4] || "").trim(),
        condition: String(rows[i][5] || "").trim(),
        status: String(rows[i][6] || "").trim(),
        notes: String(rows[i][7] || "").trim(),
        dateAdded: rows[i][8]
          ? Utilities.formatDate(
              new Date(rows[i][8]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
        lastModified: rows[i][9]
          ? Utilities.formatDate(
              new Date(rows[i][9]),
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
        assetTag: String(row[0] || "").trim(),
        serialNumber: String(row[1] || "").trim(),
        deviceName: String(row[2] || "").trim(),
        model: String(row[3] || "").trim(),
        category: String(row[4] || "").trim(),
        condition: String(row[5] || "").trim(),
        status: String(row[6] || "").trim(),
        notes: String(row[7] || "").trim(),
        dateAdded: row[8]
          ? Utilities.formatDate(
              new Date(row[8]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd",
            )
          : "",
        lastModified: row[9]
          ? Utilities.formatDate(
              new Date(row[9]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
      };
    })
    .filter(function (a) {
      return a.assetTag !== "";
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
        assetTag: String(row[1] || "").trim(),
        deviceName: String(row[2] || "").trim(),
        category: String(row[3] || "").trim(),
        condition: String(row[4] || "").trim(),
        borrowerName: String(row[5] || "").trim(),
        borrowerEmail: String(row[6] || "").trim(),
        loanDate: row[7]
          ? Utilities.formatDate(
              new Date(row[7]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
        returnDate: row[8]
          ? Utilities.formatDate(
              new Date(row[8]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd HH:mm:ss",
            )
          : "",
        type: String(row[9] || "").trim(),
        adminNotes: String(row[10] || "").trim(),
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
 * assetData: { assetTag, serialNumber, deviceName, model, category, condition, notes }
 */
function addAsset(assetData) {
  var tag = String(assetData.assetTag || "").trim();
  if (!tag) throw new Error("AssetTag is required.");

  // Check for duplicate
  var existing = getAssetByTag(tag);
  if (existing) throw new Error('Asset tag "' + tag + '" already exists.');

  var now = new Date();
  var sheet = getAssetsSheet_();
  sheet.appendRow([
    tag,
    String(assetData.serialNumber || "").trim(),
    String(assetData.deviceName || "").trim(),
    String(assetData.model || "").trim(),
    String(assetData.category || "").trim(),
    String(assetData.condition || "Good").trim(),
    "Available",
    String(assetData.notes || "").trim(),
    now,
    now,
  ]);
  return { success: true, assetTag: tag };
}

/**
 * Updates specified fields for an existing asset.
 * Called from frontend as updateAsset({ assetTag, updatedFields }).
 * params: { assetTag, updatedFields: { serialNumber?, deviceName?, model?, category?, condition?, status?, notes? } }
 */
function updateAsset(params) {
  var assetTag = params.assetTag;
  var updatedFields = params.updatedFields;
  var tag = String(assetTag || "")
    .trim()
    .toUpperCase();
  var sheet = getAssetsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("No assets found.");

  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (var i = 0; i < data.length; i++) {
    if (
      String(data[i][0] || "")
        .trim()
        .toUpperCase() === tag
    ) {
      var rowNum = i + 2;
      var row = data[i];

      // Map field names to column indices (0-based)
      var fieldMap = {
        serialNumber: 1,
        deviceName: 2,
        model: 3,
        category: 4,
        condition: 5,
        status: 6,
        notes: 7,
      };

      Object.keys(fieldMap).forEach(function (field) {
        if (updatedFields.hasOwnProperty(field)) {
          row[fieldMap[field]] = updatedFields[field];
        }
      });
      row[9] = new Date(); // LastModified

      sheet.getRange(rowNum, 1, 1, 10).setValues([row]);
      return { success: true };
    }
  }
  throw new Error('Asset tag "' + assetTag + '" not found.');
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
    assetTag: headers.indexOf("assettag"),
    serialNumber: headers.indexOf("serialnumber"),
    deviceName: headers.indexOf("devicename"),
    model: headers.indexOf("model"),
    category: headers.indexOf("category"),
    condition: headers.indexOf("condition"),
    notes: headers.indexOf("notes"),
  };

  if (colIndex.assetTag === -1)
    return {
      imported: 0,
      skipped: 0,
      errors: ['CSV missing required "AssetTag" column.'],
    };

  var imported = 0;
  var skipped = 0;
  var errors = [];

  for (var i = 1; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

    // Simple CSV parse (handles quoted fields)
    var cols = parseCSVLine_(line);

    var tag =
      colIndex.assetTag >= 0
        ? String(cols[colIndex.assetTag] || "").trim()
        : "";
    if (!tag) {
      errors.push("Row " + (i + 1) + ": AssetTag is blank, skipped.");
      continue;
    }

    // Check duplicate
    var existing = getAssetByTag(tag);
    if (existing) {
      skipped++;
      continue;
    }

    try {
      addAsset({
        assetTag: tag,
        serialNumber:
          colIndex.serialNumber >= 0 ? cols[colIndex.serialNumber] || "" : "",
        deviceName:
          colIndex.deviceName >= 0 ? cols[colIndex.deviceName] || "" : "",
        model: colIndex.model >= 0 ? cols[colIndex.model] || "" : "",
        category: colIndex.category >= 0 ? cols[colIndex.category] || "" : "",
        condition:
          colIndex.condition >= 0 ? cols[colIndex.condition] || "Good" : "Good",
        notes: colIndex.notes >= 0 ? cols[colIndex.notes] || "" : "",
      });
      imported++;
    } catch (e) {
      errors.push("Row " + (i + 1) + " (" + tag + "): " + e.message);
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
 * loanData: { borrowerName, borrowerEmail, assetTag }
 */
function submitLoan(loanData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // wait up to 15 seconds

    var tag = String(loanData.assetTag || "").trim();
    if (!tag) throw new Error("Asset tag is required.");
    if (!loanData.borrowerName) throw new Error("Borrower name is required.");
    if (!loanData.borrowerEmail) throw new Error("Borrower email is required.");

    // Verify asset is Available
    var asset = getAssetByTag(tag);
    if (!asset) throw new Error('Asset "' + tag + '" not found.');
    if (asset.status !== "Available")
      throw new Error(
        'Asset "' +
          tag +
          '" is not available (current status: ' +
          asset.status +
          ").",
      );

    var now = new Date();
    var txnId = generateTxnId_();

    // Append transaction
    var txnSheet = getTransactionsSheet_();
    txnSheet.appendRow([
      txnId,
      asset.assetTag,
      asset.deviceName,
      asset.category,
      asset.condition,
      loanData.borrowerName,
      loanData.borrowerEmail,
      now,
      "", // ReturnDate blank
      "Loan",
      "", // AdminNotes
    ]);

    // Update asset status
    updateAsset({ assetTag: tag, updatedFields: { status: "On Loan" } });

    // Send email receipt
    sendLoanReceipt_({
      transactionId: txnId,
      assetTag: asset.assetTag,
      serialNumber: asset.serialNumber,
      deviceName: asset.deviceName,
      model: asset.model,
      category: asset.category,
      condition: asset.condition,
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
    var type = String(row[9] || "").trim();
    var returnDate = String(row[8] || "").trim();
    if (type !== "Loan" || returnDate !== "") return; // only open loans

    var name = String(row[5] || "").toLowerCase();
    var email = String(row[6] || "").toLowerCase();
    if (name.indexOf(q) !== -1 || email.indexOf(q) !== -1) {
      results.push({
        transactionId: String(row[0] || "").trim(),
        assetTag: String(row[1] || "").trim(),
        deviceName: String(row[2] || "").trim(),
        category: String(row[3] || "").trim(),
        condition: String(row[4] || "").trim(),
        borrowerName: String(row[5] || "").trim(),
        borrowerEmail: String(row[6] || "").trim(),
        loanDate: row[7]
          ? Utilities.formatDate(
              new Date(row[7]),
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
 * Finds the open loan for a specific asset tag.
 * Returns a single loan object or null.
 */
function lookupLoanByAssetTag(assetTag) {
  var tag = String(assetTag || "")
    .trim()
    .toUpperCase();
  if (!tag) return null;

  var sheet = getTransactionsSheet_();
  var rows = getSheetData_(sheet);

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var type = String(row[9] || "").trim();
    var returnDate = String(row[8] || "").trim();
    var rowTag = String(row[1] || "")
      .trim()
      .toUpperCase();
    if (type === "Loan" && returnDate === "" && rowTag === tag) {
      return {
        transactionId: String(row[0] || "").trim(),
        assetTag: String(row[1] || "").trim(),
        deviceName: String(row[2] || "").trim(),
        category: String(row[3] || "").trim(),
        condition: String(row[4] || "").trim(),
        borrowerName: String(row[5] || "").trim(),
        borrowerEmail: String(row[6] || "").trim(),
        loanDate: row[7]
          ? Utilities.formatDate(
              new Date(row[7]),
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
 * returnData: { transactionId, assetTag, borrowerEmail, borrowerName, adminNotes? }
 */
function submitReturn(returnData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var txnId = String(returnData.transactionId || "").trim();
    var tag = String(returnData.assetTag || "").trim();
    if (!txnId) throw new Error("Transaction ID is required.");
    if (!tag) throw new Error("Asset tag is required.");

    var sheet = getTransactionsSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("No transactions found.");

    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    var targetRow = -1;
    var loanDate = "";
    var txnData = null;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (
        String(row[0] || "").trim() === txnId &&
        String(row[9] || "").trim() === "Loan" &&
        String(row[8] || "").trim() === ""
      ) {
        targetRow = i + 2;
        loanDate = row[7]
          ? Utilities.formatDate(
              new Date(row[7]),
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
    sheet.getRange(targetRow, 9).setValue(now); // col I = ReturnDate
    if (returnData.adminNotes) {
      sheet.getRange(targetRow, 11).setValue(returnData.adminNotes); // col K
    }

    // Append a Return transaction row for full history
    var returnTxnId = generateTxnId_();
    sheet.appendRow([
      returnTxnId,
      String(txnData[1] || ""),
      String(txnData[2] || ""),
      String(txnData[3] || ""),
      String(txnData[4] || ""),
      String(txnData[5] || ""),
      String(txnData[6] || ""),
      txnData[7], // original LoanDate
      now, // ReturnDate
      "Return",
      returnData.adminNotes || "",
    ]);

    // Update asset status back to Available
    updateAsset({ assetTag: tag, updatedFields: { status: "Available" } });

    // Get full asset details for email
    var asset = getAssetByTag(tag);

    // Send email receipt
    sendReturnReceipt_({
      transactionId: txnId,
      assetTag: tag,
      serialNumber: asset ? asset.serialNumber : "",
      deviceName: asset ? asset.deviceName : String(txnData[2] || ""),
      model: asset ? asset.model : "",
      category: asset ? asset.category : String(txnData[3] || ""),
      condition: asset ? asset.condition : String(txnData[4] || ""),
      borrowerName: returnData.borrowerName || String(txnData[5] || ""),
      borrowerEmail: returnData.borrowerEmail || String(txnData[6] || ""),
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
