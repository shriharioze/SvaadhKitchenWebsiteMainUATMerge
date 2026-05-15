// ============================================================
// 09_Archive.gs
// Quarterly/monthly archive, year-folder organisation,
// cross-archive analytics lookup.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── QUARTERLY ARCHIVE ─────────────────────────────────────────────────────────
/*
  Archives SK_Orders and SK_Wallet for a given month into a new Google
  Spreadsheet, writes Balance Carry Forward snapshots so wallet balances are
  preserved, then deletes the archived rows from the main sheet.

  Runs on the 10th of every month — archives the previous calendar month.
  e.g. May 10 → archives April data.
*/
function archiveMonth(year, month) {
  if (!year || !month || month < 1 || month > 12)
    return {success:false, error:"Invalid year/month"};

  var MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  var pad = function(n) { return n < 10 ? "0"+n : String(n); };
  // Last day of month: day 0 of next month
  var lastDay = new Date(year, month, 0).getDate();
  var qr = {
    from:  year + "-" + pad(month) + "-01",
    to:    year + "-" + pad(month) + "-" + pad(lastDay),
    label: MONTH_NAMES[month - 1] + " " + year
  };

  // Single global lock for the whole archive operation. Without this, a
  // simultaneous order submission could write to SK_Orders between our
  // read and our rebuild, causing data loss.
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30 * 60 * 1000); } catch (e) {
    return {success:false, error:"Could not acquire script lock (system busy). Try again in a minute."};
  }

  try {
    var ss = getSpreadsheet();
    var fmtDate = function(v) {
      return v instanceof Date
        ? Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd")
        : String(v || "").trim().slice(0, 10);
    };

    // ── STEP 1: Create archive spreadsheet IN THE RIGHT YEAR FOLDER ─────────
    var archiveName = "Svaadh Kitchen Archive — " + qr.label;
    var archiveSS   = SpreadsheetApp.create(archiveName);
    var archiveFile = DriveApp.getFileById(archiveSS.getId());

    // Move to: Drive > WebBased Ordering > Archive > <year>/
    var yearFolder  = _getArchiveYearFolder(year);
    if (yearFolder) {
      try {
        var parents = archiveFile.getParents();
        while (parents.hasNext()) {
          var parent = parents.next();
          if (parent.getId() !== yearFolder.getId()) parent.removeFile(archiveFile);
        }
        yearFolder.addFile(archiveFile);
      } catch (e) {
        Logger.log("archiveMonth: could not move file to year folder: " + e.message);
      }
    }

    var log = [];

    // ── STEP 2: Read live SK_Orders into memory ─────────────────────────────
    var ordersWs      = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
    var allOrderData  = ordersWs.getDataRange().getValues();
    var oHeaders      = allOrderData[0];
    var oDateIdx      = oHeaders.indexOf("Order_Date");

    // Partition into archive vs keep
    var toArchiveOrders = [];
    var keepOrders      = [];
    for (var i = 1; i < allOrderData.length; i++) {
      var d = fmtDate(allOrderData[i][oDateIdx]);
      if (d >= qr.from && d <= qr.to) {
        toArchiveOrders.push(allOrderData[i]);
      } else {
        keepOrders.push(allOrderData[i]);
      }
    }

    // ── STEP 3: Write SK_Orders to archive (verify) ─────────────────────────
    if (toArchiveOrders.length > 0) {
      var archiveOrderSheet = archiveSS.getActiveSheet();
      archiveOrderSheet.setName("SK_Orders");
      archiveOrderSheet.getRange(1, 1, 1, oHeaders.length).setValues([oHeaders]);
      archiveOrderSheet.getRange(2, 1, toArchiveOrders.length, oHeaders.length)
                       .setValues(toArchiveOrders);
      SpreadsheetApp.flush();
      var oWritten = archiveOrderSheet.getLastRow() - 1;
      if (oWritten !== toArchiveOrders.length) {
        return {success:false, error:"Order archive verification failed. Expected "
          + toArchiveOrders.length + ", got " + oWritten + ". Nothing deleted from live sheet."};
      }
      log.push(toArchiveOrders.length + " orders archived ✓");
    } else {
      log.push("No orders found for this month.");
    }

    // ── STEP 4: Read live SK_Wallet into memory ─────────────────────────────
    var walletWs      = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
    var allWalletData = walletWs.getDataRange().getValues();
    var wHeaders      = allWalletData[0];
    var wTsIdx        = wHeaders.indexOf("Timestamp");
    var wPhoneIdx     = wHeaders.indexOf("Phone");
    var wNameIdx      = wHeaders.indexOf("Customer_Name");

    var toArchiveWallet = [];
    var keepWallet      = [];
    for (var j = 1; j < allWalletData.length; j++) {
      var ts = allWalletData[j][wTsIdx];
      var wd = fmtDate(ts instanceof Date ? ts : new Date(ts));
      if (wd >= qr.from && wd <= qr.to) {
        toArchiveWallet.push(allWalletData[j]);
      } else {
        keepWallet.push(allWalletData[j]);
      }
    }

    // ── STEP 5: Write SK_Wallet to archive (verify) ─────────────────────────
    if (toArchiveWallet.length > 0) {
      var archiveWalletSheet = archiveSS.insertSheet("SK_Wallet");
      archiveWalletSheet.getRange(1, 1, 1, wHeaders.length).setValues([wHeaders]);
      archiveWalletSheet.getRange(2, 1, toArchiveWallet.length, wHeaders.length)
                        .setValues(toArchiveWallet);
      SpreadsheetApp.flush();
      var wWritten = archiveWalletSheet.getLastRow() - 1;
      if (wWritten !== toArchiveWallet.length) {
        return {success:false, error:"Wallet archive verification failed. Expected "
          + toArchiveWallet.length + ", got " + wWritten + ". Nothing deleted from live sheet."};
      }
      log.push(toArchiveWallet.length + " wallet transactions archived ✓");
    } else {
      log.push("No wallet transactions found for this month.");
    }

    // ── STEP 6: Compute Balance Carry-Forward (BEFORE wallet rebuild) ───────
    var activePhones = {};
    toArchiveWallet.forEach(function(row) {
      var ph   = String(row[wPhoneIdx] || "").trim();
      var name = String(row[wNameIdx]  || "").trim();
      if (ph) activePhones[ph] = name;
    });

    var snapshotCount = 0;
    var snapTime      = new Date();
    var refId         = "ARCHIVE-" + year + "-" + pad(month);
    var carryFwdRows  = [];
    Object.keys(activePhones).forEach(function(ph) {
      var balance = _calculateWalletBalance(ph);
      if (balance > 0) {
        var newRow = new Array(wHeaders.length).fill("");
        wHeaders.forEach(function(h, idx) {
          if (h === "Phone")          newRow[idx] = ph;
          else if (h === "Customer_Name") newRow[idx] = activePhones[ph];
          else if (h === "Txn_Type")  newRow[idx] = "Balance Carry Forward";
          else if (h === "Amount")    newRow[idx] = balance;
          else if (h === "Verified")  newRow[idx] = "TRUE";
          else if (h === "Reference_ID") newRow[idx] = refId;
          else if (h === "Timestamp") newRow[idx] = snapTime;
        });
        carryFwdRows.push(newRow);
        snapshotCount++;
      }
    });
    if (snapshotCount > 0) log.push(snapshotCount + " balance snapshots prepared ✓");

    // ── STEP 7: REBUILD live sheets atomically ──────────────────────────────
    // Critical fix for the "half-deleted" bug: clear data range and re-write
    // only kept rows in one operation instead of N deleteRow() calls.
    function rebuildSheet(ws, headers, keepRows, appendRows) {
      var allKeep = keepRows.concat(appendRows || []);
      var lastRow = ws.getLastRow();
      var lastCol = ws.getLastColumn();
      if (lastRow > 1) {
        ws.getRange(2, 1, lastRow - 1, Math.max(lastCol, headers.length)).clearContent();
      }
      if (allKeep.length > 0) {
        ws.getRange(2, 1, allKeep.length, headers.length).setValues(allKeep);
      }
      SpreadsheetApp.flush();
      var nowRows = ws.getLastRow() - 1;
      return nowRows === allKeep.length
        ? {success:true, written: nowRows}
        : {success:false, expected: allKeep.length, actual: nowRows};
    }

    if (toArchiveOrders.length > 0) {
      var oRebuild = rebuildSheet(ordersWs, oHeaders, keepOrders);
      if (!oRebuild.success) {
        return {success:false,
          error:"Order rebuild verification failed. Expected " + oRebuild.expected
                + ", got " + oRebuild.actual + ". Archive file IS created — please verify manually before retrying.",
          archiveUrl: archiveSS.getUrl()};
      }
      log.push(toArchiveOrders.length + " order rows removed from live sheet ✓");
    }

    if (toArchiveWallet.length > 0 || carryFwdRows.length > 0) {
      var wRebuild = rebuildSheet(walletWs, wHeaders, keepWallet, carryFwdRows);
      if (!wRebuild.success) {
        return {success:false,
          error:"Wallet rebuild verification failed. Expected " + wRebuild.expected
                + ", got " + wRebuild.actual + ". Archive file IS created — please verify manually before retrying.",
          archiveUrl: archiveSS.getUrl()};
      }
      log.push(toArchiveWallet.length + " wallet rows removed + " + carryFwdRows.length + " carry-forward rows added ✓");
    }

    return {
      success:        true,
      archiveName:    archiveName,
      archiveUrl:     archiveSS.getUrl(),
      archiveFolder:  yearFolder ? yearFolder.getName() : "(My Drive)",
      ordersArchived: toArchiveOrders.length,
      walletArchived: toArchiveWallet.length,
      snapshots:      snapshotCount,
      log:            log
    };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}
function _getArchiveYearFolder(year) {
  var yearStr = String(year);
  try {
    var archiveFolder = null;
    var props = PropertiesService.getScriptProperties();
    var configuredId = props.getProperty("ARCHIVE_PARENT_FOLDER_ID");
    if (configuredId) {
      try { archiveFolder = DriveApp.getFolderById(configuredId); } catch(_) {}
    }

    if (!archiveFolder) {
      var rootFolders = DriveApp.getFoldersByName("WebBased Ordering");
      var webOrdering = null;
      if (rootFolders.hasNext()) webOrdering = rootFolders.next();
      if (!webOrdering) {
        Logger.log("_getArchiveYearFolder: 'WebBased Ordering' folder not found. Archive will stay in My Drive root.");
        return null;
      }
      var archiveFolders = webOrdering.getFoldersByName("Archive");
      archiveFolder = archiveFolders.hasNext() ? archiveFolders.next() : webOrdering.createFolder("Archive");
      try { props.setProperty("ARCHIVE_PARENT_FOLDER_ID", archiveFolder.getId()); } catch(_) {}
    }

    var yearFolders = archiveFolder.getFoldersByName(yearStr);
    if (yearFolders.hasNext()) return yearFolders.next();
    return archiveFolder.createFolder(yearStr);
  } catch (e) {
    Logger.log("_getArchiveYearFolder error: " + e.message);
    return null;
  }
}
// ─────────────────────────────────────────────────────────────────────────────
// ARCHIVE LOOKUP — for analytics across archived data
// ─────────────────────────────────────────────────────────────────────────────
/**
 * Lists all archive spreadsheet files whose MONTH overlaps the given date range.
 * Filename pattern: "Svaadh Kitchen Archive — <MMM> <YYYY>" (e.g. "...— Apr 2026")
 */
function _listArchiveFilesInRange(dateFrom, dateTo) {
  var out = [];
  try {
    var props = PropertiesService.getScriptProperties();
    var configuredId = props.getProperty("ARCHIVE_PARENT_FOLDER_ID");
    var archiveFolder = null;
    if (configuredId) {
      try { archiveFolder = DriveApp.getFolderById(configuredId); } catch(_) {}
    }
    if (!archiveFolder) {
      var rootFolders = DriveApp.getFoldersByName("WebBased Ordering");
      if (!rootFolders.hasNext()) return out;
      var webOrdering = rootFolders.next();
      var archiveFolders = webOrdering.getFoldersByName("Archive");
      if (!archiveFolders.hasNext()) return out;
      archiveFolder = archiveFolders.next();
    }

    var MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    var monthIdx = function(m) { return MONTH_NAMES.indexOf(m); };
    var pad = function(n) { return n < 10 ? "0"+n : String(n); };

    // Pattern: "Svaadh Kitchen Archive — <Mon> <Year>"
    var fileNameRe = /Archive\s+—?\s*([A-Z][a-z]{2})\s+(\d{4})/;

    var processFile = function(file) {
      var name = file.getName();
      var m = name.match(fileNameRe);
      if (!m) return;
      var mIdx = monthIdx(m[1]);
      if (mIdx < 0) return;
      var yr = parseInt(m[2], 10);
      var lastDay = new Date(yr, mIdx + 1, 0).getDate();
      var rFrom = yr + "-" + pad(mIdx + 1) + "-01";
      var rTo   = yr + "-" + pad(mIdx + 1) + "-" + pad(lastDay);
      if (rTo < dateFrom || rFrom > dateTo) return;
      out.push({ file: file, year: yr, month: mIdx + 1, from: rFrom, to: rTo });
    };

    var yearFolders = archiveFolder.getFolders();
    while (yearFolders.hasNext()) {
      var yearFolder = yearFolders.next();
      var files = yearFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
      while (files.hasNext()) processFile(files.next());
    }
    var looseFiles = archiveFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (looseFiles.hasNext()) processFile(looseFiles.next());

    out.sort(function(a, b) {
      if (a.year !== b.year) return a.year - b.year;
      return a.month - b.month;
    });
  } catch (e) {
    Logger.log("_listArchiveFilesInRange error: " + e.message);
  }
  return out;
}
function _readArchivedOrdersInRange(dateFrom, dateTo) {
  var archives = _listArchiveFilesInRange(dateFrom, dateTo);
  if (!archives.length) return [];

  var cache = CacheService.getScriptCache();
  var fmtDate = function(v) {
    return v instanceof Date
      ? Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd")
      : String(v || "").trim().slice(0, 10);
  };

  var allRows = [];
  archives.forEach(function(meta) {
    var cacheKey = "arch_orders_" + meta.file.getId();
    var cached;
    try { cached = cache.get(cacheKey); } catch(_) {}
    var rows;
    if (cached) {
      try { rows = JSON.parse(cached); } catch(_) { rows = null; }
    }
    if (!rows) {
      try {
        var aSS = SpreadsheetApp.openById(meta.file.getId());
        var sheet = aSS.getSheetByName("SK_Orders");
        if (!sheet) return;
        var data = sheet.getDataRange().getValues();
        if (data.length < 2) return;
        var headers = data[0];
        rows = [];
        for (var r = 1; r < data.length; r++) {
          var obj = {};
          for (var c = 0; c < headers.length; c++) obj[headers[c]] = data[r][c];
          rows.push(obj);
        }
        try {
          var serialised = JSON.stringify(rows);
          if (serialised.length <= 95 * 1024) cache.put(cacheKey, serialised, 600);
        } catch(_) {}
      } catch (e) {
        Logger.log("_readArchivedOrdersInRange: could not read " + meta.file.getName() + ": " + e.message);
        return;
      }
    }
    rows.forEach(function(row) {
      var d = fmtDate(row.Order_Date);
      if (d >= dateFrom && d <= dateTo) allRows.push(row);
    });
  });

  return allRows;
}
function getOrdersInRangeWithArchive(dateFrom, dateTo) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var fmtDate = function(v) {
    return v instanceof Date
      ? Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd")
      : String(v || "").trim().slice(0, 10);
  };
  var liveRows = getAllRows(ws).filter(function(r) {
    var d = fmtDate(r.Order_Date);
    return d >= dateFrom && d <= dateTo;
  });
  var archivedRows = _readArchivedOrdersInRange(dateFrom, dateTo);
  var seen = {};
  var combined = [];
  archivedRows.concat(liveRows).forEach(function(r) {
    var id = String(r.Submission_ID || "").trim();
    if (id && seen[id]) return;
    if (id) seen[id] = true;
    combined.push(r);
  });
  return combined;
}
// Called by admin UI — wraps archiveMonth with PIN check (handled by router)
function triggerManualArchive(body) {
  var year  = parseInt(body.year);
  var month = parseInt(body.month);
  if (!year || !month) return {success:false, error:"year and month required"};
  return archiveMonth(year, month);
}
// ── Time-based trigger: auto-archive previous month on the 10th ──────────
// Run setupMonthlyArchiveTrigger() once from Apps Script editor to register.
// Also registered from admin UI via the setupQuarterlyArchiveTrigger action name (kept for compat).
function setupMonthlyArchiveTrigger() {
  // Remove any existing trigger for runScheduledArchive
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "runScheduledArchive") ScriptApp.deleteTrigger(t);
  });
  // Fire on the 10th of every month at 21:00 UTC = 02:30 IST (safe off-peak window)
  ScriptApp.newTrigger("runScheduledArchive")
    .timeBased()
    .onMonthDay(10)
    .atHour(21)
    .create();
  return "Monthly archive trigger set — fires on the 10th of every month at ~2 AM IST.";
}
// Keep old name working (admin UI may still call this action)
function setupQuarterlyArchiveTrigger() {
  return setupMonthlyArchiveTrigger();
}
function runScheduledArchive() {
  // Archive the previous calendar month.
  // e.g. trigger fires May 10 → archive April (month 4, year this year)
  var now   = new Date();
  var year  = now.getFullYear();
  var month = now.getMonth() + 1; // 1–12, current month

  // Previous month
  var archiveMonth_ = month - 1;
  var archiveYear   = year;
  if (archiveMonth_ === 0) { archiveMonth_ = 12; archiveYear = year - 1; }

  var result = archiveMonth(archiveYear, archiveMonth_);
  Logger.log("Monthly archive result: " + JSON.stringify(result));
}
