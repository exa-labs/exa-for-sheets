/**
 * SheetWriter.gs
 *
 * Manages the "Webset Data" sheet: writes new items, updates existing rows,
 * and deduplicates by Item ID (Column A).
 *
 * Item shape:
 *   { id, name, title, url, status, verifiedAt,
 *     enrichments: { key: value, … },
 *     entity:      { key: value, … } }
 */

/* ------------------------------------------------------------------ */
/*  Constants                                                         */
/* ------------------------------------------------------------------ */

/** @const {string} Name of the target sheet. */
var SHEET_NAME_ = 'Webset Data';

/** @const {string[]} Core columns that always appear first, in order. */
var CORE_COLUMNS_ = ['Item ID', 'Name', 'URL', 'Status', 'Verified At'];

/** @const {string} Final column appended after all dynamic columns. */
var LAST_SYNCED_HEADER_ = 'Last Synced';

/* ------------------------------------------------------------------ */
/*  Private helpers                                                   */
/* ------------------------------------------------------------------ */

/**
 * Returns the "Webset Data" sheet, creating it if it does not exist.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 * @private
 */
function websets_getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_);
  }
  return sheet;
}

/**
 * Builds and/or updates the header row so it contains every column needed
 * to represent the current batch of items.
 *
 * Column order:
 *   1. Core columns  (Item ID, Name, URL, Status, Verified At)
 *   2. Enrichment: {key}  — one per unique enrichment key across all items
 *   3. Entity: {key}      — one per unique entity key (skip id, name, url)
 *   4. Last Synced         — always the final column
 *
 * Existing columns are never removed or reordered; new columns are only
 * appended (before the trailing "Last Synced" column).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  The target sheet.
 * @param {Object[]} items  Array of item objects.
 * @return {string[]} The full, ordered headers array after sync.
 * @private
 */
function websets_syncHeaders_(sheet, items) {
  var existing = websets_getExistingHeaders_(sheet);

  // --- Collect all enrichment & entity keys from the items ----------------
  var enrichmentKeys = {};
  var entityKeys = {};
  var entitySkip = { id: true, name: true, url: true };

  items.forEach(function (item) {
    if (item.enrichments && typeof item.enrichments === 'object') {
      Object.keys(item.enrichments).forEach(function (k) {
        enrichmentKeys[k] = true;
      });
    }
    if (item.entity && typeof item.entity === 'object') {
      Object.keys(item.entity).forEach(function (k) {
        if (!entitySkip[k]) {
          entityKeys[k] = true;
        }
      });
    }
  });

  // --- Build the "desired" set of headers (preserving existing order) -----
  // Start with a set for fast lookup.
  var existingSet = {};
  existing.forEach(function (h) {
    existingSet[h] = true;
  });

  // We will build an array that equals existing (minus trailing Last Synced),
  // then append any new columns, then re-add Last Synced at the end.
  var headers = existing.slice(); // copy

  // Remove "Last Synced" if it is currently the last header so we can
  // re-append it at the very end after any new dynamic columns.
  var lastSyncedIdx = headers.indexOf(LAST_SYNCED_HEADER_);
  if (lastSyncedIdx !== -1) {
    headers.splice(lastSyncedIdx, 1);
  }

  // Ensure core columns are present (in order) at the front.
  // Only add if not already present; we never reorder existing headers.
  if (headers.length === 0) {
    // Brand-new sheet — just set the core columns.
    headers = CORE_COLUMNS_.slice();
  } else {
    CORE_COLUMNS_.forEach(function (col) {
      if (!existingSet[col]) {
        headers.push(col);
      }
    });
  }

  // Rebuild the lookup after potential core-column additions.
  var headersSet = {};
  headers.forEach(function (h) {
    headersSet[h] = true;
  });

  // Append new enrichment columns (sorted for determinism).
  Object.keys(enrichmentKeys)
    .sort()
    .forEach(function (k) {
      var col = 'Enrichment: ' + k;
      if (!headersSet[col]) {
        headers.push(col);
        headersSet[col] = true;
      }
    });

  // Append new entity columns (sorted for determinism).
  Object.keys(entityKeys)
    .sort()
    .forEach(function (k) {
      var col = 'Entity: ' + k;
      if (!headersSet[col]) {
        headers.push(col);
        headersSet[col] = true;
      }
    });

  // Always finish with "Last Synced".
  headers.push(LAST_SYNCED_HEADER_);

  // --- Write the header row -----------------------------------------------
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // Style: bold, blue background, white text.
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  return headers;
}

/**
 * Reads the current header values from row 1 of the given sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {string[]} Array of header strings (may be empty).
 * @private
 */
function websets_getExistingHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    return [];
  }
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
}

/**
 * Builds a Map of Item ID → row number for all existing data rows.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} headers  Current header array (used to confirm col A is
 *   Item ID, though by convention it always is).
 * @param {number} lastRow  The last row with content (from sheet.getLastRow()).
 * @return {Map<string, number>} itemId → 1-based row number.
 * @private
 */
function websets_buildIdIndex_(sheet, headers, lastRow) {
  var index = new Map();
  if (lastRow <= 1) {
    // Only the header row (or empty sheet).
    return index;
  }

  // Read all values in Column A, rows 2 … lastRow.
  var numRows = lastRow - 1;
  var ids = sheet.getRange(2, 1, numRows, 1).getValues(); // [[id], [id], …]

  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0]);
    if (id) {
      index.set(id, i + 2); // row number is 1-based; data starts at row 2
    }
  }

  return index;
}

/* ------------------------------------------------------------------ */
/*  Public API                                                        */
/* ------------------------------------------------------------------ */

/**
 * Synchronises an array of webset items to the "Webset Data" sheet.
 *
 * - Creates the sheet & header row if they don't exist.
 * - Updates rows in-place when an item's ID already exists in Column A.
 * - Appends new rows for items not yet in the sheet.
 * - Freezes the header row and auto-resizes columns on first sync.
 *
 * @param {Object[]} items  Array of item objects.
 *   Each item: { id, name, title, url, status, verifiedAt,
 *                enrichments: {key: value}, entity: {key: value} }
 * @return {{ added: number, updated: number }}
 */
function syncItemsToSheet(items) {
  if (!items || items.length === 0) {
    return { added: 0, updated: 0 };
  }

  var sheet = websets_getOrCreateSheet_();
  var isNewSheet = sheet.getLastColumn() === 0;

  // 1. Sync headers.
  var headers = websets_syncHeaders_(sheet, items);

  // 2. Build dedup index.
  var lastRow = sheet.getLastRow();
  var idIndex = websets_buildIdIndex_(sheet, headers, lastRow);

  // 3. Pre-build a column-index lookup: header name → 0-based index.
  var colMap = {};
  headers.forEach(function (h, i) {
    colMap[h] = i;
  });

  // 4. Prepare row data for each item.
  var now = new Date().toISOString();
  var rowsToAppend = [];
  var rowsToUpdate = []; // { row: <1-based>, values: [ … ] }

  var added = 0;
  var updated = 0;

  items.forEach(function (item) {
    var rowData = new Array(headers.length).fill('');

    // Core columns.
    rowData[colMap['Item ID']] = item.id || '';
    rowData[colMap['Name']] = item.name || item.title || '';
    rowData[colMap['URL']] = item.url || '';
    rowData[colMap['Status']] = item.status || '';
    rowData[colMap['Verified At']] = item.verifiedAt || '';

    // Enrichment columns.
    if (item.enrichments && typeof item.enrichments === 'object') {
      Object.keys(item.enrichments).forEach(function (k) {
        var header = 'Enrichment: ' + k;
        if (colMap[header] !== undefined) {
          var val = item.enrichments[k];
          rowData[colMap[header]] =
            typeof val === 'object' ? JSON.stringify(val) : val;
        }
      });
    }

    // Entity columns (skip id, name, url to avoid duplicates).
    var entitySkip = { id: true, name: true, url: true };
    if (item.entity && typeof item.entity === 'object') {
      Object.keys(item.entity).forEach(function (k) {
        if (entitySkip[k]) return;
        var header = 'Entity: ' + k;
        if (colMap[header] !== undefined) {
          var val = item.entity[k];
          rowData[colMap[header]] =
            typeof val === 'object' ? JSON.stringify(val) : val;
        }
      });
    }

    // Last Synced.
    rowData[colMap[LAST_SYNCED_HEADER_]] = now;

    // Dedup: update or append.
    var itemId = String(item.id || '');
    if (itemId && idIndex.has(itemId)) {
      rowsToUpdate.push({ row: idIndex.get(itemId), values: rowData });
      updated++;
    } else {
      rowsToAppend.push(rowData);
      if (itemId) {
        // Track in index so later duplicates within the same batch update
        // rather than create a second row.
        var nextRow = lastRow + rowsToAppend.length;
        idIndex.set(itemId, nextRow);
      }
      added++;
    }
  });

  // 5. Batch-write updates.
  rowsToUpdate.forEach(function (entry) {
    sheet
      .getRange(entry.row, 1, 1, headers.length)
      .setValues([entry.values]);
  });

  // 6. Batch-append new rows.
  if (rowsToAppend.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet
      .getRange(startRow, 1, rowsToAppend.length, headers.length)
      .setValues(rowsToAppend);
  }

  // 7. Freeze header row.
  sheet.setFrozenRows(1);

  // 8. Auto-resize columns on first sync.
  if (isNewSheet) {
    for (var c = 1; c <= headers.length; c++) {
      sheet.autoResizeColumn(c);
    }
  }

  return { added: added, updated: updated };
}
