/***** MASTER-BOUND CONFIG *****/
const AREA_SHEETS = [
  { name: "DRYDEN",     id: "1hq5bsfev_cQnLu6qKh9GVv3A8ylZMmaFilIG5ee1PZ0" },
  { name: "KENORA",     id: "1RtL_UAxCQUREAUKK8afq0dtosl4WDo2wA39BNDtNb5o" },
  { name: "FORT",       id: "1Lc8txfpctWwWnzyZ8hqMmbfG3nqgiVtc4pdZiU_xizw" },
  { name: "NORTHEAST",  id: "1oubi5xgvznfNE50Zldn1lRO3unIHbyY2Dz6RcbSDbJ0" },
  { name: "NORTHWEST",  id: "16B-F87eH1nK4TFYBkjdtbKIw3Zby9ZVIje5R9Q1ig9M" },
  { name: "REDLAKE",    id: "1pI_J7gqfQUQXBUTOMxExl3f2bOC5sBv0DyU3GpZMi80" },
  { name: "SIOUX",      id: "1UTN3BHglFoi3vvNNwBcKjZqIpZQgg3c7zqm5ZOFMlLM" },
  { name: "THOMPSON",   id: "1OaQSQgTPnoNgMk888HXHVsJ9naGk8KiEsLZecqCkqqg" },
  { name: "TBAY",       id: "1RJ3GVKRYn9whzxcWC4qP2IcFpfxm-rYm9Tp89cAxKvc" },
];

// Names / layout (identical across all area sheets)
const SOURCE_EQUIPMENT_TAB = "EQUIPMENT";
const SOURCE_OTHER_TAB     = "ALL OTHER";
const SOURCE_START_ROW     = 4;

// Master tabs
const MASTER_AREAS_TAB = "ALL AREAS"; // EQUIPMENT -> ALL AREAS
const MASTER_OLD_TAB   = "ALL OLD";   // ALL OTHER -> ALL OLD
const MASTER_FORM_TAB  = "FORM";

// Columns (1-based)
const COL_A = 1;
const COL_B = 2;
const COL_C = 3;
const COL_D = 4;
const COL_E = 5;     // E
const COL_F = 6;     // F
const COL_H = 8;     // H
const COL_J = 10;    // J (serial in master)
const COL_N = 14;    // N (serial in area sheets)
const COL_BP = 68;   // BP
const WIDTH_EH = 4;  // E:H width

// Reads
// EQUIPMENT: read E..N (E:H + N) => width 10 (E=5 .. N=14)
const EQUIP_READ_COL = COL_E;
const EQUIP_READ_WIDTH = (COL_N - COL_E + 1); // 10

// ALL OTHER: read A..N (includes A, B, E:H, N)
const OTHER_READ_COL = COL_A;
const OTHER_READ_WIDTH = COL_N; // 14

// Delete trigger
const DELETE_TRIGGER_VALUE = "$$$";

// Master headers
const MASTER_HEADER_ROWS = 1;
const FORM_START_ROW = 2;
const AREAS_MATCH_START_ROW = 7;

// RangeList chunking
const RANGE_LIST_CHUNK = 200;

/***** ENTRYPOINT (set your time trigger to this) *****/
function masterSyncAreasToMaster() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    console.log("masterSyncAreasToMaster: Could not obtain ScriptLock (another run in progress).");
    return;
  }

  try {
    const masterSS = SpreadsheetApp.getActive(); // because this is bound to master
    const masterAreasSheet = mustGetSheet_(masterSS, MASTER_AREAS_TAB);
    const masterOldSheet   = mustGetSheet_(masterSS, MASTER_OLD_TAB);
    const masterFormSheet  = mustGetSheet_(masterSS, MASTER_FORM_TAB);

    // Build master serial -> row maps once
    const areasSerialToRow = buildSerialToRowMap_(masterAreasSheet);
    const oldSerialToRow   = buildSerialToRowMap_(masterOldSheet);

    // Accumulate updates by master row
    const updatesAreasByRow = new Map(); // row -> [e,f,g,h] (nonblank merges)
    const updatesOldByRow   = new Map(); // row -> [e,f,g,h]

    // Accumulate deletes (serials)
    const deleteSerials = [];

    // For clearing source cells later (after master writes, before master deletes)
    const clearPlans = []; // one per area: {name, ss, equipSheet, otherSheet, equipClearRows, otherClearRows, triggerRows}

    // === 1) Read/snapshot every area sheet and accumulate intended actions ===
    for (const area of AREA_SHEETS) {
      const areaSS = SpreadsheetApp.openById(area.id);
      const equipSheet = mustGetSheet_(areaSS, SOURCE_EQUIPMENT_TAB);
      const otherSheet = mustGetSheet_(areaSS, SOURCE_OTHER_TAB);

      const plan = {
        name: area.name,
        ss: areaSS,
        equipSheet,
        otherSheet,
        equipClearRows: [],
        otherClearRows: [],
        triggerRows: [],
      };

      // --- EQUIPMENT -> master ALL AREAS updates ---
      accumulateEquipmentUpdates_(
        area.name,
        equipSheet,
        areasSerialToRow,
        updatesAreasByRow,
        plan.equipClearRows
      );

      // --- ALL OTHER -> master ALL OLD updates + $$$ triggers ---
      accumulateOtherUpdatesAndTriggers_(
        area.name,
        otherSheet,
        oldSerialToRow,
        updatesOldByRow,
        plan.otherClearRows,
        plan.triggerRows,
        deleteSerials
      );

      clearPlans.push(plan);
    }

    // === 2) Apply writes to master (no row shifting) ===
    applyRowUpdates_(masterAreasSheet, updatesAreasByRow, "MASTER ALL AREAS");
    applyRowUpdates_(masterOldSheet,   updatesOldByRow,   "MASTER ALL OLD");

    // === 3) Clear source E:H and clear "$$$" triggers (do this BEFORE master deletions) ===
    for (const plan of clearPlans) {
      // Clear $$$ triggers in ALL OTHER column A
      if (plan.triggerRows.length) {
        const a1 = plan.triggerRows.map(r => `A${r}`);
        clearA1InChunks_(plan.otherSheet, a1);
      }

      // Clear E:H in EQUIPMENT
      if (plan.equipClearRows.length) {
        const a1 = plan.equipClearRows.map(r => `E${r}:H${r}`);
        clearA1InChunks_(plan.equipSheet, a1);
      }

      // Clear E:H in ALL OTHER
      if (plan.otherClearRows.length) {
        const a1 = plan.otherClearRows.map(r => `E${r}:H${r}`);
        clearA1InChunks_(plan.otherSheet, a1);
      }
    }

    // === 4) Process FORM-driven add/delete in ALL AREAS (after clears, before ALL OLD deletions) ===
    processFormRows_(masterAreasSheet, masterFormSheet);

    // === 5) Delete rows in master ALL OLD bottom-up (row shifting happens here, after clears) ===
    if (deleteSerials.length) {
      const rowsToDelete = [];
      for (const s of deleteSerials) {
        const row = oldSerialToRow.get(s);
        if (row) rowsToDelete.push(row);
      }

      const uniqueRowsDesc = [...new Set(rowsToDelete)].sort((a, b) => b - a);
      for (const r of uniqueRowsDesc) {
        masterOldSheet.deleteRow(r);
      }

      console.log(`Deleted ${uniqueRowsDesc.length} row(s) from ${MASTER_OLD_TAB}.`);
    }

    console.log("masterSyncAreasToMaster complete.");

  } catch (err) {
    console.error("masterSyncAreasToMaster error:", err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/***** ACCUMULATION HELPERS *****/
function accumulateEquipmentUpdates_(areaName, sheet, masterSerialToRow, updatesByRow, clearRowsOut) {
  const lastRow = sheet.getLastRow();
  if (lastRow < SOURCE_START_ROW) return;

  const numRows = lastRow - SOURCE_START_ROW + 1;

  // Snapshot E..N (E:H + serial N)
  const data = sheet.getRange(SOURCE_START_ROW, EQUIP_READ_COL, numRows, EQUIP_READ_WIDTH).getValues();
  // Indices within this snapshot:
  // E=0, F=1, G=2, H=3, N=9

  for (let i = 0; i < data.length; i++) {
    const row = SOURCE_START_ROW + i;

    const e = data[i][0];
    const f = data[i][1];
    const g = data[i][2];
    const h = data[i][3];
    const hasAny = (e !== "" && e !== null) || (f !== "" && f !== null) || (g !== "" && g !== null) || (h !== "" && h !== null);
    if (!hasAny) continue;

    const serial = String(data[i][9] ?? "").trim();
    const masterRow = masterSerialToRow.get(serial);
    if (!masterRow) {
      console.log(`[${areaName} EQUIPMENT] Serial not found in master (unexpected): ${serial} at source row ${row}`);
      continue;
    }

    mergeNonblankUpdate_(updatesByRow, masterRow, [e, f, g, h]);
    clearRowsOut.push(row);
  }
}

function accumulateOtherUpdatesAndTriggers_(areaName, sheet, masterSerialToRow, updatesByRow, clearRowsOut, triggerRowsOut, deleteSerialsOut) {
  const lastRow = sheet.getLastRow();
  if (lastRow < SOURCE_START_ROW) return;

  const numRows = lastRow - SOURCE_START_ROW + 1;

  // Snapshot A..N
  const data = sheet.getRange(SOURCE_START_ROW, OTHER_READ_COL, numRows, OTHER_READ_WIDTH).getValues();
  // Indices within this snapshot:
  // A=0, B=1, E=4, F=5, G=6, H=7, N=13

  for (let i = 0; i < data.length; i++) {
    const row = SOURCE_START_ROW + i;

    // $$$ triggers
    const a = String(data[i][0] ?? "");
    if (a === DELETE_TRIGGER_VALUE) {
      triggerRowsOut.push(row);

      const b = String(data[i][1] ?? "").trim();
      if (b && !b.startsWith("*")) {
        const serial = String(data[i][13] ?? "").trim();
        if (serial) {
          // Serial is guaranteed to exist; still safe to collect
          deleteSerialsOut.push(serial);
        } else {
          console.log(`[${areaName} ALL OTHER $$$] Blank serial at row ${row} (unexpected).`);
        }
      }
    }

    // E:H updates -> master ALL OLD
    const e = data[i][4];
    const f = data[i][5];
    const g = data[i][6];
    const h = data[i][7];
    const hasAny = (e !== "" && e !== null) || (f !== "" && f !== null) || (g !== "" && g !== null) || (h !== "" && h !== null);
    if (!hasAny) continue;

    const serial = String(data[i][13] ?? "").trim();
    const masterRow = masterSerialToRow.get(serial);
    if (!masterRow) {
      console.log(`[${areaName} ALL OTHER] Serial not found in master (unexpected): ${serial} at source row ${row}`);
      continue;
    }

    mergeNonblankUpdate_(updatesByRow, masterRow, [e, f, g, h]);
    clearRowsOut.push(row);
  }
}

/***** MASTER WRITE HELPERS *****/
function buildSerialToRowMap_(masterSheet) {
  const startRow = MASTER_HEADER_ROWS + 1;
  const lastRow = masterSheet.getLastRow();
  const map = new Map();
  if (lastRow < startRow) return map;

  const vals = masterSheet.getRange(startRow, COL_J, lastRow - startRow + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const serial = String(vals[i][0] ?? "").trim();
    if (!serial) continue;
    // user said no duplicates; if duplicates ever appear, later ones would overwrite
    map.set(serial, startRow + i);
  }
  return map;
}

function mergeNonblankUpdate_(updatesByRow, masterRow, newValsEH) {
  if (!updatesByRow.has(masterRow)) {
    updatesByRow.set(masterRow, ["", "", "", ""]);
  }
  const cur = updatesByRow.get(masterRow);
  for (let c = 0; c < WIDTH_EH; c++) {
    const v = newValsEH[c];
    if (v !== "" && v !== null) cur[c] = v;
  }
}

function applyRowUpdates_(masterSheet, updatesByRow, label) {
  if (updatesByRow.size === 0) {
    console.log(`${label}: no updates.`);
    return;
  }

  const rows = [...updatesByRow.keys()].sort((a, b) => a - b);
  const groups = groupConsecutive_(rows);

  for (const [startRow, endRow] of groups) {
    const height = endRow - startRow + 1;
    const range = masterSheet.getRange(startRow, COL_E, height, WIDTH_EH);
    const current = range.getValues();

    for (let i = 0; i < height; i++) {
      const rowNum = startRow + i;
      const upd = updatesByRow.get(rowNum);
      if (!upd) continue;

      const merged = current[i].slice();
      for (let c = 0; c < WIDTH_EH; c++) {
        const v = upd[c];
        if (v !== "" && v !== null) merged[c] = v;
      }
      current[i] = merged;
    }

    range.setValues(current);
  }

  console.log(`${label}: updated ${updatesByRow.size} row(s).`);
}

/***** FORM PROCESSING *****/
function processFormRows_(masterAreasSheet, masterFormSheet) {
  const lastRow = masterFormSheet.getLastRow();
  if (lastRow < FORM_START_ROW) return;

  const numRows = lastRow - FORM_START_ROW + 1;
  const formData = masterFormSheet.getRange(FORM_START_ROW, COL_B, numRows, 5).getValues();
  // Indices within this snapshot:
  // B=0, C=1, D=2, E=3, F=4

  const areasSerialToRow = buildSerialToRowMap_(masterAreasSheet);
  const areaTextToRow = buildAreaTextToRowMap_(masterAreasSheet);
  const formRowsToMarkDone = [];
  let rowsAddedToAreas = 0;
  let rowsRemovedFromAreas = 0;

  for (let i = 0; i < formData.length; i++) {
    const row = FORM_START_ROW + i;
    const action = String(formData[i][0] ?? "").trim().toUpperCase();
    const areaText = String(formData[i][1] ?? "").trim();
    const serial = String(formData[i][2] ?? "").trim();
    const colE = formData[i][3];
    const status = String(formData[i][4] ?? "").trim().toUpperCase();

    const hasAny = formData[i].some(v => v !== "" && v !== null);
    if (!hasAny) continue;
    if (status === "DONE") continue;
    formRowsToMarkDone.push(row);

    if (!serial) continue;
    const serialUpper = serial.toUpperCase();
    const serialRow = areasSerialToRow.get(serial) ?? areasSerialToRow.get(serialUpper);

    if (action === "ADD") {
      if (serialRow) {
        continue;
      }

      const areaRow = areaTextToRow.get(areaText);
      if (!areaRow) {
        continue;
      }

      const areaCode = areaText.slice(1, 5);
      const insertRow = areaRow + 1;
      masterAreasSheet.insertRowAfter(areaRow);
      shiftMapRows_(areasSerialToRow, insertRow, 1);
      shiftMapRows_(areaTextToRow, insertRow, 1);

      masterAreasSheet
        .getRange(8, COL_A, 1, COL_BP)
        .copyFormatToRange(masterAreasSheet, COL_A, COL_BP, insertRow, insertRow);

      if (areaCode) {
        masterAreasSheet.getRange(insertRow, COL_A).setValue(areaCode);
      }
      if (colE !== "" && colE !== null) {
        masterAreasSheet.getRange(insertRow, COL_B).setValue(String(colE).toUpperCase());
      }
      masterAreasSheet.getRange(insertRow, COL_J).setValue(serialUpper);
      areasSerialToRow.set(serialUpper, insertRow);
      rowsAddedToAreas += 1;
      continue;
    }

    const deleteRow = serialRow;
    if (!deleteRow) {
      continue;
    }

    masterAreasSheet.deleteRow(deleteRow);
    areasSerialToRow.delete(serialUpper);
    deleteMapRow_(areaTextToRow, deleteRow);
    shiftMapRows_(areasSerialToRow, deleteRow + 1, -1);
    shiftMapRows_(areaTextToRow, deleteRow + 1, -1);
    rowsRemovedFromAreas += 1;
  }

  if (formRowsToMarkDone.length) {
    const uniqueRows = [...new Set(formRowsToMarkDone)];
    for (const r of uniqueRows) {
      masterFormSheet.getRange(r, COL_F).setValue("DONE");
    }
  }

  console.log(`FORM processing: added ${rowsAddedToAreas} row(s) to ${MASTER_AREAS_TAB}.`);
  console.log(`FORM processing: removed ${rowsRemovedFromAreas} row(s) from ${MASTER_AREAS_TAB}.`);
}

/***** SMALL UTILITIES *****/
function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function groupConsecutive_(sortedRowsAsc) {
  const groups = [];
  let start = sortedRowsAsc[0];
  let prev = sortedRowsAsc[0];

  for (let i = 1; i < sortedRowsAsc.length; i++) {
    const r = sortedRowsAsc[i];
    if (r === prev + 1) {
      prev = r;
    } else {
      groups.push([start, prev]);
      start = r;
      prev = r;
    }
  }
  groups.push([start, prev]);
  return groups;
}

function buildAreaTextToRowMap_(masterSheet) {
  const lastRow = masterSheet.getLastRow();
  const map = new Map();
  if (lastRow < AREAS_MATCH_START_ROW) return map;

  const numRows = lastRow - AREAS_MATCH_START_ROW + 1;
  const values = masterSheet.getRange(AREAS_MATCH_START_ROW, COL_B, numRows, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const areaText = String(values[i][0] ?? "").trim();
    if (!areaText) continue;
    map.set(areaText, AREAS_MATCH_START_ROW + i);
  }
  return map;
}

function shiftMapRows_(rowMap, startRow, delta) {
  if (delta === 0) return;
  for (const [key, row] of rowMap.entries()) {
    if (row >= startRow) {
      rowMap.set(key, row + delta);
    }
  }
}

function deleteMapRow_(rowMap, rowToDelete) {
  for (const [key, row] of rowMap.entries()) {
    if (row === rowToDelete) {
      rowMap.delete(key);
    }
  }
}

function clearA1InChunks_(sheet, a1Ranges) {
  for (let i = 0; i < a1Ranges.length; i += RANGE_LIST_CHUNK) {
    const chunk = a1Ranges.slice(i, i + RANGE_LIST_CHUNK);
    sheet.getRangeList(chunk).clearContent();
  }
}
