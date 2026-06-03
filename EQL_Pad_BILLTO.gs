function padColumnKToFiveDigits_ALL_OLD_current() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ALL OLD current");
  if (!sheet) throw new Error('Sheet "ALL OLD current" not found.');

  const startRow = 8;                 // start at row 8
  const colK = 11;                    // column K
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const range = sheet.getRange(startRow, colK, lastRow - startRow + 1, 1);
  const values = range.getValues();

  const output = values.map(([v]) => {
    if (v === "" || v === null) return [v];

    const s = String(v).trim();

    // If the cell contains only digits, pad to 5 chars with leading zeros
    if (/^\d+$/.test(s)) {
      return [s.padStart(5, "0")];
    }
    return [v];
  });

  range.setValues(output);
}
