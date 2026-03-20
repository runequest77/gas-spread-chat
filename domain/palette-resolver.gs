const paletteImageColSpacing = 2;
const namePaletteNoPattern = /^_{0,1}([0-9]{1,2}[1-9]{1}):/;
const npcPaletteNoPattern = /^_([1-9][0-9])[0-9]:/;

function getNamePaletteNo(value) {
  var match = namePaletteNoPattern.exec(value || "");
  return match == null ? null : Number(match[1]);
}

function getNpcPaletteNoFromName(value) {
  var match = npcPaletteNoPattern.exec(value || "");
  return match == null ? null : Number(match[1]);
}

function resolvePaletteNoFromNameColumn(value) {
  var paletteNo = getNamePaletteNo(value);
  if (paletteNo != null) return paletteNo;
  return getNpcPaletteNoFromName(value);
}

function getPaletteImageOffsets(paletteNo) {
  return {
    row: Number(palette_nameRow) + Number(String(paletteNo).slice(-1)) - 1,
    col: Number(Math.floor(paletteNo / 10)) * paletteImageColSpacing,
  };
}

function buildNameImageFormula(value) {
  var paletteNo = getNamePaletteNo(value);
  if (paletteNo == null) return "";
  var offsets = getPaletteImageOffsets(paletteNo);
  var escapedPaletteSheetName = String(PALETTE_SHEET_NAME).replace(/'/g, "''");
  return `=OFFSET('${escapedPaletteSheetName}'!A1,${offsets.row},${offsets.col})`;
}

function adjustRowHeightForNameColumn(sheet, row, nameValue) {
  var paletteNo = resolvePaletteNoFromNameColumn(nameValue);
  if (paletteNo != null) {
    sheet.setRowHeight(row, sheet.getColumnWidth(imageCol));
  } else {
    sheet.autoResizeRows(row, 1);
  }
}
