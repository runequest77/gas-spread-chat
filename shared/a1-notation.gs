/**
 * Converting colum letter to column index. Start of column index is 0.
 * Ref: https://tanaikech.github.io/2022/05/01/increasing-column-letter-by-one-using-google-apps-script/
 * @param {String} letter Column letter.
 * @return {Number} Column index.
 */
function columnLetterToIndex_(letter = null) {
  letter = letter.toUpperCase();
  return [...letter].reduce(
    (c, e, i, a) =>
      (c += (e.charCodeAt(0) - 64) * Math.pow(26, a.length - i - 1)),
    -1
  );
}

/**
 * Converting colum index to column letter. Start of column index is 0.
 * Ref: https://stackoverflow.com/a/53678158
 * @param {Number} index Column index.
 * @return {String} Column letter.
 */
function columnIndexToLetter_(index = null) {
  return (a = Math.floor(index / 26)) >= 0
    ? columnIndexToLetter_(a - 1) + String.fromCharCode(65 + (index % 26))
    : "";
}

/**
 * Converting 1-based column numbers to a column-only A1 selector.
 * @param {Number} startColumnNumber Start column number.
 * @param {Number} endColumnNumber End column number.
 * @return {String} Column selector like "A:B" or "H:H".
 */
function getColSelector(startColumnNumber, endColumnNumber) {
  return `${columnIndexToLetter_(startColumnNumber - 1)}:${columnIndexToLetter_(endColumnNumber - 1)}`;
}

/**
 * Converting a1Notation to gridrange.
 * Ref: https://tanaikech.github.io/2017/07/31/converting-a1notation-to-gridrange-for-google-sheets-api/
 * @param {String} a1Notation A1Notation of range.
 * @param {Number} sheetId Sheet ID of the range.
 * @return {Object} Gridrange.
 */
function convA1NotationToGridRange_(a1Notation, sheetId) {
  const { col, row } = a1Notation
    .toUpperCase()
    .split("!")
    .map((f) => f.split(":"))
    .pop()
    .reduce(
      (o, g) => {
        var [r1, r2] = ["[A-Z]+", "[0-9]+"].map((h) => g.match(new RegExp(h)));
        o.col.push(r1 && columnLetterToIndex_(r1[0]));
        o.row.push(r2 && Number(r2[0]));
        return o;
      },
      { col: [], row: [] }
    );
  col.sort((a, b) => (a > b ? 1 : -1));
  row.sort((a, b) => (a > b ? 1 : -1));
  const [start, end] = col.map((e, i) => ({ col: e, row: row[i] }));
  const gridrange = {
    sheetId,
    startRowIndex: start?.row && start.row - 1,
    endRowIndex: end?.row ? end.row : start.row,
    startColumnIndex: start && start.col,
    endColumnIndex: end ? end.col + 1 : 1,
  };
  if (gridrange.startRowIndex === null) {
    gridrange.startRowIndex = 0;
    delete gridrange.endRowIndex;
  }
  if (gridrange.startColumnIndex === null) {
    gridrange.startColumnIndex = 0;
    delete gridrange.endColumnIndex;
  }
  return gridrange;
}
//gridrangeは「半開区間」なので、gasで扱いやすいように座標に変換
//{"columnEnd":Number,"columnStart":Number,"rowEnd":Number,"rowStart":Number}
function convA1NotationToPos(a1Notation) {
  var gridrange = convA1NotationToGridRange_(a1Notation,0);
  return {"columnEnd":gridrange.endColumnIndex ,"columnStart":gridrange.startColumnIndex +1 ,"rowEnd":gridrange.endRowIndex ,"rowStart":gridrange.startRowIndex +1};
}



/**
 * Converting gridrange to a1Notation.
 * Ref: https://tanaikech.github.io/2017/07/31/converting-a1notation-to-gridrange-for-google-sheets-api/
 * @param {Object} gridrange Gridrange of range.
 * @param {String} sheetName Sheet name of the range.
 * @return {Object} A1Notation.
 */
function convGridRangeToA1Notation_(gridrange) {
  const start = {};
  const end = {};
  if (gridrange.hasOwnProperty("startColumnIndex")) {
    start.col = columnIndexToLetter_(gridrange.startColumnIndex);
  } else if (
    !gridrange.hasOwnProperty("startColumnIndex") &&
    gridrange.hasOwnProperty("endColumnIndex")
  ) {
    start.col = "A";
  }
  if (gridrange.hasOwnProperty("startRowIndex")) {
    start.row = gridrange.startRowIndex + 1;
  } else if (
    !gridrange.hasOwnProperty("startRowIndex") &&
    gridrange.hasOwnProperty("endRowIndex")
  ) {
    start.row = gridrange.endRowIndex;
  }
  if (gridrange.hasOwnProperty("endColumnIndex")) {
    end.col = columnIndexToLetter_(gridrange.endColumnIndex - 1);
  } else if (!gridrange.hasOwnProperty("endColumnIndex")) {
    end.col = "{Here, please set the max column letter.}";
  }
  if (gridrange.hasOwnProperty("endRowIndex")) {
    end.row = gridrange.endRowIndex;
  }
  const k = ["col", "row"];
  const st = k.map((e) => start[e]).join("");
  const en = k.map((e) => end[e]).join("");
  const a1Notation =
    st == en ? `${st}` : `${st}:${en}`;
  return a1Notation;
}
//gridrangeは「半開区間」なので、gasで扱いやすいように座標に変換
//{"columnEnd":Number,"columnStart":Number,"rowEnd":Number,"rowStart":Number}
function convPosToA1Notation(pos) {
  var gridrange = {sheetId:0, startRowIndex: pos.rowStart -1, endRowIndex:pos.rowEnd, startColumnIndex:pos.columnStart -1, endColumnIndex:pos.columnEnd};
  var a1Notation = convGridRangeToA1Notation_(gridrange);
  return a1Notation;
}
