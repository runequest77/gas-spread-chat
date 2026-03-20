/**
 * アクティブなGoogle Spread Sheetを戻します。
 * オブジェクトをキャッシュすることで、API呼び出しをスクリプトの実行あたり1回に抑制しています。
 * @return {SpreadsheetApp.Spreadsheet}
 */
function c_ss() {
  if (_ss == undefined) {
    _ss = SpreadsheetApp.getActiveSpreadsheet();
    return _ss;
  } else {
    return _ss;
  }
}
let _ss;

/**
 * 指定された名前のシートを戻します。
 * オブジェクトをキャッシュすることで、API呼び出しをスクリプトの実行あたり1回に抑制しています。
 * @param {string} name
 * @return {SpreadsheetApp.Sheet}
 */
function c_sheet(name) {
  if (_sheet[name]) {
    return _sheet[name];
  } else {
    _sheet[name] = c_ss().getSheetByName(name);
    return _sheet[name];
  }
}
let _sheet = {};

/**
 * スプレッドシートの編集された行の値を配列として戻します。
 * オブジェクトをキャッシュすることで、API呼び出しをスクリプトの実行あたり1回に抑制しています。
 * @param {number} row
 * @return {string[]}
 */
function c_rowData(row) {
  if (_rowData[row]) {
    return _rowData[row];
  } else {
    let r = c_sheet(CHAT_SHEET_NAME).getRange(row,1,1,CHAT_SHEET.mapColStart);
    //行列の2次元配列で戻ってきているので1行目のみにして戻す。
    _rowData[row] = r.getValues()[0];
    return _rowData[row];
  }
}
let _rowData = {};
