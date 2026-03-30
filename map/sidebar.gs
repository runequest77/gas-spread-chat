/**
 * マップサイドバー用サーバーサイド関数
 *
 * MAPシートを読み込み、サイドバー上にグリッドマップを描画する機能を提供する。
 * ユニット位置は UserProperties に保存し、ユーザーごとに分離される。
 */

/** サイドバーで使用するMAPシート名 */
const MAP_SIDEBAR_SHEET_NAME = 'MAP';

/**
 * マップサイドバーを表示する
 */
function showMapSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('map/MapSidebar')
    .setTitle('マップUI');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * MAPシートのデータを2次元配列として返す
 * 各セルの値を文字列として返す（空文字も有効な地形）
 *
 * @return {Object} { data: string[][], rows: number, cols: number }
 */
function getMapData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MAP_SIDEBAR_SHEET_NAME);
  if (!sheet) {
    throw new Error('シート "' + MAP_SIDEBAR_SHEET_NAME + '" が見つかりません。');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    return { data: [], rows: 0, cols: 0 };
  }

  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getValues();

  // すべての値を文字列に変換
  const data = values.map(function(row) {
    return row.map(function(cell) {
      return cell === null || cell === undefined ? '' : String(cell);
    });
  });

  return { data: data, rows: lastRow, cols: lastCol };
}

/**
 * 現在のユーザーのユニット位置を取得する
 * 未配置の場合は null を返す
 *
 * @return {Object|null} { row: number, col: number } (1始まり) または null
 */
function getUnitPosition() {
  const userProps = PropertiesService.getUserProperties();
  const posJson = userProps.getProperty('mapUnitPosition');

  if (!posJson) {
    return null;
  }

  try {
    const pos = JSON.parse(posJson);
    if (typeof pos.row === 'number' && typeof pos.col === 'number') {
      return pos;
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * 現在のユーザーのユニット位置を保存する
 *
 * @param {number} row 行番号（1始まり）
 * @param {number} col 列番号（1始まり）
 * @return {Object} { row: number, col: number } 保存した位置
 */
function saveUnitPosition(row, col) {
  row = parseInt(row, 10);
  col = parseInt(col, 10);

  if (isNaN(row) || isNaN(col) || row < 1 || col < 1) {
    throw new Error('無効な座標です: row=' + row + ', col=' + col);
  }

  // MAPシートの範囲チェック
  const mapInfo = getMapData();
  if (row > mapInfo.rows || col > mapInfo.cols) {
    throw new Error('マップ範囲外です: row=' + row + ', col=' + col);
  }

  const pos = { row: row, col: col };
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('mapUnitPosition', JSON.stringify(pos));

  return pos;
}
