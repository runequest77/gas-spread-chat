/**
 * Chart ボットで使用する Charts シートの固定定義です。
 * 表のサンプルとして1 列・20 行に留めています。
 */
const SPREADCHAT_SHEET_CHARTS = {
  // シート名
  name: "Charts",
  // シート全体の初期サイズです。
  initialRowCount: 20,
  initialColumnCount: 1,
  // 1 行目は見出しとして使います。
  values: [["部位"]]
}

const SPREADCHAT_SHEET_CHARTS_DEFAULT_ROW_HEIGHT = 21
const SPREADCHAT_SHEET_CHARTS_HEADER_FILL = "#4a86e8"
const SPREADCHAT_SHEET_CHARTS_HEADER_TEXT = "#ffffff"

const SPREADCHAT_SHEET_CHARTS_SPEC = {
  // シート名
  sheetName: SPREADCHAT_SHEET_CHARTS.name,

  // 初期行数/列数
  initialRowCount: SPREADCHAT_SHEET_CHARTS.initialRowCount,
  initialColumnCount: SPREADCHAT_SHEET_CHARTS.initialColumnCount,

  // 固定行数/列数
  frozenRows: 1,
  frozenColumns: 0,

  // 既定の行の高さ
  defaultRowHeight: SPREADCHAT_SHEET_CHARTS_DEFAULT_ROW_HEIGHT,

  // シート全体の基本スタイルです。
  styleRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_CHARTS.initialRowCount,
    columnCount: SPREADCHAT_SHEET_CHARTS.initialColumnCount,
    verticalAlignment: "top"
  }],

  // 折り返しは行わずクリップ表示にします。
  wrapStrategyRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_CHARTS.initialRowCount,
    columnCount: SPREADCHAT_SHEET_CHARTS.initialColumnCount,
    wrapStrategy: "CLIP"
  }],

  // 初期ヘッダ値
  values: SPREADCHAT_SHEET_CHARTS.values,

  // 見出し行の背景色
  backgroundRanges: [{
    rowStart: 1,
    rowCount: 1,
    columnStart: 1,
    columnCount: SPREADCHAT_SHEET_CHARTS.initialColumnCount,
    color: SPREADCHAT_SHEET_CHARTS_HEADER_FILL
  }],

  // 見出し行の文字色
  fontColorRanges: [{
    rowStart: 1,
    rowCount: 1,
    columnStart: 1,
    columnCount: SPREADCHAT_SHEET_CHARTS.initialColumnCount,
    color: SPREADCHAT_SHEET_CHARTS_HEADER_TEXT
  }],

  // 参加者別色分けは適用しません。
  participantColorBands: []
}
