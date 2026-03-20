/**
 * RollLog シートの固定定義です。
 * ロール履歴を時系列で残す補助シートなので、必要列だけに絞った 5 列・100 行の素直な構成にしています。
 */
const SPREADCHAT_SHEET_ROLLLOG = {
  // シート名
  name: "RollLog",
  // シート全体の初期サイズです。
  initialRowCount: 100,
  initialColumnCount: 5,
  // 1 行目はヘッダ固定です。
  values: [[
    "RollID",
    "CreateDate",
    "Name",
    "Command",
    "Result"
  ]]
}

const SPREADCHAT_SHEET_ROLLLOG_DEFAULT_ROW_HEIGHT = 21
const SPREADCHAT_SHEET_ROLLLOG_HEADER_FILL = "#4a86e8"
const SPREADCHAT_SHEET_ROLLLOG_HEADER_TEXT = "#ffffff"

const SPREADCHAT_SHEET_ROLLLOG_SPEC = {
  // シート名
  sheetName: SPREADCHAT_SHEET_ROLLLOG.name,
  // 初期行数/列数
  initialRowCount: SPREADCHAT_SHEET_ROLLLOG.initialRowCount,
  initialColumnCount: SPREADCHAT_SHEET_ROLLLOG.initialColumnCount,
  // 固定行数/列数
  frozenRows: 1,
  frozenColumns: 0,
  // 既定の行の高さ
  defaultRowHeight: SPREADCHAT_SHEET_ROLLLOG_DEFAULT_ROW_HEIGHT,
  // シート全体の基本スタイルです。
  styleRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_ROLLLOG.initialRowCount,
    columnCount: SPREADCHAT_SHEET_ROLLLOG.initialColumnCount,
    verticalAlignment: "top"
  }],
  // 折り返しは行わずクリップ表示にします。
  wrapStrategyRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_ROLLLOG.initialRowCount,
    columnCount: SPREADCHAT_SHEET_ROLLLOG.initialColumnCount,
    wrapStrategy: "CLIP"
  }],
  // 初期ヘッダ値
  values: SPREADCHAT_SHEET_ROLLLOG.values,
  // 見出し行の背景色
  backgroundRanges: [{
    rowStart: 1,
    rowCount: 1,
    columnStart: 1,
    columnCount: SPREADCHAT_SHEET_ROLLLOG.initialColumnCount,
    color: SPREADCHAT_SHEET_ROLLLOG_HEADER_FILL
  }],
  // 見出し行の文字色
  fontColorRanges: [{
    rowStart: 1,
    rowCount: 1,
    columnStart: 1,
    columnCount: SPREADCHAT_SHEET_ROLLLOG.initialColumnCount,
    color: SPREADCHAT_SHEET_ROLLLOG_HEADER_TEXT
  }],
  // 参加者別色分けは適用しません。
  participantColorBands: []
}
