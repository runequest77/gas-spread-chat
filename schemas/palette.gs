/**
 * パレットシートの固定定義です。
 * 10 人ぶんを Key/Valueとして2 列ずつ持つ 20 列構成で、上 5 行を見出し・表示名・画像欄として使います。
 * 5 行目～13 行目はアイコン行。
 * 14 行目以降がデータ行。
 */
const SPREADCHAT_SHEET_PALETTE = {
  // シート名
  name: "パレット",
  // シート全体の初期サイズです。
  initialRowCount: 20,
  initialColumnCount: 20,
  // 先頭の見出し領域を固定します。
  frozenRows: 5,
  frozenColumns: 2,
  // 列幅とアイコン行の高さです。
  columnWidth: 90,
  imageRowHeight: 90,
  // 固定レイアウトで参照する行範囲です。
  rows: {
    headerRows: "1:2",
    imageRows: "5:13"
  },
  // 条件付き書式で参照する全体範囲です。
  ranges: {
    fullSheet: "A1:T20"
  }
}

const SPREADCHAT_SHEET_PALETTE_FIXED_RANGES = {
  // シート全体の A1 範囲です。
  fullSheet: getSpreadChatA1RangeSpec_(SPREADCHAT_SHEET_PALETTE.ranges.fullSheet),
  // 見出し・画像行の行範囲です。
  headerRows: getSpreadChatRowRangeSpec_(SPREADCHAT_SHEET_PALETTE.rows.headerRows),
  imageRows: getSpreadChatRowRangeSpec_(SPREADCHAT_SHEET_PALETTE.rows.imageRows)
}

const SPREADCHAT_SHEET_PALETTE_DEFAULT_ROW_HEIGHT = 21
const SPREADCHAT_SHEET_PALETTE_SECRET_MARKER = "■"

const SPREADCHAT_SHEET_PALETTE_SPEC = {
  // シート名
  sheetName: SPREADCHAT_SHEET_PALETTE.name,
  // 初期行数/列数
  initialRowCount: SPREADCHAT_SHEET_PALETTE.initialRowCount,
  initialColumnCount: SPREADCHAT_SHEET_PALETTE.initialColumnCount,
  // 固定行数/列数
  frozenRows: SPREADCHAT_SHEET_PALETTE.frozenRows,
  frozenColumns: SPREADCHAT_SHEET_PALETTE.frozenColumns,
  // 既定の行の高さ
  defaultRowHeight: SPREADCHAT_SHEET_PALETTE_DEFAULT_ROW_HEIGHT,
  // 全列を同じ幅にそろえます。
  columnWidths: [{
    columnStart: 1,
    columnCount: SPREADCHAT_SHEET_PALETTE.initialColumnCount,
    width: SPREADCHAT_SHEET_PALETTE.columnWidth
  }],
  // アイコン表示行のみ高さを広げます。
  rowHeights: [Object.assign({ height: SPREADCHAT_SHEET_PALETTE.imageRowHeight }, SPREADCHAT_SHEET_PALETTE_FIXED_RANGES.imageRows)],
  // ヘッダ行の誤編集を防ぐため保護します。
  protectedRowRanges: [Object.assign({ protectionKey: "PaletteRowHeader" }, SPREADCHAT_SHEET_PALETTE_FIXED_RANGES.headerRows)],
  // シート全体の基本スタイルです。
  styleRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_PALETTE.initialRowCount,
    columnCount: SPREADCHAT_SHEET_PALETTE.initialColumnCount,
    verticalAlignment: "top"
  }],
  // 折り返しはしないでクリップ表示にします。
  wrapStrategyRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_PALETTE.initialRowCount,
    columnCount: SPREADCHAT_SHEET_PALETTE.initialColumnCount,
    wrapStrategy: "CLIP"
  }],
  // 秘匿行は背景色と文字色を同じにして見えなくします。
  conditionalFormatRules: [{
    ranges: [SPREADCHAT_SHEET_PALETTE_FIXED_RANGES.fullSheet],
    criteriaType: "CUSTOM_FORMULA",
    criteriaValue: `=LEFT($A:$A,1)="${SPREADCHAT_SHEET_PALETTE_SECRET_MARKER}"`,
    backgroundColor: "#434343",
    fontColor: "#434343"
  }]
}

/**
 * 1 行目のパレット見出しを返します。
 * 各参加者を 2 列 1 組で扱うため、「palette / 番号」の並びを繰り返します。
 * @param {object[]} participantSlots
 * @return {string[]}
 */
function buildSpreadChatPaletteHeaderRow_(participantSlots) {
  const row = []
  participantSlots.forEach(slot => {
    row.push("palette", slot.paletteHeaderLabel)
  })
  return row
}

/**
 * 2～3 行目のキー行を返します。
 * GM の 1 列目だけは秘匿パレット識別用の記号を付け、pname 行には初期値として GM を入れます。
 * @param {string} key
 * @param {object[]} participantSlots
 * @param {boolean} markSecretFirstSlot
 * @return {string[]}
 */
function buildSpreadChatPaletteKeyRow_(key, participantSlots, markSecretFirstSlot) {
  const row = []
  participantSlots.forEach(slot => {
    row.push(
      markSecretFirstSlot && slot.userNo === 0 ? `${SPREADCHAT_SHEET_PALETTE_SECRET_MARKER}${key}` : key,
      slot.userNo === 0 && key === "pname" ? "GM" : ""
    )
  })
  return row
}

/**
 * 4 行目のキャラクター名行を返します。
 * GM 枠だけは描写用として使う想定なので、初期表示を「描写」にしています。
 * @param {object[]} participantSlots
 * @return {string[]}
 */
function buildSpreadChatPaletteCnameRow_(participantSlots) {
  const row = []
  participantSlots.forEach(slot => {
    row.push("cname", slot.userNo === 0 ? "描写" : "")
  })
  return row
}
