/**
 * チャットシートの固定レイアウト定義です。
 * 先頭 3 行をヘッダ専用に使い、4 行目以降を会話・判定ログ本体として 500 行ぶん確保します。
 * 列数を 40 にしているのは、右側にマップ表現用の細いセル群を残すためです。
 */
const SPREADCHAT_SHEET_CHAT = {
  // シート名
  name: "チャット",
  // シート全体の初期サイズです。
  initialRowCount: 500,
  initialColumnCount: 40,
  // 先頭のヘッダ領域を固定します。
  frozenRows: 3,
  frozenColumns: 7,
  // チャットシート上の主要列レイアウトです。
  layout: {
    rollColumnStart: 8,
    rollColumnEnd: 17,
    timestampColumn: 1,
    rollIdColumn: 2,
    imageColumn: 4,
    actionColumn: 6,
    mapColumnStart: 18
  },
  // タイトルとヘッダの固定文字列です。
  titleText: "SPREAD CHAT",
  headerRow: ["TimeStamp", "RollId", "シーン", "画像", "名前", "行動", "判定式"],
  // 2 行目の現在値表示だけ高さを広げます。
  currentResultRowHeight: 80,
  // 固定レイアウトで参照する行範囲です。
  rows: {
    counterRow: "1:1",
    resultRow: "2:2",
    headerRows: "1:3"
  },
  // 固定レイアウトで参照する列範囲です。
  columns: {
    timestampColumn: "A:A",
    rollIdColumn: "B:B",
    sceneColumn: "C:C",
    imageColumn: "D:D",
    nameColumn: "E:E",
    actionColumn: "F:F",
    formulaColumn: "G:G",
    rollColumns: "H:Q",
    mapEdgeColumn: "R:R",
    mapAreaColumns: "S:AN"
  },
  // ヘッダ内で個別参照する A1 範囲です。
  ranges: {
    title: "C1:G2",
    headerTitleBand: "C1:G3",
    headerStatusSampleCell: "J2:J2"
  },
  // 特定表示に使うフォントサイズです。
  fontSizes: {
    resultRow: 40,
    title: 44,
    sceneColumn: 20
  },
  // 列ごとの折り返し方針です。
  wrapStrategies: {
    defaultSheet: "CLIP",
    sceneColumn: "OVERFLOW",
    actionColumn: "WRAP"
  }
}

// タイムスタンプ列から RollId 列までを初期非表示列として扱います。
SPREADCHAT_SHEET_CHAT.layout.initialHiddenColumnCount =
  SPREADCHAT_SHEET_CHAT.layout.rollIdColumn - SPREADCHAT_SHEET_CHAT.layout.timestampColumn + 1

const SPREADCHAT_SHEET_CHAT_FIXED_RANGES = {
  // 行系の固定範囲
  counterRow: getSpreadChatRowRangeSpec_(SPREADCHAT_SHEET_CHAT.rows.counterRow),
  resultRow: getSpreadChatRowRangeSpec_(SPREADCHAT_SHEET_CHAT.rows.resultRow),
  headerRows: getSpreadChatRowRangeSpec_(SPREADCHAT_SHEET_CHAT.rows.headerRows),
  // タイトル周辺の固定範囲
  title: getSpreadChatA1RangeSpec_(SPREADCHAT_SHEET_CHAT.ranges.title),
  headerTitleBand: getSpreadChatA1RangeSpec_(SPREADCHAT_SHEET_CHAT.ranges.headerTitleBand),
  headerStatusSampleCell: getSpreadChatA1RangeSpec_(SPREADCHAT_SHEET_CHAT.ranges.headerStatusSampleCell)
}

const SPREADCHAT_SHEET_CHAT_DEFAULT_ROW_HEIGHT = 21
const SPREADCHAT_SHEET_CHAT_SECRET_PREFIX =
  typeof SPREADCHAT_SHEET_GENERATOR !== "undefined" && SPREADCHAT_SHEET_GENERATOR.markers
    ? SPREADCHAT_SHEET_GENERATOR.markers.secretPrefix
    : (typeof secretPrefix !== "undefined" ? secretPrefix : "秘匿")
const SPREADCHAT_SHEET_CHAT_HEADER_FILL = "#fff2cc"
const SPREADCHAT_SHEET_CHAT_DARK_FILL = "#434343"
const SPREADCHAT_SHEET_CHAT_MID_FILL = "#999999"
const SPREADCHAT_SHEET_CHAT_LIGHT_TEXT = "#ffffff"
const SPREADCHAT_SHEET_CHAT_PALETTE_NAME =
  typeof SPREADCHAT_SHEET_PALETTE !== "undefined" ? SPREADCHAT_SHEET_PALETTE.name : "パレット"

/**
 * 1 行目のカウンタ行を組み立てます。
 * C1:G2 のタイトル表示と、各参加者欄の未処理件数表示の土台を兼ねています。
 * @param {object[]} participantSlots
 * @return {string[]}
 */
function buildSpreadChatChatCounterRow_(participantSlots) {
  const row = Array(SPREADCHAT_SHEET_CHAT.layout.rollColumnEnd).fill("")
  row[SPREADCHAT_SHEET_CHAT_FIXED_RANGES.title.columnStart - 1] = SPREADCHAT_SHEET_CHAT.titleText
  participantSlots.forEach(slot => {
    row[slot.chatColumn - 1] = "0"
  })
  return row
}

/**
 * ロール列のヘッダ 3 行目を返します。
 * 各参加者の表示は「パレット番号 + キャラクター名」の並びにします。
 * @param {object[]} participantSlots
 * @return {string[]}
 */
function buildSpreadChatChatHeaderRow_(participantSlots) {
  const row = SPREADCHAT_SHEET_CHAT.headerRow.slice()
  participantSlots.forEach(slot => row.push(slot.chatHeaderLabel))
  return row
}

const SPREADCHAT_SHEET_CHAT_RANGES = getSpreadChatChatRangeMap_(
  SPREADCHAT_SHEET_CHAT.initialRowCount,
  SPREADCHAT_SHEET_CHAT.initialColumnCount
)

// チャット本文の強調・秘匿表示に使う条件付き書式ルール群です。
const SPREADCHAT_SHEET_CHAT_CONDITIONAL_FORMAT_RULES = [
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.chatBody],
      criteriaType: "CUSTOM_FORMULA",
      criteriaValue: "=LEN($C:$C)>0",
      backgroundColor: "#00ff00"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.headerStatusSampleCell, SPREADCHAT_SHEET_CHAT_RANGES.actionBodyColumn],
      criteriaType: "TEXT_ENDS_WITH",
      criteriaValue: "→",
      backgroundColor: "#ffff00",
      fontColor: "#000000"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.rollBody],
      criteriaType: "TEXT_STARTS_WITH",
      criteriaValue: "FB",
      backgroundColor: "#ff0000",
      fontColor: "#ffffff"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.rollBody],
      criteriaType: "TEXT_STARTS_WITH",
      criteriaValue: "決定",
      backgroundColor: "#ffff00",
      fontColor: "#000000"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.rollBody],
      criteriaType: "TEXT_STARTS_WITH",
      criteriaValue: "効果",
      backgroundColor: "#00ffff",
      fontColor: "#ff0000"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.rollBody],
      criteriaType: "TEXT_STARTS_WITH",
      criteriaValue: "成功",
      backgroundColor: "#00ff00",
      fontColor: "#000000"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.rollBody],
      criteriaType: "TEXT_STARTS_WITH",
      criteriaValue: "失敗",
      backgroundColor: "#ffff00",
      fontColor: "#ff0000"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.headerStatusSampleCell, SPREADCHAT_SHEET_CHAT_RANGES.secretSensitiveBody],
      criteriaType: "TEXT_STARTS_WITH",
      criteriaValue: SPREADCHAT_SHEET_CHAT_SECRET_PREFIX,
      backgroundColor: "#434343",
      fontColor: "#434343"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.hiddenMarkerArea],
      criteriaType: "CUSTOM_FORMULA",
      criteriaValue: '=REGEXMATCH($E3,"^__")',
      backgroundColor: "#434343",
      fontColor: "#ffffff"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.rollBody],
      criteriaType: "CUSTOM_FORMULA",
      criteriaValue: '=(REGEXMATCH(H4,"^(失敗|成功|決定|効果|FB|秘匿|$)")=FALSE)',
      backgroundColor: "#ff9900",
      fontColor: "#000000"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.resultRollRow],
      criteriaType: "NUMBER_GREATER_THAN",
      criteriaValue: 0,
      backgroundColor: "#ffff00",
      fontColor: "#ff0000"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.headerStatusSampleCell, SPREADCHAT_SHEET_CHAT_RANGES.actionBodyColumn],
      criteriaType: "TEXT_STARTS_WITH",
      criteriaValue: "（",
      backgroundColor: "#efefef",
      fontColor: "#000000"
    },
    {
      ranges: [SPREADCHAT_SHEET_CHAT_RANGES.headerStatusSampleCell, SPREADCHAT_SHEET_CHAT_RANGES.actionBodyColumn],
      criteriaType: "TEXT_STARTS_WITH",
      criteriaValue: "「",
      backgroundColor: "#ffe599",
      fontColor: "#000000"
    }
  ]

const SPREADCHAT_SHEET_CHAT_SPEC = {
  // シート名
  sheetName: SPREADCHAT_SHEET_CHAT.name,
  // 初期行数/列数
  initialRowCount: SPREADCHAT_SHEET_CHAT.initialRowCount,
  initialColumnCount: SPREADCHAT_SHEET_CHAT.initialColumnCount,
  // 固定行数/列数
  frozenRows: SPREADCHAT_SHEET_CHAT.frozenRows,
  frozenColumns: SPREADCHAT_SHEET_CHAT.frozenColumns,
  // 既定の行の高さ
  defaultRowHeight: SPREADCHAT_SHEET_CHAT_DEFAULT_ROW_HEIGHT,
  // 列幅レイアウト
  columnWidths: [
    { columnStart: 1, columnCount: 1, width: 60 },
    { columnStart: 2, columnCount: 1, width: 30 },
    { columnStart: 3, columnCount: 1, width: 21 },
    { columnStart: 4, columnCount: 1, width: 80 },
    { columnStart: 5, columnCount: 1, width: 100 },
    { columnStart: 6, columnCount: 1, width: 560 },
    { columnStart: 7, columnCount: 1, width: 100 },
    {
      columnStart: SPREADCHAT_SHEET_CHAT.layout.rollColumnStart,
      columnCount: SPREADCHAT_SHEET_CHAT.layout.rollColumnEnd - SPREADCHAT_SHEET_CHAT.layout.rollColumnStart + 1,
      width: 80
    },
    {
      columnStart: SPREADCHAT_SHEET_CHAT.layout.mapColumnStart,
      columnCount: SPREADCHAT_SHEET_CHAT.initialColumnCount - SPREADCHAT_SHEET_CHAT.layout.mapColumnStart + 1,
      width: 21
    }
  ],
  // 2 行目の現在値行だけ高さを調整します。
  rowHeights: [{
    rowStart: SPREADCHAT_SHEET_CHAT_FIXED_RANGES.resultRow.rowStart,
    rowCount: 1,
    height: SPREADCHAT_SHEET_CHAT.currentResultRowHeight
  }],
  // 内部管理用の先頭列は非表示にします。
  hiddenColumns: [{
    columnStart: SPREADCHAT_SHEET_CHAT.layout.timestampColumn,
    columnCount: SPREADCHAT_SHEET_CHAT.layout.initialHiddenColumnCount
  }],
  // カウンタ行は表示対象外です。
  hiddenRows: [SPREADCHAT_SHEET_CHAT_FIXED_RANGES.counterRow],
  // ヘッダ行を保護して誤編集を防ぎます。
  protectedRowRanges: [Object.assign({ protectionKey: "ChatRowHeader" }, SPREADCHAT_SHEET_CHAT_FIXED_RANGES.headerRows)],
  // タイトル領域を結合表示します。
  mergedRanges: [SPREADCHAT_SHEET_CHAT_FIXED_RANGES.title],
  // シート全体および強調表示のスタイルです。
  styleRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_CHAT.initialRowCount,
    columnCount: SPREADCHAT_SHEET_CHAT.initialColumnCount,
    verticalAlignment: "top"
  }].concat([
    Object.assign({}, SPREADCHAT_SHEET_CHAT_RANGES.resultRollRow, {
      verticalAlignment: "middle",
      horizontalAlignment: "center",
      backgroundColor: "#ffff00",
      fontColor: "#ff0000",
      fontSize: SPREADCHAT_SHEET_CHAT.fontSizes.resultRow,
      fontWeight: "bold"
    }),
    Object.assign({}, SPREADCHAT_SHEET_CHAT_RANGES.participantHeaderRollRow, {
      fontWeight: "bold"
    }),
    Object.assign({}, SPREADCHAT_SHEET_CHAT_RANGES.sceneColumn, {
      fontSize: SPREADCHAT_SHEET_CHAT.fontSizes.sceneColumn
    }),
    Object.assign({}, SPREADCHAT_SHEET_CHAT_RANGES.title, {
      verticalAlignment: "middle",
      horizontalAlignment: "left",
      fontSize: SPREADCHAT_SHEET_CHAT.fontSizes.title,
      fontFamily: "Dela Gothic One",
      fontStyle: "italic"
    })
  ]),
  // 基本の折り返しと列別上書きです。
  wrapStrategyRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_CHAT.initialRowCount,
    columnCount: SPREADCHAT_SHEET_CHAT.initialColumnCount,
    wrapStrategy: SPREADCHAT_SHEET_CHAT.wrapStrategies.defaultSheet
  }].concat([
    Object.assign({}, SPREADCHAT_SHEET_CHAT_RANGES.sceneColumn, {
      wrapStrategy: SPREADCHAT_SHEET_CHAT.wrapStrategies.sceneColumn
    }),
    Object.assign({}, SPREADCHAT_SHEET_CHAT_RANGES.actionColumn, {
      wrapStrategy: SPREADCHAT_SHEET_CHAT.wrapStrategies.actionColumn
    })
  ]),
  // 背景色の配色設定
  backgroundRanges: [
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_HEADER_FILL }, SPREADCHAT_SHEET_CHAT_RANGES.imageColumn),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_HEADER_FILL }, SPREADCHAT_SHEET_CHAT_RANGES.nameColumn),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_HEADER_FILL }, SPREADCHAT_SHEET_CHAT_RANGES.formulaColumn),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_DARK_FILL }, SPREADCHAT_SHEET_CHAT_RANGES.timestampColumn),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_DARK_FILL }, SPREADCHAT_SHEET_CHAT_RANGES.rollIdColumn),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_DARK_FILL }, SPREADCHAT_SHEET_CHAT_RANGES.mapEdgeColumn),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_MID_FILL }, SPREADCHAT_SHEET_CHAT_RANGES.mapAreaBody),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_DARK_FILL }, SPREADCHAT_SHEET_CHAT_RANGES.headerTitleBand)
  ],
  // 文字色の配色設定
  fontColorRanges: [
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_LIGHT_TEXT }, SPREADCHAT_SHEET_CHAT_RANGES.timestampColumn),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_LIGHT_TEXT }, SPREADCHAT_SHEET_CHAT_RANGES.rollIdColumn),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_LIGHT_TEXT }, SPREADCHAT_SHEET_CHAT_RANGES.mapEdgeColumn),
    Object.assign({ color: SPREADCHAT_SHEET_CHAT_LIGHT_TEXT }, SPREADCHAT_SHEET_CHAT_RANGES.headerTitleBand)
  ],
  // 条件付き書式ルール
  conditionalFormatRules: SPREADCHAT_SHEET_CHAT_CONDITIONAL_FORMAT_RULES
}

function getSpreadChatChatRangeMap_(rowCount, columnCount) {
  const totalRows = typeof rowCount === "number" ? rowCount : SPREADCHAT_SHEET_CHAT.initialRowCount
  const totalColumns = typeof columnCount === "number" ? columnCount : SPREADCHAT_SHEET_CHAT.initialColumnCount
  const counterRow = SPREADCHAT_SHEET_CHAT_FIXED_RANGES.counterRow
  const resultRow = SPREADCHAT_SHEET_CHAT_FIXED_RANGES.resultRow
  const headerRows = SPREADCHAT_SHEET_CHAT_FIXED_RANGES.headerRows
  const timestampColumn = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.timestampColumn, totalRows)
  const rollIdColumn = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.rollIdColumn, totalRows)
  const sceneColumn = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.sceneColumn, totalRows)
  const imageColumn = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.imageColumn, totalRows)
  const nameColumn = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.nameColumn, totalRows)
  const actionColumn = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.actionColumn, totalRows)
  const formulaColumn = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.formulaColumn, totalRows)
  const rollColumns = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.rollColumns, totalRows)
  const mapEdgeColumn = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.mapEdgeColumn, totalRows)
  const mapAreaColumns = getSpreadChatColumnRangeSpec_(SPREADCHAT_SHEET_CHAT.columns.mapAreaColumns, totalRows)
  const title = SPREADCHAT_SHEET_CHAT_FIXED_RANGES.title
  const headerTitleBand = SPREADCHAT_SHEET_CHAT_FIXED_RANGES.headerTitleBand
  const headerStatusSampleCell = SPREADCHAT_SHEET_CHAT_FIXED_RANGES.headerStatusSampleCell
  const headerRow = headerRows.rowStart + headerRows.rowCount - 1
  const bodyRowStart = headerRow + 1
  const bodyRowCount = Math.max(0, totalRows - headerRows.rowCount)
  const sceneToSheetEndColumnCount = Math.max(0, totalColumns - (sceneColumn.columnStart - 1))
  const nameToSheetEndColumnCount = Math.max(0, totalColumns - (nameColumn.columnStart - 1))
  return {
    counterRow: counterRow,
    resultRow: resultRow,
    headerRows: headerRows,
    title: title,
    headerTitleBand: headerTitleBand,
    chatBody: {
      rowStart: bodyRowStart,
      columnStart: sceneColumn.columnStart,
      rowCount: bodyRowCount,
      columnCount: sceneToSheetEndColumnCount
    },
    resultRollRow: {
      rowStart: resultRow.rowStart,
      columnStart: rollColumns.columnStart,
      rowCount: 1,
      columnCount: rollColumns.columnCount
    },
    participantHeaderRollRow: {
      rowStart: headerRow,
      columnStart: rollColumns.columnStart,
      rowCount: 1,
      columnCount: rollColumns.columnCount
    },
    headerStatusSampleCell: headerStatusSampleCell,
    actionBodyColumn: {
      rowStart: bodyRowStart,
      columnStart: actionColumn.columnStart,
      rowCount: bodyRowCount,
      columnCount: 1
    },
    rollBody: {
      rowStart: bodyRowStart,
      columnStart: rollColumns.columnStart,
      rowCount: bodyRowCount,
      columnCount: rollColumns.columnCount
    },
    secretSensitiveBody: {
      rowStart: bodyRowStart,
      columnStart: nameColumn.columnStart,
      rowCount: bodyRowCount,
      columnCount: (rollColumns.columnStart + rollColumns.columnCount - 1) - nameColumn.columnStart + 1
    },
    hiddenMarkerArea: {
      rowStart: headerRow,
      columnStart: nameColumn.columnStart,
      rowCount: Math.max(0, totalRows - resultRow.rowStart),
      columnCount: nameToSheetEndColumnCount
    },
    timestampColumn: timestampColumn,
    rollIdColumn: rollIdColumn,
    sceneColumn: sceneColumn,
    actionColumn: actionColumn,
    imageColumn: imageColumn,
    nameColumn: nameColumn,
    formulaColumn: formulaColumn,
    mapEdgeColumn: mapEdgeColumn,
    mapAreaBody: {
      rowStart: mapAreaColumns.rowStart,
      rowCount: totalRows,
      columnStart: mapAreaColumns.columnStart,
      columnCount: Math.min(mapAreaColumns.columnCount, Math.max(0, totalColumns - (mapAreaColumns.columnStart - 1)))
    }
  }
}

/**
 * ヘッダ部に埋め込む数式セルを返します。
 * 1 行目は進行管理用カウンタ、2 行目は現在値、3 行目はパレット由来の参加者ラベルを表示します。
 * @param {object[]} participantSlots
 * @return {object[]}
 */
function buildSpreadChatChatHeaderFormulaCells_(participantSlots) {
  const firstParticipantColumn = SPREADCHAT_SHEET_CHAT.layout.rollColumnStart
  const lastParticipantColumn = SPREADCHAT_SHEET_CHAT.layout.rollColumnEnd
  const firstPlayerColumn = firstParticipantColumn + 1
  const firstPlayerColumnLetter = getSpreadChatColumnLetter_(firstPlayerColumn)
  const lastParticipantColumnLetter = getSpreadChatColumnLetter_(lastParticipantColumn)
  const paletteSheetName = getSpreadChatEscapedSheetReference_(SPREADCHAT_SHEET_CHAT_PALETTE_NAME)
  const counterRowStart = SPREADCHAT_SHEET_CHAT_FIXED_RANGES.counterRow.rowStart
  const resultRowStart = SPREADCHAT_SHEET_CHAT_FIXED_RANGES.resultRow.rowStart
  const participantHeaderRowStart =
    SPREADCHAT_SHEET_CHAT_FIXED_RANGES.headerRows.rowStart + SPREADCHAT_SHEET_CHAT_FIXED_RANGES.headerRows.rowCount - 1
  const formulaCells = [{
    rowStart: counterRowStart,
    columnStart: firstParticipantColumn,
    formula: `=countif(${firstPlayerColumnLetter}3:${lastParticipantColumnLetter},"=a")`
  }]
  participantSlots
    .filter(slot => slot.chatColumn >= firstPlayerColumn)
    .forEach(slot => {
      const columnLetter = getSpreadChatColumnLetter_(slot.chatColumn)
      formulaCells.push({
        rowStart: counterRowStart,
        columnStart: slot.chatColumn,
        formula: `=countif(${columnLetter}3:${columnLetter},"=q")`
      })
    })
  participantSlots.forEach(slot => {
    const columnLetter = getSpreadChatColumnLetter_(slot.chatColumn)
    const paletteValueColumnLetter = getSpreadChatColumnLetter_(slot.paletteColumnStart)
    const paletteHeaderColumnLetter = getSpreadChatColumnLetter_(slot.paletteColumnEnd)
    const paletteLabelRow = slot.userNo === 0 ? 3 : 4
    formulaCells.push({
      rowStart: resultRowStart,
      columnStart: slot.chatColumn,
      formula: `=if(${columnLetter}1>0,${columnLetter}1,${paletteSheetName}!${paletteValueColumnLetter}5)`
    })
    formulaCells.push({
      rowStart: participantHeaderRowStart,
      columnStart: slot.chatColumn,
      formula: `="${slot.userNo} "&${paletteSheetName}!${paletteHeaderColumnLetter}${paletteLabelRow}`
    })
  })
  return formulaCells
}
