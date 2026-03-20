/**
 * SelectList シートの固定定義です。
 * 画面には見せない前処理列も含めて 12 列を確保し、他シートの入力補助候補を集約します。
 * 1 行目は列の役割説明、2 行目以降は数式で自動展開される前提です。
 */
const SPREADCHAT_SHEET_SELECTLIST = {
  // シート名
  name: "SelectList",
  // シート全体の初期サイズです。
  initialRowCount: 20,
  initialColumnCount: 12,
  // 1 行目は候補列の役割を示すヘッダです。
  values: [[
    "前処理1(名前)",
    "NPCName",
    "MapName",
    "前処理4(パレット名)",
    "ChartName",
    "名前",
    "パレット名",
    "",
    "PaletteKey1",
    "PaletteKey2",
    "PaletteKey3",
    "PaletteKey4"
  ]]
}

const SPREADCHAT_SHEET_SELECTLIST_DEFAULT_ROW_HEIGHT = 21
const SPREADCHAT_SHEET_SELECTLIST_SECRET_MARKER = "■"
const SPREADCHAT_SHEET_SELECTLIST_HEADER_FILL = "#4a86e8"
const SPREADCHAT_SHEET_SELECTLIST_HEADER_TEXT = "#ffffff"

const SPREADCHAT_SHEET_SELECTLIST_FORMULA_CELLS = [
    // 1 列目: パレット 3 行目から「番号:名前」の候補を作成します。
    {
      rowStart: 2,
      columnStart: 1,
      formula: '=SORT(UNIQUE(FLATTEN(ARRAYFORMULA(IF(LEN(\'パレット\'!A3:T3)=0,"",IF(ISODD(column(\'パレット\'!A3:T3)),"",(column(\'パレット\'!A3:T3)/2-1)&":"&\'パレット\'!A3:T3))))))'
    },
    // 2 列目: アイコン欄から NPC 名候補を抽出します。
    {
      rowStart: 2,
      columnStart: 2,
      formula: '=SORT(UNIQUE(FLATTEN(FILTER(ARRAYFORMULA(IF(LEN(\'パレット\'!4:13)=0,"",IF(ISODD(column(\'パレット\'!4:13)),"",IF(column(\'パレット\'!4:13)>20,"_","")&(column(\'パレット\'!4:13)/2-1)&(ROW(\'パレット\'!4:13)-4)&":"&\'パレット\'!4:13))),LEN(\'パレット\'!3:3)>0))))'
    },
    // 4 列目: データ行からパレット名の候補を収集します。
    {
      rowStart: 2,
      columnStart: 4,
      formula: '=UNIQUE(FLATTEN({FILTER(OFFSET(\'パレット\'!A1,14,0,ROWS(\'パレット\'!A:A),COLUMNS(\'パレット\'!1:1)),MOD(COLUMN(OFFSET(\'パレット\'!A1,14,0,ROWS(\'パレット\'!A:A),COLUMNS(\'パレット\'!1:1))),2)=1)}))'
    },
    // 5 列目: Charts シート由来のコマンド候補です。
    {
      rowStart: 2,
      columnStart: 5,
      formula: '=ARRAYFORMULA("c:" & FLATTEN({Charts!1:1}))'
    },
    // 6 列目: 名前候補を統合して重複除去します。
    {
      rowStart: 2,
      columnStart: 6,
      formula: '=FILTER(UNIQUE({A2:A;B2:B;C2:C}),UNIQUE({A2:A;B2:B;C2:C})<>"")'
    },
    // 7 列目: パレット候補を統合し、秘匿・空白先頭を除外します。
    {
      rowStart: 2,
      columnStart: 7,
      formula: `=FILTER({D2:D;E2:E},LEN({D2:D;E2:E})>1,LEFT({D2:D;E2:E},1)<>"${SPREADCHAT_SHEET_SELECTLIST_SECRET_MARKER}",LEFT({D2:D;E2:E},1)<>" ")`
    },
    // 9 列目: 1 人目パレットキー候補
    {
      rowStart: 2,
      columnStart: 9,
      formula: '=ARRAYFORMULA("/" & FILTER(UNIQUE(\'パレット\'!C14:C),LEFT(UNIQUE(\'パレット\'!C14:C),1)<>"*"))'
    },
    // 10 列目: 2 人目パレットキー候補
    {
      rowStart: 2,
      columnStart: 10,
      formula: '=ARRAYFORMULA("/" & FILTER(UNIQUE(\'パレット\'!E14:E),LEFT(UNIQUE(\'パレット\'!E14:E),1)<>"*"))'
    },
    // 11 列目: 3 人目パレットキー候補
    {
      rowStart: 2,
      columnStart: 11,
      formula: '=ARRAYFORMULA("/" & FILTER(UNIQUE(\'パレット\'!G14:G),LEFT(UNIQUE(\'パレット\'!G14:G),1)<>"*"))'
    },
    // 12 列目: 4 人目パレットキー候補
    {
      rowStart: 2,
      columnStart: 12,
      formula: '=ARRAYFORMULA("/" & FILTER(UNIQUE(\'パレット\'!I14:I),LEFT(UNIQUE(\'パレット\'!I14:I),1)<>"*"))'
    }
  ]

const SPREADCHAT_SHEET_SELECTLIST_SPEC = {
  // シート名
  sheetName: SPREADCHAT_SHEET_SELECTLIST.name,
  // 初期行数/列数
  initialRowCount: SPREADCHAT_SHEET_SELECTLIST.initialRowCount,
  initialColumnCount: SPREADCHAT_SHEET_SELECTLIST.initialColumnCount,
  // 固定行数/列数
  frozenRows: 1,
  frozenColumns: 0,
  // 既定の行の高さ
  defaultRowHeight: SPREADCHAT_SHEET_SELECTLIST_DEFAULT_ROW_HEIGHT,
  // シート全体の基本スタイルです。
  styleRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_SELECTLIST.initialRowCount,
    columnCount: SPREADCHAT_SHEET_SELECTLIST.initialColumnCount,
    verticalAlignment: "top"
  }],
  // 折り返しは行わずクリップ表示にします。
  wrapStrategyRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_SELECTLIST.initialRowCount,
    columnCount: SPREADCHAT_SHEET_SELECTLIST.initialColumnCount,
    wrapStrategy: "CLIP"
  }],
  // 初期ヘッダ値
  values: SPREADCHAT_SHEET_SELECTLIST.values,
  // 2 行目以降の自動展開数式
  formulaCells: SPREADCHAT_SHEET_SELECTLIST_FORMULA_CELLS,
  // 見出し行の背景色
  backgroundRanges: [{
    rowStart: 1,
    rowCount: 1,
    columnStart: 1,
    columnCount: SPREADCHAT_SHEET_SELECTLIST.initialColumnCount,
    color: SPREADCHAT_SHEET_SELECTLIST_HEADER_FILL
  }],
  // 見出し行の文字色
  fontColorRanges: [{
    rowStart: 1,
    rowCount: 1,
    columnStart: 1,
    columnCount: SPREADCHAT_SHEET_SELECTLIST.initialColumnCount,
    color: SPREADCHAT_SHEET_SELECTLIST_HEADER_TEXT
  }],
  // 参加者別色分けは適用しません。
  participantColorBands: []
}
