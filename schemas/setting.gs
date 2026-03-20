/**
 * Setting シートの固定定義です。
 */
const SPREADCHAT_SHEET_SETTING = {
  // シート名
  name: "Setting",
  // シート全体の初期サイズです。列は「設定名 / 値 / 説明」の 3 列で固定し、
  // 行は既定項目より少し余裕を持たせて追加設定を書き足せるようにしています。
  initialRowCount: 20,
  initialColumnCount: 3,
  values: [
    // 「設定名 / 初期値 / 説明」
    ["settingName", "settingValue", "settingDescription"],
    // 省略記法の既定値です。チャット欄で bot 指定を毎回書かなくて済むよう、よく使う RQG を初期値にしています。
    ["DEFAULT_BOT", "rqg:", "ダイスボットマーカーを省略したときに使用されるダイスボットマーカー"],
    // スラッシュだけが入力されたときの既定コマンドです。RuneQuest 系の基本判定に合わせて 1d100 を採用しています。
    ["DEFAULT_COMMAND", "1d100", "判定式列に入力がないときのスラッシュのみのロール動作"],
    // 公開 BCDice API の接続先です。導入直後にそのまま使える代表的な URL を既定値にしています。
    ["BCDICEAPI_URL", "https://bcdice.onlinesession.app", "BCDICEを利用する場合、APIエンドポイントを指定します。"],
    // 秘匿発言の接頭辞です。実際の値は共通設定側の marker を反映するため、ここでは null プレースホルダーを置きます。
    ["SECRET_PREFIX", null, "秘匿情報として扱われる文字。アクション列・判定式列・PC列の背景色と文字色を同一にし、見えなくします。\n内容が見えてしまうため、PC列のノートは生成されなくなります。"],
    // チャットシートの参照先です。利用者がシート名を変更できるよう、実際の値は共通設定から差し込みます。
    ["CHAT_SHEET_NAME", null, "チャットシート名。気に入らない時だけ変更。他と被らないように注意。"],
    // パレットシートの参照先です。チャットシートと同様に、実行時の正式名称をここへ反映します。
    ["PALETTE_SHEET_NAME", null, "パレットシート名。気に入らない時だけ変更。他と被らないように注意。"]
  ]
}

const SPREADCHAT_SHEET_SETTING_DEFAULT_ROW_HEIGHT = 21
const SPREADCHAT_SHEET_SETTING_SECRET_PREFIX =
  typeof SPREADCHAT_SHEET_GENERATOR !== "undefined" && SPREADCHAT_SHEET_GENERATOR.markers
    ? SPREADCHAT_SHEET_GENERATOR.markers.secretPrefix
    : (typeof secretPrefix !== "undefined" ? secretPrefix : "秘匿")

const SPREADCHAT_SHEET_SETTING_SPEC = {
  // シート名
  sheetName: SPREADCHAT_SHEET_SETTING.name,
  // 初期行数/列数
  initialRowCount: SPREADCHAT_SHEET_SETTING.initialRowCount,
  initialColumnCount: SPREADCHAT_SHEET_SETTING.initialColumnCount,
  // 固定行数/列数
  frozenRows: 1,
  frozenColumns: 0,
  // 既定の行の高さ
  defaultRowHeight: SPREADCHAT_SHEET_SETTING_DEFAULT_ROW_HEIGHT,
  // シート全体の基本スタイルです。
  styleRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_SETTING.initialRowCount,
    columnCount: SPREADCHAT_SHEET_SETTING.initialColumnCount,
    verticalAlignment: "top"
  }],
  // 折り返しは行わずクリップ表示にします。
  wrapStrategyRanges: [{
    rowStart: 1,
    columnStart: 1,
    rowCount: SPREADCHAT_SHEET_SETTING.initialRowCount,
    columnCount: SPREADCHAT_SHEET_SETTING.initialColumnCount,
    wrapStrategy: "CLIP"
  }],
  // 参加者別色分けは適用しません。
  participantColorBands: []
}

/**
 * Setting シートに書き込む行配列を返します。
 * シート上に「現在有効な設定値」を見せるため、共通設定で上書きされうる項目だけは
 * headerRows のテンプレート値ではなく実行時の値へ差し替えます。
 * @return {Array<Array<*>>}
 */
function buildSpreadChatSettingRows_() {
  return SPREADCHAT_SHEET_SETTING.values.map(row => {
    const cloned = row.slice()
    if (cloned[0] === "SECRET_PREFIX") {
      cloned[1] = SPREADCHAT_SHEET_SETTING_SECRET_PREFIX
    } else if (cloned[0] === "CHAT_SHEET_NAME") {
      cloned[1] = SPREADCHAT_SHEET_CHAT.name
    } else if (cloned[0] === "PALETTE_SHEET_NAME") {
      cloned[1] = SPREADCHAT_SHEET_PALETTE.name
    }
    return cloned
  })
}
