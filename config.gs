const version = "4.05"

//選択可能なダイスボット
//チャット上で省略名:で指定すると、その後の文字列がコマンドとしてbotに渡される。
const botKey = {
  "base":"botBase",
  "rq3":"botRq3",
  "c":"botChart",
  "rqg":"botRqg"
}

//bot指定がない場合に付加されるdicebot
const DEFAULT_BOT_KEY = "rq3";
const defaultBot = getDefaultBot();
function getDefaultBot() {
  // スクリプトプロパティを取得
  const scriptProperties = PropertiesService.getScriptProperties();
  var defaultBot = scriptProperties.getProperty('DefaultBot');

  if (defaultBot == "rqn" || !defaultBot || botKey[defaultBot] === undefined) {
    defaultBot = DEFAULT_BOT_KEY;
  }

  scriptProperties.setProperty('DefaultBot', defaultBot);
  return defaultBot;
}

//判定式列に何もないときのスラッシュのみの動作
const defaultCommand ="1d100";

//defaultBotが対応しないコマンドに対して呼び出されるdicebot
const baseBot = "base";

//コマンド文字列を複数コマンドとして分割する正規表現
//デフォルトはスペースですべて分割します。必要に応じて改行区切りにも変更できます。
//他の正規表現はあまり推奨しません。スプレッドチャットのコードが完全に読み解ける方だけお好きにどうぞ。
const cmdSplitter = "\s";
//const cmdSplitter = "\n";

//チャットシートの補助処理を使うには、シンプルトリガーからトリガーに切り替える必要があります。
const useTrigger = 1;

//秘匿情報として扱われる文字。アクション列・判定式列・PC列の背景色と文字色を同一にし、見えなくします。
//また、内容が見えてしまうため、PC列のノートは生成されなくなります。
const secretPrefix = "秘匿";

//ロール結果に付加される入力コマンドの区切り文字。
const joinResultChar = " ";
//\nを指定すると改行される。チャットのスペースを喰うので好み。
//const joinResultChar = "\n ";
//0文字列を指定すると付加されなくなる。おすすめはしない。
//const joinResultChar = "";

//チャットシート上のMap行を判定する接頭辞
const mapPrefix = "__";

//チャットシートの残り行数がaddRowPoint以下になるとaddRowNumber行を追加します。
const addRowPoint = 30;
const addRowNumber = 200;

//列幅と文字のバイト数を比較して、ノートの有無を決定する値。
//バイト数×font_byteWidth>列幅　ならばノートが表示される。
//例えば幅80の列に
//「あいうえお」なら　10byte×8>80　はfalseでノートなし。
//「あいうえおか」なら　12byte×8>80　はtrueでノートあり。
//実際のフォントや幅は計算しないのであくまで目安。
const font_byteWidth = 8;

//以下は自分でデザイン変更する人以外は変更しないでください。

//チャットシートに対する設定
// const chatSheetName = 'チャット';
const CHAT_SHEET_NAME = 'チャット';
const timestampCol = 1;
const rollIDCol = 2;
const imageCol = 4;
const nameCol = 5;
const actionCol = 6;
const formulaCol = 7;
const rollColStart = 8;
const rollColEnd = 17;
const mapColStart = 18;
const rollRowStart = 4;

const CHAT_SHEET = {
  name          : 'チャット',
  timestampCol  : 1,
  rollIDCol     : 2,
  imageCol      : 4,
  nameCol       : 5,
  actionCol     : 6,
  formulaCol    : 7,
  rollColStart  : 8,  //GM列が8、プレイヤー列が9～17
  rollColEnd    : 17,
  mapColStart   : 18, //Mapの行ヘッダ(Map本体は19から)
  rollRowStart  : 4
}

//パレットシートに対する設定
// const paletteSheetName = 'パレット';
const PALETTE_SHEET_NAME = 'パレット';
const palette_emailRow = 2
const palette_nameRow = 4;//パレットシートの名前行
const palette_imageRow = 5;//パレットシートの基本画像行
const palette_elementsStartRow = 14;//PCシートの技能の始まる行

const MAP_SHEET = {
  name          : "Map"
}


//MapDialogに対する設定
const SETTING_SHEET_NAME = 'Setting';
const CHARTS_SHEET_NAME = 'Charts';
