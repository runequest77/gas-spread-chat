//[]は内部を計算するの意味
//XdYはすべて自動計算する
//演算可能な数値はすべて演算する
//[種別:XX]は判定系を表す

// Object.prototype.getName = function() { 
//    var funcNameRegex = /function (.{1,})\(/;
//    var results = (funcNameRegex).exec((this).constructor.toString());
//    return (results && results.length > 1) ? results[1] : "";
// };
//シンプルトリガー
/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEdit(e) {
  logFunctionNameAndArgs('e省略');
  console.log("isOwner():" + isOwner());
  console.log("JSON.stringify(e.range):" + JSON.stringify(e.range));

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
    console.log("sheetName:" + sheetName);
  //キャッシュされているシートは編集後にキャッシュをクリアする。
  //連続して編集されたときにキャッシュ生成は無駄なので、クリアのみ。生成は使用時。
  if (sheetName == CHAT_SHEET_NAME) {
    onEditChatSheet(e);
  } else if (sheetName == PALETTE_SHEET_NAME) {
    onEditPaletteSheet(e);
  } else if (sheetName == SETTING_SHEET_NAME) {
    CacheService.getScriptCache().remove("getDictArrayFromSheetChash_" + SETTING_SHEET_NAME);
  } else if (sheetName == CHARTS_SHEET_NAME) {
    botChartClear();
    CacheService.getScriptCache().remove("getDictArrayFromSheetChash_" + CHARTS_SHEET_NAME);
  }
}
/**
 * チャットシート
 */
function onEditChatSheet(e) {
  const sheet = c_sheet(CHAT_SHEET_NAME);
  const cell = e.range;
  const e_pos = getEventPosition(cell);
  var row = e_pos.rowStart;

  if (row < rollRowStart) return;  //ヘッダ行は処理しない

  console.time("3a★");
  var rowData = c_rowData(row);
  console.timeEnd("3a★");

  var editAction = getChatSheetEditAction(rowData, e_pos.columnStart);
  console.log("editAction:" + editAction);

  switch (editAction) {
    case "name":
      onEditNameCol(e, e_pos);
      return;
    case "roll":
      onEditRollCol(sheet, e_pos, cell, e.value, rowData);
      return;
    case "map":
      onEditChatMap(e);
      return;
    default:
      return;
  }
}

function getChatSheetEditAction(rowData, col) {
  if (isChatMapRow(rowData)) return "map";
  if (col == nameCol) return "name";
  if (isRollColumn(col)) return "roll";
  if (col >= CHAT_SHEET.mapColStart) return "map";
  return "";
}

function isChatMapRow(rowData) {
  return String(rowData[nameCol - 1] || "").startsWith(mapPrefix);
}

function isRollColumn(col) {
  return col >= rollColStart && col <= rollColEnd;
}
/**
 * チャットシート/名前列
 */
function onEditNameCol(e,e_pos) {
  //名前列の場合、先頭数字X1-X9なら画像表示して終了。数字一桁はプレイヤー名、下桁0はPC名のみ。
  //0はGM、1-9プレイヤー、10-99がNPC。下一桁が行OFFSET　0の時は何もしない
  var sheet = c_sheet(CHAT_SHEET_NAME);
  var row = e_pos.rowStart;
  var imageFormula = buildNameImageFormula(e.value);
  if (imageFormula != "") sheet.getRange(row, imageCol).setValue(imageFormula);
}

/**
 * チャットシート/ロール列
 */
function onEditRollCol(sheet, pos, cell, value, rowData) {
  //ノートをクリア
  if (value === undefined) {
    cell.clearNote();
    return;
  }
  //ロールコマンドでなければノートを生成して終了
  if (value.charAt(0) != "/") {
    setNoteOverWidth(sheet,cell,value);
    return;
  }

  console.time("4b");
  var command = buildRollCommand(value, rowData[formulaCol - 1]);
  console.timeEnd("4b");

  console.time("5");
  var row = pos.rowStart;
  var userNo = getUserNoFromRollColumn(pos.columnStart);
  // 現時点でpermissionsは参照していないが、slash-command のロール経路ではキャッシュを温めておく。
  c_permissions(userNo);

  var rollContext = resolveRollCommandContext(command, userNo, rowData[nameCol - 1]);
  command = rollContext.command;
  console.timeEnd("5");

  console.time("6");
  var pc = buildRollPc(userNo, rollContext.paletteNo);
  console.log('pc:' + JSON.stringify(pc));

  var replacedCommand = replaceRollCommand(command, pc.elements);
  console.timeEnd("6");

  console.time("7");
  console.log("replacedCommand:" + replacedCommand);
  //実行本体　pcいる？
  var result = executeCommand(replacedCommand,pc);
  console.timeEnd("7");

  console.time("8");
  //スプレッドチャットは行の追加削除ができてしまうので、あとでチャットとログの照らし合わせができるように、RollIDを記録する。
  //速度向上のための仕様変更→ダイスをロールした時のRowでいいこととする。
  //RollId/CreateDate/Name/Command/Result
  var rollLog = createRollLog(row, pc.name, command, result);
  writeRollLog(rollLog);  //★
  console.timeEnd("8");

  //timestamp未記録ならセット 
  console.time("9");
  writeTimestampIfMissing(sheet, row, rowData);
  console.timeEnd("9");

  console.time("10");
  cell.setValue(rollLog.Result);  //★
  console.timeEnd("10");

}

function createRollLog(rollID, pcName, command, result) {
  return {
    'RollID': rollID,
    'CreateDate': new Date(),
    'Name': pcName,
    'Command': command,
    'Result': result,
  };
}

function writeTimestampIfMissing(sheet, row, rowData) {
  let ts = [rowData[timestampCol - 1], rowData[rollIDCol - 1]];
  //いずれかがなければ書き込む
  if (ts[0] == "" || ts[1] == "") {
    if (ts[0] == "") ts[0] = new Date();
    if (ts[1] == "") ts[1] = row;
    sheet.getRange(row, timestampCol, 1, 2).setValues([ts]);
  }
}

/**
 * パレットシートが編集されたら、当該ユーザーのパレットキャッシュ破棄
 */
function onEditPaletteSheet(e) {
  const e_pos = getEventPosition(e.range);
  var col = e_pos.columnStart;
  let cNo = Math.floor((col-1) / 2); 
  clearPaletteCache(cNo);
}

function getEventPosition(range) {
  if (range && typeof range === "object" && typeof range.getRow === "function" && typeof range.getColumn === "function") {
    const rowStart = range.getRow();
    const columnStart = range.getColumn();
    const rowEnd = (typeof range.getLastRow === "function") ? range.getLastRow() : rowStart;
    const columnEnd = (typeof range.getLastColumn === "function") ? range.getLastColumn() : columnStart;
    return { rowStart, rowEnd, columnStart, columnEnd };
  }
  return {
    rowStart: range && range.rowStart,
    rowEnd: range && range.rowEnd,
    columnStart: range && range.columnStart,
    columnEnd: range && range.columnEnd,
  };
}

/**
 * ユーザーがどのuserNo(0-9)に対して権限を持っているかをパレットシートの書き込み権限を元に返す
 * onEditもしくはメニュー経由で呼ばれる。
 */
function getPermissions() {
  // パレットシートを取得
  const sheet = c_sheet(PALETTE_SHEET_NAME);
  const range = sheet.getRange("A2:T3");//2行20列
  const values = range.getValues();
  const permissions = [];
  
  for (let userNo = 0; userNo < 10; userNo++) {
    const col = userNo * 2 + 2;
    const email = values[0][col - 1];
    if (email) {
      // emailが存在する場合のみ、ユーザーを追加する
      if (sheet.getRange(3, col).canEdit()) {
      // 実行ユーザーの権限をチェック
        permissions.push(userNo);
      }
    }
  }
  return permissions;
}
/**
 * userNoで生成したパーミッションのキャッシュを返します。
 * オーナーが編集したときはキャッシュしないようにしています。
 */
function c_permissions(userNo) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `c_permissions_${userNo}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    console.log("キャッシュを返す");  
    return JSON.parse(cached);
  } else if (isOwner()) {
    console.log("オーナー実行");  
    const ui = SpreadsheetApp.getUi();
//    ui.alert('オーナー実行', `userNo:${userNo}`,ui.ButtonSet.OK);
    return userNo;
  } else {
    console.log("キャッシュして返す");  
    const permissions =getPermissions();
    addScriptCache(cacheKey, permissions);
    return permissions;
  }
}
function isOwner() {
  return (SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail()==Session.getEffectiveUser().getEmail());
}

function getRollID(){
  const cache = CacheService.getScriptCache();
  var rollID = cache.get('rollID');
  if (rollID) {
    console.log("キャッシュ効いた！");
    //キャッシュをカウントアップして戻す。
    rollID = Number(rollID) + 1;
    cache.put('rollID', rollID, 60 * 60 * 6);
    return rollID;
  }
  console.timeEnd("キャッシュ戻し");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_SHEET_NAME);  //★
  var values = sheet.getRange(rollRowStart,rollIDCol,sheet.getMaxRows(),1).getValues();//★
  var ids = transpose(values)[0];

  console.log("JSON.stringify(ids):" + JSON.stringify(ids));
  let idsf = ids.filter((v) => v.toString().match(/[0-9]{1,4}/g));
  console.log("JSON.stringify(idsf):" + JSON.stringify(idsf));
  let max = Math.max(...idsf);
  console.log("max:" + max);
//  return 0;
} 

// function executeCommand(command,pc){
//   console.log("executeCommand-command:" + command);

//   //空白文字（スペースや改行）で複数コマンドに分割
//   var sepS = command.split(/\s/);
//   var result = [];
//   var secret = "";
//   for (let i = 0; i < sepS.length; ++i) {
//     if (sepS[i].length>0) {

//       //コロンで分割してbot指定の有無を判定      
//       var cmd = sepS[i].split(':');
//       if (cmd.length==1) {
//         cmd.unshift(defaultBot);
//       }
//       console.log("cmd:" + cmd);
//       const addResult = (joinResultChar) ? joinResultChar + cmd[0] + ":" + cmd[1]:"";
//       cmd[1] = cmd[1].replace(/\[[^\[\]]*\]/,"");
//       result.push(botRoll(cmd[0],cmd[1]) + addResult);
//     }

//   }

//   console.log("executeCommand-result:" + result);
//   return result.join("\n");
// }

function writeRollLog(rollLog){
  console.log("writeRollLog:" + JSON.stringify(rollLog));
  //RollId/CreateDate/Name/Command/Result
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RollLog");
  var output = [rollLog.RollID
                ,rollLog.CreateDate
                ,rollLog.Name
                ,rollLog.Command
                ,rollLog.Result];
  sheet.appendRow(output);
}

//中括弧{}が含まれていれば項目名を値に置換
function replaceCommandByPalette(command,palette) {
  console.log("replaceCommandByPalette:" + [command,palette]);
  var sepCommand = command.split(/({[^{}]*})/);
  if (sepCommand.length == 1) return command;
  var result = sepCommand.map(item => replaceStringByPalette(item,palette)).join("");

  //ボット指定のコロンが連続している場合、ボットを除去(ボットの差し替えに使用)、
  console.log("replaceCommandByPalette/result:" + result);
  result2 = result.replace(/([^\s:]*::)/, '')
  console.log("replaceCommandByPalette/result2:" + result2);
  return result2;
}

//{}で括られていた場合、paletteのItem0と一致するItem1で置換する
function replaceStringByPalette(command,palette){
  var key = /{(.*)}/.exec(command); //  console.log("key:" + key);
  if (key == null) return command;
  var value = getItem1ByItem0(palette,key[1]); //  console.log("value:" + value);
  if (value == null) return key[1] + "は設定されていません。";
  return value;
}

function getPaletteElements(paletteNo) {
  var sheet = c_sheet(PALETTE_SHEET_NAME);
  var col = paletteNo * 2+1;
  var d2Array = sheet.getRange(palette_nameRow,col,sheet.getMaxRows(),2).getDisplayValues();
  return d2Array;
}

function safetyArrayToHash(d2Array){
  var result = {};
    for (let i = 0; i < d2Array.length; i++) {
        var k = d2Array[i][0];
        var v = d2Array[i][1];
    if (k != undefined && v != undefined) result[replaceMultibyteForCalc(k)] = v;
  }
  return result;
}

function getItem1ByItem0(d2Array,key){
  //パレットシートは自由入力なのでkeyが重複する可能性がある。
  //仕様として先頭を取得。keyに入る文字も不定なのでハッシュにはしにくい。
  var pair = d2Array.find(item => (item[0] == key));
  return pair === undefined ? undefined : pair[1] ;
}

function registerFunction(name , response){
  Function('return this')()[name](response);
}

function setNoteOverWidth(sheet, cell, text) {
  //秘匿情報はNoteを作成しない。
  var regex = new RegExp("^" + secretPrefix);
  if (regex.test(text)) return;

  var w = sheet.getColumnWidth(cell.getColumn());

  var arr_bytelength = text.split(/\n/).map(t => getBytes(t));
  var max_bytelength = Math.max(...arr_bytelength);

  if (max_bytelength * font_byteWidth > w) {
    cell.setNote(text);
  } else {
    cell.clearNote();
  }
}
function getBytes(text) {
  var length = 0;
  for (var i = 0; i < text.length; i++) {
    var c = text.charCodeAt(i);
    if ((c >= 0x0 && c < 0x81) || (c === 0xf8f0) || (c >= 0xff61 && c < 0xffa0) || (c >= 0xf8f1 && c < 0xf8f4)) {
      length += 1;
    } else {
      length += 2;
    }
  }
  return length;
};
