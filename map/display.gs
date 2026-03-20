function onEditChatMap(e) {
  logFunctionNameAndArgs();

  if (e.range.columnStart === CHAT_SHEET.nameCol && e.value === "__Map") {
    //名前列に管理者が__Mapを入力したらカレントマップ切り替え(追加)
    drawChatMap2(e.range.rowStart);
    return;
  }
}

/**
 * チャットシートの選択範囲をカレントマップとして登録
 */ 
function setCurrentMap() {
  // 選択された範囲を取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selectedRange = sheet.getActiveRange();
  const a1Notation = selectedRange.getA1Notation();
  
  // スクリプトプロパティに登録
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('CurrentMap', a1Notation);
  
  // ユーザーに通知
  SpreadsheetApp.getUi().alert('CurrentMapが登録されました: ' + a1Notation);
}

/**
 * 
 */ 
function copyCurrentMap(row) {
  // スクリプトプロパティを取得
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentMapRange = scriptProperties.getProperty('CurrentMap');
  
  // 「CurrentMap」の範囲が設定されているか確認
  if (!currentMapRange) {
    SpreadsheetApp.getUi().alert('CurrentMapを登録してください');
    return;
  }

  // 現在のシートを取得
  const sheet = c_sheet(CHAT_SHEET.name);
  const col = CHAT_SHEET.mapColStart;

  // 「CurrentMap」の範囲を取得してコピー
  // rowとcolはヘッダ部を含んでいるので、それぞれ+1
  const rangeToCopy = sheet.getRange(currentMapRange);
  const numRows = rangeToCopy.getNumRows();
  const numCols = rangeToCopy.getNumColumns();
  const destinationRange = sheet.getRange(row + 1, col +1 , numRows, numCols);

  // 値とフォーマットをコピー
  rangeToCopy.copyTo(destinationRange);

  //オーナーしか使用しないので行の高さ固定も可能。
  sheet.setRowHeightsForced(row, numRows, 21);  

  // スクリプトプロパティを更新
  const newA1Notation = destinationRange.getA1Notation();
  scriptProperties.setProperty('CurrentMap', newA1Notation);
}

function drawChatMap2(row) {
  copyCurrentMap(row);
  setRoundDeclare(row);
}

function openTile(row, col) {
  const SPREADSHEET_ID_A = '1J5vznBixtoUZz_ri8uSxFtA079bEjkLVy1ptfnab-D8';
  const SPREADSHEET_ID_B = '1dTpcJH0U_VOVvEZn-10Z7gEfD7eZADXGKBNMh0P80Aw';
  const SHEET_ID_X = '1983702346';
  const SHEET_ID_Y = '2107637972';
  
  const startRow = row - 1;
  const startCol = col - 1;
  const numRows = 3;
  const numCols = 3;

  const r1c1 = `Map!R${startRow}C${startCol}:R${startRow -1 + numRows}C${startCol -1 + numCols}`;
  console.log("r1c1:" + r1c1);

  // Fetch the format data from Spreadsheet B
  const sourceRange = Sheets.Spreadsheets.get(SPREADSHEET_ID_B, {
    ranges: [r1c1],
    fields: 'sheets.data.rowData.values.userEnteredFormat'
  }).sheets[0].data[0].rowData;

  // console.log(JSON.stringify(sourceRange));

  // Prepare the format data to paste into Spreadsheet A
 const rows = sourceRange.map(rowData => ({
    values: (rowData.values || []).map(cell => ({
      userEnteredFormat: cell ? cell.userEnteredFormat : {}
    }))
  }));

  // Create the update request
  const requests = [{
    updateCells: {
      range: {
        sheetId: SHEET_ID_X,
        startRowIndex: startRow - 1,
        endRowIndex: startRow -1 + numRows,
        startColumnIndex: startCol -1 ,
        endColumnIndex: startCol -1 + numCols
      },
      rows: rows,
      fields: 'userEnteredFormat'
    }
  }];

  // Apply the format to the destination sheet in Spreadsheet A
  Sheets.Spreadsheets.batchUpdate({
    requests: requests
  }, SPREADSHEET_ID_A);
}

function maskMapSheet() {
  // シートを取得
  const sheet = c_sheet(MAP_SHEET.name);

  // シートの全データを取得
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  const dataRange = sheet.getRange(1, 1, maxRows, maxColumns);
  const data = dataRange.getValues();

  // データをループして空の要素に"■"をセット
  for (let row = 0; row < data.length; row++) {
    for (let col = 0; col < data[row].length; col++) {
      if (data[row][col] === '') {
        data[row][col] = '■';
      }
    }
  }

  // シートにデータを貼り付け（A1から開始）
  dataRange.setValues(data);

  // 条件付き書式を適用
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('■')
    .setBackground('#333333')
    .setRanges([dataRange])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

const ROUND_DECLARE_PATTERN = getRoundDeclare();
function getRoundDeclare() {
  // スクリプトプロパティを取得
  const scriptProperties = PropertiesService.getScriptProperties();
  var pattern = scriptProperties.getProperty('RoundReclarePatternName');
  
  if (!pattern) {
    pattern="rqg";
    scriptProperties.setProperty('RoundReclarePatternName', pattern);
  }

  if (pattern == "rqg") {
    return [
          ["行動宣言", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:移動が終了したらq消去", "", "q", "q", "q", "q", "q", "q", "q", "q", "q", "q"],
          ["SR:01", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:02", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:03", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:04", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:05", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:06", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:07", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:08", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:09", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:10", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:11", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:12", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:行動宣言が終了したらq消去", "", "q", "q", "q", "q", "q", "q", "q", "q", "q", "q"],
          ["SR:全行動終了したらq消去", "", "q", "q", "q", "q", "q", "q", "q", "q", "q", "q"]
      ];
  }
  if (pattern == "rq3") {
    return [
          ["行動宣言", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:移動が終了したらq消去", "", "q", "q", "q", "q", "q", "q", "q", "q", "q", "q"],
          ["SR:01", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:02", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:03", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:04", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:05", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:06", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:07", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:08", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:09", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:10", "", "", "", "", "", "", "", "", "", "", ""],
          ["SR:行動宣言が終了したらq消去", "", "q", "q", "q", "q", "q", "q", "q", "q", "q", "q"],
          ["SR:全行動終了したらq消去", "", "q", "q", "q", "q", "q", "q", "q", "q", "q", "q"]
      ];
  }

  return [["roundReclarePatternNameが不正です。rqgもしくはrq3を設定してください。"]]
}

function setRoundDeclare(row) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentMapRange = scriptProperties.getProperty('CurrentMap');

  var sheet=c_sheet(CHAT_SHEET.name);
  sheet.getRange(row,CHAT_SHEET.actionCol,ROUND_DECLARE_PATTERN.length,12).setValues(ROUND_DECLARE_PATTERN);
}
