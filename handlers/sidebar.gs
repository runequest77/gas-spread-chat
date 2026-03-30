/**
 * キャラクターパレットサイドバーを表示します。
 */
function showCharacterPalette() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('キャラクターパレット')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * パレットシートに登録されているキャラクターの一覧を返します。
 * パレット番号 0（GM）～ 9（プレイヤー）のうち cname が設定されているものを返します。
 * @return {Array<{no: number, name: string}>}
 */
function getCharacterList() {
  var characters = [];
  for (var i = 0; i <= 9; i++) {
    try {
      var palette = getPalette(i);
      var cname = palette['cname'];
      if (cname && String(cname).trim() !== '') {
        characters.push({ no: i, name: String(cname).trim() });
      }
    } catch (e) {
      console.log('getCharacterList error at paletteNo=' + i + ': ' + e.message);
    }
  }
  return characters;
}

/**
 * 指定したパレット番号のキャラクターの技能・属性リストを返します。
 * パレットシートから直接読み取り、行番号付きで返します。
 * 内部管理キー（cname / pname / email / palette）と空キーは除外します。
 * @param {number} paletteNo - パレット番号（0～9）
 * @return {Array<{row: number, key: string, value: string}>}
 */
function getCharacterSkills(paletteNo) {
  var excludeKeys = ['palette', 'cname', 'pname', 'email'];
  var sheet = c_sheet(PALETTE_SHEET_NAME);
  var col = paletteNo * 2 + 1;
  var d2Array = sheet.getRange(1, col, sheet.getMaxRows(), 2).getDisplayValues();
  var skills = [];
  for (var i = 0; i < d2Array.length; i++) {
    var key = d2Array[i][0];
    var value = d2Array[i][1];
    if (!key || key.trim() === '') continue;
    // 秘匿マーカー「■」が先頭についているキーはスキップ
    if (key.charAt(0) === '\u25a0') continue;
    if (excludeKeys.indexOf(key) !== -1) continue;
    if (value === undefined || value === null) continue;
    skills.push({ row: i + 1, key: key, value: String(value) });
  }
  return skills;
}

/**
 * サイドバーから編集された技能データをパレットシートに保存します。
 * 各エントリの row をキーとして、該当セルのキー・値を上書きします。
 * 保存後はパレットキャッシュをクリアします。
 * @param {number} paletteNo - パレット番号（0～9）
 * @param {Array<{row: number, key: string, value: string}>} skills - 保存する技能データ
 */
function saveCharacterSkills(paletteNo, skills) {
  var sheet = c_sheet(PALETTE_SHEET_NAME);
  var col = paletteNo * 2 + 1;
  for (var i = 0; i < skills.length; i++) {
    var s = skills[i];
    var row = s.row;
    if (row < palette_elementsStartRow) continue; // ヘッダ・画像行は保護
    sheet.getRange(row, col).setValue(s.key);
    sheet.getRange(row, col + 1).setValue(s.value);
  }
  clearPaletteCache(paletteNo);
}

/**
 * アクティブセルにロールコマンドを入力します。
 * @param {string} command - 入力するロールコマンド（例: /剣）
 */
function insertRollCommandToActiveCell(command) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getActiveCell();
  cell.setValue(command);
}
