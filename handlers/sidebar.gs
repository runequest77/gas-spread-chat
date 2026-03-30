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
 * 内部管理キー（cname / pname / email / palette）と空キーは除外します。
 * @param {number} paletteNo - パレット番号（0～9）
 * @return {Array<{key: string, value: string}>}
 */
function getCharacterSkills(paletteNo) {
  var excludeKeys = ['palette', 'cname', 'pname', 'email'];
  var palette = getPalette(paletteNo);
  var skills = [];
  for (var key in palette) {
    if (!key || key.trim() === '') continue;
    // 秘匿マーカー「■」が先頭についているキーはスキップ
    if (key.charAt(0) === '\u25a0') continue;
    if (excludeKeys.indexOf(key) !== -1) continue;
    var value = palette[key];
    if (value === undefined || value === null) continue;
    skills.push({ key: key, value: String(value) });
  }
  return skills;
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
