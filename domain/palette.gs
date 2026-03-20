//高速化のためにパレットシートは起動時に読み込む
//必要に応じてリフレッシュする
const palette = [];

function getPalette(paletteNo) {
  console.time("キャッシュ戻し");
  const cache = CacheService.getScriptCache();
//  var cache = CacheService.getScriptCache();
  var paletteJson = cache.get('palette' + paletteNo);
  if (paletteJson) {
    //キャッシュがあればそれを戻す。
    console.log("キャッシュ効いた！");
    return JSON.parse(paletteJson);
  }
  console.timeEnd("キャッシュ戻し");

  console.time("キャッシュ作成");
  //なければシートから値を取得し、キャッシュ化してそれを戻す
  var sheet = c_sheet(PALETTE_SHEET_NAME);
  var col = paletteNo *2 +1;
  var d2Array = sheet.getRange(1,col,sheet.getMaxRows(),2).getDisplayValues();
  var palette = safetyArrayToHash(d2Array);

  cache.put('palette' + paletteNo, JSON.stringify(palette), 60 * 60 * 6);
  console.timeEnd("キャッシュ作成");

  return palette;
}

function clearPaletteCache(paletteNo) {
  var cache = CacheService.getScriptCache();
  cache.remove("palette" + paletteNo);
}

