function googleDriveImageUrlLink(url) {
  var regex = /^https:\/\/drive\.google\.com\/file\/d\/(.+?)\/view\?usp=share_link$/;
  var match = url.match(regex);

  if (match) {
    var fileId = match[1];
    return 'https://drive.google.com/uc?id=' + fileId;
  } else {
    return url;
  }
}

function logFunctionNameAndArgs() {
  const functionName = logFunctionNameAndArgs.caller?.name || "anonymous";
  const args = [...arguments];
  console.log(`called:${functionName}`);
  if (args.length > 0) {
    console.log(`args:`, ...args);
  }
}

//行列変換
const transpose = a => a[0].map((_, c) => a.map(r => r[c]));

/**
 * シートからデータを取得し、キャッシュに格納して返します。
 * 以前に同じシート名で呼び出された場合は、キャッシュからデータを取得します。
 * @param {string} sheetName - データを取得するシートの名前
 * @param {boolean} [forceRefresh=false] - trueに設定すると、キャッシュを無視して新しいデータを取得します。デフォルトはfalseです。
 * @returns {object} シートから取得したデータ
 */
function getDictArrayFromSheetCache(sheetName, forceRefresh = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `getDictArrayFromSheetChash_${sheetName}`;
  
  if (forceRefresh) {
    cache.remove(cacheKey);
  }

  const cached = cache.get(cacheKey);
  if (cached !== null) {
    return JSON.parse(cached);
  }
  
  const result = getDictArrayFromSheet(sheetName);
  addScriptCache(cacheKey, result);
  
  return result;
}

function getDictFromSheetCache(sheetName, forceRefresh = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `getDictArrayFromSheetChash_${sheetName}`;
  
  if (forceRefresh) {
    cache.remove(cacheKey);
  }

  const cached = cache.get(cacheKey);
  if (cached !== null) {
    return JSON.parse(cached);
  }
  
  const result = getDictFromSheet(sheetName);
  addScriptCache(cacheKey, result);
  
  return result;
}

/**
 * 指定されたスプレッドシートの2列のデータをKey:Valueの連想配列で返します。
 */

function getDictFromSheet(sheetName) {
  const sheet = c_sheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const dict = {};
  for (let i = 0; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    dict[key] = value;
  }
  return dict;
}
/**
 * 指定されたスプレッドシートのデータを連想配列の配列で返します。
 * スプレッドシートの1行目を連想配列のキーとして使用します。
 */
function getDictArrayFromSheet(sheetName) {
  const sheet = c_sheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const result = data.slice(1).map(row => {
    const obj = {};
    row.forEach((value, index) => obj[headers[index]] = value);
    return obj;
  });
  return result;
}

/**
 * 指定されたスプレッドシートのデータを配列の連想配列で返します。
 * スプレッドシートの1行目を連想配列のキーとして使用します。
 */
function getNamedArraysFromSheet(sheetName) {
  const data = c_sheet(sheetName).getDataRange().getValues();
  const result = {};

  for (let j = 0; j < data[0].length; j++) {
    const header = data[0][j];
    result[header] = [];

    // 各行をループし、値がある場合のみ配列に追加
    for (let i = 1; i < data.length; i++) {
      if (data[i][j] === "" || data[i][j] === undefined) {
        break;
      }
      result[header].push(data[i][j]);
    }
  }

  return result;
}

/**
 * スクリプトキャッシュにデータを追加し、キーのリストにキーを登録します。
 * GASでは登録されているスクリプトキャッシュを列挙する方法がないため、
 * 対策として登録時にこの関数を使用することで、スクリプトキャッシュを一括でクリアできるようにします。
 * @param {string} key - キャッシュデータのキー
 * @param {*} value - キャッシュに保存するデータ
 */
function addScriptCache(key, value) {
  var cache = CacheService.getScriptCache();
  var keys = JSON.parse(cache.get('scriptCacheKeys')) || [];
  if (keys.indexOf(key) === -1) {
    keys.push(key);
  }
  cache.put('scriptCacheKeys', JSON.stringify(keys), 21600);
  cache.put(key, JSON.stringify(value), 21600);
}

/**
 * スクリプトキャッシュ内のすべてのデータを削除します。
 * addScriptCacheで登録したKeyが列挙されて削除されます。
 */
function clearScriptCache() {
  var cache = CacheService.getScriptCache();
  var keys = JSON.parse(cache.get('scriptCacheKeys')) || [];
  for (var i = 0; i < keys.length; i++) {
    cache.remove(keys[i]);
  }
  cache.remove('scriptCacheKeys');
}

/**
 * 選択範囲の文字配置指定をリセットします。 
 */
function resetAlignment() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  range.setHorizontalAlignment(null);
}
/**
 * 二次元配列を作成
 */
function create2DArray(numRow, numColumn) {
    var array = [];

    for (var i = 0; i < numRow; i++) {
        var row = [];

        for (var j = 0; j < numColumn; j++) {
            row.push('');
        }

        array.push(row);
    }

    return array;
}
