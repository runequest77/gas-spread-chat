
function botChart(command){
  console.log("command:" + JSON.stringify(command) );

  const cache = CacheService.getScriptCache();
  const cacheKey = `BOT_CHARTS_${command}`;
  let chart = [];

  const cached = cache.get(cacheKey);
  if (cached !== null) {
    chart = JSON.parse(cached);
  } else {
    const dic = getNamedArraysFromSheet(CHARTS_SHEET_NAME);
    if (dic[command]) {
      addScriptCache(cacheKey, dic[command]);
      chart = dic[command];
    }
  }
  
  if (chart.length > 0) {
    var roll = XdY(1,chart.length);
    return chart[roll - 1];
  } else {
    return `チャート[ ${command} ]は登録されていません。`;
  }
}

function botChartClear() {
  var cache = CacheService.getScriptCache();
  var keys = JSON.parse(cache.get('scriptCacheKeys')) || [];
  const filteredKeys = keys.filter(key => key.startsWith('BOT_CHARTS_'));
  for (var i = 0; i < filteredKeys.length; i++) {
    if (filteredKeys[i]) {
      console.log("cache.remove:" + filteredKeys[i]);
      cache.remove(filteredKeys[i]);
    }
  }
}


// function botChart(command) {
//   const d2Array = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chart").getDataRange().getValues();
//   //チャート名を取得
//   var chartName = "c:" + command;
//   console.log("chartName:" + chartName);
//   if (chartName == null) return "コマンド[ " + command + " ]は解析できませんでした。";
//   var header = d2Array[0];
//   console.log("header:" + header);
//   var col;
//   for (var i=0; i<header.length; i++) {
//     if (header[i] == chartName) {
//       col = i;
//       break;
//     }
//   }
//   console.log("col:" + col);
  
//   var chart = [];  
//   for (var i=0; i<d2Array.length; i++) {
//     if (d2Array[i][col] != "") chart.push(d2Array[i][col]);
//   }

//   var roll = XdY(1,chart.length);
//   result = chart[roll - 1];  
  
//   return result;
// }
