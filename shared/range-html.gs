const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

const htmlEscape = (str) => {
  if (!str) return;
  return str.replace(/[<>&"'`]/g, (match) => {
    const escape = {
      '<': '&lt;',
      '>': '&gt;',
      '&': '&amp;',
      '"': '&quot;',
      "'": '&#39;',
      '`': '&#x60;'
    };
    return escape[match];
  });
}

/**
 * アクティブなGoogle Spread Sheetの指定されたRangeをHtmlのtableに成型して戻します。
 * @param {integer} spreadsheetId
 * @param {text} rangeNotation
 * @return {text} tableHtml
 */
function getHtmlFromRange(spreadsheetId, rangeNotation) {
  if (rangeNotation == undefined) return "";

  var res = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: rangeNotation,
//      ranges: "チャット!R5C4",
      fields: "sheets/data/rowData/values(formattedValue,effectiveFormat/backgroundColorStyle/rgbColor,effectiveFormat/borders)"
    });
  var rowData = res.sheets[0].data[0].rowData;
  //  console.log(JSON.stringify(rowData));

  var tableHtml = `<table style="border-spacing:0px";>`;
  for (r=0; r<rowData.length; r++) {
    var row = rowData[r].values;
    tableHtml += "\n<tr>";
    for (c=0; c<row.length; c++) {
      var cell = row[c];      
      var rgbColor = cell?.effectiveFormat?.backgroundColorStyle?.rgbColor;
      var background = rgbColor? `background:${convertGoogleRGB(rgbColor)};`:""; 
      var border = convertGoogleBorders(cell?.effectiveFormat?.borders);
      var value = cell?.formattedValue ? htmlEscape(cell.formattedValue).replace(/\n/g, "</br>") : "";
        tableHtml += `\n<td style="${background} ${border}">${value}</td>`;
    }
    tableHtml += "\n</tr>";
  }
  tableHtml += "\n</table>";
  return tableHtml;
  //  console.log(tableHtml);
}

/**
 * Google Sheets APIのColorオブジェクトをCSSのStyle表記に変換して戻します。
 * @param {Sheets_v4.Sheets.V4.Schema.Color} color
 * @return {text}
 */
function convertGoogleRGB(color) {
  if (color == undefined) return "";
  var r = color?.red ? Math.floor(color.red*255) : 0;
  var g = color?.green ? Math.floor(color.green*255) : 0;
  var b = color?.blue ? Math.floor(color.blue*255) : 0;
  return `rgb(${r},${g},${b})`;
}

/**
 * Google Sheets APIのBordersオブジェクトをCSSのStyle表記に変換して戻します。
 * @param {Sheets_v4.Sheets.V4.Schema.Borders} borders
 * @return {text}
 */
function convertGoogleBorders(borders) {
  if (borders == undefined) return "";
  var text = "";
  if (borders.bottom) {
    text += "border-bottom:solid "
    text += borders.bottom?.width ? `${borders.bottom.width}px ` : "";
    text += borders.bottom?.colorStyle ? `${convertGoogleRGB(borders.bottom.colorStyle.rgbColor)} ` : "";
    text += ";";
  }
  if (borders.left) {
    text += "border-left:solid "
    text += borders.left?.width ? `${borders.left.width}px ` : "";
    text += borders.left?.colorStyle ? `${convertGoogleRGB(borders.left.colorStyle.rgbColor)} ` : "";
    text += ";";
  }
  if (borders.right) {
    text += "border-right:solid "
    text += borders.right?.width ? `${borders.right.width}px ` : "";
    text += borders.right?.colorStyle ? `${convertGoogleRGB(borders.right.colorStyle.rgbColor)} ` : "";
    text += ";";
  }
  if (borders.top) {
    text += "border-top:solid "
    text += borders.top?.width ? `${borders.top.width}px ` : "";
    text += borders.top?.colorStyle ? `${convertGoogleRGB(borders.top.colorStyle.rgbColor)} ` : "";
    text += ";";
  }
  return text;
}
