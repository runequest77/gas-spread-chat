const SPREADSHEET_INSPECTOR = {
  exportSheetName: "_SchemaExport",
  sampleRowCount: 20,
  sampleColumnCount: 26,
  chunkSize: 40000,
  snapshotFilePath: "inspection/schema-snapshot.gs",
  driveSnapshotFileName: "spreadsheet-schema-snapshot.gs",
  driveSnapshotFileIdPropertyKey: "spreadsheetSchemaSnapshotDriveFileId",
  appsScriptApiBaseUrl: "https://script.googleapis.com/v1/projects",
  supportedApiMethods: ["get", "put"]
}

/**
 * リファクタリング用に、シート構造のサマリをJSONでログ出力します。
 * 既定では値を伏せ、数式・書式・列幅などの構造情報だけを出力します。
 */
function exportSpreadsheetSchemaToLog() {
  const json = getSpreadsheetSchemaSnapshotJson()
  console.log(json)
  return json
}

/**
 * GitHub Assistantで同期しやすい .gs ファイル内容をログ出力します。
 * 出力結果を inspection/schema-snapshot.gs に貼り付ける運用を想定します。
 * @param {object=} options
 * @return {string}
 */
function exportSpreadsheetSchemaSourceToLog(options) {
  const source = getSpreadsheetSchemaSnapshotSource(options)
  console.log(source)
  return source
}

/**
 * スキーマスナップショットを Apps Script プロジェクト内の .gs ファイルへ直接反映します。
 * Apps Script API が使えない場合は Drive 上のテキストファイルへ自動フォールバックします。
 * @param {object=} options
 * @return {{mode:string, scriptId:(string|undefined), fileName:string, sourceLength:number, fileId:(string|undefined), url:(string|undefined), fallbackReason:(string|undefined)}}
 */
function exportSpreadsheetSchemaToProjectFile(options) {
  const source = getSpreadsheetSchemaSnapshotSource(options)
  try {
    const result = updateAppsScriptProjectFile_(SPREADSHEET_INSPECTOR.snapshotFilePath, source)
    return Object.assign({ mode: "project" }, result)
  } catch (error) {
    if (!shouldFallbackToDriveFile_(error)) {
      throw error
    }
    return exportSpreadsheetSchemaToDriveFileFromSource_(source, error.message)
  }
}

/**
 * スキーマスナップショットを Drive 上のテキストファイルへ書き出します。
 * Apps Script API が使えない環境向けのフォールバックです。
 * @param {object=} options
 * @return {{mode:string, fileId:string, fileName:string, url:string, sourceLength:number, fallbackReason:(string|null)}}
 */
function exportSpreadsheetSchemaToDriveFile(options) {
  const source = getSpreadsheetSchemaSnapshotSource(options)
  return exportSpreadsheetSchemaToDriveFileFromSource_(source)
}

/**
 * リファクタリング用に、シート構造のサマリを専用シートへ出力します。
 * JSONが長くなる場合は複数行に分割します。
 */
function exportSpreadsheetSchemaToSheet() {
  const json = getSpreadsheetSchemaSnapshotJson()
  const ss = c_ss()
  const sheetName = SPREADSHEET_INSPECTOR.exportSheetName
  let sheet = ss.getSheetByName(sheetName)
  if (!sheet) {
    sheet = ss.insertSheet(sheetName)
  }

  sheet.clearContents()
  sheet.getRange(1,1,1,2).setValues([["part", "json"]])

  const chunks = splitTextByLength_(json, SPREADSHEET_INSPECTOR.chunkSize)
  if (chunks.length > 0) {
    const values = chunks.map((chunk, index) => [index + 1, chunk])
    sheet.getRange(2,1,values.length,2).setValues(values)
  }

  sheet.autoResizeColumns(1,2)
  return json
}

function getSpreadsheetSchemaSnapshotJson(options) {
  const snapshot = getSpreadsheetSchemaSnapshot(options)
  return JSON.stringify(snapshot, null, 2)
}

/**
 * スナップショットを Git 管理しやすい整形済み .gs ソース文字列へ変換します。
 * @param {object=} options
 * @return {string}
 */
function getSpreadsheetSchemaSnapshotSource(options) {
  const snapshot = getSpreadsheetSchemaSnapshot(options)
  return getSpreadsheetSchemaSnapshotSourceFromObject_(snapshot)
}

function getSpreadsheetSchemaSnapshot(options) {
  const normalizedOptions = normalizeSpreadsheetInspectorOptions_(options)
  const ss = c_ss()

  return {
    exportedAt: new Date().toISOString(),
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    directSheetAccessFromAgent: false,
    options: normalizedOptions,
    namedRanges: ss.getNamedRanges().map(namedRange => ({
      name: namedRange.getName(),
      sheetName: namedRange.getRange().getSheet().getName(),
      a1Notation: namedRange.getRange().getA1Notation()
    })),
    sheets: ss.getSheets().map(sheet => describeSheetForRefactoring_(sheet, normalizedOptions))
  }
}

function describeSheetForRefactoring_(sheet, options) {
  const lastRow = sheet.getLastRow()
  const lastColumn = sheet.getLastColumn()
  const sampledRowCount = Math.min(lastRow, options.sampleRowCount)
  const sampledColumnCount = Math.min(lastColumn, options.sampleColumnCount)

  const summary = {
    name: sheet.getName(),
    sheetId: sheet.getSheetId(),
    hidden: sheet.isSheetHidden(),
    lastRow: lastRow,
    lastColumn: lastColumn,
    maxRows: sheet.getMaxRows(),
    maxColumns: sheet.getMaxColumns(),
    frozenRows: sheet.getFrozenRows(),
    frozenColumns: sheet.getFrozenColumns(),
    sampledRange: null,
    columnWidths: getColumnWidths_(sheet, sampledColumnCount),
    rows: []
  }

  if (sampledRowCount === 0 || sampledColumnCount === 0) {
    return summary
  }

  const range = sheet.getRange(1, 1, sampledRowCount, sampledColumnCount)
  const displayValues = range.getDisplayValues()
  const formulas = range.getFormulas()
  const backgrounds = range.getBackgrounds()
  const fontColors = range.getFontColors()
  const fontFamilies = range.getFontFamilies()
  const fontSizes = range.getFontSizes()
  const fontWeights = range.getFontWeights()
  const numberFormats = range.getNumberFormats()
  const horizontalAlignments = range.getHorizontalAlignments()
  const verticalAlignments = range.getVerticalAlignments()
  const wrapStrategies = range.getWrapStrategies()
  const notes = range.getNotes()

  summary.sampledRange = buildSheetRangeA1_(sheet.getName(), sampledRowCount, sampledColumnCount)
  summary.rows = displayValues.map((row, rowIndex) => ({
    row: rowIndex + 1,
    cells: row.map((displayValue, columnIndex) => describeCellForRefactoring_({
      row: rowIndex + 1,
      column: columnIndex + 1,
      displayValue: displayValue,
      formula: formulas[rowIndex][columnIndex],
      background: backgrounds[rowIndex][columnIndex],
      fontColor: fontColors[rowIndex][columnIndex],
      fontFamily: fontFamilies[rowIndex][columnIndex],
      fontSize: fontSizes[rowIndex][columnIndex],
      fontWeight: fontWeights[rowIndex][columnIndex],
      numberFormat: numberFormats[rowIndex][columnIndex],
      horizontalAlignment: horizontalAlignments[rowIndex][columnIndex],
      verticalAlignment: verticalAlignments[rowIndex][columnIndex],
      wrapStrategy: String(wrapStrategies[rowIndex][columnIndex] || ""),
      hasNote: notes[rowIndex][columnIndex] !== ""
    }, options))
  }))

  return summary
}

function describeCellForRefactoring_(cell, options) {
  const hasFormula = cell.formula !== ""
  const isBlank = cell.displayValue === ""

  return {
    a1Notation: `${columnNumberToLetter_(cell.column)}${cell.row}`,
    value: redactCellValue_(cell.displayValue, hasFormula, isBlank, options.includeLiteralValues),
    formula: hasFormula ? cell.formula : "",
    format: {
      background: cell.background,
      fontColor: cell.fontColor,
      fontFamily: cell.fontFamily,
      fontSize: cell.fontSize,
      fontWeight: cell.fontWeight,
      numberFormat: cell.numberFormat,
      horizontalAlignment: cell.horizontalAlignment,
      verticalAlignment: cell.verticalAlignment,
      wrapStrategy: cell.wrapStrategy,
      hasNote: cell.hasNote
    }
  }
}

function redactCellValue_(displayValue, hasFormula, isBlank, includeLiteralValues) {
  if (isBlank) {
    return ""
  }
  if (includeLiteralValues) {
    return displayValue
  }
  if (hasFormula) {
    return "[formula-result]"
  }
  if (/^(TRUE|FALSE)$/i.test(displayValue)) {
    return "[boolean]"
  }
  if (/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(displayValue)) {
    return "[number]"
  }
  return "[text]"
}

function normalizeSpreadsheetInspectorOptions_(options) {
  const opts = options || {}
  const sampleRowCount = Number(opts.sampleRowCount)
  const sampleColumnCount = Number(opts.sampleColumnCount)
  return {
    includeLiteralValues: opts.includeLiteralValues === true,
    sampleRowCount: Number.isFinite(sampleRowCount) && sampleRowCount > 0 ? sampleRowCount : SPREADSHEET_INSPECTOR.sampleRowCount,
    sampleColumnCount: Number.isFinite(sampleColumnCount) && sampleColumnCount > 0 ? sampleColumnCount : SPREADSHEET_INSPECTOR.sampleColumnCount
  }
}

function getColumnWidths_(sheet, columnCount) {
  const widths = []
  for (let column = 1; column <= columnCount; column++) {
    widths.push({
      column: columnNumberToLetter_(column),
      width: sheet.getColumnWidth(column)
    })
  }
  return widths
}

function buildSheetRangeA1_(sheetName, rowCount, columnCount) {
  const escapedSheetName = String(sheetName).replace(/'/g, "''")
  return `'${escapedSheetName}'!A1:${columnNumberToLetter_(columnCount)}${rowCount}`
}

function columnNumberToLetter_(columnNumber) {
  let n = columnNumber
  let result = ""

  while (n > 0) {
    const mod = (n - 1) % 26
    result = String.fromCharCode(65 + mod) + result
    n = Math.floor((n - 1) / 26)
  }

  return result
}

function splitTextByLength_(text, chunkSize) {
  const chunks = []
  for (let start = 0; start < text.length; start += chunkSize) {
    chunks.push(text.slice(start, start + chunkSize))
  }
  return chunks
}

function updateAppsScriptProjectFile_(filePath, source) {
  const scriptId = ScriptApp.getScriptId()
  const content = callAppsScriptProjectContentApi_(scriptId, "get")
  const updatedFiles = replaceAppsScriptProjectFile_(content.files || [], filePath, source)
  callAppsScriptProjectContentApi_(scriptId, "put", { files: updatedFiles })
  const file = findAppsScriptProjectFile_(updatedFiles, filePath)
  return {
    scriptId: scriptId,
    fileName: file ? file.name : buildAppsScriptFileNameFromPath_(filePath),
    sourceLength: source.length
  }
}

function exportSpreadsheetSchemaToDriveFileFromSource_(source, fallbackReason) {
  const file = upsertSpreadsheetSchemaDriveFile_(source)
  return {
    mode: "drive",
    fileId: file.getId(),
    fileName: file.getName(),
    url: file.getUrl(),
    sourceLength: source.length,
    fallbackReason: fallbackReason || null
  }
}

function callAppsScriptProjectContentApi_(scriptId, method, payload) {
  const normalizedMethod = normalizeAppsScriptApiMethod_(method)
  const displayMethod = normalizedMethod.toUpperCase()
  const accessToken = getAppsScriptApiAccessToken_()
  const response = UrlFetchApp.fetch(buildAppsScriptProjectContentUrl_(scriptId), {
    method: normalizedMethod,
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${accessToken}`
    },
    payload: payload ? JSON.stringify(payload) : undefined,
    muteHttpExceptions: true
  })
  const status = response.getResponseCode()
  const text = response.getContentText()
  if (status < 200 || status >= 300) {
    const error = new Error(`Apps Script API ${displayMethod} failed (${status}): ${extractAppsScriptApiErrorMessage_(text)}`)
    error.statusCode = status
    error.apiErrorMessage = extractAppsScriptApiErrorMessage_(text)
    throw error
  }
  return text ? JSON.parse(text) : {}
}

function buildAppsScriptProjectContentUrl_(scriptId) {
  return `${SPREADSHEET_INSPECTOR.appsScriptApiBaseUrl}/${encodeURIComponent(scriptId)}/content`
}

function replaceAppsScriptProjectFile_(files, filePath, source) {
  const sanitizedFiles = files.map(sanitizeAppsScriptProjectFile_)
  const sanitizedExistingFile = findAppsScriptProjectFile_(sanitizedFiles, filePath)
  const baseName = sanitizedExistingFile ? null : getAppsScriptFileNameCandidates_(filePath).slice(-1)[0]
  const nextFile = {
    name: sanitizedExistingFile ? sanitizedExistingFile.name : baseName,
    type: sanitizedExistingFile ? sanitizedExistingFile.type : "SERVER_JS",
    source: source
  }

  if (sanitizedExistingFile) {
    return sanitizedFiles.map(file => file.name === sanitizedExistingFile.name ? nextFile : file)
  }

  return sanitizedFiles.concat(nextFile)
}

function findAppsScriptProjectFile_(files, filePath) {
  const candidates = getAppsScriptFileNameCandidates_(filePath)
  return files.find(file => file && candidates.includes(file.name)) || null
}

function getAppsScriptFileNameCandidates_(filePath) {
  const fileName = buildAppsScriptFileNameFromPath_(filePath)
  if (!fileName) {
    throw new Error("Apps Script の更新対象ファイル名が空です。snapshotFilePath を確認してください。")
  }
  const baseName = fileName.split("/").pop()
  return [fileName, baseName]
}

function buildAppsScriptFileNameFromPath_(filePath) {
  return String(filePath).replace(/\.gs$/, "").replace(/\/+$/, "")
}

function sanitizeAppsScriptProjectFile_(file) {
  return {
    name: file.name,
    type: file.type,
    source: file.source || ""
  }
}

function upsertSpreadsheetSchemaDriveFile_(source) {
  const scriptProperties = PropertiesService.getScriptProperties()
  const fileIdKey = SPREADSHEET_INSPECTOR.driveSnapshotFileIdPropertyKey
  const storedFileId = scriptProperties.getProperty(fileIdKey)
  const file = storedFileId ? getDriveFileByIdOrNull_(storedFileId) : null
  const targetFile = file || DriveApp.createFile(SPREADSHEET_INSPECTOR.driveSnapshotFileName, source, MimeType.PLAIN_TEXT)
  if (file) {
    targetFile.setContent(source)
  }
  scriptProperties.setProperty(fileIdKey, targetFile.getId())
  return targetFile
}

function getDriveFileByIdOrNull_(fileId) {
  try {
    return DriveApp.getFileById(fileId)
  } catch (error) {
    return null
  }
}

function shouldFallbackToDriveFile_(error) {
  const statusCode = Number(error?.statusCode)
  const message = String(error?.apiErrorMessage || error?.message || error)
  return statusCode === 403 &&
    /(has not been used in project .* before or it is disabled|Apps Script API .* disabled)/i.test(message)
}

function getAppsScriptApiAccessToken_() {
  try {
    const accessToken = ScriptApp.getOAuthToken()
    if (!accessToken) {
      throw new Error("空のアクセストークンです")
    }
    return accessToken
  } catch (error) {
    throw new Error(`Apps Script API の利用権限が未承認です。初回実行時の承認ダイアログで許可したあと、呼び出し元の関数をもう一度実行してください: ${error.message}`)
  }
}

function extractAppsScriptApiErrorMessage_(text) {
  try {
    const parsed = JSON.parse(text)
    return parsed && parsed.error && parsed.error.message ? parsed.error.message : text
  } catch (error) {
    return text
  }
}

function normalizeAppsScriptApiMethod_(method) {
  const normalizedMethod = String(method).toLowerCase()
  if (!SPREADSHEET_INSPECTOR.supportedApiMethods.includes(normalizedMethod)) {
    throw new Error(`未対応の Apps Script API method です: ${sanitizeAppsScriptApiMethodForError_(method)}`)
  }
  return normalizedMethod
}

function sanitizeAppsScriptApiMethodForError_(method) {
  return String(method).replace(/[^\w-]/g, "?").slice(0, 40)
}


/**
 * スナップショットオブジェクトを .gs ソース文字列へ変換します。
 * @param {object} snapshot
 * @return {string}
 */
function getSpreadsheetSchemaSnapshotSourceFromObject_(snapshot) {
  const json = JSON.stringify(snapshot, null, 2)
  const jsonChunks = splitTextByLength_(json, SPREADSHEET_INSPECTOR.chunkSize)
  return [
    "const SPREADSHEET_SCHEMA_SNAPSHOT_JSON_PARTS = [",
    jsonChunks.map((chunk) => `  ${JSON.stringify(chunk)}`).join(",\n"),
    "]",
    "",
    "let SPREADSHEET_SCHEMA_SNAPSHOT = null",
    "",
    "function getStoredSpreadsheetSchemaSnapshot() {",
    "  if (SPREADSHEET_SCHEMA_SNAPSHOT === null) {",
    "    SPREADSHEET_SCHEMA_SNAPSHOT = JSON.parse(SPREADSHEET_SCHEMA_SNAPSHOT_JSON_PARTS.join(\"\"))",
    "  }",
    "  return SPREADSHEET_SCHEMA_SNAPSHOT",
    "}"
  ].join("\n") + "\n"
}
