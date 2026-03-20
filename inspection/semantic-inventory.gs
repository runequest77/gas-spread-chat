const SPREADSHEET_SEMANTIC_INVENTORY = {
  sampleRowCount: 5,
  sampleColumnCount: 12,
  formulaScanRowCount: 50,
  headerRowCount: 2,
  settingsRowCount: 100,
  maxValueLength: 120,
  driveJsonFileName: "spreadsheet-semantic-inventory.json",
  driveJsonFileIdPropertyKey: "spreadsheetSemanticInventoryJsonDriveFileId",
  driveMarkdownFileName: "spreadsheet-semantic-inventory.md",
  driveMarkdownFileIdPropertyKey: "spreadsheetSemanticInventoryMarkdownDriveFileId",
  driveConditionalFormatRulesFileName: "spreadsheet-conditional-format-rules.json",
  driveConditionalFormatRulesFileIdPropertyKey: "spreadsheetConditionalFormatRulesJsonDriveFileId",
  driveDataValidationRulesFileName: "spreadsheet-data-validation-rules.json",
  driveDataValidationRulesFileIdPropertyKey: "spreadsheetDataValidationRulesJsonDriveFileId"
}

/**
 * スプレッドシートの「意味」を議論しやすい棚卸し情報を JSON / Markdown で出力します。
 * 完全な復元用ダンプではなく、役割・設定・構造の把握に必要な最小限の情報へ絞ります。
 * @param {object=} options
 * @return {{json:string, markdown:string, inventory:object}}
 */
function exportSpreadsheetSemanticInventoryToLog(options) {
  const inventory = getSpreadsheetSemanticInventory(options)
  const json = JSON.stringify(inventory, null, 2)
  const markdown = getSpreadsheetSemanticInventoryMarkdownFromObject_(inventory)
  console.log(json)
  if ((options || {}).includeMarkdown !== false) {
    console.log(markdown)
  }
  return {
    json: json,
    markdown: markdown,
    inventory: inventory
  }
}

/**
 * semantic inventory を Drive 上のテキストファイルへ保存します。
 * JSON は必須で保存し、Markdown も既定で保存します。
 * @param {object=} options
 * @return {{jsonFile:object, markdownFile:(object|null), inventory:object}}
 */
function exportSpreadsheetSemanticInventoryToDriveFile(options) {
  const opts = options || {}
  let stage = "export_start"
  try {
    logSemanticInventoryExport_("log", "[semantic-inventory] Drive export started")
    stage = "inventory_extraction"
    const inventory = getSpreadsheetSemanticInventory(options)
    logSemanticInventoryExport_("log", `[semantic-inventory] Inventory extracted: sheets=${(inventory.sheets || []).length}`)
    stage = "json_serialization"
    const json = JSON.stringify(inventory, null, 2)
    logSemanticInventoryExport_("log", `[semantic-inventory] JSON serialized: length=${json.length}`)
    stage = "json_drive_write"
    const jsonFile = upsertSemanticInventoryDriveFile_(
      SPREADSHEET_SEMANTIC_INVENTORY.driveJsonFileName,
      SPREADSHEET_SEMANTIC_INVENTORY.driveJsonFileIdPropertyKey,
      json
    )
    logSemanticInventoryExport_("log", `[semantic-inventory] JSON Drive file saved: name=${SPREADSHEET_SEMANTIC_INVENTORY.driveJsonFileName}, id=${jsonFile.getId()}`)
    stage = "markdown_drive_write"
    let markdownFile = null
    let markdownContentLength = 0
    if (opts.includeMarkdown !== false) {
      stage = "markdown_render"
      const markdown = getSpreadsheetSemanticInventoryMarkdownFromObject_(inventory)
      markdownContentLength = markdown.length
      logSemanticInventoryExport_("log", `[semantic-inventory] Markdown rendered: includeMarkdown=true, length=${markdownContentLength}`)
      stage = "markdown_drive_write"
      markdownFile = upsertSemanticInventoryDriveFile_(
        SPREADSHEET_SEMANTIC_INVENTORY.driveMarkdownFileName,
        SPREADSHEET_SEMANTIC_INVENTORY.driveMarkdownFileIdPropertyKey,
        markdown
      )
    }
    if (markdownFile) {
      logSemanticInventoryExport_("log", `[semantic-inventory] Markdown Drive file saved: name=${SPREADSHEET_SEMANTIC_INVENTORY.driveMarkdownFileName}, id=${markdownFile.getId()}`)
    } else {
      logSemanticInventoryExport_("log", "[semantic-inventory] Markdown Drive write skipped")
    }
    logSemanticInventoryExport_("log", "[semantic-inventory] Drive export completed")
    return {
      jsonFile: buildSemanticInventoryDriveFileResult_(jsonFile, json.length),
      markdownFile: markdownFile ? buildSemanticInventoryDriveFileResult_(markdownFile, markdownContentLength) : null,
      inventory: inventory
    }
  } catch (error) {
    const errorMessage = error && error.message ? error.message : String(error)
    logSemanticInventoryExport_("error", `[semantic-inventory] Drive export failed at stage=${stage}: ${errorMessage}`)
    throw error
  }
}

function getSpreadsheetSemanticInventoryJson(options) {
  return JSON.stringify(getSpreadsheetSemanticInventory(options), null, 2)
}

function getSpreadsheetSemanticInventoryMarkdown(options) {
  return getSpreadsheetSemanticInventoryMarkdownFromObject_(getSpreadsheetSemanticInventory(options))
}

/**
 * スプレッドシートの全シートから条件付き書式のコレクションを JSON 文字列として返します。
 * @return {string} JSON 文字列
 */
function getSpreadsheetConditionalFormatRulesJson() {
  return JSON.stringify(getSpreadsheetConditionalFormatRulesObject_(), null, 2)
}

/**
 * JSON 文字列を受け取り、スプレッドシートの各シートへ条件付き書式を適用します。
 * @param {string} json getSpreadsheetConditionalFormatRulesJson が返す JSON 文字列
 */
function applySpreadsheetConditionalFormatRulesFromJson(json) {
  const data = typeof json === "string" ? JSON.parse(json) : json
  const ss = getSpreadsheetForSemanticInventory_()
  const sheets = data && data.sheets
  if (!Array.isArray(sheets)) {
    return
  }
  sheets.forEach(function (entry) {
    if (!entry || !entry.sheetName || !Array.isArray(entry.conditionalFormatRules)) {
      return
    }
    const sheet = ss.getSheetByName(entry.sheetName)
    if (!sheet) {
      return
    }
    applySpreadChatConditionalFormatRules_(sheet, entry.conditionalFormatRules)
  })
}

/**
 * スプレッドシートの全シートからデータ入力規則のコレクションを JSON 文字列として返します。
 * @return {string} JSON 文字列
 */
function getSpreadsheetDataValidationRulesJson() {
  return JSON.stringify(getSpreadsheetDataValidationRulesObject_(), null, 2)
}

/**
 * JSON 文字列を受け取り、スプレッドシートの各シートへデータ入力規則を適用します。
 * @param {string} json getSpreadsheetDataValidationRulesJson が返す JSON 文字列
 */
function applySpreadsheetDataValidationRulesFromJson(json) {
  const data = typeof json === "string" ? JSON.parse(json) : json
  const ss = getSpreadsheetForSemanticInventory_()
  const sheets = data && data.sheets
  if (!Array.isArray(sheets)) {
    return
  }
  sheets.forEach(function (entry) {
    if (!entry || !entry.sheetName || !Array.isArray(entry.dataValidationRules)) {
      return
    }
    const sheet = ss.getSheetByName(entry.sheetName)
    if (!sheet) {
      return
    }
    applySpreadChatDataValidationRules_(sheet, entry.dataValidationRules)
  })
}

/**
 * 条件付き書式とデータ入力規則の両方を JSON で Google Drive に出力します。
 * getSpreadsheetConditionalFormatRulesJson と getSpreadsheetDataValidationRulesJson をラップします。
 * @param {object=} options
 * @return {{conditionalFormatRulesFile:object, dataValidationRulesFile:object}}
 */
function exportSpreadsheetRulesToDriveFile(options) {
  let stage = "export_start"
  try {
    logSemanticInventoryExport_("log", "[rules-export] Drive export started")
    stage = "conditional_format_extraction"
    const cfObject = getSpreadsheetConditionalFormatRulesObject_()
    const cfJson = JSON.stringify(cfObject, null, 2)
    logSemanticInventoryExport_("log", "[rules-export] Conditional format rules extracted: sheets=" + (cfObject.sheets || []).length)
    stage = "conditional_format_drive_write"
    const cfFile = upsertSemanticInventoryDriveFile_(
      SPREADSHEET_SEMANTIC_INVENTORY.driveConditionalFormatRulesFileName,
      SPREADSHEET_SEMANTIC_INVENTORY.driveConditionalFormatRulesFileIdPropertyKey,
      cfJson
    )
    logSemanticInventoryExport_("log", "[rules-export] Conditional format rules saved: name=" + SPREADSHEET_SEMANTIC_INVENTORY.driveConditionalFormatRulesFileName + ", id=" + cfFile.getId())
    stage = "data_validation_extraction"
    const dvObject = getSpreadsheetDataValidationRulesObject_()
    const dvJson = JSON.stringify(dvObject, null, 2)
    logSemanticInventoryExport_("log", "[rules-export] Data validation rules extracted: sheets=" + (dvObject.sheets || []).length)
    stage = "data_validation_drive_write"
    const dvFile = upsertSemanticInventoryDriveFile_(
      SPREADSHEET_SEMANTIC_INVENTORY.driveDataValidationRulesFileName,
      SPREADSHEET_SEMANTIC_INVENTORY.driveDataValidationRulesFileIdPropertyKey,
      dvJson
    )
    logSemanticInventoryExport_("log", "[rules-export] Data validation rules saved: name=" + SPREADSHEET_SEMANTIC_INVENTORY.driveDataValidationRulesFileName + ", id=" + dvFile.getId())
    logSemanticInventoryExport_("log", "[rules-export] Drive export completed")
    return {
      conditionalFormatRulesFile: buildSemanticInventoryDriveFileResult_(cfFile, cfJson.length),
      dataValidationRulesFile: buildSemanticInventoryDriveFileResult_(dvFile, dvJson.length)
    }
  } catch (error) {
    const errorMessage = error && error.message ? error.message : String(error)
    logSemanticInventoryExport_("error", "[rules-export] Drive export failed at stage=" + stage + ": " + errorMessage)
    throw error
  }
}

function getSpreadsheetConditionalFormatRulesObject_() {
  const ss = getSpreadsheetForSemanticInventory_()
  const sheets = ss.getSheets()
  return {
    exportedAt: new Date().toISOString(),
    spreadsheetId: typeof ss.getId === "function" ? ss.getId() : "",
    spreadsheetName: typeof ss.getName === "function" ? ss.getName() : "",
    sheets: sheets.map(function (sheet) {
      return {
        sheetName: sheet.getName(),
        conditionalFormatRules: summarizeConditionalFormatRules_(sheet)
      }
    })
  }
}

function getSpreadsheetDataValidationRulesObject_() {
  const ss = getSpreadsheetForSemanticInventory_()
  const sheets = ss.getSheets()
  return {
    exportedAt: new Date().toISOString(),
    spreadsheetId: typeof ss.getId === "function" ? ss.getId() : "",
    spreadsheetName: typeof ss.getName === "function" ? ss.getName() : "",
    sheets: sheets.map(function (sheet) {
      return {
        sheetName: sheet.getName(),
        dataValidationRules: summarizeDataValidationRules_(sheet)
      }
    })
  }
}

/**
 * スプレッドシート全体の semantic inventory を構築します。
 * workbook レベルの棚卸しと、各シートの役割推定・設定候補・ runtime 候補をまとめます。
 * @param {object=} options
 * @return {object}
 */
function getSpreadsheetSemanticInventory(options) {
  const normalizedOptions = normalizeSemanticInventoryOptions_(options)
  const ss = getSpreadsheetForSemanticInventory_()
  const namedRanges = getNamedRangesForSemanticInventory_(ss)
  const scriptProperties = getScriptPropertiesForSemanticInventory_()
  const sheets = ss.getSheets()
  const sheetInventories = sheets.map(sheet => describeSheetSemanticInventory_(sheet, namedRanges, normalizedOptions))
  const settingSheetName = getSemanticInventorySettingSheetName_()
  const settingSheet = typeof ss.getSheetByName === "function" ? ss.getSheetByName(settingSheetName) : null
  const settingSheetEntries = settingSheet ? getSettingSheetEntries_(settingSheet, normalizedOptions) : []
  const scriptPropertyEntries = Object.keys(scriptProperties).sort().map(key => buildScriptPropertySemanticEntry_(key, scriptProperties[key], normalizedOptions))

  return {
    exportedAt: new Date().toISOString(),
    spreadsheetId: typeof ss.getId === "function" ? ss.getId() : "",
    spreadsheetName: typeof ss.getName === "function" ? ss.getName() : "",
    workbook: {
      sheetNames: sheets.map(sheet => sheet.getName()),
      namedRanges: namedRanges,
      scriptPropertyKeys: Object.keys(scriptProperties).sort(),
      installableTriggers: getInstallableTriggersForSemanticInventory_(),
      generatedFrom: "semantic_inventory"
    },
    settings: {
      settingSheetName: settingSheet ? settingSheet.getName() : null,
      settingSheetEntries: settingSheetEntries,
      scriptPropertyEntries: scriptPropertyEntries,
      entries: settingSheetEntries.concat(scriptPropertyEntries)
    },
    sheets: sheetInventories,
    options: normalizedOptions
  }
}

function normalizeSemanticInventoryOptions_(options) {
  const opts = options || {}
  return {
    sampleRowCount: normalizePositiveInteger_(opts.sampleRowCount, SPREADSHEET_SEMANTIC_INVENTORY.sampleRowCount),
    sampleColumnCount: normalizePositiveInteger_(opts.sampleColumnCount, SPREADSHEET_SEMANTIC_INVENTORY.sampleColumnCount),
    formulaScanRowCount: normalizePositiveInteger_(opts.formulaScanRowCount, SPREADSHEET_SEMANTIC_INVENTORY.formulaScanRowCount),
    headerRowCount: normalizePositiveInteger_(opts.headerRowCount, SPREADSHEET_SEMANTIC_INVENTORY.headerRowCount),
    settingsRowCount: normalizePositiveInteger_(opts.settingsRowCount, SPREADSHEET_SEMANTIC_INVENTORY.settingsRowCount),
    maxValueLength: normalizePositiveInteger_(opts.maxValueLength, SPREADSHEET_SEMANTIC_INVENTORY.maxValueLength),
    includeMarkdown: opts.includeMarkdown !== false,
    maskSensitiveValues: opts.maskSensitiveValues !== false
  }
}

function normalizePositiveInteger_(value, fallbackValue) {
  const normalized = Number(value)
  return Number.isFinite(normalized) && normalized > 0 ? Math.floor(normalized) : fallbackValue
}

function getSpreadsheetForSemanticInventory_() {
  if (typeof c_ss === "function") {
    return c_ss()
  }
  return SpreadsheetApp.getActiveSpreadsheet()
}

function getSemanticInventorySettingSheetName_() {
  return typeof SETTING_SHEET_NAME !== "undefined" ? SETTING_SHEET_NAME : "Setting"
}

function getSemanticInventoryMapPrefix_() {
  return typeof mapPrefix !== "undefined" ? mapPrefix : "__"
}

function getSemanticInventorySecretPrefix_() {
  return typeof secretPrefix !== "undefined" ? secretPrefix : "秘匿"
}

function getNamedRangesForSemanticInventory_(ss) {
  if (!ss || typeof ss.getNamedRanges !== "function") {
    return []
  }
  return ss.getNamedRanges().map(namedRange => {
    const range = namedRange.getRange()
    return {
      name: namedRange.getName(),
      sheetName: range.getSheet().getName(),
      a1Notation: range.getA1Notation()
    }
  })
}

function getScriptPropertiesForSemanticInventory_() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties()
    if (scriptProperties && typeof scriptProperties.getProperties === "function") {
      return scriptProperties.getProperties() || {}
    }
  } catch (error) {
  }
  return {}
}

function getInstallableTriggersForSemanticInventory_() {
  try {
    if (typeof ScriptApp === "undefined" || typeof ScriptApp.getProjectTriggers !== "function") {
      return []
    }
    return ScriptApp.getProjectTriggers().map(trigger => ({
      handlerFunction: getObjectMethodValue_(trigger, "getHandlerFunction"),
      eventType: stringifySemanticInventoryValue_(getObjectMethodValue_(trigger, "getEventType")),
      triggerSource: stringifySemanticInventoryValue_(getObjectMethodValue_(trigger, "getTriggerSource")),
      triggerSourceId: getObjectMethodValue_(trigger, "getTriggerSourceId"),
      uniqueId: getObjectMethodValue_(trigger, "getUniqueId")
    }))
  } catch (error) {
    return [{
      error: error.message
    }]
  }
}

function getObjectMethodValue_(target, methodName) {
  return target && typeof target[methodName] === "function" ? target[methodName]() : null
}

function stringifySemanticInventoryValue_(value) {
  if (value === null || value === undefined) {
    return null
  }
  return String(value)
}

function describeSheetSemanticInventory_(sheet, namedRanges, options) {
  const sheetName = sheet.getName()
  const lastRow = typeof sheet.getLastRow === "function" ? sheet.getLastRow() : 0
  const lastColumn = typeof sheet.getLastColumn === "function" ? sheet.getLastColumn() : 0
  const conditionalFormatRules = summarizeConditionalFormatRules_(sheet)
  const dataValidationRules = summarizeDataValidationRules_(sheet)
  const headerRowCount = inferSemanticHeaderRowCount_(sheet, options.headerRowCount, lastRow)
  const sampledColumnCount = Math.min(lastColumn, options.sampleColumnCount)
  const sampleRowStart = Math.min(lastRow + 1, headerRowCount + 1)
  const sampleRowCount = Math.max(0, Math.min(options.sampleRowCount, lastRow - headerRowCount))
  const formulaScanRowCount = Math.min(lastRow, options.formulaScanRowCount)
  const headerValues = getDisplayValuesSlice_(sheet, 1, headerRowCount, sampledColumnCount)
  const sampleValues = sampleRowCount > 0 ? getDisplayValuesSlice_(sheet, sampleRowStart, sampleRowCount, sampledColumnCount) : []
  const formulaValues = formulaScanRowCount > 0 && sampledColumnCount > 0 ? getFormulaValuesSlice_(sheet, 1, formulaScanRowCount, sampledColumnCount) : []
  const backgroundValues = sampleRowCount > 0 ? getBackgroundValuesSlice_(sheet, sampleRowStart, sampleRowCount, sampledColumnCount) : []
  const specialMarkers = collectSemanticSpecialMarkers_(headerValues.concat(sampleValues), {
    rowOffset: 1,
    mapPrefix: getSemanticInventoryMapPrefix_(),
    secretPrefix: getSemanticInventorySecretPrefix_(),
    maxMarkers: 15
  })
  const classification = inferSheetSemanticClassification_({
    sheetName: sheetName,
    headerValues: headerValues,
    sampleValues: sampleValues,
    specialMarkers: specialMarkers
  })

  return {
    sheetName: sheetName,
    maxRows: typeof sheet.getMaxRows === "function" ? sheet.getMaxRows() : 0,
    maxColumns: typeof sheet.getMaxColumns === "function" ? sheet.getMaxColumns() : 0,
    lastRow: lastRow,
    lastColumn: lastColumn,
    frozenRows: typeof sheet.getFrozenRows === "function" ? sheet.getFrozenRows() : 0,
    frozenColumns: typeof sheet.getFrozenColumns === "function" ? sheet.getFrozenColumns() : 0,
    headerRowValues: headerValues,
    sampleRows: buildSemanticSampleRows_(sampleValues, sampleRowStart, options),
    nonEmptyColumnIndexes: getNonEmptyColumnIndexes_(headerValues.concat(sampleValues)),
    columnSummaries: buildSemanticColumnSummaries_(headerValues, sampleValues, formulaValues, options),
    notes: classification.notes,
    schemaCandidate: classification.schemaCandidate,
    runtimeCandidate: classification.runtimeCandidate,
    specialMarkers: specialMarkers,
    settingsCandidate: classification.settingsCandidate,
    formulaSummary: summarizeSemanticFormulaColumns_(formulaValues, headerValues),
    conditionalFormatRuleCount: conditionalFormatRules.length,
    conditionalFormatRules: conditionalFormatRules,
    conditionalFormatSummary: conditionalFormatRules,
    dataValidationRuleCount: dataValidationRules.length,
    dataValidationRules: dataValidationRules,
    protectionCount: getProtectionCount_(sheet),
    backgroundColorSummary: summarizeSemanticBackgrounds_(backgroundValues),
    columnWidthSummary: summarizeSemanticColumnWidths_(sheet, sampledColumnCount),
    namedRanges: namedRanges.filter(namedRange => namedRange.sheetName === sheetName),
    sampledRanges: {
      headers: headerValues.length > 0 ? buildSemanticSheetRangeA1_(sheetName, 1, headerValues.length, sampledColumnCount) : null,
      sampleRows: sampleRowCount > 0 ? buildSemanticSheetRangeA1_(sheetName, sampleRowStart, sampleRowCount, sampledColumnCount) : null,
      formulaScan: formulaScanRowCount > 0 ? buildSemanticSheetRangeA1_(sheetName, 1, formulaScanRowCount, sampledColumnCount) : null
    }
  }
}

function inferSemanticHeaderRowCount_(sheet, defaultHeaderRowCount, lastRow) {
  if (lastRow <= 0) {
    return 0
  }
  const frozenRows = typeof sheet.getFrozenRows === "function" ? sheet.getFrozenRows() : 0
  return Math.max(1, Math.min(lastRow, frozenRows || defaultHeaderRowCount))
}

function getDisplayValuesSlice_(sheet, rowStart, rowCount, columnCount) {
  if (rowCount <= 0 || columnCount <= 0) {
    return []
  }
  return sheet.getRange(rowStart, 1, rowCount, columnCount).getDisplayValues()
}

function getFormulaValuesSlice_(sheet, rowStart, rowCount, columnCount) {
  if (rowCount <= 0 || columnCount <= 0) {
    return []
  }
  return sheet.getRange(rowStart, 1, rowCount, columnCount).getFormulas()
}

function getBackgroundValuesSlice_(sheet, rowStart, rowCount, columnCount) {
  if (rowCount <= 0 || columnCount <= 0) {
    return []
  }
  const range = sheet.getRange(rowStart, 1, rowCount, columnCount)
  return typeof range.getBackgrounds === "function" ? range.getBackgrounds() : []
}

function buildSemanticSampleRows_(sampleValues, rowStart, options) {
  return sampleValues.map((row, index) => ({
    row: rowStart + index,
    values: row.map(value => compactSemanticCellValue_(value, options.maxValueLength))
  }))
}

function compactSemanticCellValue_(value, maxValueLength) {
  const normalized = String(value || "")
  if (normalized.length <= maxValueLength) {
    return normalized
  }
  return `${normalized.slice(0, Math.max(0, maxValueLength - 1))}…`
}

function getNonEmptyColumnIndexes_(rows) {
  const nonEmptyColumns = []
  const maxColumnCount = rows.reduce((maxCount, row) => Math.max(maxCount, row.length), 0)
  for (let columnIndex = 0; columnIndex < maxColumnCount; columnIndex++) {
    const hasValue = rows.some(row => String(row[columnIndex] || "") !== "")
    if (hasValue) {
      nonEmptyColumns.push(columnIndex + 1)
    }
  }
  return nonEmptyColumns
}

function buildSemanticColumnSummaries_(headerValues, sampleValues, formulaValues, options) {
  const rows = headerValues.concat(sampleValues)
  const maxColumnCount = rows.reduce((maxCount, row) => Math.max(maxCount, row.length), 0)
  const headerRow = headerValues[0] || []
  const summaries = []

  for (let columnIndex = 0; columnIndex < maxColumnCount; columnIndex++) {
    const sampleColumnValues = sampleValues
      .map(row => compactSemanticCellValue_(row[columnIndex], options.maxValueLength))
      .filter(value => value !== "")
    const formulaCount = formulaValues.reduce((count, row) => count + (String(row[columnIndex] || "") !== "" ? 1 : 0), 0)
    if (sampleColumnValues.length === 0 && String(headerRow[columnIndex] || "") === "" && formulaCount === 0) {
      continue
    }
    summaries.push({
      columnIndex: columnIndex + 1,
      columnLetter: semanticInventoryColumnNumberToLetter_(columnIndex + 1),
      header: compactSemanticCellValue_(headerRow[columnIndex], options.maxValueLength),
      nonEmptySampleCount: sampleColumnValues.length,
      formulaCount: formulaCount,
      sampleValues: sampleColumnValues.slice(0, 3)
    })
  }

  return summaries
}

function summarizeSemanticFormulaColumns_(formulaValues, headerValues) {
  const headerRow = headerValues[0] || []
  if (formulaValues.length === 0) {
    return []
  }
  const maxColumnCount = formulaValues.reduce((maxCount, row) => Math.max(maxCount, row.length), 0)
  const summaries = []
  for (let columnIndex = 0; columnIndex < maxColumnCount; columnIndex++) {
    const formulas = formulaValues.map(row => String(row[columnIndex] || "")).filter(Boolean)
    if (formulas.length === 0) {
      continue
    }
    summaries.push({
      columnIndex: columnIndex + 1,
      columnLetter: semanticInventoryColumnNumberToLetter_(columnIndex + 1),
      header: String(headerRow[columnIndex] || ""),
      sampledFormulaCount: formulas.length,
      representativeFormula: formulas[0]
    })
  }
  return summaries
}

function summarizeSemanticColumnWidths_(sheet, columnCount) {
  const widthMap = {}
  for (let column = 1; column <= columnCount; column++) {
    const width = typeof sheet.getColumnWidth === "function" ? sheet.getColumnWidth(column) : null
    const key = String(width)
    if (!widthMap[key]) {
      widthMap[key] = {
        width: width,
        columns: []
      }
    }
    widthMap[key].columns.push(semanticInventoryColumnNumberToLetter_(column))
  }
  return Object.keys(widthMap)
    .map(key => widthMap[key])
    .sort((a, b) => Number(a.width) - Number(b.width))
}

function getConditionalFormatRuleCount_(sheet) {
  try {
    return typeof sheet.getConditionalFormatRules === "function" ? sheet.getConditionalFormatRules().length : 0
  } catch (error) {
    return 0
  }
}

function summarizeConditionalFormatRules_(sheet, options) {
  try {
    if (typeof sheet.getConditionalFormatRules !== "function") {
      return []
    }
    return (sheet.getConditionalFormatRules() || []).map((rule, index) => summarizeConditionalFormatRule_(rule, index))
  } catch (error) {
    return []
  }
}

function summarizeConditionalFormatRule_(rule, index) {
  const ranges = typeof rule.getRanges === "function" ? (rule.getRanges() || []) : []
  const booleanCondition = typeof rule.getBooleanCondition === "function" ? rule.getBooleanCondition() : null
  const gradientCondition = typeof rule.getGradientCondition === "function" ? rule.getGradientCondition() : null
  const format = summarizeConditionalFormatStyle_(rule)
  const summary = {
    index: index + 1,
    ranges: ranges.map(range => buildSemanticRangeLabel_(range)),
    ruleType: gradientCondition ? "gradient" : (booleanCondition ? "boolean" : "unknown"),
    criteriaType: booleanCondition && typeof booleanCondition.getCriteriaType === "function"
      ? stringifySemanticInventoryValue_(booleanCondition.getCriteriaType())
      : null,
    criteriaValues: booleanCondition && typeof booleanCondition.getCriteriaValues === "function"
      ? serializeSemanticCriteriaValues_(booleanCondition.getCriteriaValues())
      : null
  }
  if (format) {
    summary.format = format
  }
  return summary
}

function summarizeConditionalFormatStyle_(rule) {
  const format = {}
  appendSemanticConditionalFormatStyleValue_(format, "backgroundColor", normalizeSemanticVisibleColor_(getObjectMethodValue_(rule, "getVisibleBackgroundColor")))
  appendSemanticConditionalFormatStyleValue_(format, "fontColor", normalizeSemanticVisibleColor_(getObjectMethodValue_(rule, "getVisibleFontColor")))
  appendSemanticConditionalFormatStyleValue_(format, "bold", getObjectMethodValue_(rule, "getVisibleBold"))
  appendSemanticConditionalFormatStyleValue_(format, "italic", getObjectMethodValue_(rule, "getVisibleItalic"))
  appendSemanticConditionalFormatStyleValue_(format, "underline", getObjectMethodValue_(rule, "getVisibleUnderline"))
  appendSemanticConditionalFormatStyleValue_(format, "strikethrough", getObjectMethodValue_(rule, "getVisibleStrikethrough"))
  appendSemanticConditionalFormatStyleValue_(format, "fontFamily", stringifySemanticInventoryValue_(getObjectMethodValue_(rule, "getVisibleFontFamily")))
  appendSemanticConditionalFormatStyleValue_(format, "fontSize", getObjectMethodValue_(rule, "getVisibleFontSize"))
  return Object.keys(format).length > 0 ? format : null
}

function appendSemanticConditionalFormatStyleValue_(target, key, value) {
  if (value === null || value === undefined || value === "") {
    return
  }
  target[key] = value
}

function normalizeSemanticVisibleColor_(color) {
  const normalized = String(color || "").trim().toLowerCase()
  return normalized || null
}

function summarizeDataValidationRules_(sheet) {
  try {
    if (!sheet || typeof sheet.getRange !== "function" || typeof sheet.getMaxRows !== "function" || typeof sheet.getMaxColumns !== "function") {
      return []
    }
    const rowCount = getSemanticDataValidationScanCount_(sheet, "getLastRow", "getMaxRows")
    const columnCount = getSemanticDataValidationScanCount_(sheet, "getLastColumn", "getMaxColumns")
    if (rowCount <= 0 || columnCount <= 0) {
      return []
    }
    const range = sheet.getRange(1, 1, rowCount, columnCount)
    if (!range || typeof range.getDataValidations !== "function") {
      return []
    }
    return summarizeDataValidationMatrix_(sheet.getName(), range.getDataValidations() || [])
  } catch (error) {
    return []
  }
}

function getSemanticDataValidationScanCount_(sheet, lastMethodName, maxMethodName) {
  const lastCount = sheet && typeof sheet[lastMethodName] === "function" ? Number(sheet[lastMethodName]()) : 0
  if (Number.isFinite(lastCount) && lastCount > 0) {
    return Math.floor(lastCount)
  }
  const maxCount = sheet && typeof sheet[maxMethodName] === "function" ? Number(sheet[maxMethodName]()) : 0
  return Number.isFinite(maxCount) && maxCount > 0 ? Math.floor(maxCount) : 0
}

function summarizeDataValidationMatrix_(sheetName, matrix) {
  const rules = []
  const visited = matrix.map(row => new Array((row || []).length).fill(false))
  for (let rowIndex = 0; rowIndex < matrix.length; rowIndex++) {
    const row = matrix[rowIndex] || []
    for (let columnIndex = 0; columnIndex < row.length; columnIndex++) {
      if (visited[rowIndex][columnIndex]) {
        continue
      }
      const validation = row[columnIndex]
      if (!validation) {
        visited[rowIndex][columnIndex] = true
        continue
      }
      const summary = summarizeDataValidationRule_(validation)
      const signature = JSON.stringify(summary)
      const width = getSemanticValidationRegionWidth_(matrix, visited, rowIndex, columnIndex, signature)
      const height = getSemanticValidationRegionHeight_(matrix, visited, rowIndex, columnIndex, width, signature)
      markSemanticValidationRegionVisited_(visited, rowIndex, columnIndex, height, width)
      rules.push(Object.assign({
        ranges: [buildSemanticSheetRectRangeA1_(sheetName, rowIndex + 1, columnIndex + 1, height, width)]
      }, summary))
    }
  }
  return rules
}

function summarizeDataValidationRule_(validation) {
  return {
    criteriaType: typeof validation.getCriteriaType === "function"
      ? stringifySemanticInventoryValue_(validation.getCriteriaType())
      : null,
    criteriaValues: typeof validation.getCriteriaValues === "function"
      ? serializeSemanticCriteriaValues_(validation.getCriteriaValues())
      : null,
    allowInvalid: typeof validation.getAllowInvalid === "function" ? validation.getAllowInvalid() : null,
    helpText: typeof validation.getHelpText === "function" ? validation.getHelpText() : null
  }
}

function getSemanticValidationRegionWidth_(matrix, visited, rowIndex, columnIndex, signature) {
  const row = matrix[rowIndex] || []
  let width = 0
  while (columnIndex + width < row.length) {
    if (visited[rowIndex][columnIndex + width]) {
      break
    }
    const validation = row[columnIndex + width]
    if (!validation || JSON.stringify(summarizeDataValidationRule_(validation)) !== signature) {
      break
    }
    width++
  }
  return width
}

function getSemanticValidationRegionHeight_(matrix, visited, rowIndex, columnIndex, width, signature) {
  let height = 0
  while (rowIndex + height < matrix.length) {
    const row = matrix[rowIndex + height] || []
    let matches = true
    for (let offset = 0; offset < width; offset++) {
      if (visited[rowIndex + height][columnIndex + offset]) {
        matches = false
        break
      }
      const validation = row[columnIndex + offset]
      if (!validation || JSON.stringify(summarizeDataValidationRule_(validation)) !== signature) {
        matches = false
        break
      }
    }
    if (!matches) {
      break
    }
    height++
  }
  return height
}

function markSemanticValidationRegionVisited_(visited, rowIndex, columnIndex, height, width) {
  for (let rowOffset = 0; rowOffset < height; rowOffset++) {
    const visitedRow = visited[rowIndex + rowOffset]
    for (let columnOffset = 0; columnOffset < width; columnOffset++) {
      visitedRow[columnIndex + columnOffset] = true
    }
  }
}

function buildSemanticRangeLabel_(range) {
  try {
    const sheet = typeof range.getSheet === "function" ? range.getSheet() : null
    const sheetName = sheet && typeof sheet.getName === "function" ? sheet.getName() : ""
    const a1Notation = typeof range.getA1Notation === "function" ? range.getA1Notation() : ""
    return sheetName ? `${sheetName}!${a1Notation}` : a1Notation
  } catch (error) {
    return ""
  }
}

function buildSemanticSheetRectRangeA1_(sheetName, rowStart, columnStart, rowCount, columnCount) {
  const escapedSheetName = String(sheetName).replace(/'/g, "''")
  const rowEnd = rowStart + rowCount - 1
  const columnEnd = columnStart + columnCount - 1
  return `'${escapedSheetName}'!${semanticInventoryColumnNumberToLetter_(columnStart)}${rowStart}:${semanticInventoryColumnNumberToLetter_(columnEnd)}${rowEnd}`
}

function serializeSemanticInventoryValue_(value) {
  if (value === null || value === undefined) {
    return null
  }
  if (Array.isArray(value)) {
    return value.map(entry => serializeSemanticInventoryValue_(entry))
  }
  if (Object.prototype.toString.call(value) === "[object Date]" && typeof value.toISOString === "function") {
    return value.toISOString()
  }
  if (typeof value.getA1Notation === "function") {
    return buildSemanticRangeLabel_(value)
  }
  if (typeof value === "object") {
    const keys = Object.keys(value).sort()
    if (keys.length === 0) {
      return stringifySemanticInventoryValue_(value)
    }
    const result = {}
    keys.forEach(key => {
      result[key] = serializeSemanticInventoryValue_(value[key])
    })
    return result
  }
  return value
}

function serializeSemanticCriteriaValues_(values) {
  if (!Array.isArray(values)) {
    return serializeSemanticInventoryValue_(values)
  }
  return values.map(value => serializeSemanticCriteriaValue_(value))
}

function serializeSemanticCriteriaValue_(value) {
  if (value && typeof value.getUserEnteredValue === "function") {
    return serializeSemanticInventoryValue_(value.getUserEnteredValue())
  }
  if (value && typeof value.getRelativeDate === "function") {
    return serializeSemanticInventoryValue_(value.getRelativeDate())
  }
  return serializeSemanticInventoryValue_(value)
}

function summarizeSemanticBackgrounds_(backgroundValues) {
  const counts = {}
  backgroundValues.forEach(row => {
    row.forEach(color => {
      const normalized = normalizeSemanticBackgroundColor_(color)
      if (!normalized) {
        return
      }
      counts[normalized] = (counts[normalized] || 0) + 1
    })
  })
  return Object.keys(counts)
    .map(color => ({ color: color, sampleCellCount: counts[color] }))
    .sort((a, b) => b.sampleCellCount - a.sampleCellCount || a.color.localeCompare(b.color))
}

function normalizeSemanticBackgroundColor_(color) {
  const normalized = String(color || "").trim().toLowerCase()
  if (normalized === "" || normalized === "#ffffff" || normalized === "white" || normalized === "null") {
    return ""
  }
  return normalized
}

function getProtectionCount_(sheet) {
  try {
    if (typeof sheet.getProtections !== "function" || typeof SpreadsheetApp === "undefined" || !SpreadsheetApp.ProtectionType) {
      return 0
    }
    const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || []
    const sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET) || []
    return rangeProtections.length + sheetProtections.length
  } catch (error) {
    return 0
  }
}

function collectSemanticSpecialMarkers_(rows, options) {
  const opts = options || {}
  const mapPrefix = String(opts.mapPrefix || "__")
  const secretPrefix = String(opts.secretPrefix || "秘匿")
  const markers = []
  const seenKeys = {}

  rows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      const text = String(value || "")
      if (text === "") {
        return
      }
      if (text === "__Map") {
        pushSemanticMarker_(markers, seenKeys, {
          type: "map_row_marker",
          value: text,
          row: (opts.rowOffset || 1) + rowIndex,
          column: columnIndex + 1
        }, opts.maxMarkers)
      } else if (text.indexOf(mapPrefix) === 0) {
        pushSemanticMarker_(markers, seenKeys, {
          type: row[0] === text ? "row_prefix" : "double_underscore_value",
          value: text,
          row: (opts.rowOffset || 1) + rowIndex,
          column: columnIndex + 1
        }, opts.maxMarkers)
      }
      if (text.indexOf(secretPrefix) === 0) {
        pushSemanticMarker_(markers, seenKeys, {
          type: "secret_prefix",
          value: text,
          row: (opts.rowOffset || 1) + rowIndex,
          column: columnIndex + 1
        }, opts.maxMarkers)
      }
      if (/^=IMAGE\(/i.test(text) || /^IMAGE\(/i.test(text)) {
        pushSemanticMarker_(markers, seenKeys, {
          type: "image_formula_like",
          value: compactSemanticCellValue_(text, 40),
          row: (opts.rowOffset || 1) + rowIndex,
          column: columnIndex + 1
        }, opts.maxMarkers)
      }
    })
  })

  return markers
}

function pushSemanticMarker_(markers, seenKeys, marker, maxMarkers) {
  const key = JSON.stringify(marker)
  if (seenKeys[key] || (maxMarkers && markers.length >= maxMarkers)) {
    return
  }
  seenKeys[key] = true
  markers.push(marker)
}

function inferSheetSemanticClassification_(context) {
  const sheetName = String(context.sheetName || "")
  const normalizedSheetName = sheetName.toLowerCase()
  const headerText = flattenSemanticRows_(context.headerValues).join(" ").toLowerCase()
  const sampleText = flattenSemanticRows_(context.sampleValues).join(" ").toLowerCase()
  const specialMarkerTypes = (context.specialMarkers || []).map(marker => marker.type)
  const settingSheetName = getSemanticInventorySettingSheetName_().toLowerCase()
  const isSettingLikeName = normalizedSheetName === settingSheetName || /(setting|settings|config|master|schema|palette|パレット|定義|設定|一覧|chart|charts)/i.test(sheetName)
  const isRuntimeLikeName = /(chat|log|history|current|situation|unit|map|session|cache|state|runtime|チャット|ログ|履歴|現在)/i.test(sheetName)
  const looksKeyValue = /^(key|name|setting|config|項目|設定|value|値)$/i.test(String((context.headerValues[0] || [])[0] || "")) ||
    (/\bdefault\b|\bconfig\b|設定|しきい値|sheet|column/.test(headerText) && headerText.indexOf("value") >= 0)
  const hasMapMarkers = specialMarkerTypes.includes("map_row_marker") || specialMarkerTypes.includes("row_prefix")
  const hasRuntimeMarkers = hasMapMarkers || sampleText.indexOf(getSemanticInventorySecretPrefix_().toLowerCase()) >= 0

  if (isSettingLikeName || looksKeyValue) {
    return {
      schemaCandidate: true,
      runtimeCandidate: false,
      settingsCandidate: true,
      notes: "Likely a settings / schema sheet because the name or sampled headers look like key-value configuration."
    }
  }

  if (isRuntimeLikeName || hasRuntimeMarkers) {
    return {
      schemaCandidate: false,
      runtimeCandidate: true,
      settingsCandidate: false,
      notes: "Likely a runtime / operational sheet because the name or sampled markers look like chat logs, map state, or live records."
    }
  }

  return {
    schemaCandidate: true,
    runtimeCandidate: true,
    settingsCandidate: false,
    notes: "Mixed or unclear role. Treat as a sheet that may contain both schema hints and runtime data until clarified."
  }
}

function flattenSemanticRows_(rows) {
  return (rows || []).reduce((allValues, row) => allValues.concat((row || []).map(value => String(value || ""))), [])
}

function getSettingSheetEntries_(sheet, options) {
  const lastRow = typeof sheet.getLastRow === "function" ? sheet.getLastRow() : 0
  const rowCount = Math.min(lastRow, options.settingsRowCount)
  if (rowCount <= 0) {
    return []
  }
  const lastColumn = typeof sheet.getLastColumn === "function" ? sheet.getLastColumn() : 0
  const columnCount = Math.min(Math.max(lastColumn, 2), 3)
  const values = sheet.getRange(1, 1, rowCount, columnCount).getDisplayValues()
  return extractSemanticSettingsEntries_(values, options)
}

function extractSemanticSettingsEntries_(values, options) {
  const entries = []
  values.forEach((row, index) => {
    const key = compactSemanticCellValue_(row[0], options.maxValueLength)
    if (key === "" || isSemanticSettingsHeaderRow_(row)) {
      return
    }
    const value = compactSemanticCellValue_((row[1] || row.slice(1).join(" | ")), options.maxValueLength)
    entries.push(buildSemanticSettingsEntry_(key, value, "sheet", index + 1, options))
  })
  return entries
}

function isSemanticSettingsHeaderRow_(row) {
  const key = String((row || [])[0] || "").toLowerCase()
  const value = String((row || [])[1] || "").toLowerCase()
  return /^(key|name|setting|config|項目|設定)$/.test(key) && /^(value|値|内容)$/.test(value)
}

function buildSemanticSettingsEntry_(key, value, source, row, options) {
  const normalizedKey = String(key || "")
  const lowerKey = normalizedKey.toLowerCase()
  const runtimeCandidate = isSemanticSensitiveKey_(normalizedKey) || /current|last|session|cache|token|state|active|runtime|user|id/.test(lowerKey)
  const schemaCandidate = !runtimeCandidate || /default|sheet|column|prefix|name|config|setting|bot|map|chart|table/.test(lowerKey)
  return {
    key: normalizedKey,
    value: maskSemanticInventoryValue_(normalizedKey, value, options),
    source: source,
    row: row,
    runtimeCandidate: runtimeCandidate,
    schemaCandidate: schemaCandidate
  }
}

function buildScriptPropertySemanticEntry_(key, value, options) {
  return Object.assign(
    buildSemanticSettingsEntry_(key, value, "script_properties", null, options),
    {
      runtimeCandidate: true,
      schemaCandidate: /default|sheet|column|prefix|name|config|setting|bot/.test(String(key || "").toLowerCase())
    }
  )
}

function maskSemanticInventoryValue_(key, value, options) {
  const normalizedValue = compactSemanticCellValue_(value, options.maxValueLength)
  if (normalizedValue === "") {
    return ""
  }
  if (options.maskSensitiveValues !== false && isSemanticSensitiveKey_(key)) {
    return `[masked:${normalizedValue.length}]`
  }
  return normalizedValue
}

function isSemanticSensitiveKey_(key) {
  return /(token|secret|password|passwd|webhook|bearer|oauth|credential|private|api.?key|client.?secret)/i.test(String(key || ""))
}

function buildSemanticSheetRangeA1_(sheetName, rowStart, rowCount, columnCount) {
  const escapedSheetName = String(sheetName).replace(/'/g, "''")
  const rowEnd = rowStart + rowCount - 1
  return `'${escapedSheetName}'!A${rowStart}:${semanticInventoryColumnNumberToLetter_(columnCount)}${rowEnd}`
}

function semanticInventoryColumnNumberToLetter_(columnNumber) {
  let n = columnNumber
  let result = ""
  while (n > 0) {
    const mod = (n - 1) % 26
    result = String.fromCharCode(65 + mod) + result
    n = Math.floor((n - 1) / 26)
  }
  return result
}

function upsertSemanticInventoryDriveFile_(fileName, propertyKey, content) {
  const scriptProperties = PropertiesService.getScriptProperties()
  const storedFileId = scriptProperties.getProperty(propertyKey)
  const file = storedFileId ? getSemanticInventoryDriveFileByIdOrNull_(storedFileId) : null
  const targetFile = file || DriveApp.createFile(fileName, content, MimeType.PLAIN_TEXT)
  if (file) {
    targetFile.setContent(content)
  }
  scriptProperties.setProperty(propertyKey, targetFile.getId())
  return targetFile
}

function logSemanticInventoryExport_(level, message) {
  if (typeof console === "undefined") {
    return
  }
  const logger = typeof console[level] === "function"
    ? console[level]
    : (typeof console.log === "function" ? console.log : null)
  if (logger) {
    logger.call(console, message)
  }
}

function getSemanticInventoryDriveFileByIdOrNull_(fileId) {
  try {
    return DriveApp.getFileById(fileId)
  } catch (error) {
    return null
  }
}

function buildSemanticInventoryDriveFileResult_(file, contentLength) {
  return {
    fileId: file.getId(),
    fileName: file.getName(),
    url: file.getUrl(),
    contentLength: contentLength
  }
}

function getSpreadsheetSemanticInventoryMarkdownFromObject_(inventory) {
  const workbook = inventory.workbook || {}
  const settings = inventory.settings || {}
  const lines = [
    "# Spreadsheet Semantic Inventory",
    "",
    `- Exported At: ${inventory.exportedAt}`,
    `- Spreadsheet ID: ${inventory.spreadsheetId}`,
    `- Spreadsheet Name: ${inventory.spreadsheetName}`,
    `- Sheets: ${(workbook.sheetNames || []).join(", ")}`,
    `- Script Property Keys: ${(workbook.scriptPropertyKeys || []).join(", ") || "(none)"}`,
    `- Installable Triggers: ${(workbook.installableTriggers || []).length}`,
    "",
    "## Named Ranges",
    ""
  ]

  const namedRanges = workbook.namedRanges || []
  if (namedRanges.length === 0) {
    lines.push("- (none)")
  } else {
    namedRanges.forEach(namedRange => {
      lines.push(`- ${namedRange.name}: ${namedRange.sheetName}!${namedRange.a1Notation}`)
    })
  }

  lines.push("", "## Settings / Properties", "")
  const entries = (settings.entries || []).slice(0, 30)
  if (entries.length === 0) {
    lines.push("- (none)")
  } else {
    entries.forEach(entry => {
      lines.push(`- [${entry.source}] ${entry.key} = ${entry.value}`)
    })
    if ((settings.entries || []).length > entries.length) {
      lines.push(`- ... ${settings.entries.length - entries.length} more entries`)
    }
  }

  lines.push("", "## Sheets", "")
  ;(inventory.sheets || []).forEach(sheet => {
    lines.push(`### ${sheet.sheetName}`)
    lines.push("")
    lines.push(`- Size: lastRow=${sheet.lastRow}, lastColumn=${sheet.lastColumn}, maxRows=${sheet.maxRows}, maxColumns=${sheet.maxColumns}`)
    lines.push(`- Frozen: rows=${sheet.frozenRows}, columns=${sheet.frozenColumns}`)
    lines.push(`- Classification: schemaCandidate=${sheet.schemaCandidate}, runtimeCandidate=${sheet.runtimeCandidate}, settingsCandidate=${sheet.settingsCandidate}`)
    lines.push(`- Notes: ${sheet.notes}`)
    lines.push(`- Non-empty columns: ${(sheet.nonEmptyColumnIndexes || []).join(", ") || "(none)"}`)
    lines.push(`- Conditional format rules: ${sheet.conditionalFormatRuleCount}`)
    lines.push(`- Conditional format summary: ${(sheet.conditionalFormatSummary || []).map(rule => `${rule.ruleType}:${(rule.ranges || []).join(",") || "(no range)"}`).join(" / ") || "(none)"}`)
    lines.push(`- Data validation rules: ${sheet.dataValidationRuleCount || 0}`)
    lines.push(`- Protections: ${sheet.protectionCount}`)
    lines.push(`- Background colors: ${(sheet.backgroundColorSummary || []).map(entry => `${entry.color} x${entry.sampleCellCount}`).join(", ") || "(none)"}`)
    lines.push(`- Named ranges: ${(sheet.namedRanges || []).map(range => range.name).join(", ") || "(none)"}`)
    lines.push(`- Special markers: ${(sheet.specialMarkers || []).map(marker => `${marker.type}:${marker.value}`).join(", ") || "(none)"}`)
    lines.push("")
    lines.push("Header rows:")
    lines.push("")
    if ((sheet.headerRowValues || []).length === 0) {
      lines.push("- (none)")
    } else {
      sheet.headerRowValues.forEach((row, index) => {
        lines.push(`- Row ${index + 1}: ${row.join(" | ")}`)
      })
    }
    lines.push("")
    lines.push("Sample rows:")
    lines.push("")
    if ((sheet.sampleRows || []).length === 0) {
      lines.push("- (none)")
    } else {
      sheet.sampleRows.forEach(row => {
        lines.push(`- Row ${row.row}: ${row.values.join(" | ")}`)
      })
    }
    lines.push("")
    lines.push("Formula columns:")
    lines.push("")
    if ((sheet.formulaSummary || []).length === 0) {
      lines.push("- (none)")
    } else {
      sheet.formulaSummary.forEach(formula => {
        lines.push(`- ${formula.columnLetter} (${formula.header || "(no header)"}): ${formula.sampledFormulaCount} sampled formulas, e.g. ${formula.representativeFormula}`)
      })
    }
    lines.push("")
  })

  return lines.join("\n") + "\n"
}
