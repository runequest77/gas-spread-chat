const SPREADCHAT_SHEET_GENERATOR = {
  // 共通で使う配色パレットです。
  standardPalette: [
    ["#000000", "#434343", "#666666", "#999999", "#b7b7b7", "#cccccc", "#d9d9d9", "#efefef", "#f3f3f3", "#ffffff"],
    ["#980000", "#ff0000", "#ff9900", "#ffff00", "#00ff00", "#00ffff", "#4a86e8", "#0000ff", "#9900ff", "#ff00ff"],
    ["#e6b8af", "#f4cccc", "#fce5cd", "#fff2cc", "#d9ead3", "#d0e0e3", "#c9daf8", "#cfe2f3", "#d9d2e9", "#ead1dc"],
    ["#dd7e6b", "#ea9999", "#f9cb9c", "#ffe599", "#b6d7a8", "#a2c4c9", "#a4c2f4", "#9fc5e8", "#b4a7d6", "#d5a6bd"],
    ["#cc4125", "#e06666", "#f6b26b", "#ffd966", "#93c47d", "#76a5af", "#6d9eeb", "#6fa8dc", "#8e7cc3", "#c27ba0"],
    ["#a61c00", "#cc0000", "#e69138", "#f1c232", "#6aa84f", "#45818e", "#3c78d8", "#3d85c6", "#674ea7", "#a64d79"],
    ["#85200c", "#990000", "#b45f06", "#bf9000", "#38761d", "#134f5c", "#1155cc", "#0b5394", "#351c75", "#741b47"],
    ["#5b0f00", "#660000", "#783f04", "#7f6000", "#274e13", "#0c343d", "#1c4587", "#073763", "#20124d", "#4c1130"]
  ],
  // 参加者スロットは 0(GM) を含めて 10 枠です。
  slotCount: 10,
  // 参加者列の既定背景色です。
  defaultParticipantSlotColors: [
    "#e6b8af",
    "#f4cccc",
    "#fce5cd",
    "#fff2cc",
    "#d9ead3",
    "#d0e0e3",
    "#c9daf8",
    "#cfe2f3",
    "#d9d2e9",
    "#ead1dc"
  ],
  // 秘匿やパレット判定に使う記号です。
  markers: {
    secretPrefix: typeof secretPrefix !== "undefined" ? secretPrefix : "秘匿",
    paletteSecret: "■"
  },
  // 全シート共通の既定行高です。
  defaultRowHeight: 21
}

/**
 * 必要シート生成用の定義一覧を返します。
 * semantic inventory をもとに、現状必要な固定構造のみをコード化しています。
 * @param {object=} options
 * @return {{participantSlots: object[], sheets: object[]}}
 */
function getSpreadChatRequiredSheetSpecs(options) {
  const normalizedOptions = normalizeSpreadChatSheetGeneratorOptions_(options)
  const participantSlots = buildSpreadChatParticipantSlots_(normalizedOptions)
  return {
      participantSlots: participantSlots,
      sheets: [
        Object.assign({}, SPREADCHAT_SHEET_CHAT_SPEC, {
        values: [
          buildSpreadChatChatCounterRow_(participantSlots),
          Array(SPREADCHAT_SHEET_CHAT.layout.rollColumnEnd).fill(""),
          buildSpreadChatChatHeaderRow_(participantSlots)
        ],
        formulaCells: buildSpreadChatChatHeaderFormulaCells_(participantSlots),
        participantColorBands: normalizedOptions.applyParticipantColumnColors === false ? [] : participantSlots.map(slot => ({
          rowStart: 1,
          rowCount: SPREADCHAT_SHEET_CHAT.initialRowCount,
          columnStart: slot.chatColumn,
          columnCount: 1,
          color: slot.color
        }))
      }),
      Object.assign({}, SPREADCHAT_SHEET_PALETTE_SPEC, {
        values: [
          buildSpreadChatPaletteHeaderRow_(participantSlots),
          buildSpreadChatPaletteKeyRow_("email", participantSlots, true),
          buildSpreadChatPaletteKeyRow_("pname", participantSlots, false),
          buildSpreadChatPaletteCnameRow_(participantSlots)
        ],
        participantColorBands: normalizedOptions.applyParticipantColumnColors === false ? [] : participantSlots.map(slot => ({
          rowStart: 1,
          rowCount: SPREADCHAT_SHEET_PALETTE.initialRowCount,
          columnStart: slot.paletteColumnStart,
          columnCount: 2,
          color: slot.color
        }))
      }),
      SPREADCHAT_SHEET_ROLLLOG_SPEC,
      SPREADCHAT_SHEET_SELECTLIST_SPEC,
      Object.assign({}, SPREADCHAT_SHEET_SETTING_SPEC, {
        values: buildSpreadChatSettingRows_()
      }),
      SPREADCHAT_SHEET_CHARTS_SPEC
    ]
  }
}

/**
 * 必要シートの骨格を作成または更新します。
 * 既存データの破壊を避けるため、既定では missing sheet を設計どおりのサイズ/レイアウトで作成し、
 * 既存 sheet へのサイズ縮小やヘッダ上書きは overwriteExisting=true のときだけ行います。
 * @param {object=} options
 * @return {{participantSlots: object[], sheets: object[]}}
 */
function createSpreadChatRequiredSheets(options) {
  const normalizedOptions = normalizeSpreadChatSheetGeneratorOptions_(options)
  const plan = getSpreadChatRequiredSheetSpecs(normalizedOptions)
  const spreadsheet = normalizedOptions.spreadsheet || getSpreadChatGeneratorSpreadsheet_()
  const existingSheetMap = {}
  if (typeof spreadsheet.getSheetByName === "function") {
    plan.sheets.forEach(spec => {
      existingSheetMap[spec.sheetName] = spreadsheet.getSheetByName(spec.sheetName)
    })
  }
  const sheetMap = {}
  const layoutOptions = Object.assign({}, normalizedOptions, { skipFormulas: true })

  plan.sheets.forEach(spec => {
    sheetMap[spec.sheetName] = applySpreadChatSheetSpec_(spreadsheet, spec, layoutOptions)
  })
  applySpreadChatPlanFormulaCells_(plan.sheets, sheetMap, existingSheetMap, normalizedOptions)
  return plan
}

function normalizeSpreadChatSheetGeneratorOptions_(options) {
  const opts = options || {}
  const spreadsheet = opts.spreadsheet || null
  return {
    spreadsheet: spreadsheet,
    overwriteExisting: opts.overwriteExisting === true,
    skipFormulas: opts.skipFormulas === true,
    applyParticipantColumnColors: opts.applyParticipantColumnColors !== false,
    participantSlotColors: resolveSpreadChatParticipantSlotColors_(opts.participantSlotColors, spreadsheet)
  }
}

function resolveSpreadChatParticipantSlotColors_(colors, spreadsheet) {
  if (Array.isArray(colors) && colors.length > 0) {
    return normalizeSpreadChatParticipantSlotColors_(colors)
  }
  const extractedColors = extractSpreadChatParticipantSlotColorsFromSpreadsheet_(spreadsheet)
  if (extractedColors.length > 0) {
    return normalizeSpreadChatParticipantSlotColors_(extractedColors.concat(SPREADCHAT_SHEET_GENERATOR.defaultParticipantSlotColors))
  }
  return normalizeSpreadChatParticipantSlotColors_(SPREADCHAT_SHEET_GENERATOR.defaultParticipantSlotColors)
}

function normalizeSpreadChatParticipantSlotColors_(colors) {
  const normalized = (Array.isArray(colors) ? colors : SPREADCHAT_SHEET_GENERATOR.defaultParticipantSlotColors)
    .map(color => normalizeSpreadChatColor_(color))
    .filter(Boolean)
  if (normalized.length === 0) {
    return SPREADCHAT_SHEET_GENERATOR.defaultParticipantSlotColors.slice()
  }
  while (normalized.length < SPREADCHAT_SHEET_GENERATOR.slotCount) {
    normalized.push(normalized[normalized.length - 1])
  }
  return normalized.slice(0, SPREADCHAT_SHEET_GENERATOR.slotCount)
}

function normalizeSpreadChatColor_(color) {
  const normalized = String(color || "").trim().toLowerCase()
  if (normalized === "" || normalized === "#ffffff" || normalized === "white") {
    return ""
  }
  return normalized
}

function buildSpreadChatParticipantSlots_(options) {
  const colors = normalizeSpreadChatParticipantSlotColors_(options && options.participantSlotColors)
  const slots = []
  for (let userNo = 0; userNo < SPREADCHAT_SHEET_GENERATOR.slotCount; userNo++) {
    slots.push({
      userNo: userNo,
      chatColumn: SPREADCHAT_SHEET_CHAT.layout.rollColumnStart + userNo,
      paletteColumnStart: userNo * 2 + 1,
      paletteColumnEnd: userNo * 2 + 2,
      chatHeaderLabel: userNo === 0 ? "0 GM" : `${userNo} `,
      paletteHeaderLabel: String(userNo),
      color: colors[userNo]
    })
  }
  return slots
}

/**
 * 条件付き書式の REGEXMATCH 文字列リテラル向けに値をエスケープします。
 * まずバックスラッシュを二重化して後続のメタ文字エスケープと干渉しないようにし、
 * その後で正規表現メタ文字を無効化します。最後に Google Sheets の数式文字列リテラル用として
 * ダブルクオートを `""` に変換し、固定プレフィックスを安全に `=REGEXMATCH(..., "^...")`
 * へ埋め込めるようにします。
 * Escape order matters: backslashes must be doubled first, then regex metacharacters escaped,
 * and finally quotes doubled for Google Sheets string literal syntax.
 * @param {string} text 条件付き書式の正規表現へ埋め込む固定文字列。
 * @return {string} REGEXMATCH と Sheets 数式文字列の両方で安全に使えるエスケープ済み文字列。
 */
function escapeSpreadChatConditionalFormatRegexLiteral_(text) {
  return String(text || "")
    // Preserve literal backslashes before escaping regex metacharacters.
    .replace(/\\/g, "\\\\")
    // Escape regex metacharacters so the prefix is matched literally.
    .replace(/([.^$*+?()[\]{}|])/g, "\\$1")
    // Escape double quotes for Google Sheets formula string literals.
    .replace(/"/g, '""')
}

function applySpreadChatSheetSpec_(spreadsheet, spec, options) {
  const existingSheet = typeof spreadsheet.getSheetByName === "function" ? spreadsheet.getSheetByName(spec.sheetName) : null
  const sheet = existingSheet || spreadsheet.insertSheet(spec.sheetName)
  const shouldApplyLayout = !existingSheet || options.overwriteExisting
  const shouldWriteValues = !existingSheet || options.overwriteExisting
  const initialRowCount = getSpreadChatSpecInitialRowCount_(spec)
  const initialColumnCount = getSpreadChatSpecInitialColumnCount_(spec)

  ensureSpreadChatSheetSize_(sheet, initialRowCount, initialColumnCount, shouldApplyLayout)
  if (typeof sheet.setFrozenRows === "function") {
    sheet.setFrozenRows(spec.frozenRows || 0)
  }
  if (typeof sheet.setFrozenColumns === "function") {
    sheet.setFrozenColumns(spec.frozenColumns || 0)
  }
  if (shouldWriteValues) {
    writeSpreadChatSheetValues_(sheet, getSpreadChatSpecValues_(spec))
    if (!options.skipFormulas) {
      applySpreadChatFormulaCells_(sheet, spec.formulaCells || [])
    }
  }
  applySpreadChatParticipantColorBands_(sheet, spec.participantColorBands || [])
  if (shouldApplyLayout) {
    applySpreadChatSheetLayout_(sheet, spec)
    applySpreadChatMergedRanges_(sheet, spec.mergedRanges || [])
    applySpreadChatStyleRanges_(sheet, spec.styleRanges || [])
    applySpreadChatWrapStrategyRanges_(sheet, spec.wrapStrategyRanges || [])
    applySpreadChatBackgroundRanges_(sheet, spec.backgroundRanges || [])
    applySpreadChatFontColorRanges_(sheet, spec.fontColorRanges || [])
    applySpreadChatConditionalFormatRules_(sheet, spec.conditionalFormatRules || [])
    applySpreadChatDataValidationRules_(sheet, spec.dataValidationRules || [])
  }
  return sheet
}

function getSpreadChatSpecInitialRowCount_(spec) {
  const explicitRowCount = typeof spec.initialRowCount === "number" ? spec.initialRowCount : spec.maxRows
  return Math.max(explicitRowCount || 0, getSpreadChatSpecValueRowCount_(spec))
}

function getSpreadChatSpecInitialColumnCount_(spec) {
  const explicitColumnCount = typeof spec.initialColumnCount === "number" ? spec.initialColumnCount : spec.maxColumns
  return Math.max(explicitColumnCount || 0, getSpreadChatSpecValueColumnCount_(spec))
}

function getSpreadChatSpecValues_(spec) {
  return spec.values || []
}

function getSpreadChatSpecValueRowCount_(spec) {
  const values = getSpreadChatSpecValues_(spec)
  return Array.isArray(values) ? values.length : 0
}

function getSpreadChatSpecValueColumnCount_(spec) {
  const values = getSpreadChatSpecValues_(spec)
  if (!Array.isArray(values) || values.length === 0) {
    return 0
  }
  return values.reduce((maxCount, row) => Math.max(maxCount, Array.isArray(row) ? row.length : 0), 0)
}

function applySpreadChatPlanFormulaCells_(specs, sheetMap, existingSheetMap, options) {
  specs.forEach(spec => {
    const existedBeforeApply = !!existingSheetMap[spec.sheetName]
    if (existedBeforeApply && !options.overwriteExisting) {
      return
    }
    const sheet = sheetMap[spec.sheetName]
    if (!sheet) {
      return
    }
    applySpreadChatFormulaCells_(sheet, spec.formulaCells || [])
  })
}

function ensureSpreadChatSheetSize_(sheet, rowCount, columnCount, shouldTrimToExactSize) {
  const currentMaxRows = typeof sheet.getMaxRows === "function" ? sheet.getMaxRows() : 0
  const currentMaxColumns = typeof sheet.getMaxColumns === "function" ? sheet.getMaxColumns() : 0
  if (rowCount > currentMaxRows && typeof sheet.insertRowsAfter === "function") {
    sheet.insertRowsAfter(Math.max(1, currentMaxRows), rowCount - currentMaxRows)
  }
  if (shouldTrimToExactSize && rowCount < currentMaxRows && typeof sheet.deleteRows === "function") {
    sheet.deleteRows(rowCount + 1, currentMaxRows - rowCount)
  }
  if (columnCount > currentMaxColumns && typeof sheet.insertColumnsAfter === "function") {
    sheet.insertColumnsAfter(Math.max(1, currentMaxColumns), columnCount - currentMaxColumns)
  }
  if (shouldTrimToExactSize && columnCount < currentMaxColumns && typeof sheet.deleteColumns === "function") {
    sheet.deleteColumns(columnCount + 1, currentMaxColumns - columnCount)
  }
}

function writeSpreadChatSheetValues_(sheet, values) {
  if (!values || values.length === 0) {
    return
  }
  const rowCount = values.length
  const columnCount = values.reduce((maxCount, row) => Math.max(maxCount, row.length), 0)
  const normalizedValues = values.map(row => {
    const cloned = row.slice()
    while (cloned.length < columnCount) {
      cloned.push("")
    }
    return cloned
  })
  sheet.getRange(1, 1, rowCount, columnCount).setValues(normalizedValues)
}

function applySpreadChatParticipantColorBands_(sheet, bands) {
  bands.forEach(band => {
    if (!band.color) {
      return
    }
    const range = sheet.getRange(band.rowStart, band.columnStart, band.rowCount, band.columnCount)
    if (typeof range.setBackground === "function") {
      range.setBackground(band.color)
    }
  })
}

function applySpreadChatFormulaCells_(sheet, formulaCells) {
  formulaCells.forEach(spec => {
    if (!spec || !spec.formula) {
      return
    }
    const range = sheet.getRange(spec.rowStart, spec.columnStart, 1, 1)
    if (typeof range.setFormula === "function") {
      range.setFormula(spec.formula)
      return
    }
    if (typeof range.setFormulas === "function") {
      range.setFormulas([[spec.formula]])
    }
  })
}

function applySpreadChatBackgroundRanges_(sheet, backgroundRanges) {
  applySpreadChatColorRanges_(sheet, backgroundRanges, "setBackground")
}

function applySpreadChatFontColorRanges_(sheet, fontColorRanges) {
  applySpreadChatColorRanges_(sheet, fontColorRanges, "setFontColor")
}

function applySpreadChatColorRanges_(sheet, ranges, setterName) {
  ranges.forEach(spec => {
    if (!spec.color) {
      return
    }
    const range = sheet.getRange(spec.rowStart, spec.columnStart, spec.rowCount, spec.columnCount)
    if (typeof range[setterName] === "function") {
      range[setterName](spec.color)
    }
  })
}

function applySpreadChatSheetLayout_(sheet, spec) {
  applySpreadChatDefaultRowHeight_(sheet, getSpreadChatSpecInitialRowCount_(spec), spec.defaultRowHeight)
  applySpreadChatColumnWidths_(sheet, spec.columnWidths || [])
  applySpreadChatRowHeights_(sheet, spec.rowHeights || [])
  applySpreadChatHiddenColumns_(sheet, spec.hiddenColumns || [])
  applySpreadChatHiddenRows_(sheet, spec.hiddenRows || [])
}

function applySpreadChatDefaultRowHeight_(sheet, rowCount, rowHeight) {
  if (!rowHeight) {
    return
  }
  if (typeof sheet.setRowHeights === "function") {
    sheet.setRowHeights(1, rowCount, rowHeight)
    return
  }
  if (typeof sheet.setRowHeight === "function") {
    for (let row = 1; row <= rowCount; row++) {
      sheet.setRowHeight(row, rowHeight)
    }
  }
}

function applySpreadChatColumnWidths_(sheet, columnWidths) {
  columnWidths.forEach(spec => {
    if (typeof sheet.setColumnWidth !== "function") {
      return
    }
    for (let column = 0; column < spec.columnCount; column++) {
      sheet.setColumnWidth(spec.columnStart + column, spec.width)
    }
  })
}

function applySpreadChatRowHeights_(sheet, rowHeights) {
  rowHeights.forEach(spec => {
    if (typeof sheet.setRowHeights === "function") {
      sheet.setRowHeights(spec.rowStart, spec.rowCount, spec.height)
      return
    }
    if (typeof sheet.setRowHeight === "function") {
      for (let row = 0; row < spec.rowCount; row++) {
        sheet.setRowHeight(spec.rowStart + row, spec.height)
      }
    }
  })
}

function applySpreadChatHiddenColumns_(sheet, hiddenColumns) {
  hiddenColumns.forEach(spec => {
    if (typeof sheet.hideColumns === "function") {
      sheet.hideColumns(spec.columnStart, spec.columnCount)
    }
  })
}

function applySpreadChatHiddenRows_(sheet, hiddenRows) {
  hiddenRows.forEach(spec => {
    if (typeof sheet.hideRows === "function") {
      sheet.hideRows(spec.rowStart, spec.rowCount)
    }
  })
}

function applySpreadChatMergedRanges_(sheet, mergedRanges) {
  mergedRanges.forEach(spec => {
    const range = sheet.getRange(spec.rowStart, spec.columnStart, spec.rowCount, spec.columnCount)
    if (typeof range.merge === "function") {
      range.merge()
    }
  })
}

function applySpreadChatStyleRanges_(sheet, styleRanges) {
  styleRanges.forEach(spec => {
    const range = sheet.getRange(spec.rowStart, spec.columnStart, spec.rowCount, spec.columnCount)
    if (spec.verticalAlignment && typeof range.setVerticalAlignment === "function") {
      range.setVerticalAlignment(spec.verticalAlignment)
    }
    if (spec.horizontalAlignment && typeof range.setHorizontalAlignment === "function") {
      range.setHorizontalAlignment(spec.horizontalAlignment)
    }
    if (spec.fontSize && typeof range.setFontSize === "function") {
      range.setFontSize(spec.fontSize)
    }
    if (spec.fontFamily && typeof range.setFontFamily === "function") {
      range.setFontFamily(spec.fontFamily)
    }
    if (spec.fontStyle && typeof range.setFontStyle === "function") {
      range.setFontStyle(spec.fontStyle)
    }
    if (spec.fontWeight && typeof range.setFontWeight === "function") {
      range.setFontWeight(spec.fontWeight)
    }
    if (spec.backgroundColor && typeof range.setBackground === "function") {
      range.setBackground(spec.backgroundColor)
    }
    if (spec.fontColor && typeof range.setFontColor === "function") {
      range.setFontColor(spec.fontColor)
    }
  })
}

function applySpreadChatWrapStrategyRanges_(sheet, wrapStrategyRanges) {
  wrapStrategyRanges.forEach(spec => {
    const wrapStrategy = resolveSpreadChatWrapStrategy_(spec.wrapStrategy)
    const range = sheet.getRange(spec.rowStart, spec.columnStart, spec.rowCount, spec.columnCount)
    if (wrapStrategy && typeof range.setWrapStrategy === "function") {
      range.setWrapStrategy(wrapStrategy)
    }
  })
}

function resolveSpreadChatWrapStrategy_(wrapStrategy) {
  if (!wrapStrategy) {
    return null
  }
  const enumMap = typeof SpreadsheetApp !== "undefined" ? SpreadsheetApp.WrapStrategy : null
  if (enumMap && enumMap[wrapStrategy]) {
    return enumMap[wrapStrategy]
  }
  return wrapStrategy
}

function applySpreadChatConditionalFormatRules_(sheet, rules) {
  if (!canApplySpreadChatConditionalFormatRules_(sheet)) {
    return
  }
  const builtRules = []
  rules.forEach(ruleSpec => {
    const ruleBuilder = SpreadsheetApp.newConditionalFormatRule()
    const ruleRanges = resolveSpreadChatRuleRanges_(sheet, ruleSpec)
    const criteriaValues = normalizeSpreadChatRuleCriteriaValues_(ruleSpec)
    const firstCriteriaValue = criteriaValues[0]
    if (ruleRanges.length === 0) {
      return
    }
    if (ruleSpec.criteriaType === "CUSTOM_FORMULA" && typeof ruleBuilder.whenFormulaSatisfied === "function") {
      ruleBuilder.whenFormulaSatisfied(firstCriteriaValue)
    } else if (ruleSpec.criteriaType === "TEXT_STARTS_WITH" && typeof ruleBuilder.whenTextStartsWith === "function") {
      ruleBuilder.whenTextStartsWith(firstCriteriaValue)
    } else if (ruleSpec.criteriaType === "TEXT_ENDS_WITH" && typeof ruleBuilder.whenTextEndsWith === "function") {
      ruleBuilder.whenTextEndsWith(firstCriteriaValue)
    } else if (ruleSpec.criteriaType === "CELL_NOT_EMPTY" && typeof ruleBuilder.whenCellNotEmpty === "function") {
      ruleBuilder.whenCellNotEmpty()
    } else if (ruleSpec.criteriaType === "CELL_EMPTY" && typeof ruleBuilder.whenCellEmpty === "function") {
      ruleBuilder.whenCellEmpty()
    } else if (ruleSpec.criteriaType === "NUMBER_GREATER_THAN" && typeof ruleBuilder.whenNumberGreaterThan === "function") {
      ruleBuilder.whenNumberGreaterThan(firstCriteriaValue)
    } else {
      return
    }
    applySpreadChatConditionalFormatStyle_(ruleBuilder, ruleSpec)
    if (typeof ruleBuilder.setRanges === "function") {
      ruleBuilder.setRanges(ruleRanges)
    }
    if (typeof ruleBuilder.build === "function") {
      builtRules.push(ruleBuilder.build())
    }
  })
  sheet.setConditionalFormatRules(builtRules)
}

function applySpreadChatConditionalFormatStyle_(ruleBuilder, ruleSpec) {
  const format = getSpreadChatConditionalFormatStyleSpec_(ruleSpec)
  if (format.backgroundColor && typeof ruleBuilder.setBackground === "function") {
    ruleBuilder.setBackground(format.backgroundColor)
  }
  if (format.fontColor && typeof ruleBuilder.setFontColor === "function") {
    ruleBuilder.setFontColor(format.fontColor)
  }
  if (typeof format.bold === "boolean" && typeof ruleBuilder.setBold === "function") {
    ruleBuilder.setBold(format.bold)
  }
  if (typeof format.italic === "boolean" && typeof ruleBuilder.setItalic === "function") {
    ruleBuilder.setItalic(format.italic)
  }
  if (typeof format.underline === "boolean" && typeof ruleBuilder.setUnderline === "function") {
    ruleBuilder.setUnderline(format.underline)
  }
  if (typeof format.strikethrough === "boolean" && typeof ruleBuilder.setStrikethrough === "function") {
    ruleBuilder.setStrikethrough(format.strikethrough)
  }
  if (format.fontFamily && typeof ruleBuilder.setFontFamily === "function") {
    ruleBuilder.setFontFamily(format.fontFamily)
  }
  if (typeof format.fontSize === "number" && typeof ruleBuilder.setFontSize === "function") {
    ruleBuilder.setFontSize(format.fontSize)
  }
}

function getSpreadChatConditionalFormatStyleSpec_(ruleSpec) {
  const format = ruleSpec && typeof ruleSpec.format === "object" && ruleSpec.format
    ? Object.assign({}, ruleSpec.format)
    : {}
  if (ruleSpec && typeof ruleSpec.backgroundColor !== "undefined" && typeof format.backgroundColor === "undefined") {
    format.backgroundColor = ruleSpec.backgroundColor
  }
  if (ruleSpec && typeof ruleSpec.fontColor !== "undefined" && typeof format.fontColor === "undefined") {
    format.fontColor = ruleSpec.fontColor
  }
  if (ruleSpec && typeof ruleSpec.bold === "boolean" && typeof format.bold === "undefined") {
    format.bold = ruleSpec.bold
  }
  if (ruleSpec && typeof ruleSpec.italic === "boolean" && typeof format.italic === "undefined") {
    format.italic = ruleSpec.italic
  }
  if (ruleSpec && typeof ruleSpec.underline === "boolean" && typeof format.underline === "undefined") {
    format.underline = ruleSpec.underline
  }
  if (ruleSpec && typeof ruleSpec.strikethrough === "boolean" && typeof format.strikethrough === "undefined") {
    format.strikethrough = ruleSpec.strikethrough
  }
  if (ruleSpec && typeof ruleSpec.fontFamily !== "undefined" && typeof format.fontFamily === "undefined") {
    format.fontFamily = ruleSpec.fontFamily
  }
  if (ruleSpec && typeof ruleSpec.fontSize !== "undefined" && typeof format.fontSize === "undefined") {
    format.fontSize = ruleSpec.fontSize
  }
  return format
}

function applySpreadChatDataValidationRules_(sheet, rules) {
  if (!canApplySpreadChatDataValidationRules_(sheet)) {
    return
  }
  rules.forEach(ruleSpec => {
    const rule = buildSpreadChatDataValidationRule_(ruleSpec)
    if (!rule) {
      return
    }
    resolveSpreadChatRuleRanges_(sheet, ruleSpec).forEach(range => {
      if (typeof range.setDataValidation === "function") {
        range.setDataValidation(rule)
      }
    })
  })
}

function buildSpreadChatDataValidationRule_(ruleSpec) {
  if (typeof SpreadsheetApp === "undefined" || typeof SpreadsheetApp.newDataValidation !== "function") {
    return null
  }
  const builder = SpreadsheetApp.newDataValidation()
  const criteriaType = ruleSpec && ruleSpec.criteriaType
  const criteriaValues = normalizeSpreadChatRuleCriteriaValues_(ruleSpec)
  const criteriaEnum = SpreadsheetApp.DataValidationCriteria && criteriaType
    ? SpreadsheetApp.DataValidationCriteria[criteriaType]
    : null
  if (!criteriaEnum || typeof builder.withCriteria !== "function") {
    return null
  }
  builder.withCriteria(criteriaEnum, criteriaValues)
  if (typeof ruleSpec.allowInvalid === "boolean" && typeof builder.setAllowInvalid === "function") {
    builder.setAllowInvalid(ruleSpec.allowInvalid)
  }
  if (ruleSpec.helpText && typeof builder.setHelpText === "function") {
    builder.setHelpText(ruleSpec.helpText)
  }
  return typeof builder.build === "function" ? builder.build() : null
}

function normalizeSpreadChatRuleCriteriaValues_(ruleSpec) {
  if (!ruleSpec) {
    return []
  }
  if (Array.isArray(ruleSpec.criteriaValues)) {
    return ruleSpec.criteriaValues.slice()
  }
  return typeof ruleSpec.criteriaValue === "undefined" ? [] : [ruleSpec.criteriaValue]
}

function resolveSpreadChatRuleRanges_(sheet, ruleSpec) {
  return ((ruleSpec && ruleSpec.ranges) || [])
    .map(rangeSpec => resolveSpreadChatRuleRange_(sheet, rangeSpec))
    .filter(range => !!range)
}

function resolveSpreadChatRuleRange_(sheet, rangeSpec) {
  if (!sheet || typeof sheet.getRange !== "function" || !rangeSpec) {
    return null
  }
  if (typeof rangeSpec === "string") {
    const a1Notation = getSpreadChatLocalRangeA1Notation_(rangeSpec)
    if (!a1Notation) {
      return null
    }
    return sheet.getRange(a1Notation)
  }
  if (typeof rangeSpec.rowStart === "number"
    && typeof rangeSpec.columnStart === "number"
    && typeof rangeSpec.rowCount === "number"
    && typeof rangeSpec.columnCount === "number") {
    return sheet.getRange(rangeSpec.rowStart, rangeSpec.columnStart, rangeSpec.rowCount, rangeSpec.columnCount)
  }
  return null
}

function getSpreadChatLocalRangeA1Notation_(rangeLabel) {
  const text = String(rangeLabel || "").trim()
  if (!text) {
    return ""
  }
  if (text.charAt(0) === "'") {
    for (let index = 1; index < text.length; index++) {
      if (text.charAt(index) !== "'") {
        continue
      }
      if (text.charAt(index + 1) === "'") {
        index++
        continue
      }
      if (text.charAt(index + 1) === "!") {
        return text.slice(index + 2)
      }
      break
    }
  }
  const separatorIndex = text.indexOf("!")
  return separatorIndex >= 0 ? text.slice(separatorIndex + 1) : text
}

/**
 * 条件付き書式 rule builder を安全に呼び出せる実行環境かどうかを返します。
 * @param {object} sheet
 * @return {boolean}
 */
function canApplySpreadChatConditionalFormatRules_(sheet) {
  return !!sheet
    && typeof sheet.setConditionalFormatRules === "function"
    && typeof SpreadsheetApp !== "undefined"
    && typeof SpreadsheetApp.newConditionalFormatRule === "function"
}

function canApplySpreadChatDataValidationRules_(sheet) {
  return !!sheet
    && typeof sheet.getRange === "function"
    && typeof SpreadsheetApp !== "undefined"
    && typeof SpreadsheetApp.newDataValidation === "function"
}

function extractSpreadChatParticipantSlotColorsFromSpreadsheet_(spreadsheet) {
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== "function") {
    return []
  }
  const chatColors = extractSpreadChatChatSlotColors_(spreadsheet.getSheetByName(SPREADCHAT_SHEET_CHAT.name))
  if (chatColors.length > 0) {
    return chatColors
  }
  return extractSpreadChatPaletteSlotColors_(spreadsheet.getSheetByName(SPREADCHAT_SHEET_PALETTE.name))
}

function extractSpreadChatChatSlotColors_(sheet) {
  if (!sheet || typeof sheet.getRange !== "function") {
    return []
  }
  try {
    const backgrounds = sheet.getRange(1, SPREADCHAT_SHEET_CHAT.layout.rollColumnStart, 1, SPREADCHAT_SHEET_GENERATOR.slotCount).getBackgrounds()
    return (backgrounds[0] || [])
      .map(color => normalizeSpreadChatColor_(color))
      .filter(Boolean)
  } catch (error) {
    return []
  }
}

function extractSpreadChatPaletteSlotColors_(sheet) {
  if (!sheet || typeof sheet.getRange !== "function") {
    return []
  }
  try {
    const backgrounds = sheet.getRange(1, 1, 1, SPREADCHAT_SHEET_GENERATOR.slotCount * 2).getBackgrounds()
    const firstRow = backgrounds[0] || []
    const colors = []
    for (let index = 0; index < firstRow.length; index += 2) {
      colors.push(firstRow[index])
    }
    return colors
      .map(color => normalizeSpreadChatColor_(color))
      .filter(Boolean)
  } catch (error) {
    return []
  }
}

function getSpreadChatGeneratorSpreadsheet_() {
  if (typeof c_ss === "function") {
    return c_ss()
  }
  return SpreadsheetApp.getActiveSpreadsheet()
}

function getSpreadChatColumnLetter_(columnNumber) {
  let normalized = Number(columnNumber) || 0
  let letters = ""
  while (normalized > 0) {
    const remainder = (normalized - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    normalized = Math.floor((normalized - 1) / 26)
  }
  return letters
}

function getSpreadChatColumnNumber_(columnLetter) {
  return String(columnLetter || "")
    .toUpperCase()
    .split("")
    .reduce((columnNumber, letter) => {
      return (columnNumber * 26) + (letter.charCodeAt(0) - 64)
    }, 0)
}

function getSpreadChatRowRangeSpec_(rowNotation) {
  const match = /^(\d+):(\d+)$/.exec(String(rowNotation || ""))
  if (!match) {
    throw new Error(`Invalid row notation: ${rowNotation}`)
  }
  const rowStart = Number(match[1])
  const rowEnd = Number(match[2])
  return {
    rowStart: rowStart,
    rowCount: Math.max(0, rowEnd - rowStart + 1)
  }
}

function getSpreadChatColumnRangeSpec_(columnNotation, totalRows) {
  const match = /^([A-Z]+):([A-Z]+)$/i.exec(String(columnNotation || ""))
  if (!match) {
    throw new Error(`Invalid column notation: ${columnNotation}`)
  }
  const columnStart = getSpreadChatColumnNumber_(match[1])
  const columnEnd = getSpreadChatColumnNumber_(match[2])
  return {
    rowStart: 1,
    rowCount: totalRows,
    columnStart: columnStart,
    columnCount: Math.max(0, columnEnd - columnStart + 1)
  }
}

function getSpreadChatA1RangeSpec_(a1Notation) {
  const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(String(a1Notation || ""))
  if (!match) {
    throw new Error(`Invalid A1 range notation: ${a1Notation}`)
  }
  const rowStart = Number(match[2])
  const rowEnd = Number(match[4])
  const columnStart = getSpreadChatColumnNumber_(match[1])
  const columnEnd = getSpreadChatColumnNumber_(match[3])
  return {
    rowStart: rowStart,
    columnStart: columnStart,
    rowCount: Math.max(0, rowEnd - rowStart + 1),
    columnCount: Math.max(0, columnEnd - columnStart + 1)
  }
}

function getSpreadChatEscapedSheetReference_(sheetName) {
  return `'${String(sheetName || "").replace(/'/g, "''")}'`
}
