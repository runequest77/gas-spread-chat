/** スプレッドチャットメニュー　チャット用レイアウト */
function applyBaseLayout() {
  const sheet = c_sheet(CHAT_SHEET_NAME);
  sheet.hideColumns(timestampCol);
  sheet.hideColumns(rollIDCol);
  sheet.showColumns(imageCol);
  sheet.setColumnWidth(actionCol, 560);
}

/** スプレッドチャットメニュー　MAP用レイアウト */
function applyMapLayout() {
  const sheet = c_sheet(CHAT_SHEET_NAME);
  sheet.hideColumns(timestampCol);
  sheet.hideColumns(rollIDCol);
  sheet.hideColumns(imageCol);
  sheet.setColumnWidth(actionCol, 280);
}
