function syncEditInstallableTrigger() {
  // このスクリプトでのトリガーを全削除
  var allTriggers = ScriptApp.getProjectTriggers();
  for( var i = 0; i < allTriggers.length; ++i ){
    ScriptApp.deleteTrigger(allTriggers[i]);
  }

  if (useTrigger==1) {
    var sheet = SpreadsheetApp.getActive();
    ScriptApp.newTrigger("onEditInstallable")
      .forSpreadsheet(sheet)
      .onEdit()
      .create();
    Browser.msgBox("トリガーを有効にしました");
  } else {
    Browser.msgBox("トリガーを無効にしました");
  }
}

function setTriggerOnEdit() {
  return syncEditInstallableTrigger();
}

/**
 * トリガーから実行されるのでオーナー権限で実行される
 * 1.チャットシートの行の高さを変更
 * 2.Webマップのポップアップ/タブ表示
 */
function onEditInstallable(e) {
  logFunctionNameAndArgs('e省略');
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  // if (sheetName == MAP_SHEET.name) {
  //   onEditMapSheet(e);
  // }

  if (sheetName != CHAT_SHEET_NAME) return;  //★チャットシート以外処理しない

  const cell = e.range;
  const e_pos = getEventPosition(cell);
  const col = e_pos.columnStart;
  const row = e_pos.rowStart;

  if (col == nameCol) {
    adjustRowHeightForNameColumn(sheet, row, e.value);
  }

}

function onEdit_Trigger(e) {
  return onEditInstallable(e);
}
