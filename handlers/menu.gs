function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('スプレッドチャット',
    [
      {name: '表示(チャット主体)', functionName: 'applyBaseLayout'},
      {name: '表示切替(Map主体)', functionName: 'applyMapLayout'},
      null,
      {name: 'CurrentMapを登録', functionName: 'setCurrentMap'},
      null,
      {name: 'キャッシュクリア', functionName: 'clearScriptCache'},
      {name: 'トリガーON/OFF', functionName: 'syncEditInstallableTrigger'},
      {name: '権限を更新', functionName: 'setEditorsByList'},
      {name: '権限をチェック', functionName: 'getPermissionsTest'},
    ]
  );

  var dic = getDictFromSheet(SETTING_SHEET_NAME);
    
}
