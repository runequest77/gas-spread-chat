
/**
 * usersではなくusersDataなのは、取得するデータとMAP上での要素名に区別をつけるため
 */
function getUsersData() {
  // パレットシートを取得
  const sheet = c_sheet('パレット');
  const range = sheet.getRange("A2:T4");
  const values = range.getValues();
  const usersData = [];
  
  for (let userNo = 0; userNo < 10; userNo++) {
    const index = userNo * 2;
    const email = values[0][index + 1];
    if (email) { // emailが存在する場合のみ、ユーザーを追加する
      const user = {
        userNo: userNo,
        email: email,
        pname: values[1][index + 1],
        cname: values[2][index + 1],
      };
      usersData.push(user);
    }
  }
  return usersData;
}

function clearPaletteCacheMenu(){
  // const ui = SpreadsheetApp.getUi();
  // const paletteNo = ui.prompt("パレットキャッシュをクリアします。キャラクター番号を入力して下さい。");
  clearPaletteCache(0);
  clearPaletteCache(1);
  clearPaletteCache(2);
  clearPaletteCache(3);
  clearPaletteCache(4);
  clearPaletteCache(5);
  clearPaletteCache(6);
  clearPaletteCache(7);
  clearPaletteCache(8);
  clearPaletteCache(9);
}

function onInstall(e) {
  onOpen(e);
}

function setEditorsByList() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var owner = ss.getOwner();  
  var activeUser = Session.getActiveUser().getEmail();
  if (owner != activeUser) {
    Browser.msgBox(`権限の更新はオーナーのみが行えます。\nowner:${owner}`);
    return;
  }

  //すべての範囲保護を除去
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  }

  var chat = c_sheet(CHAT_SHEET_NAME);  
  var palette = c_sheet(PALETTE_SHEET_NAME);  

  for (i=0; i<10; i++) {
    var email = palette.getRange(palette_emailRow, i*2+2).getValue();

    if (email.length > 5) {
      ss.addEditor(email);

      var dataCol = palette.getRange(getColSelector(i*2+1,i*2+2));
      var p1 = dataCol.protect().setDescription('DataCol'+i);
      //初期値は編集者全員が含まれているので、いったんクリアしないといけない
      p1.removeEditors(p1.getEditors());
      p1.addEditor(email);

      var chatCol = chat.getRange(getColSelector(rollColStart+i,rollColStart+i));
      console.log('ChatCol'+i);
      var p2 = chatCol.protect().setDescription('ChatCol'+i);
      p2.removeEditors(p2.getEditors());
      p2.addEditor(email)
    } else {
      console.log('NotSet'+i);
    }
  }

  //ヘッダーからはEditorを抜く（管理者は常に編集できる）
  var chatRowHeader = chat.getRange("1:3");
  var p3 = chatRowHeader.protect().setDescription('ChatRowHeader');
  p3.removeEditors(p3.getEditors());

  var pcRowHeader = palette.getRange("1:2");
  var p4 = pcRowHeader.protect().setDescription('PaletteRowHeader');
  p4.removeEditors(p4.getEditors());

  var rollLog = ss.getSheetByName("RollLog");  
  var rollLogRowHeader = rollLog.getRange("1:1");
  var p5 = rollLogRowHeader.protect().setDescription('RollLogRowHeader');
  p5.removeEditors(p5.getEditors());

  Browser.msgBox("パレットシートに記載されたgoogleアカウントに従って編集権限を設定しました。");
}

