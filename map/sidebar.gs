/**
 * map/sidebar.gs
 *
 * Firebase Realtime Database と連携するリアルタイム共有マップサイドバー。
 * MAP シートのグリッドをサイドバーに表示し、複数ユーザーが
 * ユニットをクリック配置・ドラッグ移動できる。
 * ユニット位置は Firebase Realtime Database に保存され、
 * 他ユーザーのサイドバーにもリアルタイムで反映される。
 */

/** MAP シートの名前（既存定数と合わせる） */
const MAP_SIDEBAR_SHEET_NAME = "MAP";

/**
 * サイドバーを表示する。
 * メニューから呼び出される。
 */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile("Sidebar")
    .evaluate()
    .setTitle("共有マップ")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * MAP シートの使用範囲を 2 次元配列として返す。
 * 行・列は 1 始まりでサイドバー側座標と一致させる。
 *
 * @returns {{ values: string[][], numRows: number, numCols: number }}
 * @throws エラーメッセージ文字列（シートが存在しない場合など）
 */
function getMapData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MAP_SIDEBAR_SHEET_NAME);
    if (!sheet) {
      throw new Error(
        "「" + MAP_SIDEBAR_SHEET_NAME + "」シートが見つかりません。" +
        "シート名が正しいか確認してください。"
      );
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // Date オブジェクトなど JSON 非対応の値を文字列へ変換する
    var sanitized = values.map(function(row) {
      return row.map(function(cell) {
        if (cell === null || cell === undefined) {
          return "";
        }
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        return String(cell);
      });
    });

    return {
      values: sanitized,
      numRows: sanitized.length,
      numCols: sanitized.length > 0 ? sanitized[0].length : 0
    };
  } catch (e) {
    // エラーをクライアントへそのまま返せるよう文字列化する
    throw e.message || String(e);
  }
}

/**
 * Firebase 設定とユーザー識別子をクライアントへ渡す設定オブジェクトを返す。
 * Firebase の設定値は Script Properties から読む。
 * ユーザー識別子は Session.getTemporaryActiveUserKey() を優先する。
 *
 * Script Properties に設定するキー:
 *   FIREBASE_API_KEY
 *   FIREBASE_AUTH_DOMAIN
 *   FIREBASE_DATABASE_URL
 *   FIREBASE_PROJECT_ID
 *   FIREBASE_APP_ID
 *
 * @returns {{
 *   firebaseConfig: {
 *     apiKey: string,
 *     authDomain: string,
 *     databaseURL: string,
 *     projectId: string,
 *     appId: string
 *   },
 *   userKey: string,
 *   displayName: string,
 *   roomId: string
 * }}
 * @throws エラーメッセージ文字列（必須設定が未設定の場合など）
 */
function getClientConfig() {
  try {
    var props = PropertiesService.getScriptProperties();

    var apiKey       = props.getProperty("FIREBASE_API_KEY");
    var authDomain   = props.getProperty("FIREBASE_AUTH_DOMAIN");
    var databaseURL  = props.getProperty("FIREBASE_DATABASE_URL");
    var projectId    = props.getProperty("FIREBASE_PROJECT_ID");
    var appId        = props.getProperty("FIREBASE_APP_ID");

    if (!apiKey || !databaseURL || !projectId || !appId) {
      throw new Error(
        "Firebase の設定が Script Properties に登録されていません。\n" +
        "setupFirebaseProperties() を実行して設定してください。\n" +
        "必須: FIREBASE_API_KEY, FIREBASE_DATABASE_URL, FIREBASE_PROJECT_ID, FIREBASE_APP_ID"
      );
    }

    // ユーザー識別子の取得
    // getTemporaryActiveUserKey() はスプレッドシート閲覧者にも一意のキーを返す。
    // ただし匿名でメールは露出しない。
    var userKey = "";
    try {
      userKey = Session.getTemporaryActiveUserKey();
    } catch (keyErr) {
      // フォールバック: スクリプトプロパティにキャッシュしたキーを使う（タブをまたいで同一ユーザーと判定するため）
      // これは完全な代替ではないが最小構成として許容する。
      var email = Session.getActiveUser().getEmail();
      if (email) {
        // SHA-256 相当の簡易ハッシュ化は GAS では使えないので、
        // Utilities.computeDigest で MD5 ハッシュを生成してエンコードする。
        var digest = Utilities.computeDigest(
          Utilities.DigestAlgorithm.MD5,
          email,
          Utilities.Charset.UTF_8
        );
        userKey = digest
          .map(function(b) { return ("0" + (b & 0xff).toString(16)).slice(-2); })
          .join("")
          .substring(0, 20);
      }
    }

    // 最後の手段: ランダムキー（同一ユーザーでもタブごとに別ユーザーとなる）
    if (!userKey) {
      userKey = "anon_" + Math.random().toString(36).substring(2, 14);
    }

    // 表示名: ユーザーキーの先頭 8 文字を使う（メールは露出しない）
    var displayName = "User-" + userKey.substring(0, 8);

    return {
      firebaseConfig: {
        apiKey:      apiKey,
        authDomain:  authDomain || (projectId + ".firebaseapp.com"),
        databaseURL: databaseURL,
        projectId:   projectId,
        appId:       appId
      },
      userKey:     userKey,
      displayName: displayName,
      roomId:      "default"
    };
  } catch (e) {
    throw e.message || String(e);
  }
}

/**
 * Script Properties に Firebase 設定値を一括登録する補助関数。
 * GAS エディタの「実行」メニューから直接実行して初期設定に使う。
 * 各値を自分の Firebase プロジェクトのものに書き換えること。
 */
function setupFirebaseProperties() {
  var props = PropertiesService.getScriptProperties();

  // ▼ ここを自分の Firebase プロジェクトの値に書き換えてください ▼
  props.setProperties({
    "FIREBASE_API_KEY":       "YOUR_API_KEY",
    "FIREBASE_AUTH_DOMAIN":   "YOUR_PROJECT_ID.firebaseapp.com",
    "FIREBASE_DATABASE_URL":  "https://YOUR_PROJECT_ID-default-rtdb.firebaseio.com",
    "FIREBASE_PROJECT_ID":    "YOUR_PROJECT_ID",
    "FIREBASE_APP_ID":        "YOUR_APP_ID"
  });
  // ▲ ここまで書き換えてください ▲

  SpreadsheetApp.getUi().alert(
    "Firebase Script Properties を設定しました。\n" +
    "setupFirebaseProperties() 内の値が正しいか確認してください。"
  );
}
