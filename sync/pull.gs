/**
 * GitHubから全階層のファイルを再帰的に取得し、自分自身を更新する
 */
function pullFromGitHubRecursive() {
  const user = "runequest77";
  const repo = "gas-spread-chat";
  const branch = "main";
  const githubToken = "ghp_XXXXXXXXXXXXXXXXXXXX";

  const rootUrl = `https://api.github.com/repos/${user}/${repo}/contents/?ref=${branch}`;
  const scriptId = ScriptApp.getScriptId();
  const gasToken = ScriptApp.getOAuthToken();

  const githubHeaders = {
    "Authorization": "token " + githubToken,
    "Accept": "application/vnd.github.v3+json"
  };

  // 1. GitHubから全ファイルを再帰的に取得
  const projectFiles = fetchFilesFromGitHub(rootUrl, githubHeaders);

  // --- 2. 現在のプロジェクトのマニフェスト(appsscript.json)を確実に取得して追加 ---
  const currentContent = JSON.parse(UrlFetchApp.fetch(
    `https://script.googleapis.com/v1/projects/${scriptId}/content`,
    { headers: { Authorization: "Bearer " + gasToken } }
  ).getContentText());
  
  // "appsscript.json" または "appsscript" という名前のJSONファイルを探す
  const manifest = currentContent.files.find(f => f.type === "JSON");

  if (manifest) {
    // APIが受け付ける形式に整形（名前を強制的に "appsscript" にする）
    projectFiles.push({
      name: "appsscript",
      type: "JSON",
      source: manifest.source
    });
  } else {
    throw new Error("現在のプロジェクトにマニフェストファイルが見つかりません。");
  }

  const manifest = currentContent.files.find(f => f.name === "appsscript.json");
  if (manifest) projectFiles.push(manifest);

  // 3. Apps Script APIで自分自身を更新
  const updateUrl = `https://script.googleapis.com/v1/projects/${scriptId}/content`;
  const options = {
    method: "put",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + gasToken },
    payload: JSON.stringify({ files: projectFiles }),
    muteHttpExceptions: true
  };

  const result = UrlFetchApp.fetch(updateUrl, options);
  if (result.getResponseCode() === 200) {
    console.log("Pull完了。エディタをリロードしてください。");
  } else {
    console.error("更新エラー: " + result.getContentText());
  }
}

/**
 * 再帰関数（API経由で中身を取得する安定版）
 */
function fetchFilesFromGitHub(url, headers, allFiles = []) {
  const response = UrlFetchApp.fetch(url, { headers: headers });
  const items = JSON.parse(response.getContentText());

  items.forEach(item => {
    if (item.type === "dir") {
      fetchFilesFromGitHub(item.url, headers, allFiles);
    } else if (item.type === "file") {
      if ((item.name.endsWith('.gs') || item.name.endsWith('.html')) && item.name !== "appsscript.json") {
        
        // --- 修正箇所：直接ダウンロードURLではなく、APIから詳細(Content)を取得 ---
        const fileResponse = UrlFetchApp.fetch(item.url, { headers: headers });
        const fileData = JSON.parse(fileResponse.getContentText());
        
        // Base64で返ってくるのでデコードする（改行が含まれることがあるので除去）
        const decodedContent = Utilities.newBlob(
          Utilities.base64Decode(fileData.content.replace(/\n/g, ''))
        ).getDataAsString();

        const type = item.name.endsWith('.gs') ? 'SERVER_JS' : 'HTML';
        const gasName = item.path.replace(/\.(gs|html)$/, "");

        allFiles.push({
          name: gasName,
          type: type,
          source: decodedContent
        });
        console.log("取得済み: " + gasName);
      }
    }
  });
  return allFiles;
}
