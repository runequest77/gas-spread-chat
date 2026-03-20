function getUserNoFromRollColumn(col) {
  return col - rollColStart;
}

function resolveRollCommandContext(command, userNo, nameValue) {
  var paletteNo = Number(userNo);
  if (userNo != 0) {
    return { command: command, paletteNo: paletteNo };
  }

  var npcCommand = /^([0-9]{1,2}):(.*)/.exec(command);
  if (npcCommand) {
    return {
      command: npcCommand[2],
      paletteNo: Number(npcCommand[1]),
    };
  }

  var npcPaletteNo = getNpcPaletteNoFromName(nameValue);
  console.log("npcPaletteNo:" + npcPaletteNo);
  if (npcPaletteNo != null) paletteNo = npcPaletteNo;
  return { command: command, paletteNo: paletteNo };
}

function buildRollPc(userNo, paletteNo) {
  var palette = getPalette(paletteNo);  //★
  return {
    no: userNo,
    name: palette["cname"],
    elements: palette,
  };
}

function replaceRollCommand(command, elements) {
  var replacedCommand = replaceCommandKeyValue(command, elements);

  //値から取得した内容も一回だけ置き換え(値にKEYを設定して　例えば 数値-{ENC} のような設定が可能)
  return replaceCommandKeyValue(replacedCommand, elements);
}

function buildRollCommand(value, formulaValue) {
  var command = "";

  if (value == "/") {
    // 入力がスラッシュのみなら、判定式列からコマンド生成。
    var formulaString = replaceMultibyteForCalc(formulaValue);
    command = String(formulaString);
    if (formulaString == "") command = defaultCommand;
  } else {
    // 入力がスラッシュのみでなければ、判定式列は使用せず、そのままコマンドと見なす
    var suffixMatch = value.match(/^\/([+-/*].+)/);
    if (suffixMatch) {
      // スラッシュの後ろに +-*/ で始まる文字列がある場合
      var formulaString = replaceMultibyteForCalc(formulaValue);
      command = ("{" + formulaString + "}" || defaultCommand) + suffixMatch[1];
    } else {
      // それ以外は従来通りスラッシュを外してコマンド扱い
      command = String(value.slice(1));
    }
  }
  //コマンドの特定記号と数字を半角化
  return replaceMultibyteForCalc(command);
}

function replaceCommandKeyValue(command, keyValue) {
  console.log("replaceCommandKeyValue-command:" + command);

  const lines = command.split(/\r?\n/);
  for (let li = 0; li < lines.length; ++li) {
    let line = lines[li];

    var sepS = line.split(/\s/);  // スペース／タブなどで分割
    for (let i = 0; i < sepS.length; ++i) {
      if (sepS[i].length > 0) {
        var addBot = "";
        var sepC = sepS[i].split(/,/);  // カンマで分割
        for (let j = 0; j < sepC.length; ++j) {
          if (sepC[j].length > 0) {
            var parsedToken = parseAndFormatToken(sepC[j], keyValue, j, addBot);
            sepC[j] = parsedToken.token;
            addBot = parsedToken.addBot;
          }
        }
        if (/:/.test(sepC[0]) == false) {
          // 第一引数に:がなければaddBotを結合(""の場合Bot指定なし扱い)
          sepC[0] = addBot + sepC[0];
        }
        sepS[i] = sepC.join(",");
      }
    }

    lines[li] = sepS.join(" ");
  }

  let result = lines.join("\n");
  console.log("replaceCommandKeyValue-result:" + result);
  return result;
}

function parseAndFormatToken(token, keyValue, tokenIndex, addBot) {
  var nextToken = token;
  var nextAddBot = addBot;

  if (tokenIndex == 0) {
    // 配列の先頭のみBot判定あり
    console.log("replaceCommandKeyValue-sepS[i]:" + nextToken);
    var sepCN = nextToken.split(/:/, 2);
    console.log("replaceCommandKeyValue-sepCN:" + sepCN);
    if (sepCN.length == 2) {
      nextToken = sepCN[1];
      if (sepCN[0].length == 0) {
        nextAddBot = defaultBot + ":"; // :のみ指定なら既定のBot
      } else {
        nextAddBot = sepCN[0] + ":";   // Bot指定あり
      }
    }
    console.log("replaceCommandKeyValue-addBot:" + nextAddBot);
  }

  // {}で分割（{}は除去）
  var sepM = nextToken.split(/{([^{}]*)}/);
  console.log("replaceCommandKeyValue-sepM:" + sepM);
  for (let k = 0; k < sepM.length; ++k) {
    if (sepM[k].length > 0) {
      if (/[^0-9]+/.test(sepM[k])) {
        if (sepM[k] in keyValue) {
          if (nextAddBot == "") nextAddBot = defaultBot + ":"; // チャパレから取得した場合も既定Bot
          sepM[k] = "[" + sepM[k] + "]" + keyValue[sepM[k]];
        }
      }
    }
  }

  return {
    token: sepM.join(""),
    addBot: nextAddBot,
  };
}

function executeCommand(command, pc) {
  console.log("executeCommand-command:" + command);

  // 1) 改行で分割
  const lines = String(command).split(/\n/);
  console.log("lines.length:" + lines.length);

  // 2) 各行をスペースで分割→処理→スペースで結合
  const processedLines = lines.map((line) => {
    const tokens = line.trim().split(/\s+/).filter(Boolean);
    console.log("tokens.length:" + tokens.length);

    const processedTokens = tokens.map((tok) => {
      // コロンで分割して bot 指定の有無を判定
      let cmd = tok.split(':');
      if (cmd.length === 1) {
        cmd.unshift(defaultBot);
      }
      console.log("cmd:", cmd);
      const addResult = (joinResultChar) ? (joinResultChar + cmd[0] + ":" + cmd[1]) : "";
      cmd[1] = cmd[1].replace(/\[[^\[\]]*\]/, "");
      return botRoll(cmd[0], cmd[1]) + addResult;
    });

    return processedTokens.join(" ");
  });

  // 3) 改行で結合
  const out = processedLines.join("\n");
  console.log("executeCommand-result:" + out);
  return out;
}

function botRoll(key,command) {
//  console.log("key:" + key + " command:" + command);
  if (botKey[key] == undefined) return `ERROR:botKey[${key}]は登録されていません。省略名と関数名を間違えていませんか？`; 
  let botValue = botKey[key].split(/:/);
  var botname = botValue.shift();
  botValue.push(command);
  try {
    var result = Function('return this')()[botname](...botValue);
    return result;
  } catch(e) {
    console.log( e.message );
    return `ERROR:botKey{${key}:${botKey[key]}}に登録された関数が存在しない可能性が高いです。\n実際のエラーメッセージは` + e.message;
  }
}

function addDefaultBotCommand(command) {
  var sepSpace = command.split(/\s/);      //スペースでセパレート
  var sepConnma = sepSpace[0].split(/,/);  //第一ブロックをカンマでセパレート
  var sepColon = sepConnma[0].split(/:/);  //第一引数をコロンでセパレート
  if (sepColon.length == 1) {
    //コロンがついていない=既定のダイスボットを付加して{}で囲む
    sepColon[0] = defaultBot + ":{" + sepColon[0] + "}"
  } else if (sepColon[0]=="") {
    sepColon.shift();    //コロン先頭であればsepColon[0]を除去
  }
  sepConnma[0] = sepColon.join(":");
  sepSpace[0] = sepConnma.join(",");
  var result = sepSpace.join(" ");
  return result;
}

function replaceMultibyteForCalc(str) {
  if (!str) return str;
  return String(str).replace(/[Ａ-Ｚａ-ｚ０-９＋－／＊（）｛｝［］＜＞＝＆｜％！]/g, function(s) {
        return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });
}
