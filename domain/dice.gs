function rollString(string) {
    if (string == undefined) return string;

  //[]で囲まれた部分を除去
  var string = string.replace(/\[[^\[\]]*\]/,"");
  
  //まずダイス部分を抜き出してロール。元の文字列を数値に置換した文字列とダイス目の列挙を作る
  const regexp = RegExp('([0-9]{1,})[d|D]([0-9]{1,})','g');
  console.log('6:string:' + string);
  const matches = String(string).matchAll(regexp);

  var calcString = '';
  var resultString = '';
  var currentIndex = 0;
  var lastIndex = 0;
  var eyes = [];

  for (const match of matches) {
    var roll = XdYArray(match[1],match[2]);
    console.log("17 roll:" + roll);
    eyes = eyes.concat(roll.eyes);
    if (match.index > currentIndex) {
      calcString = calcString + string.slice(currentIndex, match.index);
      resultString = resultString // + string.slice(currentIndex, match.index)
    }
    calcString = calcString + roll.sum;
    resultString = resultString + " [" + roll.eyes.join(",") + "]";
    currentIndex = match.index + match[0].length;
  }

  if (currentIndex < string.length) {
    calcString = calcString + string.slice(currentIndex);
    resultString = resultString //+ string.slice(currentIndex);
  }

  return { 'rolledResultString':calcString,'joinEyesString':resultString ,"eyes":eyes };
}

function calcString(string) {
  if (string == undefined) return string;
  //演算  
  //数式 の定義「(連続するカッコ)があれば(+-)があれば、その後に数字が続けば」
  //  const regexpCalc = RegExp('([(]{0,}[+\-]{0,}[0-9][0-9()+*\/\-]{0,})','g');
  const regexpCalc = RegExp('([(]{0,}[+\-]{0,}[0-9][0-9()!=><%&|+*\/\-]{0,})','g');
  const matchesCalc = string.matchAll(regexpCalc);
  var currentIndexCalc = 0;
  var calcStringCalc = '';  

  for (const match of matchesCalc) {
    if (match.index > currentIndexCalc) {
      calcStringCalc = calcStringCalc + string.slice(currentIndexCalc, match.index);
    }
    try {
      calcStringCalc = calcStringCalc + Function('"use strict";return ('+match[0]+')')();
    } catch (err) {
      calcStringCalc = calcStringCalc + 'error';
    }
    currentIndexCalc = match.index + match[0].length;
  }

  if (currentIndexCalc < string.length) {
    calcStringCalc = calcStringCalc + string.slice(currentIndexCalc);
  }

  return calcStringCalc;
}

function XdYArray(x,y) {
  var sum = 0;
  var eyes = [];
  for (var i=1; i<=x; i++) {
    var eye = Math.floor( Math.random() * (y) ) + 1
    sum = sum + eye;
    eyes.push(eye);
  }

  console.log('eyes:' + eyes);
  return {'sum':sum,'eyes':eyes};
}

function XdY(x,y) {
  var sum = 0;
  for (var i=1; i<=x; i++) {
    sum = sum + Math.floor( Math.random() * (y) ) + 1 ;
  }
  return sum;
}

function botBase(command) {
//  return { 'rolledResultString':calcString,'joinEyesString':resultString ,"eyes":eyes };

  var rs = rollString(command);
  return calcString(rs.rolledResultString) + rs.joinEyesString ;
}
