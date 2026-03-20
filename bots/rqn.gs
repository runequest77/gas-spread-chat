function botRq3(command) {

  //[]で囲まれた部分を除去
  var command = command.replace(/\[[^\[\]]*\]/,"");
  const regexp = RegExp('([0-9]{1,})[d|D]([0-9]{1,})','g');
  console.log("2022/05/05 22:27 command:"+command);

  if (command.match(/[0-9]{1,}[d|D][0-9]{1,}/)) {
    console.log("throw botBase");
    return botBase(command);
  }

  console.log("botRq3-command:" + command);
  var resultMark = ['FB-6','FB-5','FB-4','FB-3','FB-2','FB-1','失敗','成功','効果','決定','決定4','決定5']
  
  var roll = XdY(1,100);
  if (roll==77) return "★77";
  var result = "";

  //■倍率■
  var param = /(.*),\*([0-9]{1,})$/.exec(command);
  if (param != null) {
    console.log("botRq3倍率-param:" + param);
    //末尾に*Xなら倍率
    var rs = rollString(param[1]);
    var ability = calcString(rs.rolledResultString);
    var multipleResult = 0;
    if (ability != 0) multipleResult = Math.ceil(roll / ability)
    var lv = getRq3SuccessLV(Number(ability) * Number(param[2]), roll);
    result = resultMark[lv+6] + ':' + multipleResult + '倍 [' + roll + "]÷" + ability;
    console.log("botRq3倍率-result:" + result);
    return result
  }

  //■抵抗■
  param = /(.*),vs([0-9]{1,})$/.exec(command);
  if (param != null) {
    console.log("botRq3抵抗-param:" + JSON.stringify(param));
    //末尾にvsXなら抵抗
    var rs = rollString(param[1]);
    console.log("JSON.stringify(rs):" + JSON.stringify(rs));
    var ability = calcString(rs.rolledResultString);
    console.log("ability:" + ability);
    var targetValue = Number(param[2]);
    var resultScore = Number(ability) + Number(getRegistBonus(roll));
    console.log("resultScore:" + resultScore);
    if (resultScore >= targetValue) {
      result = '成功:' + resultScore + '>=' + targetValue + ' [' + roll + "]";
    } else {
      result = '失敗:' + resultScore + '>=' + targetValue + ' [' + roll + "]";
    }
    console.log("botRq3抵抗-result:" + result);
    return result;
  }

  param = command.split(/,/);
  console.log("botRq3通常-param[0]:" + param[0]);
  console.log("botRq3通常-param[1]:" + param[1]);

  var bonusText = "";
  var bonus = 0;
  if (param.length > 1) {
    bonusText = calcString(rollString(param[1]).rolledResultString);
    bonus = Number(bonusText);
  }

  //■通常■
  var rs = rollString(param[0]);
  console.log("rs.rolledResultString:" + rs.rolledResultString);
  var per = Math.ceil(Number(calcString(rs.rolledResultString)) + bonus) ; //技能成功率

  if (/^[0-9]+$/.test(per)) {
    var lv = getRq3SuccessLV(per,Number(roll));
    result = resultMark[lv+6] + ': [' + roll + ']<=' + per;
    return result;  
  }
  
  return command;
}


function getRq3SuccessLV(per,roll) {
  console.log("getRq3SuccessLV-per:" + per);
  if (per<5) per=5;
  if (roll == 100) return XdY(1,6) * -1;
  if (roll == 99 && per<70) return XdY(1,6) * -1;
  if (roll == 98 && per<50) return XdY(1,6) * -1;
  if (roll == 97 && per<30) return XdY(1,6) * -1;
  if (roll == 96 && per<15) return XdY(1,6) * -1;
  if (roll >= 96) return 0;
  if (roll == 1 && per >= 100) return 5;
  if (roll == 1 && per >= 20) return 4;
  if (roll == 1) return 3;
//  if (roll <= Math.round(per/20.0001)) return 3;
  if (roll <= Math.round(per/20)) return 3;
  if (roll <= Math.round(per/5)) return 2;
  if (roll <= Math.round(per)) return 1;
  return 0;
}

function getRegistBonus(roll) {
  console.log("getRegistBonus-roll:" + roll);
  if (roll == 1) return 99;
  if (roll <= 5) return 19;
  if (roll >= 96) return -19;
  if (roll == 100) return -99;
  return 10-Math.ceil(roll/5);
}

// function getD20LV(roll) {
//   if (roll == 1) return '-5';
//   if (roll <= 2) return '-4';
//   if (roll <= 3) return '-3';
//   if (roll <= 5) return '-2';
//   if (roll <= 8) return '-1';
//   if (roll <= 12) return '0';
//   if (roll <= 15) return '+1';
//   if (roll <= 17) return '+2';
//   if (roll <= 18) return '+3';
//   if (roll <= 19) return '+4';
//   if (roll <= 20) return '+5';
//   return '0';
// }
