function rqgDamage(command) {
  const pattern = /(?<weaponDamage>[1-9][0-9]{0,2}(?<weaponType>[TCBtcb])[1-9][0-9]{0,2})(?<addtionalDamage>[+-][1-9][0-9]*[A][1-9][0-9]*)?,(?<attackLv>[0-3])[Pp](?<parryLv>[0-3])(?<otherCommand>.*)/;
  //縦軸が攻撃成功度、横軸が防御成功度
  const damageArray = [
    ['0', '0', '0', '0'],
    ['N', 'N', '0', '0'],
    ['S', 'S', 'N', '0'],
    ['M', 'S', 'S', 'N']
  ];

    const match = command.match(pattern);
    
    if (match) {
        const {weaponDamage, weaponType, addtionalDamage, attackLv, parryLv, otherCommand} = match.groups;
        const damageLv =  damageArray[attackLv,parryLv];
        var damage = 0;

        if (damageLv == '0') {
          
        } else if (damageLv == 'N') {

        } else {
          if (weaponType=='C') {
            if (damageLv='S') {
              //効果なので武器ダメージ最大
            } else {
              //決定なので
            }
          } else if (weaponType=='B') {

          } else if (weaponType=='T') {

          } else {
            return command;
          }
        }
    } else {
      return command;
    }
}

function botRqg(command) {

  //[]で囲まれた部分を除去
  var command = command.replace(/\[[^\[\]]*\]/,"");
  const regexp = RegExp('([0-9]{1,})[d|D]([0-9]{1,})','g');
  console.log("2024/02/03 08:51 command:"+command);

  if (command.match(/[0-9]{1,}[d|D][0-9]{1,}/)) {
    console.log("throw botBase");
    return botBase(command);
  }

  console.log("botRqg-command:" + command);
  var resultMark = ['FB','失敗','成功','効果','決定']
  
  var roll = XdY(1,100);
  var result = "";

  //■倍率■
  var param = /(.*),\*([0-9]{1,})$/.exec(command);
  if (param != null) {
    console.log("botRqg倍率-param:" + param);
    //末尾に*Xなら倍率
    var rs = rollString(param[1]);
    var ability = calcString(rs.rolledResultString);
    var multipleResult = 0;
    if (ability != 0) multipleResult = Math.ceil(roll / ability)
    var lv = getRqgSuccessLV(Number(ability) * Number(param[2]), roll);
    result = resultMark[lv+1] + ':' + multipleResult + '倍 [' + roll + "]÷" + ability;
    console.log("botRqg倍率-result:" + result);
    return result
  }

  //■抵抗■
  param = /(.*),vs([0-9]{1,})$/.exec(command);
  if (param != null) {
    console.log("botRqg抵抗-param:" + JSON.stringify(param));
    //末尾にvsXなら抵抗
    var rs = rollString(param[1]);
    console.log("JSON.stringify(rs):" + JSON.stringify(rs));
    var ability = calcString(rs.rolledResultString);
    console.log("ability:" + ability);
    var targetValue = Number(param[2]);
    var resultScore = Number(ability) + Number(getRqgRegistBonus(roll));
    console.log("resultScore:" + resultScore);
    if (roll<=5) {
      result = '成功:' + targetValue + '>=' + targetValue + ' [' + roll + "]";
    } else if (roll>95) {
      result = '失敗:' + 0 + '>=' + targetValue + ' [' + roll + "]";
    } else if (resultScore >= targetValue) {
      result = '成功:' + resultScore + '>=' + targetValue + ' [' + roll + "]";
    } else {
      result = '失敗:' + resultScore + '>=' + targetValue + ' [' + roll + "]";
    }
    console.log("botRqg抵抗-result:" + result);
    return result;
  }

  param = command.split(/,/);
  console.log("botRqg通常-param[0]:" + param[0]);
  console.log("botRqg通常-param[1]:" + param[1]);

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
    var lv = getRqgSuccessLV(per,Number(roll));
    result = resultMark[lv+1] + ': [' + roll + ']<=' + per;
    return result;  
  }
  
  return command;
}


function getRqgSuccessLV(per,roll) {
  console.log("getRqgSuccessLV-per:" + per);
  if (per<5) per=5;
  if (roll == 100) return -1;
  if (roll == 99 && per<71) return -1;
  if (roll == 98 && per<51) return -1;
  if (roll == 97 && per<31) return -1;
  if (roll == 96 && per<11) return -1;
  if (roll >= 96) return 0;
  if (roll == 1) return 3;
  if (roll <= Math.round(per/20)) return 3;
  if (roll <= Math.round(per/5)) return 2;
  if (roll <= per) return 1;
  return 0;
}

function getRqgRegistBonus(roll) {
  console.log("getRqgRegistBonus-roll:" + roll);
  return 10-Math.ceil(roll/5);
}


function getAvsD(perA,rollA,perB,rollB) {
  let lvA =  getRqgSuccessLV(perA,rollA);
  let lvB =  getRqgSuccessLV(perA,rollA);
  //攻撃が失敗


}

//ノチェット便り
//http://news-from-nochet.blogspot.com/p/rqg.html

function getBattleResult(attackLevel, defenseLevel) {
    const resultTable = [
        ['無効','反撃1','反撃2','反撃2c'],
        ['攻撃1','拮抗1','反撃1','反撃2'],
        ['攻撃2','拮抗2','拮抗1','反撃1'],
        ['攻撃3','拮抗3','拮抗3c','拮抗1']
    ];

    return resultTable[attackLevel][defenseLevel];
}
function processResult(result, a_params, b_params) {

    result.forEach(function(element) {
      let a_OverWepHpDm, a_OverPartApDm, d_OverWepHpDm;
      switch (element) {
          case '無効':
              console.log('無効');
              break;
          case '攻撃1':
              a_OverPartApDm  = a_params.damage1 - b_params.partArmor;
              if (a_OverPartApDm > 0) b_params.partHP = b_params.partHP - a_OverPartApDm;
              console.log('攻撃1');
              break;
          case '攻撃2':
              a_OverPartApDm = a_params.damage2 - a_params.partArmor;
              if (a_OverPartApDm > 0) b_params.partHP = b_params.partHP - a_OverPartApDm;
              console.log('攻撃2');
              break;
          case '攻撃3':
              b_params.partHP = b_params.partHP - a_params.damage3;
              console.log('攻撃3');
              break;
          case '拮抗1':
              a_OverWepHpDm = a_params.damage1 - b_params.weaponHP;
              if (a_OverWepHpDm > 0) b_params.weaponHP = b_params.weaponHP - 1;
              a_OverPartApDm  = a_OverWepHpDm - b_params.partArmor;
              if (a_OverPartApDm > 0) b_params.partHP = b_params.partHP - a_OverPartApDm;
              console.log('攻撃を受け流しました。');
              break;
          case '拮抗2':
              a_OverWepHpDm = a_params.damage2 - b_params.weaponHP;
              if (a_OverWepHpDm > 0) b_params.weaponHP = b_params.weaponHP - a_OverWepHpDm;
              a_OverPartApDm = a_OverWepHpDm - a_params.partArmor;
              if (a_OverPartApDm > 0) b_params.partHP = b_params.partHP - a_OverPartApDm;
              console.log('攻撃を弱く受け流しました。');
              break;
          case '拮抗3':
              a_OverWepHpDm = a_params.damage3 - b_params.weaponHP;
              if (a_OverWepHpDm > 0) b_params.weaponHP = b_params.weaponHP - a_OverWepHpDm;
              a_OverPartApDm = a_OverWepHpDm - 0;
              if (a_OverPartApDm > 0) b_params.partHP = b_params.partHP - a_OverPartApDm;
              console.log('攻撃を弱く受け流しました。');
              break;
          case '拮抗3c':
              b_params.weaponHP = b_params.weaponHP - a_params.damage3;
              a_OverWepHpDm = a_params.damage3 - b_params.weaponHP;
              if (a_OverWepHpDm > 0) b_params.partHP = b_params.partHP - a_OverWepHpDm
              console.log('攻撃をとても弱く受け流しました。');
              break;
          case '反撃1':
              d_OverWepHpDm = b_params.damage1 - a_params.weaponHP;
              if (d_OverWepHpDm > 0) a_params.weaponHP = a_params.weaponHP - 1;
              console.log('反撃(通常)が発動しました。');
              break;
          case '反撃2':
              d_OverWepHpDm = b_params.damage2 - a_params.weaponHP;
              if (d_OverWepHpDm > 0) a_params.weaponHP = a_params.weaponHP - d_OverWepHpDm;
              console.log('反撃(強)が発動しました。');
              break;
          case '反撃2c':
              a_params.weaponHP = a_params.weaponHP - b_params.damage2;
              console.log('反撃(強強)が発動しました。');
              break;
          default:
              console.log('不明な結果です。');
              break;
      }
    });
}

