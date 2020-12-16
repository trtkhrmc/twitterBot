var ALL = '*';
var LIMIT_MATCHES = {'minute': '[0-5]?[0-9]', 'hour': '[01]?[0-9]|2[0-3]', 'day': '0?[1-9]|[1-2][0-9]|3[01]', 'month': '0?[1-9]|1[0-2]', 'week': '[0-6]'};
var MAX_NUMBERS   = {'minute': 59, 'hour': 23, 'day': 31, 'month': 12, 'week': 6}; // 本物のcronは曜日の7を日曜と判定するが、手間なのでmax6とする
var COLUMNS       = ['minute', 'hour', 'day', 'month', 'week'];

const TITTLE_ROW = 4;
const MARGIN_ROW = TITTLE_ROW -1;
const MARGIN_COL = 1;
const TWITTER_ID_COL = 3;
const CRON_MIN = 4;
const CRON_DAY_OF_WEEK = 8;
const FLUCTUATION = 9;
const RANDOM_REPEAT = 10;
const LAST_REQ_DATE = 12;
const LAST_FEEDBACK_DATA = 13;
const TWITTER_ACCESS_TOKEN = 15;
const TWITTER_ACCESS_TOKEN_SECRET = 16;
const TWITTER_EMAIL_COL = 17; //CronListの最後
const TWITTER_PASS_COL = 18; 
const ABS_2_CRON_LIST = -3;

function myFunction() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TOP');
  var cronList = getCronList(sheet);
  var currentTime = new Date();
  var times  = {
    'minute': Utilities.formatDate(currentTime, 'Asia/Tokyo', 'm'),
    'hour':   Utilities.formatDate(currentTime, 'Asia/Tokyo', 'H'),
    'day':    Utilities.formatDate(currentTime, 'Asia/Tokyo', 'd'),
    'month':  Utilities.formatDate(currentTime, 'Asia/Tokyo', 'M'),
    'week':   Utilities.formatDate(currentTime, 'Asia/Tokyo', 'u')
  };
  for (var i = 1; i < cronList.length; i++) { // スプレッドシートから取得した一行目(key:0)はラベルなので、key1から実行
    executeIfNeeded(cronList[i], times, sheet, i + TITTLE_ROW);
  }
}

// シートからCronの一覧を取得し配列で返す
function getCronList(sheet) {
  //var columnCVals = sheet.getRange('C4:C200').getValues();
  //var lastRow = columnCVals.filter(String).length;
  var lastRow = sheet.getRange(TITTLE_ROW, TWITTER_ID_COL).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() - MARGIN_ROW;
  return sheet.getRange(TITTLE_ROW, TWITTER_ID_COL, lastRow , TWITTER_EMAIL_COL).getValues(); 
}

// 実行すべきタイミングか判定し、必要であれば実行
function executeIfNeeded(cron, times, sheet, row) {
  for( var i=0; i<COLUMNS.length; i++ ){   // minute, hour, day, month, weekを順番にチェックして全て条件にマッチするようならcron実行
    var timeType = COLUMNS[i];
    
    var timingList = getTimingList(cron[i + 1], timeType); //先頭はtwitterアカウント名のため+1
    if (!isMatch(timingList, times[timeType])) {
      return false;
    } 
  }
  
  if (!(sheet.getRange(row, TWITTER_ACCESS_TOKEN).isBlank() )){ //token が設定されている場合
    if((new Date() - sheet.getRange(row, LAST_REQ_DATE).getValue())/1000/60 > 15){　//前回処理時から15分以上経過していない場合は処理しない(実行に揺らぎを与えているため)
      cronStClrFlg = postTweet(cron[TWITTER_ID_COL + ABS_2_CRON_LIST ],
                                    cron[TWITTER_ACCESS_TOKEN + ABS_2_CRON_LIST ],
                                    cron[TWITTER_ACCESS_TOKEN_SECRET + ABS_2_CRON_LIST],
                                    cron[RANDOM_REPEAT + ABS_2_CRON_LIST]
                                   ); //ツイート実行
      
      
      if (cronStClrFlg){               //順序投稿リピートなしでリスト終端までいった場合はcronStClrFlgから時刻のセッティング情報をクリアする
        sheet.getRange(row, CRON_MIN, 1 , CRON_DAY_OF_WEEK - CRON_MIN + 1).clearContent();
      }else{
        sheet.getRange(row, LAST_REQ_DATE).setValue(new Date()); // 最終リクエスト送信日時
      }
    }
  }

  //次回の実行タイミング(分)をゆらぎを与えて更新 （分セルはD列固定）
  if (cron[FLUCTUATION + ABS_2_CRON_LIST ] != 0){ 
    var commaSeparatedMinList = String(sheet.getRange(row, 4 ).getValue()).split(',');
    var minRand = Math.floor( Math.random() * 10 ) - 5;  //ゆらぎを与えるための乱数を設
    var limitPattern  = "(" + LIMIT_MATCHES['minute'] + ")"; 
    var numReg   = new RegExp("^" + limitPattern + "$"); 
    var newMinList = [];
    commaSeparatedMinList.forEach(function(value){
      if(value.match(numReg) != null){
        if (toInt(value)  + minRand < 0 ){
          value = value + 60;
        }else if (toInt(value)  + minRand > 59){
          value = value - 60;
        }
        newMinList.push(toStr((toInt(value)  + minRand) % 59));
      }else{
        newMinList.push(value);
      }
      sheet.getRange(row, 4).setValue(newMinList.join(','));
    });
  }  
}

// 中身が*もしくは指定した数字を含んでいるか
function isMatch(timingList, time) {
  return (timingList[0] === ALL || timingList.indexOf(time) !== -1);
}

// 文字列から数字のリストを返す timingListの作成
function getTimingList(timingStr, type) {
  var timingList = [];
  if (timingStr === ALL) { // * の時はそのまま配列にして返す
    timingList.push(timingStr);
    return timingList;
  }

  var limitPattern  = "(" + LIMIT_MATCHES[type] + ")";
  var numReg   = new RegExp("^" + limitPattern + "$");                      // 単一指定パターン ex) 2
  var rangeReg = new RegExp("^" + limitPattern + "-" + limitPattern + "$"); // 範囲指定パターン ex) 1-5
  var devReg   = new RegExp("^\\*\/" + limitPattern + "$");                 // 間隔指定パターン ex) */10
  var commaSeparatedList = String(timingStr).split(','); // 共存指定パターン ex) 1,3-5
    
  commaSeparatedList.forEach(function(value) {
  if (match = value.match(numReg)) { // 単一指定パターンにマッチしたら配列に追加
    timingList.push(toStr(match[1]));
  } else if ((match = value.match(rangeReg)) && toInt(match[1]) < toInt(match[2])) { // 範囲指定パターンにマッチしたら配列に追加
    for (var i = toInt(match[1]); i <= toInt(match[2]); i++) {
      timingList.push(toStr(i));
    }
  } else if ((match = value.match(devReg)) && toInt(match[1]) <= MAX_NUMBERS[type]) { // 間隔指定パターンにマッチしたら配列に追加
    var start = (type == 'day' || type == 'month') ? 1 : 0; // 月と日だけ0が存在しないので1からカウントする
    for (var i = start; i <= MAX_NUMBERS[type] / match[1]; i++) {
      timingList.push(toStr(i * match[1]));
    }
  }
  });
  return timingList;
}

// ifやforの判定を正しく行う為に文字列を10進数int型に変換
function toInt(num) {
  return parseInt(num, 10);
}
// 数値を10進数int型にして文字列に変換。実行タイミング一致判定（indexOf）で必要
function toStr(num) {
  return toInt(num).toFixed();
}
//////////////////////////////////////////////////////////

//投稿日時桁数調整
var toDoubleDigits = function(num) {
  num += "";
  if (num.length === 1) {
    num = "0" + num;
  }
 return num;     
}

function Getnow() {
  var date = new Date();
  var yyyy = date.getFullYear();
  var mm = toDoubleDigits(date.getMonth() + 1);
  var dd = toDoubleDigits(date.getDate());
  var hh = toDoubleDigits(date.getHours());
  var mi = toDoubleDigits(date.getMinutes()); 
  var ss = toDoubleDigits(date.getSeconds()); 
  return yyyy + '/' + mm + '/' + dd + ' ' + hh + ':' + mi + ':' + ss;
}

function replaceTag(doc){
  var date = new Date();
  var yyyy = date.getFullYear();
  var mm = toDoubleDigits(date.getMonth() + 1);
  var dd = toDoubleDigits(date.getDate());
  var hh = toDoubleDigits(date.getHours());
  var mi = toDoubleDigits(date.getMinutes()); 
  var ss = toDoubleDigits(date.getSeconds()); 
  var rd = toStr(Math.round(Math.random()*100));
  
  var repList = {'{year}': yyyy , '{month}': mm, '{day}': dd, '{hour}': hh, '{minute}': mi , '{second}':ss ,'{rand}':('00'+rd).slice(-2)};
  for(var key in repList){
    reg = new RegExp(key,'g');
    doc = doc.replace(reg , repList[key]);
  }
  return doc;
  
}



///////////tweet内容の設定と実行///////////  !!!ツイート内容のシート名はIDと同一にすること!!!
function postTweet(id　,twitter_access_token , twitter_access_token_secret,random_repeat) {
 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(id);
  if (sheet){ //シートがある場合
    const max_row =sheet.getRange(1, 3).getValue();
    if (random_repeat == 0 ){ //ランダム選択時
      var post_num = Math.round( Math.random() * (max_row -1) ) + 1;  //投稿ツイートをランダムに選択
      if (post_num==0){
        sheet.getRange(2,3).setValue(1);
        post_num = 1;
      }else{
        sheet.getRange(2,3).setValue(post_num);//最終投稿文章Noを転記
      } 
    }else{ 　　　　　　　　　　　//リピート選択時(1 or 2)
      var post_num = (sheet.getRange(2,3).getValue() + 1);
      if (post_num > max_row){
        
        if(random_repeat ==2 ){
          //時刻の設定情報をクリアするフラグを返す
          return true;
        }    
        post_num = (sheet.getRange(2,3).getValue() + 1) % max_row;
      }
      sheet.getRange(2,3).setValue(post_num); 　               //最終投稿文章Noを転記
    }
    
    var body = sheet.getRange(6 + post_num, 3).getValue();

  }else{ //シートがない場合
    var body = '時報：ただいまの時刻' + Getnow() + 'をお知らせします。'
  }  

  twitterPostApi(id , twitter_access_token , twitter_access_token_secret, body);
  console.log(body);
  return false;
}

///////////Post to ApiGateWay ///////////
function twitterPostApi(id, twitter_access_token, twitter_access_token_secret ,body ) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TOP');
  var data = {
    "usr": id,
    "message":replaceTag(body),
    "access_token":twitter_access_token, //
    "access_token_secret": twitter_access_token_secret
  };  

  var options = {
    "method":"POST",
    "headers": {
      "Content-Type":"application/json"
    },
    "payload":JSON.stringify(data)
  };  
  UrlFetchApp.fetch("https://1oshi.work/twtbot/post.php", options);
  
}

///////////正常終了時のPOST結果の受信///////////
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TOP');
  var jsonString = e.postData.getDataAsString();
  var data = JSON.parse(jsonString);
  var value1 = data.usr; //USERID
  var cronList = getCronList(sheet);
  var idReg   = new RegExp("^"+ value1); 
  for(var j = 1; j < cronList.length; j++) {
    if(cronList[j][0].match(idReg) != null){
      break;
    }
  }
  if (j != cronList.length){
    if ( data.access_oauth_token){
      sheet.getRange(j + TITTLE_ROW ,TWITTER_ACCESS_TOKEN).setValue(data.access_oauth_token);
      sheet.getRange(j + TITTLE_ROW ,TWITTER_ACCESS_TOKEN_SECRET).setValue(data.access_oauth_token_secret); 
    }
    else{
      sheet.getRange(j + TITTLE_ROW ,LAST_FEEDBACK_DATA).setValue(new Date()); // Twittter投稿完了フィードバック受信日時
    }
  }
}

