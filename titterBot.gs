var ALL = '*';
var LIMIT_MATCHES = {'minute': '[0-5]?[0-9]', 'hour': '[01]?[0-9]|2[0-3]', 'day': '0?[1-9]|[1-2][0-9]|3[01]', 'month': '0?[1-9]|1[0-2]', 'week': '[0-6]'};
var MAX_NUMBERS   = {'minute': 59, 'hour': 23, 'day': 31, 'month': 12, 'week': 6}; // �{����cron�͗j����7����j�Ɣ��肷�邪�A��ԂȂ̂�max6�Ƃ���
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
const TWITTER_EMAIL_COL = 17; //CronList�̍Ō�
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
  for (var i = 1; i < cronList.length; i++) { // �X�v���b�h�V�[�g����擾������s��(key:0)�̓��x���Ȃ̂ŁAkey1������s
    executeIfNeeded(cronList[i], times, sheet, i + TITTLE_ROW);
  }
}

// �V�[�g����Cron�̈ꗗ���擾���z��ŕԂ�
function getCronList(sheet) {
  //var columnCVals = sheet.getRange('C4:C200').getValues();
  //var lastRow = columnCVals.filter(String).length;
  var lastRow = sheet.getRange(TITTLE_ROW, TWITTER_ID_COL).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow() - MARGIN_ROW;
  return sheet.getRange(TITTLE_ROW, TWITTER_ID_COL, lastRow , TWITTER_EMAIL_COL).getValues(); 
}

// ���s���ׂ��^�C�~���O�����肵�A�K�v�ł���Ύ��s
function executeIfNeeded(cron, times, sheet, row) {
  for( var i=0; i<COLUMNS.length; i++ ){   // minute, hour, day, month, week�����ԂɃ`�F�b�N���đS�ď����Ƀ}�b�`����悤�Ȃ�cron���s
    var timeType = COLUMNS[i];
    
    var timingList = getTimingList(cron[i + 1], timeType); //�擪��twitter�A�J�E���g���̂���+1
    if (!isMatch(timingList, times[timeType])) {
      return false;
    } 
  }
  
  if (!(sheet.getRange(row, TWITTER_ACCESS_TOKEN).isBlank() )){ //token ���ݒ肳��Ă���ꍇ
    if((new Date() - sheet.getRange(row, LAST_REQ_DATE).getValue())/1000/60 > 15){�@//�O�񏈗�������15���ȏ�o�߂��Ă��Ȃ��ꍇ�͏������Ȃ�(���s�ɗh�炬��^���Ă��邽��)
      cronStClrFlg = postTweet(cron[TWITTER_ID_COL + ABS_2_CRON_LIST ],
                                    cron[TWITTER_ACCESS_TOKEN + ABS_2_CRON_LIST ],
                                    cron[TWITTER_ACCESS_TOKEN_SECRET + ABS_2_CRON_LIST],
                                    cron[RANDOM_REPEAT + ABS_2_CRON_LIST]
                                   ); //�c�C�[�g���s
      
      
      if (cronStClrFlg){               //�������e���s�[�g�Ȃ��Ń��X�g�I�[�܂ł������ꍇ��cronStClrFlg���玞���̃Z�b�e�B���O�����N���A����
        sheet.getRange(row, CRON_MIN, 1 , CRON_DAY_OF_WEEK - CRON_MIN + 1).clearContent();
      }else{
        sheet.getRange(row, LAST_REQ_DATE).setValue(new Date()); // �ŏI���N�G�X�g���M����
      }
    }
  }

  //����̎��s�^�C�~���O(��)����炬��^���čX�V �i���Z����D��Œ�j
  if (cron[FLUCTUATION + ABS_2_CRON_LIST ] != 0){ 
    var commaSeparatedMinList = String(sheet.getRange(row, 4 ).getValue()).split(',');
    var minRand = Math.floor( Math.random() * 10 ) - 5;  //��炬��^���邽�߂̗������
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

// ���g��*�������͎w�肵���������܂�ł��邩
function isMatch(timingList, time) {
  return (timingList[0] === ALL || timingList.indexOf(time) !== -1);
}

// �����񂩂琔���̃��X�g��Ԃ� timingList�̍쐬
function getTimingList(timingStr, type) {
  var timingList = [];
  if (timingStr === ALL) { // * �̎��͂��̂܂ܔz��ɂ��ĕԂ�
    timingList.push(timingStr);
    return timingList;
  }

  var limitPattern  = "(" + LIMIT_MATCHES[type] + ")";
  var numReg   = new RegExp("^" + limitPattern + "$");                      // �P��w��p�^�[�� ex) 2
  var rangeReg = new RegExp("^" + limitPattern + "-" + limitPattern + "$"); // �͈͎w��p�^�[�� ex) 1-5
  var devReg   = new RegExp("^\\*\/" + limitPattern + "$");                 // �Ԋu�w��p�^�[�� ex) */10
  var commaSeparatedList = String(timingStr).split(','); // �����w��p�^�[�� ex) 1,3-5
    
  commaSeparatedList.forEach(function(value) {
  if (match = value.match(numReg)) { // �P��w��p�^�[���Ƀ}�b�`������z��ɒǉ�
    timingList.push(toStr(match[1]));
  } else if ((match = value.match(rangeReg)) && toInt(match[1]) < toInt(match[2])) { // �͈͎w��p�^�[���Ƀ}�b�`������z��ɒǉ�
    for (var i = toInt(match[1]); i <= toInt(match[2]); i++) {
      timingList.push(toStr(i));
    }
  } else if ((match = value.match(devReg)) && toInt(match[1]) <= MAX_NUMBERS[type]) { // �Ԋu�w��p�^�[���Ƀ}�b�`������z��ɒǉ�
    var start = (type == 'day' || type == 'month') ? 1 : 0; // ���Ɠ�����0�����݂��Ȃ��̂�1����J�E���g����
    for (var i = start; i <= MAX_NUMBERS[type] / match[1]; i++) {
      timingList.push(toStr(i * match[1]));
    }
  }
  });
  return timingList;
}

// if��for�̔���𐳂����s���ׂɕ������10�i��int�^�ɕϊ�
function toInt(num) {
  return parseInt(num, 10);
}
// ���l��10�i��int�^�ɂ��ĕ�����ɕϊ��B���s�^�C�~���O��v����iindexOf�j�ŕK�v
function toStr(num) {
  return toInt(num).toFixed();
}
//////////////////////////////////////////////////////////

//���e������������
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



///////////tweet���e�̐ݒ�Ǝ��s///////////  !!!�c�C�[�g���e�̃V�[�g����ID�Ɠ���ɂ��邱��!!!
function postTweet(id�@,twitter_access_token , twitter_access_token_secret,random_repeat) {
 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(id);
  if (sheet){ //�V�[�g������ꍇ
    const max_row =sheet.getRange(1, 3).getValue();
    if (random_repeat == 0 ){ //�����_���I����
      var post_num = Math.round( Math.random() * (max_row -1) ) + 1;  //���e�c�C�[�g�������_���ɑI��
      if (post_num==0){
        sheet.getRange(2,3).setValue(1);
        post_num = 1;
      }else{
        sheet.getRange(2,3).setValue(post_num);//�ŏI���e����No��]�L
      } 
    }else{ �@�@�@�@�@�@�@�@�@�@�@//���s�[�g�I����(1 or 2)
      var post_num = (sheet.getRange(2,3).getValue() + 1);
      if (post_num > max_row){
        
        if(random_repeat ==2 ){
          //�����̐ݒ�����N���A����t���O��Ԃ�
          return true;
        }    
        post_num = (sheet.getRange(2,3).getValue() + 1) % max_row;
      }
      sheet.getRange(2,3).setValue(post_num); �@               //�ŏI���e����No��]�L
    }
    
    var body = sheet.getRange(6 + post_num, 3).getValue();

  }else{ //�V�[�g���Ȃ��ꍇ
    var body = '����F�������܂̎���' + Getnow() + '�����m�点���܂��B'
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

///////////����I������POST���ʂ̎�M///////////
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
      sheet.getRange(j + TITTLE_ROW ,LAST_FEEDBACK_DATA).setValue(new Date()); // Twittter���e�����t�B�[�h�o�b�N��M����
    }
  }
}

