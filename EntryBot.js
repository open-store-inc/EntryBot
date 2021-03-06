function sendMessage(){
  

  //アクティブシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet()

  

  //シート名[メルペイ・auPAY]を取得
  const sheet = ss.getSheetByName("メルペイ・auPAY")

  //シートURL取得
  const pageUrl = ss.getUrl();

  //接続先URL
  const postUrl = "https://test"
  
  
  //最終行を取得する
  const lastRow = sheet.getLastRow();

  //先頭のメッセージ
  var msgText = "下記お客様の対応が終了しておりません。速やかに対応を行ってください。\n" + pageUrl + "\n";

  var closerBuff = [];
  var submitCloser= {};

  

  //カウンター
  let k = 0;

  //下記のforで１行ずつ処理を行う。
  for(let i = 2; i <= lastRow; i++) {
    //セルのカラーを取得
    var sellColor = sheet.getRange(i, 23).getBackground();
    
    //ハッチングされていなければ実行
    if(sellColor == "#ffffff"){

      //メルペイの状況を取得
      var melpaySt = sheet.getRange(i,23).getValue();

      //決済状態のステータスを取得
      var settlementSt = sheet.getRange(i,26).getValue();

      //メルペイの状況と決済状態のステータスが空の時に実施
      if(melpayStCheck(melpaySt) && settlementStCheck(settlementSt)){
        //現在のシステム時刻を取得
        var nowDay = Moment.moment();

        //その行の日付を取得
        var inputDay = sheet.getRange(i, 1).getValue();

        //yyyy/MM/ddの形式に変更
        var inpuDayChangeBefor = Utilities.formatDate(inputDay,"JST", "yyyy/MM/dd");
        var inpuDayChangeAfter = Moment.moment(inpuDayChangeBefor,"YYYY年M月D日");

        //日付より45日をすぎていたら下記を実施
        if(nowDay.diff(inpuDayChangeAfter, "days") > 45){

          //クローザー取得
          var closer = sheet.getRange(i, 3).getValue();
        
          //テキストメッセージに追加
          // msgText = msgText + i + "行目   クローザー：" + closer + "   お客様名："  + sheet.getRange(i, 7).getValue() + "  " + sheet.getRange(i, 8).getValue() + "\n";

          //closerBuffになければ追加
          if (closerBuff.indexOf(closer) == -1){
            closerBuff.push(closer)
          }
          submitCloser[i] = closer;

          k = k + 1;
        }
      }
    }
  }

  //送信が必要な案件のクローザー分だけループ
  closerBuff.forEach( function(hitCloser) {
    msgText = msgText + "クローザー：　" + hitCloser + "\n";
    for (var key in submitCloser) {
      if(hitCloser == submitCloser[key]){
        //テキストメッセージに追加
        msgText = msgText + "   " + key + "行目" + "   お客様名："  + sheet.getRange(key, 7).getValue() + "  " + sheet.getRange(key, 8).getValue() + "\n";
      }
    }
  });

  
  if(msgText != "下記お客様の対応が終了しておりません。速やかに対応を行ってください。\n" + pageUrl + "\n"){
    const jsonData = {
      "text": msgText
    };
    const payload = JSON.stringify(jsonData);
    const options = {
      "channel": "#testchanel",
      "method": "post",
      "contentType": "application/json",
      "payload": payload
    };
    UrlFetchApp.fetch(postUrl, options);
  }
}

var melpayStCheck = function(sellWord){

  //以下のワードが含まれていたら送信を実行する
  let chackWords = ['通常エントリー','設置'];

  //からの時は送る
  if(sellWord == ''){
    return true;
  }

  //ただの日付の時は送らない
  if(typeof sellWord == 'object'){
    return false;
  }


  //含まれていたらlを足していく
  l = 0;
  chackWords.forEach( function(word) {
   if(sellWord.indexOf(word) != -1){
     l = l + 1
   }
  });
  //lが0の場合はchackWordsの文字が含まれていない
  if(l != 0){
    return true
  }

  return false;
}

var settlementStCheck = function(sellWord){

  //以下のワードが含まれていたら送信を実行する
  let chackWords = ['不備'];

  //からの時は送る
  if(sellWord == ''){
    return true;
  }

  //ただの日付の時は送らない
  if(typeof sellWord == 'object'){
    return false;
  }

  //含まれていたらlを足していく
  l = 0;
  chackWords.forEach( function(word) {
   if(sellWord.indexOf(word) != -1){
     l = l + 1
   }
  });
  //lが0の場合はchackWordsの文字が含まれていない
  if(l != 0){
    return true
  }

  return false;
}







