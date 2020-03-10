/***********************************************************************************
使用方法はこちらです
https://docs.google.com/presentation/d/1HwuOPFMc7HEpGViNZe_yPtayH1P7OJSqjQLgaOsNnTU/edit#slide=id.g80d94ce47f_0_77
***********************************************************************************/


function getCalenderOption() {
  // スプシから各種設定情報を取得
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sh.getRange("B2:E2").getValues();
  const beforeMin = sh.getRange("C4:E4").getValues(); // 送信時間 
  const mail = Session.getActiveUser().getEmail(); // メアドは入力不要
  
  const dataObj = {
    mail : mail,
    rec : data[0][0],
    str1 : [data[0][1], beforeMin[0][0]],
    str2 : [data[0][2], beforeMin[0][1]],
    str3 : [data[0][3], beforeMin[0][2]]
  }; 
  Logger.log(dataObj);
  //  Logger.log(dataObj.str2[1]);
  return dataObj;
  
}


function setReminder() {
  
  try  {
    // スプシ入力済みの値のエラーチェック
    const err = validCheck();
    if (err == Error) { throw new Error("C2,D2,E2セルの何れかに不適切な記号が入力されているか\nC4,D4,E4の何れかに不適切な数値が入力されています。") };
    const data = getCalenderOption();
    const myCal = CalendarApp.getCalendarById(data.mail);
    
    //向こう一週間のイベントリスト
    const startTime = new Date();
    const oneWeekAgo = new Date(startTime.getTime() + (7 * 24 * 60 * 60 * 1000));
    const myEvents = myCal.getEvents(startTime, oneWeekAgo);
    //  当日のみ：const myEvents = myCal.getEventsForDay(new Date());
    if (!myEvents || myEvents.length === 0) { throw new Error("カレンダーに予定がありません…！もしかしたら、良いことかもしれませんね。") };
    
    // 定期イベントについて、送信するか否か(bool)。
    const recur = data.rec;
    Logger.log(recur);
    
    // C2D2E2セルのうち何れかが空欄の場合、全ての予定に通知が送信されてしまう。
    // false,null,undefined,""等では対処できなかったため、不本意ではあるがユーザが使用することは有り得ないであろう正規表現を生成
    // これらの正規表現はtestメソッドで使用します
    const reg1 = (!data.str1[0]) ? new RegExp("/\/\/\/\/\/\/\/\/\a/") : new RegExp(data.str1[0]);
    const reg2 = (!data.str2[0]) ? new RegExp("/\/\/\/\/\/\/\/\/\a/") : new RegExp(data.str2[0]);
    const reg3 = (!data.str3[0]) ? new RegExp("/\/\/\/\/\/\/\/\/\a/") : new RegExp(data.str3[0]);
    //  Logger.log(reg1)
    //  Logger.log(reg2)
    //  Logger.log(reg3)
    
    switch (recur) {
      case true :
        
      Logger.log("**********定期イベントにも送信する**********")
      
      /*************************************************************************************************
      //なぜswitchなのか
      //本来は、recurの値（bool,定期予定か否か)に応じて呼び出すメソッドを変える、ようにしたほうがよいのかもしれません。
      //しかしその場合、実行時間が3倍くらいになります…呼び出す関数が増えるとGASは基本長引きます
      //単に視認性を向上させて行数を削減するのならば、1つのforEachの中で都度recurを判断させるのも手ですし、仕様上？速度も早いのですが、recurは不変の定数であり、都度条件判断させる必要はありません。
      //どうしようもないので、この書き方です。
      **************************************************************************************************/
      
      myEvents.forEach(mev => {    
        // 参加状況,タイトル
        const status = mev.getMyStatus();
        const title = mev.getTitle();
        
        // 現在参照中のイベントmevのタイトルの中に指定文字列reg1~3が含まれていれば、そのイベントには通知メールを送信
        if (reg1.test(title) || reg2.test(title) || reg3.test(title)) {
          // 指定文字列reg1=data.str[0]が含まれていたら、その文字列に応じた指定待ち時間=data.str[1]をbeforeMinに代入
          const beforeMin = (reg1.test(title)) ? data.str1[1]
                          : (reg2.test(title)) ? data.str2[1]
                          : (reg3.test(title)) ? data.str3[1]
                          : 30;
          // 参加拒否以外に送信
          if (status != CalendarApp.GuestStatus.NO){
            mev.addEmailReminder(beforeMin);
            Logger.log("送信したイベントのタイトル：" + title);
            Logger.log("リマインドメール送信タイミング: " + beforeMin + "分前");     
          };
          Logger.log("****************一巡****************");
        };
      });
        
        break;
        
      case false:
    
        Logger.log("**********定期イベントには送信しない**********")
        myEvents.forEach(mev => {    
          // 参加状況,タイトル
          const status = mev.getMyStatus();
          const title = mev.getTitle();
          Logger.log(title);
        
          if (reg1.test(title) || reg2.test(title) || reg3.test(title)) {
            // 参加拒否以外に送信
            const beforeMin = (reg1.test(title)) ? data.str1[1]
                            : (reg2.test(title)) ? data.str2[1]
                            : (reg3.test(title)) ? data.str3[1]
                            : 30;

          // 参加拒否でない、かつ、定期予定でない(isReccuringEvent()==false)にのみ送信         
            if (status != CalendarApp.GuestStatus.NO && mev.isRecurringEvent() == false) {
              mev.addEmailReminder(beforeMin);
              Logger.log("送信したイベントのタイトル：" + title);
              Logger.log("リマインドメール送信タイミング" + beforeMin + "分前");
            };
           Logger.log("****************一巡****************");
          } 
        })
        
        break;
  
      default:
        throw new Error("B2セルの形式が不適切と考えられます。B2セルはチェックボックスにしてください。\nB2セルをアクティブにした後、メニューバーの「挿入」から「チェックボックス」をクリックしてください。");
    };
    
   } catch(e) { 
     errorMail(Session.getActiveUser().getEmail(),e);
     Logger.log(e);
//     Browser.msgBox(e);
  }
    

}

function errorMail(mail,errMsg) {  
  // 自分で自分にエラーメールを送信する
  const fromAndTo = mail;
  const strBody = "リマインドメールの送信プログラムが失敗しました。原因は下記が考えられます。\n\n" + errMsg;
   GmailApp.sendEmail(
    fromAndTo,
    "【RPA】カレンダーの予定へのリマインドプログラムが動作しませんでした",
    strBody,
    {
      from: fromAndTo
    }
  );
};


