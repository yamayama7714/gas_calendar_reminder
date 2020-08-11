/***********************************************************************************
使用方法はこちらです
https://docs.google.com/presentation/d/1HwuOPFMc7HEpGViNZe_yPtayH1P7OJSqjQLgaOsNnTU/edit#slide=id.g80d94ce47f_0_77
***********************************************************************************/


function getCalendarOption() {
  // スプシから各種設定情報を取得
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sh.getRange("C2:E2").getDisplayValues(); //通知したいイベントに含まれる文字列。数字のみでも可。日付等は後で弾く
  const beforeMin = sh.getRange("C4:E4").getValues(); // 何分前に通知するか
  const mail = Session.getActiveUser().getEmail();

  const dataObj = {
    mail : mail,
    rec  : sh.getRange("B2").getValue(), //boolean
    str1 : [data[0][0], beforeMin[0][0]],
    str2 : [data[0][1], beforeMin[0][1]],
    str3 : [data[0][2], beforeMin[0][2]]
  };
  Logger.log(dataObj);
  //  Logger.log(dataObj.str2[1]);
  return dataObj;

}


function setReminder() {

  try  {
    // スプシ入力済みの値のエラーチェック
    const err = validCheck();
    if (err) { throw new Error(err) };
    const data = getCalendarOption();
    const myCal = CalendarApp.getCalendarById(data.mail);
    
    //向こう4週間のイベントリスト
    const date = new Date();
    const startTime = new Date(date.getFullYear(),date.getMonth(), date.getDate(), 0, 0, 0);
    const in4weeks = new Date(startTime.getTime() + (28 * 24 * 60 * 60 * 1000));
    const myEvents = myCal.getEvents(startTime, in4weeks);
    if (!myEvents || myEvents.length === 0) { throw new Error("カレンダーに予定がありません…！もしかしたら、良いことかもしれませんね。") };

    // 定期イベントについて、送信するか否か(bool)。
    const recur = data.rec;
    if (typeof recur !== "boolean") { throw new Error("B2セルの形式が不適切と考えられます。B2セルはチェックボックスにしてください。\nB2セルをアクティブにした後、メニューバーの「挿入」から「チェックボックス」をクリックしてください。"); };

    //data.strXは[送信対象,通知時間]の配列。何れか片方が空白なものは除外
    const sendTargetTitles = [data.str1, data.str2, data.str3].filter(ds => { return  ds[0] && ds[1] });
    Logger.log(sendTargetTitles);

    //送信対象を抽出する　（不参加でない　かつ　イベント（カレンダーの予定）タイトルにスプシで指定した文字列を含む）
    let sendEvents;

    if (recur) {
      console.log("定期予定を含むイベントを抽出");
      sendEvents = myEvents.filter(mev => {
        return mev.getMyStatus() !== CalendarApp.GuestStatus.NO && // 不参加でない
        sendTargetTitles.some(st => mev.getTitle().includes(st[0])); //イベント（カレンダーの予定）のタイトルに指定文字列が存在する
      });
    } else {
      console.log("定期予定を除外");
      sendEvents = myEvents.filter(mev => {
        return mev.getMyStatus() != CalendarApp.GuestStatus.NO &&
        mev.isRecurringEvent() === false && //定期予定ではない
        sendTargetTitles.some(st => mev.getTitle().includes(st[0]));
      });
    }

    // 送信対象イベントのみに送信する
    sendEvents.forEach(sendEvent => {
    // 通知を送信する時間のこと（何分前に通知するかの「分」のこと。HH:mm:ssのmm）
      const beforeMin = sendTargetTitles.filter(st => { return sendEvent.getTitle().includes(st[0]) });
      sendEvent.addEmailReminder(beforeMin[0][1]); //bebforeMinはdataObjのstr1~3と同じ
    });

   } catch(e) {
     sendErrorMailToMe(Session.getActiveUser().getEmail(),e);
     Logger.log(e);
//     Browser.msgBox(e);
  }

}


function sendErrorMailToMe(mailAd,errMsg) {
  const fromAndTo = mailAd;
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
