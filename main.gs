/***********************************************************************************
使用方法はこちらです
https://docs.google.com/presentation/d/1HwuOPFMc7HEpGViNZe_yPtayH1P7OJSqjQLgaOsNnTU/edit#slide=id.g80d94ce47f_0_77
***********************************************************************************/


function getCalenderOption() {
  // スプシから各種設定情報を取得
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sh.getRange("C2:E2").getDisplayValues(); //通知したいイベントに含まれる文字列。数字のみでも可。日付等は後で弾く
  const beforeMin = sh.getRange("C4:E4").getValues(); // 何分前に通知するか
  const mail = Session.getActiveUser().getEmail(); // メアドは入力不要
  
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
    if (typeof err !== "boolean") { throw new Error(err) };
    const data = getCalenderOption();
    const myCal = CalendarApp.getCalendarById(data.mail);
    
    //向こう4週間のイベントリスト
    const date = new Date();
    const startTime = new Date(date.getFullYear(),date.getMonth(), date.getDate(), 0, 0, 0);
    const in4weeks = new Date(startTime.getTime() + (28 * 24 * 60 * 60 * 1000));
    const myEvents = myCal.getEvents(startTime, in4weeks);
    if (!myEvents || myEvents.length === 0) { throw new Error("カレンダーに予定がありません…！もしかしたら、良いことかもしれませんね。") };
    
    // 定期イベントについて、送信するか否か(bool)。
    const recur = data.rec;
    if (typeof recur !== "boolean") { throw new Error("B2セルの形式が不適切と考えられます。B2セルはチェックボックスにしてください。\nB2セルをアクティブにした後、メニューバーの「挿入」から「チェックボックス」をクリックしてください。"); }
    
    //data.strXは[送信対象,通知時間]の配列。何れか片方が空白なものは除外
    const sendTargetTitles = [data.str1, data.str2, data.str3].filter(ds => { return  ds[0] && ds[1] });
    Logger.log(sendTargetTitles);

    //送信対象を抽出する　（不参加でない　かつ　イベント（カレンダーの予定）タイトルにスプシで指定した文字列を含む）
    let sendEvents;
    
    if (recur) {
    // 定期予定を含む　イベントを抽出
      console.log("定期予定を含む")
      sendEvents = myEvents.filter(mev => { 
        return mev.getMyStatus() !== CalendarApp.GuestStatus.NO && // 不参加でない
        sendTargetTitles.some(st => mev.getTitle().includes(st[0])) //イベント（カレンダーの予定）のタイトルに指定文字列が存在する
      });
    } else {
    // 定期予定を含まない　イベントを抽出
      console.log("定期予定を除外")
      sendEvents = myEvents.filter(mev => { 
        return mev.getMyStatus() != CalendarApp.GuestStatus.NO &&
        mev.isRecurringEvent() === false && //定期予定ではないものを抽出
        sendTargetTitles.some(st => mev.getTitle().includes(st[0])) 
      });
    }

    // 送信対象イベントのみに送信する
    sendEvents.forEach(sendEvent => {
    // 通知を送信する時間のこと（何分前に通知するかの「分」のこと。HH:mm:ssのmm）
      const beforeMin = sendTargetTitles.filter(st => { return sendEvent.getTitle().includes(st[0]) });
      sendEvent.addEmailReminder(beforeMin[0][1]); // 通知送信
    });
    
   } catch(e) { 
     errorMail(Session.getActiveUser().getEmail(),e);
     Logger.log(e);
//     Browser.msgBox(e);
  }
    
  console.log("完走！")
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

function validCheck() {
  // setReminderを起動時に実行。エラー（スプシへの不正な値の入力）があったら、setReminder内部でそれを通知するメールを自分に送信する
    let output = true;
    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const strs = sh.getRange('C2:E2').getValues();
    const mins = sh.getRange('C4:E4').getValues();
    const cols = ["C", "D", "E"]; //エラーメールにセル位置を書きたい
    const baseMsg = "エラー：  "
//    const datas = strs[0].concat(mins[0]);
//    Logger.log(datas);
    
    let cnt = 0;
    let row;
    let cellName;
    //2行目について、エラーチェック&エラーメッセを積み上げ

    for (const str of strs[0]) {
      row = 2;
      cellName = cols[cnt] + row + "セル";
      
      if ( !(typeof str === "string" || typeof str === "number") ) {
        output += baseMsg + cellName + "および2行目には文字列または数字を入力してください。\n";
      } else if  ((typeof str === "string" || typeof str === "number") && String(str).length < 3 && String(str) !== "") {
        output += baseMsg + cellName + "に入力された文字数が短すぎます。" + row + "行目は3文字以上で入力してください。\n";
      }
      cnt += 1;
    }
    
    cnt = 0;
    //4行目について、エラーチェック&エラーメッセを積み上げ
    for (const min of mins[0]) {
      row = 4;
      cellName = cols[cnt] + row + "セル";
      
      if (min !== "" && typeof min !== "number") {
        output += baseMsg + cellName +  "には数字を入力してください。\n";
      } else if (min !== "" && typeof min === "number" && (min > 40320 || 5 > min) ) {
        output += baseMsg + "指定した通知時間は無効です。" + cellName + "および" + row + "行目の数値は5～40320(分)で指定してください。\n";
      }
      cnt += 1;
    }
    
    if (output != true) {output = output.slice(4)}; // エラーが存在すると、頭4文字に"true"が含まれてしまう
    Logger.log(output);  
    return output;
}


