/***********************************************************************************
使用方法はこちらです
https://docs.google.com/presentation/d/1HwuOPFMc7HEpGViNZe_yPtayH1P7OJSqjQLgaOsNnTU/edit#slide=id.g80d94ce47f_0_77
***********************************************************************************/
function getCalendarOption() {
    // スプシから各種設定情報を取得
    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sh.getRange("C2:E2").getDisplayValues(); //通知したいイベントに含まれる文字列。数字のみでも可。日付等は後で弾く
    const beforeMin = sh.getRange("C4:E4").getValues(); // 何分前に通知するか

    const dataObj = {
        mail : Session.getActiveUser().getEmail(),
        recur : sh.getRange("B2").getValue(), //boolean
        triggerDate : sh.getRange("B4").getValue(),
        titleAndTime1 : [data[0][0], beforeMin[0][0]], //[送信対象文字列,通知時間]
        titleAndTime2 : [data[0][1], beforeMin[0][1]],
        titleAndTime3 : [data[0][2], beforeMin[0][2]]
    };
    Logger.log(dataObj);
    //  Logger.log(dataObj.titleAndTime2[1]);
    return dataObj;

}
//トリガーに登録される関数はmain()
function main() {
    setReminder(true);
}

// 手動実行の場合(isAuto==false)にエラーした際は、エラーメールではなくエラーメッセをブラウザ上に表示する
function setReminder(isAuto) {
    try  {
        // スプシ入力済みの値のエラーチェック
        delTrigger(true);
        const err = validCheck();
        if (err) { throw new Error(err) }

        const calendarOptions = getCalendarOption();
        if (!calendarOptions.triggerDate || Object.prototype.toString.call(calendarOptions.triggerDate) !== '[object Date]') {
            throw new Error("B4セルには時間を入力してください");
        }
        const myCalendar = CalendarApp.getCalendarById(calendarOptions.mail);

        //向こう4週間のイベントリスト
        const date = new Date();
        const startTime = new Date(date.getFullYear(),date.getMonth(), date.getDate(), 0, 0, 0);
        const in4weeks = new Date(startTime.getTime() + (28 * 24 * 60 * 60 * 1000));
        const myEvents = myCalendar.getEvents(startTime, in4weeks);
        if (!myEvents || !myEvents.length) { throw new Error("カレンダーに予定がありません…！もしかしたら、良いことかもしれませんね。") }
        // 定期イベントについて、送信するか否か(bool)。
        const recur = calendarOptions.recur;
        if (typeof recur !== "boolean") { throw new Error("B2セルの形式が不適切と考えられます。B2セルはチェックボックスにしてください。\nB2セルをアクティブにした後、メニューバーの「挿入」から「チェックボックス」をクリックしてください。"); }

        //calendarOptions.titleAndTimeXは[送信対象,通知時間]の配列。何れか片方が空白なものは除外
        const sendTargetTitles = [calendarOptions.titleAndTime1, calendarOptions.titleAndTime2, calendarOptions.titleAndTime3].filter(ds => { return  ds[0] && ds[1] });
        Logger.log(sendTargetTitles);

        //送信対象を抽出する　（不参加でない　かつ　イベント（カレンダーの予定）タイトルにスプシで指定した文字列を含む）
        const sendEvents = (() => {
            if (recur) {
                console.log("定期予定を含むイベントを抽出");
                return myEvents.filter(mev => {
                    return mev.getMyStatus() !== CalendarApp.GuestStatus.NO && // 不参加でない
                    sendTargetTitles.some(st => mev.getTitle().includes(st[0])); //イベント（カレンダーの予定）のタイトルに指定文字列が存在する
                });
            } else {
                console.log("定期予定を除外");
                return myEvents.filter(mev => {
                    return mev.getMyStatus() != CalendarApp.GuestStatus.NO
                    && mev.isRecurringEvent() === false //定期予定でない
                    && sendTargetTitles.some(st => mev.getTitle().includes(st[0]));
                });
            }
        })();
        Logger.log('送信数 ' + sendEvents.length);
        // 送信対象イベントのみに送信する
        sendEvents.forEach(sendEvent => {
            // 通知を送信する時間のこと（何分前に通知するかの「分」のこと。HH:mm:ssのmm）
            const beforeMin = sendTargetTitles.filter(st => new RegExp(st[0]).test(sendEvent.getTitle()));
            Logger.log([sendEvent.getTitle(), beforeMin]);
            sendEvent.addEmailReminder(beforeMin[0][1]);
        });

        //　明日のトリガー
        setTrigger(calendarOptions.triggerDate, true);

    } catch(e) {
        if (!isAuto) { //手動実行時
            Browser.msgBox(e);
            return;
        }
        sendErrorMailToMe(Session.getActiveUser().getEmail(), e);
        Logger.log(e);
        //     Browser.msgBox(e);
    }

}


function setTrigger(argRunDate, isAuto) {
    //　トリガー重複禁止
    const runDate = (isAuto) ? argRunDate : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B4").getValue();
    if (Object.prototype.toString.call(runDate) !== '[object Date]') { 
        // 手動実行時にB4セルに時間が入力されていない場合実行される）
        Browser.msgBox('B4セルに時間を入力してください。');
        return;
    }

    const triggers = ScriptApp.getProjectTriggers();
    if (triggers.length) {
        delTrigger(true);
        if (!isAuto) Browser.msgBox("既に登録されているスケジュールを削除し、新しいスケジュールを登録します");
    }

    const triggerStartTime = (() => {
        const date = new Date();
        date.setHours(runDate.getHours());
        date.setMinutes(runDate.getMinutes());
        Logger.log(date);

        if (new Date().getTime() >= date.getTime()) {
            return new Date(date.setDate(date.getDate() + 1));
        } else {
            return date;
        }
    })();

    Logger.log("***newTrig")
    Logger.log(triggerStartTime);
    ScriptApp.newTrigger('main').timeBased().at(triggerStartTime).create();

    // 手動実行時は案内メッセ
    if (!isAuto) {
        // moment.JSなどは不使用。ファイルをコピーして使う場合、momentを手動で呼び出してもらう必要が出るため。
        const settedYMDhm = Utilities.formatDate(triggerStartTime, "JST", "yyyy/MM/dd (E) HH:mm");
        Browser.msgBox("下記の日時でトリガーが登録されました。\\n" + settedYMDhm);
    }
}

function delTrigger(isAuto) {
    const triggers = ScriptApp.getProjectTriggers();
    for(const trigger of triggers){
        if(trigger.getHandlerFunction() == "main"){
            ScriptApp.deleteTrigger(trigger);
        }
    }

    //手動入力時はメッセ表示
    if (!isAuto && !triggers.length) {
        Browser.msgBox("トリガーが登録されていませんでした");
    } else if (!isAuto && triggers.length) {
        const msg = "トリガー登録を解除しました\\nこれで自動実行されることはなくなります。\\n"
        + "再度トリガーを登録したい場合は、\\n「トリガー登録」ボタンを押してください";
        Browser.msgBox(msg);
    }

}

//自分で自分にエラーメール
function sendErrorMailToMe(myMailAddress, errMsg) {
    const strBody = "リマインドメールの送信プログラムが失敗しました。原因は下記が考えられます。\n\n" + errMsg;
        GmailApp.sendEmail(
            myMailAddress,
            "【RPA】カレンダーの予定へのリマインドプログラムが動作しませんでした",
            strBody,
            {
                from: myMailAddress
            }
    );
    // 明日のトリガーをを登録
    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const trigDate = sh.getRange("B4").getValue();
    setTrigger(trigDate, true);
}

// 説明を表示ボタン押下時起動。バッククォートで囲う場合、以下の問題が発生するためこの記法。
// 1. インデントが値に含まれてしまう
// 2. コードエディタ上の見栄えが…
function showQuickManual() {
    const msg = "ご利用ありがとうございます。\\n\\n"
    + "■これは何？「向こう4週間以内のGoogleカレンダーの予定のうち、任意の文字が含まれる予定の開始X分前に、通知メールを送信するプログラムです」"
    + "※X分は、4行目C~E列に入力された時間となります。\\n\\n"
    + "■使い方は？　このスプシをコピー後、A列のボタンを押すと使えます。\\n\\n"
    + "■ボタンの説明\\n"
    + "・「通知が送信されるようにする」 : 初回実行時に押してください。認証後、もう1度押してください。\\n"
    + "※認証方法はF列のスライドに記載してあります。\\n"
    + "・「トリガー登録」：　B4セルに新たな時間を入力した場合、稼働時間が変わります。\\n"
    + "・「トリガー登録解除」：　毎日の自動実行を停止できます。\\n\\n"
    + "その他詳細は、F列のマニュアルをご確認ください。";

    Browser.msgBox(msg);
}