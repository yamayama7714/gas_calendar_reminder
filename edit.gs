/***********************************************************************************
不正な入力に対する警告を表示（表示するだけ。入力値が不正のままである場合はvalidCheck()によって対処
***********************************************************************************/
function onEdit(e) {
 
  try {
//    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = e.range;
    const cell = range.getValue();    
    
    // 指定文字列が短すぎないか。
    if (range.getRow() === 2 && range.getColumn() > 2 && !!cell) { //!!cellは cell !== ""のこと。セルを空白にした際にメッセが表示される事を回避するため、必要
      if (String(cell).length < 3 ) { Browser.msgBox("文字数が短すぎます。\n3文字以上で入力してください。") } ;
    }
    
    // 不正な時間（分）が指定されていないか
    if (range.getRow() === 4 && range.getColumn() > 2 && !!cell && ((cell > 40320 || 5 > cell ) || (typeof cell !== "number"))) {
      Browser.msgBox("入力された数字（分）が適切ではないか、数字以外が入力されています。\nB3セルを参照してください");
    }
    
    // B2セルが破壊されていないか
    if (range.getRow() === 2 && range.getColumn() === 2 && typeof cell != "boolean" ) {
      Browser.msgBox("B2セルは、チェックボックスである必要があります。");
    };
      
  } catch(e) {
    Logger.log(e);
    Browser.msgBox(e);
  }
};

function validCheck() {
  // setReminderを起動時に実行。エラー（スプシへの不正な値の入力）があったら、setReminderのtryを終了し、エラーを通知するメールを自分に送信する
    let erMsg = "";

    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const strs = sh.getRange('C2:E2').getValues();
    const mins = sh.getRange('C4:E4').getValues();
    const cols = ["C", "D", "E"]; //エラーメールにセル位置を書きたい
    const baseMsg = "エラー：  "
//    const datas = strs[0].concat(mins[0]);
//    Logger.log(datas);

// 2行目と4行目でエラーチェック
    const rows = [
      {
        rowNum: 2,
        rowRange: strs[0]
      },

      {
        rowNum: 4,
        rowRange: mins[0]
      }
    ];

    for (const row of rows) {
        let cnt = 0;

        for (let col of row.rowRange) {
            const cellName = cols[cnt] + row.rowNum + "セル";

            if (row.rowNum === 2) {
                if ( !(typeof col === "string" || typeof col === "number") ) {
                  erMsg  += baseMsg + cellName + "および2行目には文字列または数字を入力してください。\n";
                } else if  ((typeof col === "string" || typeof col === "number") && String(col).length < 3 && String(col) !== "") {
                  erMsg  += baseMsg + cellName + "に入力された文字数が短すぎます。" + row.rowNum + "行目は3文字以上で入力してください。\n";
                };

            } else if (row.rowNum === 4) {
                if (col !== "" && typeof col !== "number") {
                  erMsg += baseMsg + cellName +  "には数字を入力してください。\n";
                } else if (col !== "" && typeof col === "number" && (col > 40320 || 5 > col) ) {
                  erMsg += baseMsg + "指定した通知時間は無効です。" + cellName + "および" + row.rowNum + "行目の数値は5～40320(分)で指定してください。\n";
                };
            };
            cnt += 1;
        };

    };
    Logger.log("validCheck: " + erMsg);
    return erMsg;
}