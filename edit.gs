function onEdit(e) {
 
  try {
//    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = e.range;
    const cell = range.getValue();
    
    
    // 不正な文字列が存在しないか
    if (range.getRow() === 2 && range.getColumn() > 2 && cell.match(/\\|\*|\+|\.|\?|\{|\}|\[|\]|\(|\)|\^|\$|\-|\||\//)) {
      // 不正な文字列一覧　正規表現化前　\*+.?{}[]()^$-|/
        Browser.msgBox("C2,D2,E2の何れかのセルに使用できない文字が入力されています");
    }
    
    // 不正な時間（分）が指定されていないか
    if (range.getRow() === 4 && range.getColumn() > 2 && (cell > 10080 || cell < 5) ) {
        Browser.msgBox("入力された数字（分）が適切ではありません。\nA3セルを参照してください");
    }
    
    // B2セルが破壊されていないか
    if (range.getRow() === 2 && range.getColumn() === 2 && typeof cell != "boolean" ) {
      Browser.msgBox("B2セルは、チェックボックスである必要があります。");
    };
      
  } catch(e) {
    Logger.log(e);
    Browser.msgBox(e);
  }
}

function validCheck() {
  // setReminderを起動時に実行。エラー（スプシへの不正な値の入力）があったら、setReminder内部でそれを通知するメールを自分に送信する
    let output = true;
    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const strs = sh.getRange('C2:E2').getValues();
    const mins = sh.getRange('C4:E4').getValues();
    const datas = strs[0].concat(mins[0]);
    Logger.log(datas);
    
    datas.forEach(dt => { 
    // アウトな文字列一覧　正規表現化前　\*+.?{}[]()^$-|/
    if (typeof dt === "string") {
      if (String(dt).match(/\\|\*|\+|\.|\?|\{|\}|\[|\]|\(|\)|\^|\$|\-|\||\//)) {
        Logger.log("str!str!");
        output = Error;
        return output;
      }
    } else if (typeof dt === "number") {
      if (dt > 10080 || dt < 5) {
        Logger.log("num!num!");
        output = Error;
        return output;
      }
    }
  })
  return output;
}