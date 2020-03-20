function onEdit(e) {
 
  try {
//    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = e.range;
    const cell = range.getValue();    
    
    // 指定文字列が短すぎないか。
    if (range.getRow() === 2 && range.getColumn() > 2 && cell !== "") {
      if (String(cell).length < 3 ) { Browser.msgBox("文字数が短すぎます。\n3文字以上で入力してください。") } ;
    }
    
    // 不正な時間（分）が指定されていないか
    if (range.getRow() === 4 && range.getColumn() > 2 && cell !== "" && ((cell > 40320 || 5 > cell ) || (typeof cell !== "number"))) {
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
}