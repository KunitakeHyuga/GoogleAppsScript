function copySheetAndUpdate() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const originalSheetName = "月コピー原紙";
  const copiedSheetName = "コピー原紙のコピー";

  const currentDate = new Date();
  const year = currentDate.getFullYear();
  const month = currentDate.getMonth() + 2; // 翌月の月を取得するために+2している

  // シートをコピーして指定の場所に移動する
  const originalSheet = ss.getSheetByName(originalSheetName);
  const copiedSheet = originalSheet.copyTo(ss);
  copiedSheet.setName(copiedSheetName);
  ss.setActiveSheet(copiedSheet);
  ss.moveActiveSheet(4);

  // シート名を変更する
  const formattedMonth = (month < 10) ? `0${month}` : month;
  const newSheetName = `${year}.${formattedMonth}`;
  copiedSheet.setName(newSheetName);

  // B2セルに翌月1日の日付を書き込む
  const firstDayOfNextMonth = new Date(year, month - 1, 2);
  const formattedDate = Utilities.formatDate(firstDayOfNextMonth, ss.getSpreadsheetTimeZone(), "yyyy/MM/dd");
  copiedSheet.getRange("B2").setValue(formattedDate);

  // N1セルに翌月1日の日付を書き込み、シート見出しを変更する
  copiedSheet.getRange("N1").setValue(firstDayOfNextMonth);

  /// 固定シフト日に色をつける
  // 指定範囲を選択し、値を取得
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var range = spreadsheet.getRange("B10:AF23");
  var values = range.getValues();

  // 範囲内のセルに対してループ処理
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      // セルの値が空白でない場合、背景色を設定
      if (values[i][j] != "") {
        range.getCell(i+1, j+1).setBackground("cyan");
      }
    }
  }

  // 関数によって算出された値を貼り付ける
  var range2 = spreadsheet.getRange("B2:AF23");
  var values2 = range2.getValues();
  range2.setValues(values2);
}
