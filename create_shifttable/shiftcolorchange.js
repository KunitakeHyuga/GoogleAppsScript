function setColorBasedOnShiftTime() {
  // スプレッドシートを開いて指定のシート名を取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // F7からF62およびG7からG62までのセル値をループで処理
  for (var row = 7; row <= 62; row++) {
    // 条件ごとに日時範囲を設定
    var timeRanges = [
      { start: "1899/12/30 16:30:00", end: "1899/12/31 01:00:00", range: "P" + row +":AF" + row },
      { start: "1899/12/30 19:00:00", end: "1899/12/31 1:00:00", range: "U" + row + ":AF" + row },
      { start: "1899/12/30 19:30:00", end: "1899/12/31 1:00:00", range: "V"+ row +":AF" + row },
      { start: "1899/12/30 20:00:00", end: "1899/12/31 1:00:00", range: "W"+ row +":AF" + row },
      { start: "1899/12/30 20:30:00", end: "1899/12/31 1:00:00", range: "X"+ row +":AF" + row },
      { start: "1899/12/30 21:00:00", end: "1899/12/31 1:00:00", range: "Y"+ row +":AF" + row },
      { start: "1899/12/30 21:30:00", end: "1899/12/31 1:00:00", range: "Z"+ row +":AF" + row },
      { start: "1899/12/30 22:00:00", end: "1899/12/31 1:00:00", range: "AA" + row +":AF" + row },
      { start: "1899/12/30 22:30:00", end: "1899/12/31 1:00:00", range: "AB" + row +":AF" + row },
      { start: "1899/12/30 18:00:00", end: "1899/12/31 00:00:00", range: "S" + row +":AD" + row },
      { start: "1899/12/30 18:30:00", end: "1899/12/31 00:30:00", range: "T" + row +":AE" + row },
      { start: "1899/12/30 18:30:00", end: "1899/12/31 00:30:00", range: "T" + row +":AE" + row },
      { start: "1899/12/30 17:00:00", end: "1899/12/30 23:00:00", range: "Q" + row +":AB" + row },
      { start: "1899/12/30 13:00:00", end: "1899/12/30 19:00:00", range: "I" + row +":T" + row },
      { start: "1899/12/30 13:00:00", end: "1899/12/30 17:00:00", range: "I" + row +":P" + row },
      { start: "1899/12/30 13:00:00", end: "1899/12/30 18:00:00", range: "I" + row +":R" + row },
      { start: "1899/12/30 11:30:00", end: "1899/12/30 17:30:00", range: "I" + row +":Q" + row },
      { start: "1899/12/30 16:30:00", end: "1899/12/31 01:00:00", range: "P" + row +":AF" + row }
    ];

    var startTime = sheet.getRange("F" + row).getValue();
    var endTime = sheet.getRange("G" + row).getValue();
    var mod9 = row % 9; // 行をmod 9で計算

    // 各条件をループで処理
    for (var j = 0; j < timeRanges.length; j++) {
      var specifiedStartTime = new Date(timeRanges[j].start);
      var specifiedEndTime = new Date(timeRanges[j].end);

      // 条件をチェックして背景色を設定
       // 条件をチェックして背景色を設定
       if (
        startTime instanceof Date &&
        endTime instanceof Date &&
        startTime.getTime() == specifiedStartTime.getTime() &&
        endTime.getTime() == specifiedEndTime.getTime()
       ) {
        if (mod9 === 7) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#9999ff'); // 明るい紫2
        } else if (mod9 === 8) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#00ff00'); // 緑
        } else if (mod9 === 0) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#4B77D1'); // 濃ゆい青
        } else if (mod9 === 1) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#FFFF00'); // 黄色
        } else if (mod9 === 2) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#00FFFF'); // シアン
        } else if (mod9 === 3) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#ff9900'); // オレンジ
        } else if (mod9 === 4) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#ea9999'); // 明るい赤2
        } else if (mod9 === 5) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#f9cb9c'); // 明るいオレンジ2
        } else if (mod9 === 6) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#FF00FF'); // マゼンタ
        }
      }
    }
  }
}

