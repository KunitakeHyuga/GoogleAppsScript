function copySheetsAndRename() {
  const sheetName = "週コピー原紙"; // コピー元のシート名
  const startDate = new Date("2024-04-08"); // コピーを開始する日付（YYYY-MM-DD形式）
  const numWeeks = 2; // コピーする週数
  const daysPerWeek = 7; // 1週間の日数
  const numRowsPerWeek = 8; // 週ごとの行数

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalSheet = ss.getSheetByName(sheetName);

  for (let i = 0; i < numWeeks; i++) {
    // 新しいシートを作成
    const copiedSheet = originalSheet.copyTo(ss);

    // 新しいシートの名前を設定（週の範囲を含む日付）
    const newStartDate = new Date(startDate.getTime() + (i * daysPerWeek + 1) * 24 * 60 * 60 * 1000); // 指定日からi週間分（プラス1日）の加算
    const newEndDate = new Date(newStartDate.getTime() + (daysPerWeek - 1) * 24 * 60 * 60 * 1000); // 開始日から6日後が終了日
    const newName = Utilities.formatDate(newStartDate, ss.getSpreadsheetTimeZone(), "MM/dd") + "～" + Utilities.formatDate(newEndDate, ss.getSpreadsheetTimeZone(), "MM/dd");
    copiedSheet.setName(newName);

    // 生成したシートを左から6番目に移動
    ss.setActiveSheet(copiedSheet);
    ss.moveActiveSheet(6);



    // B7セルにstartDateを記入（日付のみを設定）
    const formattedStartDate = Utilities.formatDate(newStartDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    copiedSheet.getRange("B7").setValue(formattedStartDate);

    // シート内に関数を挿入
    const sheetMonth = Utilities.formatDate(newStartDate, ss.getSpreadsheetTimeZone(), "yyyy.MM");

    // F7:G14 に関数を挿入
    const nameF = ["重松入","松本入", "岩下入", "森入", "國武入", "中島入","徳永入","川野入"];
    const nameG = ["重松出","松本出", "岩下出", "森出", "國武出", "中島出","徳永出","川野出"];
    for (let rowIndex = 7; rowIndex <= 14; rowIndex++) {
      const formulaF = `=XLOOKUP("${nameF[rowIndex - 7]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B7 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
      const formulaG = `=XLOOKUP("${nameG[rowIndex - 7]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B7 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
      copiedSheet.getRange(rowIndex, 6).setValue(formulaF);
      copiedSheet.getRange(rowIndex, 7).setValue(formulaG);
    }
    // B15:G22 に関数を挿入（同様の処理を行う）
    for (let rowIndex = 15; rowIndex <= 22; rowIndex++) {
    // 新しい日付を計算
     const nextStartDate = new Date(newStartDate.getTime() + (1 * 24 * 60 * 60 * 1000)); // 現在の終了日から1日後
     const sheetMonth = Utilities.formatDate(nextStartDate, ss.getSpreadsheetTimeZone(), "yyyy.MM");
     const formulaF = `=XLOOKUP("${nameF[rowIndex - 15]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B15, '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     const formulaG = `=XLOOKUP("${nameG[rowIndex - 15]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B15 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     copiedSheet.getRange(rowIndex, 6).setValue(formulaF);
     copiedSheet.getRange(rowIndex, 7).setValue(formulaG);
    }
    // B23:G30 に関数を挿入（同様の処理を行う）
    for (let rowIndex = 23; rowIndex <= 30; rowIndex++) {
     const nextStartDate = new Date(newStartDate.getTime() + (2 * 24 * 60 * 60 * 1000)); // 現在の終了日から2日後
     const sheetMonth = Utilities.formatDate(nextStartDate, ss.getSpreadsheetTimeZone(), "yyyy.MM");
     const formulaF = `=XLOOKUP("${nameF[rowIndex - 23]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B23 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     const formulaG = `=XLOOKUP("${nameG[rowIndex - 23]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B23 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     copiedSheet.getRange(rowIndex, 6).setValue(formulaF);
     copiedSheet.getRange(rowIndex, 7).setValue(formulaG);
    }
    // B31:G38 に関数を挿入（同様の処理を行う）
    for (let rowIndex = 31; rowIndex <= 38; rowIndex++) {
     const nextStartDate = new Date(newStartDate.getTime() + (3 * 24 * 60 * 60 * 1000)); // 現在の終了日から3日後
     const sheetMonth = Utilities.formatDate(nextStartDate, ss.getSpreadsheetTimeZone(), "yyyy.MM");
     const formulaF = `=XLOOKUP("${nameF[rowIndex - 31]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B31 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     const formulaG = `=XLOOKUP("${nameG[rowIndex - 31]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B31 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     copiedSheet.getRange(rowIndex, 6).setValue(formulaF);
     copiedSheet.getRange(rowIndex, 7).setValue(formulaG);
    }
    // B39:G46 に関数を挿入（同様の処理を行う）
    for (let rowIndex = 39; rowIndex <= 46; rowIndex++) {
     const nextStartDate = new Date(newStartDate.getTime() + (4 * 24 * 60 * 60 * 1000)); // 現在の終了日から4日後
     const sheetMonth = Utilities.formatDate(nextStartDate, ss.getSpreadsheetTimeZone(), "yyyy.MM");
     const formulaF = `=XLOOKUP("${nameF[rowIndex - 39]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B39 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     const formulaG = `=XLOOKUP("${nameG[rowIndex - 39]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B39 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     copiedSheet.getRange(rowIndex, 6).setValue(formulaF);
     copiedSheet.getRange(rowIndex, 7).setValue(formulaG);
    }
    // B47:G54 に関数を挿入（同様の処理を行う）
    for (let rowIndex = 47; rowIndex <= 54; rowIndex++) {
     const nextStartDate = new Date(newStartDate.getTime() + (5 * 24 * 60 * 60 * 1000)); // 現在の終了日から5日後
     const sheetMonth = Utilities.formatDate(nextStartDate, ss.getSpreadsheetTimeZone(), "yyyy.MM");
     const formulaF = `=XLOOKUP("${nameF[rowIndex - 47]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B47, '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     const formulaG = `=XLOOKUP("${nameG[rowIndex - 47]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B47, '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     copiedSheet.getRange(rowIndex, 6).setValue(formulaF);
     copiedSheet.getRange(rowIndex, 7).setValue(formulaG);
    }
    // B55:G62 に関数を挿入（同様の処理を行う）
    for (let rowIndex = 55; rowIndex <= 62; rowIndex++) {
     const nextStartDate = new Date(newStartDate.getTime() + (6 * 24 * 60 * 60 * 1000)); // 現在の終了日から6日後
     const sheetMonth = Utilities.formatDate(nextStartDate, ss.getSpreadsheetTimeZone(), "yyyy.MM");
     const formulaF = `=XLOOKUP("${nameF[rowIndex - 55]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B55 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     const formulaG = `=XLOOKUP("${nameG[rowIndex - 55]}", '${sheetMonth}'!$AJ$2:$AJ$23, XLOOKUP(B55 , '${sheetMonth}'!$A$2:$AJ$2, '${sheetMonth}'!$A$2:$AJ$23))`;
     copiedSheet.getRange(rowIndex, 6).setValue(formulaF);
     copiedSheet.getRange(rowIndex, 7).setValue(formulaG);
    }

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
     var mod8 = row % 8; // 行をmod 8で計算

     // 各条件をループで処理
      for (var j = 0; j < timeRanges.length; j++) {
       var specifiedStartTime = new Date(timeRanges[j].start);
       var specifiedEndTime = new Date(timeRanges[j].end);

       // 条件をチェックして背景色を設定
       if (
        startTime instanceof Date &&
        endTime instanceof Date &&
        startTime.getTime() == specifiedStartTime.getTime() &&
        endTime.getTime() == specifiedEndTime.getTime()
       ) {
        if (mod8 === 7) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#9999ff'); // 明るい紫2
        } else if (mod8 === 0) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#f9cb9c'); // 明るいオレンジ2
        } else if (mod8 === 1) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#00ff00'); // 緑
        } else if (mod8 === 2) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#ea9999'); // 明るい赤2
        } else if (mod8 === 3) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#FFFF00'); // 黄色
        } else if (mod8 === 4) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#00FFFF'); // シアン
        } else if (mod8 === 5) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#ff9900'); // オレンジ
        } else if (mod8 === 6) {
          sheet.getRange(timeRanges[j].range.replace("11", row)).setBackground('#FF00FF'); // マゼンタ
        }
        }
     }
   }
  }
}