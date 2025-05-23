function replaceText() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getDisplayValues(); // 数式も文字列として取得

  var replacements = {
    "830": "8:30",
    "900": "9:00",
    "930": "9:30",
    "1000": "10:00",
    "1030": "10:30",
    "1100": "11:00",
    "1130": "11:30",
    "1200": "12:00",
    "1230": "12:30",
    "1300": "13:00",
    "1330": "13:30",
    "1400": "14:00",
    "1430": "14:30",
    "1500": "15:00",
    "1530": "15:30",
    "1600": "16:00",
    "1630": "16:30",
    "1700": "17:00",
    "1730": "17:30",
    "1800": "18:00",
    "1830": "18:30",
    "1900": "19:00",
    "1930": "19:30",
    "2000": "20:00",
    "2030": "20:30",
    "2100": "21:00",
    "2130": "21:30",
    "2200": "22:00"
  };

  for (var i = 5; i < values.length; i++) {//6行目2列目以降の置換
    for (var j = 1; j < values[i].length; j++) {
      if (typeof values[i][j] === 'string') { // 文字列かどうかのチェック
        for (var key in replacements) {
          values[i][j] = values[i][j].replace(new RegExp(key, 'g'), replacements[key]);
        }
      }
    }
  }

  range.setValues(values);
}

