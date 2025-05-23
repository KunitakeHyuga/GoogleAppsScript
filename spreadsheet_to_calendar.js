// カレンダーにスケジュールを登録する 
function CreateSchedule() {

  // 読み取り範囲（表の始まり行と終わり列）
  const topRow = 2
  const lastCol = 10
  const statusCellCol = 1

  // 予定の一覧バッファ内の列(0始まり)
  const statusNum = 0
  const startdayNum = 1
  const startNum = 2
  const enddayNum = 3
  const endNum = 4
  const titleNum = 5
  const locationNum = 6
  const descriptionNum = 7
  const colorNum = 8
  const calnameNum = 9    //カレンダー名の列

  // シートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  // 予定の最終行を取得
  const lastRow = sheet.getLastRow()

  //予定の一覧をバッファに取得
  const contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues()

  // googleカレンダーの取得
  let Calendar = CalendarApp.getDefaultCalendar()

  // バッファの内容に従って予定を作成
  for (let i = 0; i <= lastRow - topRow; i++) {

    //「済」の場合は無視する
    if (contents[i][statusNum] === '済') {
      continue
    }

    // 値をセット 日時はフォーマットして保持
    let startday = contents[i][startdayNum]
    let startTime = contents[i][startNum]
    let endday = contents[i][enddayNum]
    let endTime = contents[i][endNum]
    let title = contents[i][titleNum]
    let calname = contents[i][calnameNum]    //設定するカレンダー名

    // 場所と詳細をセット
    let options = { location: contents[i][locationNum], description: contents[i][descriptionNum] }

    console.log(startday + " " + contents[i][titleNum])

    try {
      let calevent

      //カレンダーの設定
      if (calname === "") {
        calendar = CalendarApp.getDefaultCalendar()
      }
      else {
        let calendars = CalendarApp.getCalendarsByName(calname);
        for (let s in calendars) {
          Calendar = calendars[s]
          console.log("カレンダー指定 -> " + calendars[s].getName())
        }
      }

      // 終了日の有無をチェック
      if (endday === '') {
        endday = startday //終了日に開始日を入れておく
      }

      let startDate = new Date(startday)
      let endDate = new Date(endday)

      // 開始終了時刻が無ければ終日で設定
      if (startTime == '' || endTime == '') {
        endDate.setDate(endDate.getDate() + 1)    //★なぜか1日プラスする
        console.log("設定する終了日 -> " + endDate)

        //終日の予定を作成
        calevent = Calendar.createAllDayEvent(
          title,
          startDate,
          endDate,
          options
        )

        // 時間指定ありで予定を作成する
      } else {
        // 開始日時を設定する
        startDate.setHours(startTime.getHours())
        startDate.setMinutes(startTime.getMinutes())

        //終了日時を設定する
        endDate.setHours(endTime.getHours())
        endDate.setMinutes(endTime.getMinutes())

        // 日時付きの予定を作成する
        calevent = Calendar.createEvent(
          title,
          startDate,
          endDate,
          options
        )
      }

      //色の設定
      if (contents[i][colorNum] != "") {
        let color = getcolornum(contents[i][colorNum])
        // console.log(title)
        calevent.setColor(color)
      }

      //予定が作成されたら「済」にする
      sheet.getRange(topRow + i, statusCellCol).setValue('済')

      // エラーの場合ログ出力する
    } catch (e) {
      Logger.log(e)
    }
  }

  // 完了通知
  // Browser.msgBox("予定を追加しました。")
}


// 色IDに応じた色の番号を返却する
function getcolornum(color) {
  let colornum

  switch (color) {
    case "PALE_BLUE": colornum = 1; break
    case "PALE_GREEN": colornum = 2; break
    case "MAUVE": colornum = 3; break
    case "PALE_RED": colornum = 4; break
    case "YELLOW": colornum = 5; break
    case "ORANGE": colornum = 6; break
    case "CYAN": colornum = 7; break
    case "GRAY": colornum = 8; break
    case "BLUE": colornum = 9; break
    case "GREEN": colornum = 10; break
    case "RED": colornum = 11; break
  }
  console.log("色の設定 " + color + " -> " + colornum)
  return (colornum)
}