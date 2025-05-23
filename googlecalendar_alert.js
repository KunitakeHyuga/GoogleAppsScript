//Googleカレンダーから当日分の予定を全て取得し、予定の配列を作成する関数
function getGoogleCalendar() {
  //B,カレンダーから取得する時間の設定
  const today = new Date();
  today.setHours(0);//当日の午前12時の「時」
  today.setMinutes(0);//「分」
  today.setSeconds(0); //「秒」
  const tomorrow = new Date(Date.parse(today) + (24 * 60 * 60 * 1000)); //翌日の午前12時を設定

  //予定を取得する日付の取得
  const monthNum = (today.getMonth())+1; //月
  const dateNum = today.getDate(); //日
  const day = today.getDay(); //曜日
  const dayArray = ['日','月','火','水','木','金','土'];
  const thisDate = monthNum + "/" + dateNum; //当日の日付を文字列に変換
  let sendMessage = "おはようございます。\n" + thisDate + "(" + dayArray[day] + ")";

  //カレンダーから予定を取得
  let myCalendar1 = CalendarApp.getCalendarById(GoogleCalender_ID_1);
  let myCalendar2 = CalendarApp.getCalendarById(GoogleCalendar_ID_2);
  let myCalendar3 = CalendarApp.getCalendarById(GoogleCalendar_ID_3);
  let events1 = myCalendar1.getEvents(today, tomorrow); //取得する当日のカレンダーの予定をすべて取得
  let events2 = myCalendar2.getEvents(today, tomorrow);
  let events3 = myCalendar3.getEvents(today, tomorrow);
  let events = [];
  events = events1.concat(events2,events3);
  let schedule = ""; //配列から文字列に変換した予定の文字列
  let messageArray = []; //取得した予定を格納する配列

  //取得した予定を配列に格納
  for (var i in events) {
    const number = "\n" + (Number(i) + 1) + "件目"; //予定の件数
    const startHours = "0" + events[i].getStartTime().getHours();
    const startMinutes = "0" + events[i].getStartTime().getMinutes();
    const startTime = startHours.slice(-2) +":"+ startMinutes.slice(-2); //開始時間
    const endHours = "0" + events[i].getEndTime().getHours();
    const endMinutes = "0" + events[i].getEndTime().getMinutes();
    const endTime = endHours.slice(-2) +":"+ endMinutes.slice(-2); //終了時間
    const time = "\n【時間】" + startTime +" ~ "+ endTime; //予定を行う日時
    const title = "\n【予定】" + events[i].getTitle(); //予定のタイトル
    let location = ""; //場所
    let description = ""; //詳細

    //空白ではないときの処理
    if(!events[i].getLocation() == null || !events[i].getLocation() == ""){
      location = "\n【場所】" + events[i].getLocation();
    }
    if(!events[i].getDescription() == null || !events[i].getDescription() == ""){
      description = "\n【詳細】" + events[i].getDescription();
    }

    //カレンダーから得たデータを文にまとめて配列に格納
    const message = number + time + title + location + description + "\n";
    messageArray.push(message);
 }

 //送信する文章を作成
 for(var j=0; j<=messageArray.length-1; j++){
   schedule += messageArray[j];
 }
 if(schedule == "" || schedule == null){
   sendMessage += "の予定はありません。\n";
 }else{
   sendMessage += "は" + events.length + "件予定があります。\n" + schedule + "\n https://calendar.google.com";
 }

 var postData = {
    to: LINE_USER_ID,
    messages: [
      {
        type: 'text',
        text: sendMessage,
      },
    ],
  };

 //LINEに送信
 const options =
  {
    "method"  : "post",
    "contentType": "application/json",
    "payload" : JSON.stringify(postData),
    "headers" : {"Authorization" : "Bearer "+ LINE_ACCESS_TOKEN}
  };
 UrlFetchApp.fetch(URL, options);
}