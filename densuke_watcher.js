function checkDensukeDateAdded() {
  const html = UrlFetchApp.fetch(DENSUKE_URL).getContentText();

  // すべての<table>から中山を含むものを抽出
  const tables = [...html.matchAll(/<table[\s\S]*?<\/table>/g)];
  let targetTable = null;
  for (const match of tables) {
    if (match[0].includes("中山")) {
      targetTable = match[0];
      break;
    }
  }

  if (!targetTable) {
    Logger.log("中山を含むテーブルが見つかりません");
    return;
  }

  // <tr>をすべて取得（1行目はヘッダー行なのでスキップ）
  const rowMatches = [...targetTable.matchAll(/<tr[\s\S]*?<\/tr>/g)];
  const dataRows = rowMatches.slice(1);

  // 各行の1列目（日程）を取得
  const currentDates = dataRows
    .map((row) => {
      const htmlRow = row[0]; // Matchオブジェクトから実際の文字列を取り出す
      const match = htmlRow.match(/<td[\s\S]*?>([\s\S]*?)<\/td>/);
      return match ? match[1].trim().replace(/<[^>]*>/g, "") : null; // HTMLタグ除去
    })
    .filter((date) => date && !date.includes("---"));

  Logger.log("現在の日程一覧: " + currentDates.join(", "));

  // 前回の日程一覧を取得
  const stored =
    PropertiesService.getScriptProperties().getProperty("densukeDates");
  const previousDates = stored ? JSON.parse(stored) : [];

  Logger.log("前回の日程一覧: " + previousDates.join(", "));

  // 新規追加された日程を抽出
  const addedDates = currentDates.filter(
    (date) => !previousDates.includes(date)
  );

  if (addedDates.length > 0) {
    Logger.log("新しい日程が追加されました: " + addedDates.join(", "));
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "【伝助】研究ゼミの日程追加",
      body: `以下の日程が新たに追加されました：\n${addedDates.join(
        "\n"
      )}\n\n${DENSUKE_URL}`,
    });
    PropertiesService.getScriptProperties().setProperty(
      "densukeDates",
      JSON.stringify(currentDates)
    );
  } else {
    Logger.log("新しい日程の追加はありません");
  }
}
