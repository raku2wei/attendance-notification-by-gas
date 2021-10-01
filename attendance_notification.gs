// 出欠変更された場合に実行する関数
function onEdit(e) {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // スプレッドシート
  let activeSheet = activeSpreadsheet.getActiveSheet(); // アクティブシート

  if (activeSheet.getName() != "出欠表") {
    return;
  }

  let activeCell = activeSheet.getActiveCell(); // アクティブセル
  let nowInputRow = activeCell.getRow(); // 入力のあった行番号
  let nowInputColumn = activeCell.getColumn(); // 入力のあった列番号

  if (nowInputRow <= 3 || nowInputColumn < 2 || nowInputColumn == 10) {
    // 変更通知対象の範囲外のセルを変更した場合は何もしない
    return;
  }

  // 変更した列名
  let columnName = activeSheet.getRange(3, nowInputColumn).getValues();
  // キャラクター名取得
  let charaName = activeSheet.getRange(nowInputRow, 2).getValues();

  let textMessage = "";
  if (
    columnName == "キャラクター名" ||
    columnName == "職業" ||
    columnName == "戦力" ||
    columnName == "コメント"
  ) {
    activeCellValue = activeCell.getValues();
    if (columnName == "キャラクター名") {
      // キャラクター名が変更された場合は変更前の名前をセット
      charaName = e.oldValue;
    } else if (columnName == "戦力") {
      // 戦力の場合は更新日時を記録
      activeSheet
        .getRange(nowInputRow, 10)
        .setValue(
          Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss")
        );
      activeCellValue = activeCell.getValues() + "万";
    }
    // 送信するテキスト
    textMessage =
      charaName +
      "さんが" +
      columnName +
      "を「" +
      activeCellValue +
      "」に変更したよ！ヨシ！";
  } else if (nowInputColumn > 4) {
    // イベント出欠が変更された場合
    let eventDate = activeSheet.getRange(1, nowInputColumn).getValues(); // イベントの曜日・時間
    let eventName = activeSheet.getRange(2, nowInputColumn).getValues(); // イベント名
    // 送信するテキスト
    textMessage =
      charaName +
      "さんが " +
      eventName +
      "（" +
      eventDate +
      "）" +
      "の出欠を「" +
      activeCell.getValues() +
      "」に変更したよ！ヨシ！";
  }
  sendDiscord(textMessage);
}

// フォームから登録あった場合の通知
function onFormPost() {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // スプレッドシート
  let activeSheet = activeSpreadsheet.getActiveSheet(); // アクティブシート

  if (activeSheet.getName() != "フォームの回答 1") {
    return;
  }

  let activeCell = activeSheet.getActiveCell(); // アクティブセル
  let nowInputRow = activeCell.getRow(); // 入力のあった行番号
  let nowInputColumn = activeCell.getColumn(); // 入力のあった列番号

  let charName = activeSheet.getRange(nowInputRow, 2).getValue(); // キャラクター名取得
  let job = activeSheet.getRange(nowInputRow, 3).getValue(); // 職業取得
  let power = activeSheet.getRange(nowInputRow, 4).getValue(); // 戦力取得
  let attend1 = activeSheet.getRange(nowInputRow, 5).getValue(); // イベント出欠取得
  let attend2 = activeSheet.getRange(nowInputRow, 6).getValue();
  let attend3 = activeSheet.getRange(nowInputRow, 7).getValue();
  let attend4 = activeSheet.getRange(nowInputRow, 8).getValue();
  let comment = activeSheet.getRange(nowInputRow, 9).getValue(); // コメント取得

  // イベント名を取得
  let eventName1 = activeSheet.getRange(1, 5).getValue();
  let eventName2 = activeSheet.getRange(1, 6).getValue();
  let eventName3 = activeSheet.getRange(1, 7).getValue();
  let eventName4 = activeSheet.getRange(1, 8).getValue();

  // 送信するテキスト
  let sendText =
    charName +
    "さんが出欠フォームから出欠登録したよ！ヨシ！\n" +
    "```" +
    "キャラ名: " +
    charName +
    "\n" +
    "職業: " +
    job +
    "\n" +
    "戦力: " +
    power +
    "\n" +
    eventName1 +
    ": " +
    attend1 +
    "\n" +
    eventName2 +
    ": " +
    attend2 +
    "\n" +
    eventName3 +
    ": " +
    attend3 +
    "\n" +
    eventName4 +
    ": " +
    attend4 +
    "\n" +
    "コメント: " +
    comment +
    "```";

  sendDiscord(sendText);

  // 以下、出欠表への反映処理

  let attendSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出欠表");

  // キャラクター名から行を検索
  let targetRow = findRow(attendSheet, charName, 2);
  if (targetRow == 0) {
    // 見つからなかった場合はメンション
    sendDiscord(
      "@here 「" +
        charName +
        "」さんが出欠表に存在しないため、自動反映できなかったよ！手動で反映してね☆"
    );
    return;
  }

  attendSheet.getRange(targetRow, 3).setValue(job); // 職業
  if (power != "") {
    // 入力がないときは反映しない
    attendSheet.getRange(targetRow, 5).setValue(power); // 戦力
    // 戦力の更新日時を記録
    attendSheet
      .getRange(targetRow, 10)
      .setValue(
        Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss")
      );
  }
  if (attend1 != "") {
    attendSheet.getRange(targetRow, 5).setValue(attend1); // 出欠1
  }
  if (attend2 != "") {
    attendSheet.getRange(targetRow, 6).setValue(attend2); // 出欠2
  }
  if (attend3 != "") {
    attendSheet.getRange(targetRow, 7).setValue(attend3); // 出欠3
  }
  if (attend4 != "") {
    attendSheet.getRange(targetRow, 8).setValue(attend4); // 出欠4
  }
  if (comment != "") {
    attendSheet.getRange(targetRow, 9).setValue(comment); // コメント
  }

  // 反映済みチェックを入れる
  activeSheet.getRange(nowInputRow, 10).setValue(true);
  sendDiscord("「" + charName + "」さんの出欠を出欠表に反映したよ！ヨシ！");
}

// 指定した値と一致する行番号を返す
function findRow(sheet, val, col) {
  // シートのデータを二次元配列として取得
  let data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][col - 1] === val) {
      return i + 1;
    }
  }
  // 見つからなかった場合は0を返す
  return 0;
}

// Discordにテキストメッセージを送信する関数
function sendDiscord(textMessage) {
  if (textMessage == "") {
    // 通知内容が空の場合は何もしない
    return;
  }

  //Webhook URLを設定
  let webHookUrl = "https://discord.com/api/webhooks/***************";

  let jsonData = {
    content: textMessage,
  };

  let payload = JSON.stringify(jsonData);

  let options = {
    method: "post",
    contentType: "application/json",
    payload: payload,
  };

  // リクエスト
  UrlFetchApp.fetch(webHookUrl, options);
}
