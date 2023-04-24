function doPost() {
  // 土日祝日の場合が早期リターン
  if (isHoliday()) {
    return false;
  }

  // スプレッドシートを指定
  const userNames = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ユーザーマスタ');
  // 当日のDJを取得
  const dj_person = getDjPerson(userNames);
  // APIの疎通に必要な認証情報を取得
  const credentials = getAuthCredentials();
  // 送信するためのテキストを生成
  const text = `本日のDJは${dj_person}さんです！！\nお好きな曲を一曲流して8Fを沸かしてください！！\n※使用後には必ずスピーカを元の位置に戻し、Bluetoothの接続を解除して下さい。`;

  // グループに送信
  LINEWORKS.channelMessageSend(credentials, text);
}

// DJを取得するための関数
function getDjPerson(userNames) {
  // スクリプトプロパティから数字を取得するためのキーを指定
  const userIdKey = 'user_id';
  // フロアに存在する社員の数を取得
  const employeeCount = userNames.getRange(41, 2).getDisplayValue();
  // スクリプトプロパティから社員IDを取得する
  let user_id = parseInt(PropertiesService.getScriptProperties().getProperty(userIdKey), 10);

  // DJが一巡したかを判定する
  if (employeeCount < user_id) {
    // 一巡した場合はリセット
    user_id = 1;
  } else {
    // 一巡していない場合はインクリメント
    user_id++;
  }
  // スクリプトプロパティに社員番号を保存
  PropertiesService.getScriptProperties().setProperty(userIdKey, user_id);
  // DJ担当の社員IDを返す
  return userNames.getRange(user_id, 2).getDisplayValue();
}

// 土日祝かを判定する関数
function isHoliday() {
  // 今日の日付を取得
  const today = new Date();
  // 曜日を表す数値を取得
  const weekInt = today.getDay();
  // 土日の場合はtrueを返す
  if (weekInt <= 0 || 6 <= weekInt) {
    return true;
  }
  // GoogleカレンダーAPIで指定日のイベントを取得する
  const calendarId = PropertiesService.getScriptProperties().getProperty('calendarId')
  const calendar = CalendarApp.getCalendarById(calendarId);
  const todayEvents = calendar.getEventsForDay(today);
  // イベントが存在する場合はtrueを返す
  if (todayEvents.length > 0) {
    return true;
  }
  // 土日でも祝日でもない場合はfalseを返す
  return false;
}

// 認証情報を取得する関数
function getAuthCredentials() {
  return {
    CLIENT_ID: PropertiesService.getScriptProperties().getProperty('CLIENT_ID'),
    CLIENT_SECRET: PropertiesService.getScriptProperties().getProperty('CLIENT_SECRET'),
    SERVICE_ACCOUNT: PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT'),
    // 改行コードがうまく読み込まれないので、エスケープして読み込み
    PRIVATE_KEY: PropertiesService.getScriptProperties().getProperty('PRIVATE_KEY').replace(/\\n/g, "\n"),
    DOMAIN_ID: PropertiesService.getScriptProperties().getProperty('DOMAIN_ID'),
    ADMIN_ID: PropertiesService.getScriptProperties().getProperty('ADMIN_ID'),
    BOT_ID: PropertiesService.getScriptProperties().getProperty('BOT_ID'),
    channelId: PropertiesService.getScriptProperties().getProperty('channelId')
  };
}
