/**
 * GanTODO 同期用 Google Apps Script
 * -------------------------------------------------
 * セットアップ手順は README.md の「端末間同期」を参照。
 * このスクリプトは、アプリの全データ(JSON)をスプレッドシートの
 * 「data」シートに分割保存するだけのシンプルな仕組みです。
 */

// ★ 任意の合言葉に変更してください（アプリの設定画面と同じ値にする）
const SHARED_KEY = "watashi-no-aikotoba";

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ error: "リクエストが不正です" });
  }
  if (req.key !== SHARED_KEY) return json_({ error: "共有キーが一致しません" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("data") || ss.insertSheet("data");

  if (req.action === "pull") {
    const last = sheet.getLastRow();
    if (!last) return json_({ data: null });
    const chunks = sheet.getRange(1, 1, last, 1).getValues().map(function (r) { return r[0]; }).join("");
    try {
      return json_({ data: JSON.parse(chunks) });
    } catch (err) {
      return json_({ data: null });
    }
  }

  if (req.action === "push") {
    const s = JSON.stringify(req.data || {});
    const SIZE = 40000; // セルの文字数上限(50,000)より小さく分割
    const rows = [];
    for (let i = 0; i < s.length; i += SIZE) rows.push([s.slice(i, i + SIZE)]);
    sheet.clearContents();
    if (rows.length) sheet.getRange(1, 1, rows.length, 1).setValues(rows);
    return json_({ ok: true, savedAt: new Date().toISOString() });
  }

  // Googleカレンダー連携: タスクの終了日時刻(endDate+endTime)を、その時刻から1時間の予定として
  // 専用の「GanTODO」カレンダー（無ければ自動作成）に作成/更新する。既存のeventIdがあれば更新、なければ新規作成。
  if (req.action === "gcalUpsert") {
    const cal = getGanTodoCalendar_();
    const start = new Date(req.startISO);
    const end = new Date(start.getTime() + (req.durationMinutes || 60) * 60000);
    let event = null;
    if (req.eventId) {
      try { event = cal.getEventById(req.eventId); } catch (err) { event = null; }
    }
    if (event) {
      event.setTime(start, end);
      event.setTitle(req.title || "(無題)");
    } else {
      event = cal.createEvent(req.title || "(無題)", start, end);
    }
    return json_({ eventId: event.getId() });
  }

  // カレンダー連携の解除（終了日時刻が消えた・タスクが削除された場合に呼ぶ）
  if (req.action === "gcalDelete") {
    try {
      const cal = getGanTodoCalendar_();
      const event = cal.getEventById(req.eventId);
      if (event) event.deleteEvent();
    } catch (err) {
      // すでに削除済み・存在しない場合は無視
    }
    return json_({ ok: true });
  }

  return json_({ error: "不明なaction: " + req.action });
}

// GanTODO専用のカレンダーを取得（無ければ作成）。マイカレンダーに「GanTODO」という名前で表示される。
function getGanTodoCalendar_() {
  const NAME = "GanTODO";
  const cals = CalendarApp.getCalendarsByName(NAME);
  return cals.length ? cals[0] : CalendarApp.createCalendar(NAME);
}

// カレンダー権限を許可するための手動実行用関数。
// エディタ上部の関数選択で「authorizeCalendarAccess」を選んで▷実行を押すと、
// 初回だけカレンダーへのアクセス許可を求める画面が出るので許可する（実行後は削除してOK）。
function authorizeCalendarAccess() {
  getGanTodoCalendar_();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
