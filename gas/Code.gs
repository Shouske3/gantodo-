/**
 * GanTODO 同期用 Google Apps Script
 * -------------------------------------------------
 * セットアップ手順は README.md の「端末間同期」を参照。
 * このスクリプトは、アプリの全データ(JSON)をスプレッドシートの
 * 「data」シートに分割保存するだけのシンプルな仕組みです。
 *
 * このファイルにはもう一つの役割があります: Claude.ai（またはClaude Desktop/Code）から
 * 「カスタムコネクタ（リモートMCPサーバー）」として接続し、Claudeとの会話の中で決まった予定を
 * そのままGanTODOのタスク・Googleカレンダーに登録できるようにする、MCP(Model Context Protocol)
 * サーバーの実装です（README.md「Claude連携」参照）。
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

  // MCP(Model Context Protocol)のJSON-RPC 2.0リクエストは既存の同期プロトコル({key,action}形式)とは
  // 別物なので、ここで先に振り分ける。認証はリクエストボディではなくURLのクエリパラメータ(?key=...)で行う
  // （Claude.aiのカスタムコネクタ設定でURLにそのまま埋め込める、最もシンプルな方式のため）。
  if (req.jsonrpc === "2.0") {
    if ((e.parameter && e.parameter.key) !== SHARED_KEY) {
      return jsonRpcError_(req.id, -32001, "共有キーが一致しません（接続用URLの末尾に ?key=... を付け忘れていませんか）");
    }
    return handleMcp_(req);
  }

  if (req.key !== SHARED_KEY) return json_({ error: "共有キーが一致しません" });

  if (req.action === "pull") {
    const data = readData_();
    return json_({ data });
  }

  if (req.action === "push") {
    writeData_(req.data || {});
    return json_({ ok: true, savedAt: new Date().toISOString() });
  }

  // Googleカレンダー連携: 開始日時刻・終了日時刻の両方があればその区間そのままの予定を、
  // 終了日時刻だけあればその時刻から1時間の予定を、時刻指定が無く開始日・終了日だけあれば
  // 開始日〜終了日の終日の予定を、専用の「GanTODO」カレンダー（無ければ自動作成）に作成/更新する。
  // 既存のeventIdがあれば更新、なければ新規作成。
  if (req.action === "gcalUpsert") {
    const eventId = upsertCalendarEvent_(req);
    return json_({ eventId });
  }

  // カレンダー連携の解除（終了日時刻が消えた・タスクが削除された場合に呼ぶ）
  if (req.action === "gcalDelete") {
    deleteCalendarEvent_(req.eventId);
    return json_({ ok: true });
  }

  return json_({ error: "不明なaction: " + req.action });
}

/* ============================================================
   データの読み書き（同期・MCP共通）
   ============================================================ */
function readData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("data") || ss.insertSheet("data");
  const last = sheet.getLastRow();
  if (!last) return null;
  const chunks = sheet.getRange(1, 1, last, 1).getValues().map(function (r) { return r[0]; }).join("");
  try {
    return JSON.parse(chunks);
  } catch (err) {
    return null;
  }
}
function writeData_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("data") || ss.insertSheet("data");
  const s = JSON.stringify(data || {});
  const SIZE = 40000; // セルの文字数上限(50,000)より小さく分割
  const rows = [];
  for (let i = 0; i < s.length; i += SIZE) rows.push([s.slice(i, i + SIZE)]);
  sheet.clearContents();
  if (rows.length) sheet.getRange(1, 1, rows.length, 1).setValues(rows);
}

/* ============================================================
   Googleカレンダー連携（gas側の共通ロジック。index.htmlのisCalendarEligible/
   syncTaskToCalendarと同じ判定基準をサーバー側でも使えるようにしたもの）
   ============================================================ */
function getGanTodoCalendar_() {
  const NAME = "GanTODO";
  const cals = CalendarApp.getCalendarsByName(NAME);
  return cals.length ? cals[0] : CalendarApp.createCalendar(NAME);
}
// "YYYY-MM-DD"をスクリプトのタイムゾーンに依存せず正しい日付のDateにする
// （new Date("YYYY-MM-DD")はUTC解釈されるため、実行環境によって前後の日にずれることがある）
function parseDateOnly_(s) {
  const parts = s.split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}
function upsertCalendarEvent_(req) {
  const cal = getGanTodoCalendar_();
  let event = null;
  if (req.eventId) {
    try { event = cal.getEventById(req.eventId); } catch (err) { event = null; }
  }
  if (req.allDay) {
    const start = parseDateOnly_(req.startDate);
    // Calendar APIの終日予定の終了日は排他的（翌日扱い）なので+1日する
    const endExclusive = new Date(parseDateOnly_(req.endDate).getTime() + 24 * 60 * 60 * 1000);
    if (event) {
      event.setAllDayDates(start, endExclusive);
      event.setTitle(req.title || "(無題)");
    } else {
      event = cal.createAllDayEvent(req.title || "(無題)", start, endExclusive);
    }
  } else {
    const start = new Date(req.startISO);
    const end = req.endISO ? new Date(req.endISO) : new Date(start.getTime() + (req.durationMinutes || 60) * 60000);
    if (event) {
      event.setTime(start, end);
      event.setTitle(req.title || "(無題)");
    } else {
      event = cal.createEvent(req.title || "(無題)", start, end);
    }
    // 時刻ありの予定は、開始時刻ちょうどにポップアップ通知（iPhoneのGoogleカレンダーアプリで
    // 通知を有効にしていれば、開始時刻にリマインドが表示される）。更新時の重複を避けるため一旦全消し。
    try {
      event.removeAllReminders();
      event.addPopupReminder(0);
    } catch (err) {
      // リマインダー設定に非対応の環境などは無視（予定自体の作成は成功させる）
    }
  }
  return event.getId();
}
function deleteCalendarEvent_(eventId) {
  try {
    const cal = getGanTodoCalendar_();
    const event = cal.getEventById(eventId);
    if (event) event.deleteEvent();
  } catch (err) {
    // すでに削除済み・存在しない場合は無視
  }
}
// index.htmlのisCalendarEligible(t)と同じ判定（未来日タスクを完了にした場合はカレンダー対象外にする等）
function isCalendarEligible_(t) {
  if (t.deletedAt || !t.endDate) return false;
  if (t.status === "completed" && t.endDate > todayStr_()) return false;
  return !!t.endTime || !!t.startDate;
}
// タスク1件ぶんのカレンダー同期（作成・更新・不要になった場合の削除）。gcalEventIdが変わった場合はtrue、
// 呼び出し側でtaskオブジェクトのgcalEventIdを更新して保存すること。
function syncTaskCalendar_(t) {
  if (!isCalendarEligible_(t)) {
    if (t.gcalEventId) {
      deleteCalendarEvent_(t.gcalEventId);
      t.gcalEventId = null;
      return true;
    }
    return false;
  }
  const req = (t.startDate && t.startTime && t.endTime)
    ? { eventId: t.gcalEventId, title: t.title, startISO: t.startDate + "T" + t.startTime + ":00+09:00", endISO: t.endDate + "T" + t.endTime + ":00+09:00" }
    : t.endTime
      ? { eventId: t.gcalEventId, title: t.title, startISO: t.endDate + "T" + t.endTime + ":00+09:00", durationMinutes: 60 }
      : { eventId: t.gcalEventId, allDay: true, title: t.title, startDate: t.startDate, endDate: t.endDate };
  const newEventId = upsertCalendarEvent_(req);
  if (newEventId !== t.gcalEventId) { t.gcalEventId = newEventId; return true; }
  return false;
}
function todayStr_() {
  const tz = Session.getScriptTimeZone() || "Asia/Tokyo";
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
}

/* ============================================================
   MCP (Model Context Protocol) サーバー
   Claude.ai/Claude Desktop/Claude Codeの「カスタムコネクタ（リモートMCP）」として、
   このWebアプリのURLに ?key=合言葉 を付けたものをそのまま接続先URLに登録して使う。
   ============================================================ */
const MCP_TOOLS = [
  {
    name: "gantodo_add_task",
    description: "GanTODOに新しいタスクを1件追加する。projectを指定すると、そのプロジェクトが無ければ自動作成される。" +
      "startDate/endDateを指定すると、endTimeの有無に応じて自動でGoogleカレンダー（GanTODOカレンダー）にも反映される" +
      "（開始・終了とも時刻ありならその区間の予定、終了時刻のみなら1時間の予定、時刻無しなら終日の予定）。",
    inputSchema: {
      type: "object",
      properties: {
        title: { type: "string", description: "タスクのタイトル（必須）" },
        project: { type: "string", description: "プロジェクト名。省略時はInbox。既存に無い名前なら新規作成される" },
        note: { type: "string", description: "メモ（任意）" },
        startDate: { type: "string", description: "開始日 YYYY-MM-DD（任意）" },
        startTime: { type: "string", description: "開始時刻 HH:MM（任意。startDateも指定されている場合のみ有効）" },
        endDate: { type: "string", description: "終了日 YYYY-MM-DD（任意）" },
        endTime: { type: "string", description: "終了時刻 HH:MM（任意。endDateも指定されている場合のみ有効）" },
        dueDate: { type: "string", description: "納期 YYYY-MM-DD（任意）" },
        priority: { type: "string", enum: ["high", "medium", "low", "none"], description: "優先度（任意、既定はnone）" },
        tags: { type: "array", items: { type: "string" }, description: "タグ（任意）" },
      },
      required: ["title"],
    },
  },
  {
    name: "gantodo_list_tasks",
    description: "GanTODOのタスクを検索・一覧表示する。既存タスクのID確認や、これから登録する内容の重複チェックに使う。",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "タイトルの部分一致検索（任意）" },
        project: { type: "string", description: "プロジェクト名の部分一致で絞り込み（任意）" },
        includeCompleted: { type: "boolean", description: "完了済み・削除済みタスクも含めるか（既定false）" },
        limit: { type: "number", description: "最大表示件数（既定30）" },
      },
    },
  },
  {
    name: "gantodo_update_task",
    description: "既存のGanTODOタスクを更新する。idはgantodo_list_tasksやgantodo_add_taskの結果に含まれる" +
      "内部IDの下6桁（例: a1b2c3）で指定できる。日時を更新すると、対応するGoogleカレンダーの予定も追従する。",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "更新対象タスクの内部ID（下6桁でも可、必須）" },
        title: { type: "string" },
        note: { type: "string" },
        startDate: { type: "string", description: "YYYY-MM-DD" },
        startTime: { type: "string", description: "HH:MM" },
        endDate: { type: "string", description: "YYYY-MM-DD" },
        endTime: { type: "string", description: "HH:MM" },
        dueDate: { type: "string", description: "YYYY-MM-DD" },
        priority: { type: "string", enum: ["high", "medium", "low", "none"] },
        completed: { type: "boolean", description: "trueで完了、falseで未完了に戻す" },
      },
      required: ["id"],
    },
  },
];

function handleMcp_(req) {
  // 通知（idが無いリクエスト。例: notifications/initialized）には応答を返さない
  const isNotification = req.id === undefined || req.id === null;

  if (req.method === "initialize") {
    return jsonRpcResult_(req.id, {
      protocolVersion: "2025-03-26",
      capabilities: { tools: { listChanged: false } },
      serverInfo: { name: "GanTODO", version: "1.0.0" },
    });
  }
  if (req.method === "notifications/initialized" || req.method === "initialized") {
    return isNotification ? emptyOutput_() : jsonRpcResult_(req.id, {});
  }
  if (req.method === "tools/list") {
    return jsonRpcResult_(req.id, { tools: MCP_TOOLS });
  }
  if (req.method === "resources/list") {
    return jsonRpcResult_(req.id, { resources: [] });
  }
  if (req.method === "prompts/list") {
    return jsonRpcResult_(req.id, { prompts: [] });
  }
  if (req.method === "tools/call") {
    try {
      const text = callMcpTool_(req.params && req.params.name, (req.params && req.params.arguments) || {});
      return jsonRpcResult_(req.id, { content: [{ type: "text", text: text }], isError: false });
    } catch (err) {
      return jsonRpcResult_(req.id, { content: [{ type: "text", text: "エラー: " + err.message }], isError: true });
    }
  }
  if (isNotification) return emptyOutput_();
  return jsonRpcError_(req.id, -32601, "不明なmethod: " + req.method);
}

function callMcpTool_(name, args) {
  if (name === "gantodo_add_task") return mcpAddTask_(args);
  if (name === "gantodo_list_tasks") return mcpListTasks_(args);
  if (name === "gantodo_update_task") return mcpUpdateTask_(args);
  throw new Error("不明なツール: " + name);
}

function uid_() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function nowISO_() {
  return new Date().toISOString();
}
function findOrCreateProject_(data, name) {
  if (!name) return null;
  let p = data.projects.find(function (x) { return x.status !== "deleted" && x.name === name; });
  if (p) return p;
  const PROJ_COLORS = ["#5B8DEF", "#22C55E", "#F5B400", "#EF4444", "#A855F7", "#14B8A6", "#F97316"];
  p = {
    id: uid_(), name: name, color: PROJ_COLORS[data.projects.length % PROJ_COLORS.length],
    icon: "", status: "active", sortOrder: data.projects.length, showCompleted: true,
    createdAt: nowISO_(), updatedAt: nowISO_(),
  };
  data.projects.push(p);
  return p;
}
function findTaskById_(data, idOrShort) {
  return data.tasks.find(function (t) {
    return !t.deletedAt && (t.id === idOrShort || t.id.slice(-6) === idOrShort);
  });
}

function mcpAddTask_(args) {
  if (!args.title) throw new Error("titleは必須です");
  const data = readData_() || { projects: [], sections: [], tasks: [], memos: [], resetAt: null };
  data.projects = data.projects || []; data.tasks = data.tasks || [];
  const proj = args.project ? findOrCreateProject_(data, args.project) : null;
  let min = 0;
  for (const t of data.tasks) if (t.sortOrder < min) min = t.sortOrder;
  const t = {
    id: uid_(), title: args.title, note: args.note || "", status: "active", progress: 0,
    projectId: proj ? proj.id : null, sectionId: null, parentTaskId: null,
    priority: args.priority || "none", tags: args.tags || [], assignee: "", category: "",
    startDate: args.startDate || null, startTime: args.startDate ? (args.startTime || null) : null,
    endDate: args.endDate || null, endTime: args.endDate ? (args.endTime || null) : null,
    dueDate: args.dueDate || null, earliestStartDate: null,
    blockedBy: [], isMilestone: false, waitingOnOther: false, nextAction: null, repeat: null,
    attachments: [], collapsed: false, gcalEventId: null,
    sortOrder: min - 1,
    createdAt: nowISO_(), updatedAt: nowISO_(), completedAt: null, deletedAt: null,
  };
  data.tasks.push(t);
  syncTaskCalendar_(t);
  writeData_(data);
  return "追加しました: #" + t.id.slice(-6) + " 「" + t.title + "」" +
    (proj ? "（プロジェクト: " + proj.name + "）" : "") +
    (t.gcalEventId ? " ／ Googleカレンダーにも反映しました" : "");
}

function mcpListTasks_(args) {
  const data = readData_() || { projects: [], tasks: [] };
  const q = (args.query || "").toLowerCase();
  const projQ = (args.project || "").toLowerCase();
  const limit = args.limit || 30;
  let tasks = data.tasks.filter(function (t) {
    if (t.deletedAt) return false;
    if (!args.includeCompleted && t.status === "completed") return false;
    if (q && t.title.toLowerCase().indexOf(q) === -1) return false;
    if (projQ) {
      const p = data.projects.find(function (x) { return x.id === t.projectId; });
      if (!p || p.name.toLowerCase().indexOf(projQ) === -1) return false;
    }
    return true;
  });
  tasks = tasks.slice(0, limit);
  if (!tasks.length) return "該当するタスクはありません";
  return tasks.map(function (t) {
    const p = data.projects.find(function (x) { return x.id === t.projectId; });
    const period = (t.startDate || t.endDate)
      ? " 期間" + (t.startDate || "?") + (t.startTime ? " " + t.startTime : "") + "〜" + (t.endDate || "?") + (t.endTime ? " " + t.endTime : "")
      : "";
    return "#" + t.id.slice(-6) + " [" + t.status + "] " + t.title +
      (p ? "（" + p.name + "）" : "") + (t.dueDate ? " 納期" + t.dueDate : "") + period;
  }).join("\n");
}

function mcpUpdateTask_(args) {
  if (!args.id) throw new Error("idは必須です");
  const data = readData_();
  if (!data) throw new Error("データがまだありません");
  const t = findTaskById_(data, args.id);
  if (!t) throw new Error("タスクが見つかりません: id=" + args.id);
  const fields = ["title", "note", "startDate", "startTime", "endDate", "endTime", "dueDate", "priority"];
  for (const f of fields) if (args[f] !== undefined) t[f] = args[f];
  if (args.completed === true) { t.status = "completed"; t.completedAt = nowISO_(); t.progress = 100; }
  else if (args.completed === false) { t.status = "active"; t.completedAt = null; if (t.progress === 100) t.progress = 0; }
  t.updatedAt = nowISO_();
  syncTaskCalendar_(t);
  writeData_(data);
  return "更新しました: #" + t.id.slice(-6) + " 「" + t.title + "」" +
    (t.gcalEventId ? " ／ Googleカレンダーも更新しました" : "");
}

function jsonRpcResult_(id, result) {
  return json_({ jsonrpc: "2.0", id: id, result: result });
}
function jsonRpcError_(id, code, message) {
  return json_({ jsonrpc: "2.0", id: id, error: { code: code, message: message } });
}
function emptyOutput_() {
  return ContentService.createTextOutput("");
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
