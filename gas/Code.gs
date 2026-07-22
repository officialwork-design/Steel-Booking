/**
 * スチール予約システム - GAS Web API（新仕様）
 *
 * データ（同一スプレッドシート内の複数シート）:
 *   users   : userId | 名前 | 有効 | 登録日時          … 予約できる人（管理者が登録）
 *   slots   : 枠ID   | 日付 | 時間 | 有効 | 作成日時     … 予約枠（1枠1人）
 *   予約     : userId | 名前 | 枠ID | 日付 | 時間 | 備考 | ステータス | 受付日時 | 更新日時
 *   設定     : キー   | 値                              … rules 等
 *
 * スクリプトプロパティ:
 *   SHEET_ID  … 省略時は下の DEFAULT_SHEET_ID を使用
 *   ADMIN_KEY … 管理画面の操作パスコード（必須。未設定だと管理操作は拒否）
 */

// SHEET_ID はスクリプトプロパティが優先。未設定ならこの既定値を使う。
var DEFAULT_SHEET_ID = '1neHxod-oOulSHjkbd41hoZ9emkkJPtHvvaQgO6wNvoE';

var SH = {
  users: { name: 'users', headers: ['userId', '名前', '有効', '登録日時'] },
  slots: { name: 'slots', headers: ['枠ID', '日付', '時間', '有効', '作成日時'] },
  resv:  { name: '予約',  headers: ['userId', '名前', '枠ID', '日付', '時間', '備考', 'ステータス', '受付日時', '更新日時'] },
  conf:  { name: '設定',  headers: ['キー', '値'] }
};

/* ---------- 基盤ヘルパー ---------- */

function prop_(key, fallback) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return (v === null || v === '') ? fallback : v;
}

function ss_() {
  return SpreadsheetApp.openById(prop_('SHEET_ID', DEFAULT_SHEET_ID));
}

function sheet_(def) {
  var ss = ss_();
  var sh = ss.getSheetByName(def.name);
  if (!sh) sh = ss.insertSheet(def.name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// シートを {header: value} の配列で返す
function rows_(def) {
  var sh = sheet_(def);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var w = def.headers.length;
  var values = sh.getRange(2, 1, last - 1, w).getValues();
  return values.map(function (r, i) {
    var o = { _row: i + 2 };
    for (var c = 0; c < w; c++) o[def.headers[c]] = r[c];
    return o;
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function nowStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

function truthy_(v) {
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes' || s === '有効' || s === 'o' || s === '○';
}

/* ---------- 業務ロジック ---------- */

function findUser_(userId) {
  var us = rows_(SH.users);
  for (var i = 0; i < us.length; i++) if (String(us[i].userId) === String(userId)) return us[i];
  return null;
}

function findResv_(userId) {
  var rs = rows_(SH.resv);
  for (var i = 0; i < rs.length; i++) if (String(rs[i].userId) === String(userId)) return rs[i];
  return null;
}

// 他人に予約されている枠IDの集合
function takenSlotIds_(exceptUserId) {
  var rs = rows_(SH.resv);
  var set = {};
  rs.forEach(function (r) {
    if (exceptUserId && String(r.userId) === String(exceptUserId)) return;
    if (r['枠ID']) set[String(r['枠ID'])] = true;
  });
  return set;
}

// 指定ユーザーが選べる枠一覧（有効かつ空き。自分の予約中の枠は含む）
function availableSlots_(userId) {
  var taken = takenSlotIds_(userId);
  return rows_(SH.slots)
    .filter(function (s) { return truthy_(s['有効']) && !taken[String(s['枠ID'])]; })
    .map(function (s) {
      return { slotId: String(s['枠ID']), date: String(s['日付']), time: String(s['時間']) };
    })
    .sort(function (a, b) {
      return (a.date + a.time < b.date + b.time) ? -1 : 1;
    });
}

function getRules_() {
  var c = rows_(SH.conf);
  for (var i = 0; i < c.length; i++) if (String(c[i]['キー']) === 'rules') return String(c[i]['値'] || '');
  return '';
}

function reservationOut_(r) {
  if (!r) return null;
  return {
    slotId: String(r['枠ID']), date: String(r['日付']), time: String(r['時間']),
    remarks: String(r['備考'] || ''), status: String(r['ステータス'] || '')
  };
}

/* ---------- LIFF向け ---------- */

// 初期状態: 登録状況・自分の予約・空き枠・ルール
function actionInit_(body) {
  var userId = body.userId;
  if (!userId) return { ok: false, error: 'userId がありません。' };
  var u = findUser_(userId);
  var registered = !!(u && truthy_(u['有効']));
  var res = {
    ok: true,
    registered: registered,
    name: u ? String(u['名前'] || '') : '',
    userId: String(userId),
    rules: getRules_()
  };
  if (registered) {
    res.myReservation = reservationOut_(findResv_(userId));
    res.slots = availableSlots_(userId);
  }
  return res;
}

// 予約の作成/更新（1人1件）
function actionBook_(body) {
  var userId = body.userId;
  if (!userId) return { ok: false, error: 'userId がありません。' };
  var u = findUser_(userId);
  if (!u || !truthy_(u['有効'])) return { ok: false, error: '未登録です。管理者に登録を依頼してください。' };
  var slotId = String(body.slotId || '');
  if (!slotId) return { ok: false, error: '予約枠を選択してください。' };

  // 枠の存在・有効チェック
  var slot = null;
  rows_(SH.slots).forEach(function (s) { if (String(s['枠ID']) === slotId) slot = s; });
  if (!slot || !truthy_(slot['有効'])) return { ok: false, error: 'その枠は選択できません。' };

  // 他人に取られていないか
  var taken = takenSlotIds_(userId);
  if (taken[slotId]) return { ok: false, error: 'その枠はすでに埋まりました。別の枠を選んでください。' };

  var name = String(u['名前'] || '');
  var date = String(slot['日付']); var time = String(slot['時間']);
  var remarks = String(body.remarks || '');
  var sh = sheet_(SH.resv);
  var existing = findResv_(userId);

  if (existing) {
    // 更新（受付日時は保持、更新日時のみ更新）
    var vals = [userId, name, slotId, date, time, remarks, existing['ステータス'] || '受付', existing['受付日時'] || nowStr_(), nowStr_()];
    sh.getRange(existing._row, 1, 1, SH.resv.headers.length).setValues([vals]);
    return { ok: true, updated: true, reservation: reservationOut_({ '枠ID': slotId, '日付': date, '時間': time, '備考': remarks, 'ステータス': existing['ステータス'] || '受付' }) };
  } else {
    sh.appendRow([userId, name, slotId, date, time, remarks, '受付', nowStr_(), nowStr_()]);
    return { ok: true, created: true, reservation: reservationOut_({ '枠ID': slotId, '日付': date, '時間': time, '備考': remarks, 'ステータス': '受付' }) };
  }
}

/* ---------- 管理向け（要 ADMIN_KEY） ---------- */

function adminAuth_(key) {
  var k = prop_('ADMIN_KEY', '');
  if (!k) return { ok: false, error: '管理機能が未設定です（ADMIN_KEY を設定してください）。' };
  if (String(key) !== k) return { ok: false, error: 'パスコードが違います。' };
  return { ok: true };
}

function adminList_() {
  return {
    ok: true,
    users: rows_(SH.users).map(function (u) {
      return { userId: String(u.userId), name: String(u['名前'] || ''), active: truthy_(u['有効']) };
    }),
    slots: rows_(SH.slots).map(function (s) {
      var taken = takenSlotIds_(null);
      return { slotId: String(s['枠ID']), date: String(s['日付']), time: String(s['時間']), active: truthy_(s['有効']), taken: !!taken[String(s['枠ID'])] };
    }),
    reservations: rows_(SH.resv).map(function (r) {
      return {
        userId: String(r.userId), name: String(r['名前'] || ''), slotId: String(r['枠ID']),
        date: String(r['日付']), time: String(r['時間']), remarks: String(r['備考'] || ''),
        status: String(r['ステータス'] || ''), receivedAt: String(r['受付日時'] || ''), updatedAt: String(r['更新日時'] || '')
      };
    }),
    rules: getRules_()
  };
}

function adminSaveUser_(b) {
  var sh = sheet_(SH.users);
  var u = findUser_(b.userId);
  var active = (b.active === false || b.active === 'false') ? false : true;
  if (u) {
    sh.getRange(u._row, 1, 1, 4).setValues([[b.userId, b.name || u['名前'] || '', active, u['登録日時'] || nowStr_()]]);
  } else {
    if (!b.userId) return { ok: false, error: 'userId が必要です。' };
    sh.appendRow([b.userId, b.name || '', active, nowStr_()]);
  }
  return { ok: true };
}

function adminDeleteUser_(b) {
  var sh = sheet_(SH.users);
  var u = findUser_(b.userId);
  if (u) sh.deleteRow(u._row);
  return { ok: true };
}

function adminSaveSlot_(b) {
  var sh = sheet_(SH.slots);
  var active = (b.active === false || b.active === 'false') ? false : true;
  if (!b.date || !b.time) return { ok: false, error: '日付と時間は必須です。' };
  var target = null;
  rows_(SH.slots).forEach(function (s) { if (b.slotId && String(s['枠ID']) === String(b.slotId)) target = s; });
  if (target) {
    sh.getRange(target._row, 1, 1, 5).setValues([[target['枠ID'], b.date, b.time, active, target['作成日時'] || nowStr_()]]);
  } else {
    var id = 'S' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + Math.floor(Math.random() * 100);
    sh.appendRow([id, b.date, b.time, active, nowStr_()]);
  }
  return { ok: true };
}

function adminDeleteSlot_(b) {
  var sh = sheet_(SH.slots);
  var target = null;
  rows_(SH.slots).forEach(function (s) { if (String(s['枠ID']) === String(b.slotId)) target = s; });
  if (target) sh.deleteRow(target._row);
  return { ok: true };
}

function adminDeleteReservation_(b) {
  var sh = sheet_(SH.resv);
  var r = findResv_(b.userId);
  if (r) sh.deleteRow(r._row);
  return { ok: true };
}

function adminSaveRules_(b) {
  var sh = sheet_(SH.conf);
  var target = null;
  rows_(SH.conf).forEach(function (c) { if (String(c['キー']) === 'rules') target = c; });
  if (target) sh.getRange(target._row, 1, 1, 2).setValues([['rules', String(b.text || '')]]);
  else sh.appendRow(['rules', String(b.text || '')]);
  return { ok: true };
}

/* ---------- ルーティング ---------- */

var ADMIN_ACTIONS = {
  admin_list: function (b) { return adminList_(); },
  admin_saveUser: adminSaveUser_,
  admin_deleteUser: adminDeleteUser_,
  admin_saveSlot: adminSaveSlot_,
  admin_deleteSlot: adminDeleteSlot_,
  admin_deleteReservation: adminDeleteReservation_,
  admin_saveRules: adminSaveRules_
};

function route_(action, body) {
  if (action === 'init') return actionInit_(body);
  if (action === 'book') return actionBook_(body);
  if (ADMIN_ACTIONS[action]) {
    var auth = adminAuth_(body.key);
    if (!auth.ok) return auth;
    return ADMIN_ACTIONS[action](body);
  }
  return { ok: false, error: '未対応のaction: ' + action };
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_(route_(body.action, body));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (!p.action) return json_({ ok: true, message: 'Steel-Booking GAS API is running.' });
    return json_(route_(p.action, p));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
