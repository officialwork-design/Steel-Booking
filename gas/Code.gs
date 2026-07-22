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

// LINEログインチャネルID（LIFF ID の先頭部分）。idToken検証に使用。
var CHANNEL_ID = '2010792348';

var SH = {
  users: { name: 'users', headers: ['userId', '名前', '有効', '登録日時', '管理者'] },
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
  if (sh.getLastRow() === 0 || sh.getLastColumn() < def.headers.length) {
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

function fmtDateCell_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v == null ? '' : v);
}
function fmtTimeCell_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Asia/Tokyo', 'HH:mm');
  var s = String(v == null ? '' : v);
  var m = /^(\d{1,2}):(\d{2})/.exec(s);
  return m ? (('0'+m[1]).slice(-2) + ':' + m[2]) : s;
}
function nextSlotId_() {
  var max = 0;
  rows_(SH.slots).forEach(function (s) { var n = parseInt(s['枠ID'], 10); if (!isNaN(n) && n > max) max = n; });
  return max + 1;
}

function slotDateTime_(dateStr, timeStr) {
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(dateStr || ''));
  if (!m) return null;
  var t = /^(\d{1,2}):(\d{2})/.exec(String(timeStr || '00:00'));
  var hh = t ? +t[1] : 23, mm = t ? +t[2] : 59;
  return new Date(+m[1], +m[2] - 1, +m[3], hh, mm);
}
function isPast_(dateStr, timeStr) {
  var d = slotDateTime_(dateStr, timeStr);
  return d ? (d.getTime() < Date.now()) : false;
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
    .filter(function (s) {
      if (!truthy_(s['有効']) || taken[String(s['枠ID'])]) return false;
      return !isPast_(fmtDateCell_(s['日付']), fmtTimeCell_(s['時間']));
    })
    .map(function (s) {
      return { slotId: String(s['枠ID']), date: fmtDateCell_(s['日付']), time: fmtTimeCell_(s['時間']) };
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
    slotId: String(r['枠ID']), date: fmtDateCell_(r['日付']), time: fmtTimeCell_(r['時間']),
    remarks: String(r['備考'] || ''), status: String(r['ステータス'] || '')
  };
}

/* ---------- LIFF向け ---------- */

// 初期状態: 登録状況・自分の予約・空き枠・ルール
function actionInit_(body) {
  var userId = body.userId;
  if (!userId) return { ok: false, error: 'userId がありません。' };
  var u = findUser_(userId);
  if (!u) {
    // 初回アクセス時は LINE名で自動登録（管理者が後で名前を変更できる）
    var nm = String(body.displayName || '');
    sheet_(SH.users).appendRow([userId, nm, true, nowStr_(), false]);
    u = findUser_(userId);
  }
  var active = truthy_(u['有効']);
  var res = {
    ok: true,
    registered: active,               // 有効なら予約可
    blocked: !active,                 // 無効化された人は予約不可
    name: String(u['名前'] || body.displayName || ''),
    userId: String(userId),
    isAdmin: !!truthy_(u['管理者']),
    rules: getRules_()
  };
  if (active) {
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
  var date = fmtDateCell_(slot['日付']); var time = fmtTimeCell_(slot['時間']);
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

// LINEのアクセストークンを検証して {ok, userId} または {ok:false, error} を返す。
// idTokenは短命ですぐ失効するため、SDKが自動更新するアクセストークンを使う。
function verifyLineUser_(accessToken) {
  if (!accessToken) return { ok: false, error: 'accessTokenがありません' };
  // 1) トークンが有効かつ自チャネル発行のものか確認
  var vRes;
  try {
    vRes = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify?access_token=' + encodeURIComponent(accessToken),
      { muteHttpExceptions: true });
  } catch (e) {
    return { ok: false, error: '外部通信の権限が未承認の可能性: ' + e };
  }
  var vd;
  try { vd = JSON.parse(vRes.getContentText()); } catch (e) { return { ok: false, error: '検証応答の解析失敗' }; }
  if (vd.error) return { ok: false, error: 'LINE: ' + vd.error + ' ' + (vd.error_description || '') };
  if (String(vd.client_id) !== String(CHANNEL_ID)) return { ok: false, error: 'client_id不一致(' + vd.client_id + ')' };
  // 2) userId を取得
  var pRes = UrlFetchApp.fetch('https://api.line.me/v2/profile',
    { headers: { Authorization: 'Bearer ' + accessToken }, muteHttpExceptions: true });
  var pd;
  try { pd = JSON.parse(pRes.getContentText()); } catch (e) { return { ok: false, error: 'プロフィール応答の解析失敗' }; }
  if (!pd.userId) return { ok: false, error: 'userIdが取得できません' };
  return { ok: true, userId: String(pd.userId) };
}

// 管理者本人であることを検証（accessToken → 管理者フラグ）
function requireAdmin_(accessToken) {
  var v = verifyLineUser_(accessToken);
  if (!v.ok) return { ok: false, error: '認証NG: ' + v.error };
  var u = findUser_(v.userId);
  if (!u || !truthy_(u['管理者'])) return { ok: false, error: '管理者権限がありません（あなたのuserId: ' + v.userId + '）' };
  return { ok: true, userId: v.userId };
}

// ▼ 初回だけエディタから手動実行して権限を承認するための関数
function 権限承認用() {
  UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify?access_token=dummy', { muteHttpExceptions: true });
  var n = SpreadsheetApp.openById(prop_('SHEET_ID', DEFAULT_SHEET_ID)).getName();
  Logger.log('OK: ' + n);
}

function adminList_() {
  return {
    ok: true,
    users: rows_(SH.users).map(function (u) {
      return { userId: String(u.userId), name: String(u['名前'] || ''), active: truthy_(u['有効']), admin: truthy_(u['管理者']) };
    }),
    slots: rows_(SH.slots).map(function (s) {
      var taken = takenSlotIds_(null);
      var _d = fmtDateCell_(s['日付']), _t = fmtTimeCell_(s['時間']);
      return { slotId: String(s['枠ID']), date: _d, time: _t, active: truthy_(s['有効']), taken: !!taken[String(s['枠ID'])], past: isPast_(_d, _t) };
    }),
    reservations: rows_(SH.resv).map(function (r) {
      var ru = findUser_(r.userId);
      var rname = ru ? String(ru['名前'] || '') : String(r['名前'] || '');
      return {
        userId: String(r.userId), name: rname, slotId: String(r['枠ID']),
        date: fmtDateCell_(r['日付']), time: fmtTimeCell_(r['時間']), remarks: String(r['備考'] || ''),
        status: String(r['ステータス'] || ''), receivedAt: String(r['受付日時'] || ''), updatedAt: String(r['更新日時'] || '')
      };
    }),
    rules: getRules_()
  };
}

function optBool_(v, dflt) {
  if (v === undefined || v === null) return dflt;
  return !(v === false || v === 'false');
}
function adminSaveUser_(b) {
  var sh = sheet_(SH.users);
  var u = findUser_(b.userId);
  var active = optBool_(b.active, u ? truthy_(u['有効']) : true);
  var admin = optBool_(b.admin, u ? truthy_(u['管理者']) : false);
  var name = (b.name === undefined || b.name === null || b.name === '') ? (u ? String(u['名前'] || '') : '') : b.name;
  if (u) {
    sh.getRange(u._row, 1, 1, 5).setValues([[b.userId, name, active, u['登録日時'] || nowStr_(), admin]]);
  } else {
    if (!b.userId) return { ok: false, error: 'userId が必要です。' };
    sh.appendRow([b.userId, name, active, nowStr_(), admin]);
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
    var id = nextSlotId_();
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

// 予約キャンセル（本人）
function actionCancel_(body) {
  var userId = body.userId;
  if (!userId) return { ok: false, error: 'userId がありません。' };
  var sh = sheet_(SH.resv);
  var r = findResv_(userId);
  if (r) sh.deleteRow(r._row);
  return { ok: true };
}

// ステータス変更（管理）
function adminSetStatus_(b) {
  var sh = sheet_(SH.resv);
  var r = findResv_(b.userId);
  if (!r) return { ok: false, error: '予約が見つかりません。' };
  sh.getRange(r._row, 7, 1, 1).setValue(String(b.status || '受付')); // 7列目=ステータス
  return { ok: true };
}

// 期間で枠を一括追加（管理）
function adminAddSlots_(b) {
  if (!b.startDate || !b.endDate) return { ok: false, error: '開始日と終了日は必須です。' };
  var times = [].concat(b.times || []).map(function (t) { return String(t).trim(); }).filter(Boolean);
  if (!times.length) return { ok: false, error: '時間を1つ以上指定してください。' };
  var start = slotDateTime_(b.startDate, '00:00'), end = slotDateTime_(b.endDate, '00:00');
  if (!start || !end || start.getTime() > end.getTime()) return { ok: false, error: '日付の指定が不正です。' };
  var sh = sheet_(SH.slots);
  var existing = {};
  rows_(SH.slots).forEach(function (s) { existing[fmtDateCell_(s['日付']) + ' ' + fmtTimeCell_(s['時間'])] = true; });
  var id = nextSlotId_(), added = 0;
  var cur = new Date(start.getTime());
  while (cur.getTime() <= end.getTime()) {
    var dstr = Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy-MM-dd');
    for (var i = 0; i < times.length; i++) {
      var tt = /^(\d{1,2}):(\d{2})/.exec(times[i]);
      var tstr = tt ? (('0' + tt[1]).slice(-2) + ':' + tt[2]) : times[i];
      if (existing[dstr + ' ' + tstr]) continue;
      sh.appendRow([id, dstr, tstr, true, nowStr_()]);
      existing[dstr + ' ' + tstr] = true; id++; added++;
    }
    cur.setDate(cur.getDate() + 1);
  }
  return { ok: true, added: added };
}

/* ---------- ルーティング ---------- */

var ADMIN_ACTIONS = {
  admin_list: function (b) { return adminList_(); },
  admin_saveUser: adminSaveUser_,
  admin_deleteUser: adminDeleteUser_,
  admin_saveSlot: adminSaveSlot_,
  admin_deleteSlot: adminDeleteSlot_,
  admin_deleteReservation: adminDeleteReservation_,
  admin_saveRules: adminSaveRules_,
  admin_setStatus: adminSetStatus_,
  admin_addSlots: adminAddSlots_
};

function route_(action, body) {
  if (action === 'init') return actionInit_(body);
  if (action === 'book') return actionBook_(body);
  if (action === 'cancel') return actionCancel_(body);
  if (ADMIN_ACTIONS[action]) {
    var auth = requireAdmin_(body.accessToken);
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
