/**
 * スチール予約システム - Google Apps Script Web API
 *
 * 役割:
 *   doPost  … LIFFから届いた予約をスプレッドシートに1行追記
 *   doGet   … 管理画面向けに予約一覧をJSONで返す
 *
 * 事前準備（スクリプトプロパティ）:
 *   SHEET_ID   … 保存先スプレッドシートのID（必須）
 *   SHEET_NAME … シート名（任意、既定 "予約"）
 *   ADMIN_KEY  … 管理画面の閲覧パスコード（任意、設定すると一覧取得に必須）
 *
 * デプロイ: 「デプロイ > 新しいデプロイ > ウェブアプリ」
 *   実行するユーザー = 自分 / アクセスできるユーザー = 全員
 *   発行される /exec URL を liff・admin の GAS_ENDPOINT に設定する。
 */

// SHEET_ID はスクリプトプロパティが優先。未設定ならこの既定値を使う。
var DEFAULT_SHEET_ID = '1neHxod-oOulSHjkbd41hoZ9emkkJPtHvvaQgO6wNvoE';

var HEADERS = ['受付番号', '受付日時', 'userId', '表示名', 'ライン名', '希望日', '希望時間', '備考', 'ステータス'];

function prop_(key, fallback) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return (v === null || v === '') ? fallback : v;
}

function getSheet_() {
  var id = prop_('SHEET_ID', DEFAULT_SHEET_ID);
  if (!id) throw new Error('スクリプトプロパティ SHEET_ID が未設定です。');
  var ss = SpreadsheetApp.openById(id);
  var name = prop_('SHEET_NAME', '予約');
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 予約の作成（LIFFから） */
function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action && body.action !== 'create') {
      return json_({ ok: false, error: '未対応のaction: ' + body.action });
    }
    if (!body.lineName || !body.date || !body.time) {
      return json_({ ok: false, error: '必須項目（ライン名・希望日・希望時間）が不足しています。' });
    }
    var sh = getSheet_();
    var id = 'R' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    sh.appendRow([
      id, now,
      body.userId || '', body.displayName || '',
      body.lineName, body.date, body.time,
      body.remarks || '', '受付'
    ]);
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 予約一覧の取得（管理画面から） / ?action=list&key=XXXX */
function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    if (params.action !== 'list') {
      return json_({ ok: true, message: 'Steel-Booking GAS API is running.' });
    }
    var adminKey = prop_('ADMIN_KEY', '');
    if (adminKey && params.key !== adminKey) {
      return json_({ ok: false, error: 'パスコードが違います。' });
    }
    var sh = getSheet_();
    var last = sh.getLastRow();
    var rows = [];
    if (last > 1) {
      var values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
      rows = values.map(function (r) {
        return {
          id: r[0], receivedAt: r[1], userId: r[2], displayName: r[3],
          lineName: r[4], date: String(r[5]), time: String(r[6]),
          remarks: r[7], status: r[8]
        };
      }).reverse(); // 新しい順
    }
    return json_({ ok: true, count: rows.length, rows: rows });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
