/**
 * 「ひなたの予定」同期用スクリプト（Google スプレッドシート ＋ Apps Script）
 *
 * ・アプリの記録（予定の達成・ポイント・ごほうび）を、スプレッドシートの「履歴」シートに 1 行ずつ追加していきます。
 * ・iPad の Safari でもホーム画面アイコンでも、このシートを通して同じ状態になります。
 * ・このファイルは公開しないでください（下の TOKEN は、あなただけの合言葉です）。
 *
 * 使い方は SYNC_SETUP.md を見てください。
 */

// ★ここを、自分で決めた長い合言葉（英数字 20 文字以上を推奨）に書き換えてください。
//   書き換えないと、すべてのアクセスを拒否します。
const TOKEN = 'ここを自分で決めた長い合言葉に書き換える';

const SHEET_NAME = '履歴';
const SUMMARY_NAME = 'サマリー';
const HEADERS = ['eid', 't', '日時', '端末', '種類', '日付', 'キー', '名前', '値'];
const KINDS = { done: 1, award: 1, adj: 1, claim: 1, bonus: 1, badj: 1, bclaim: 1 };
const MAX_PUSH = 200;
const MAX_PULL = 2000;
const TZ = 'Asia/Tokyo';

/** 最初に 1 回だけ、手動で実行します（シートを準備します）。 */
function setup() {
  getSheet_();
  refreshSummary_();
  Logger.log(TOKEN.length < 12 || TOKEN.indexOf('ここを') === 0
    ? '⚠ TOKEN（合言葉）がまだ書き換えられていません。Code.gs の先頭を直して保存してください。'
    : 'OK：シートを準備しました。次は「デプロイ」→「新しいデプロイ」→「ウェブアプリ」です。');
}

/** 「履歴」を最初から読み直して、「サマリー」シートに合計を書き出します（記録が追加されるたびに自動で実行）。 */
function refreshSummary_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sm = ss.getSheetByName(SUMMARY_NAME);
    if (!sm) sm = ss.insertSheet(SUMMARY_NAME);
    const sh = getSheet_();
    const n = Math.max(0, sh.getLastRow() - 1);
    const latest = {};
    let adj = 0, badj = 0, lastT = 0;
    if (n > 0) {
      const rows = sh.getRange(2, 1, n, 9).getValues();
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const t = Number(r[1]), k = str_(r[4]), v = Number(r[8]);
        if (t > lastT) lastT = t;
        if (k === 'adj') { adj += v; continue; }
        if (k === 'badj') { badj += v; continue; }
        const id = k + '|' + str_(r[5]) + '|' + str_(r[6]);
        if (!latest[id] || t >= latest[id].t) latest[id] = { t: t, v: v, k: k };
      }
    }
    let awards = 0, bonus = 0, bonusCount = 0;
    Object.keys(latest).forEach(function (id) {
      const o = latest[id];
      if (o.k === 'award' && o.v > 0) awards++;
      if (o.k === 'bonus' && o.v > 0) { bonus += o.v; bonusCount++; }
    });
    const upd = lastT ? Utilities.formatDate(new Date(lastT), TZ, 'yyyy-MM-dd HH:mm') : '';
    sm.getRange(1, 1, 9, 2).setValues([
      ['項目', '値'],
      ['ポイント合計', awards + adj],
      ['　全部できた日数', awards],
      ['　手動補正の合計', adj],
      ['ボーナスポイント合計', bonus + badj],
      ['　ボーナスで獲得した合計', bonus],
      ['　ボーナスの完了件数', bonusCount],
      ['　手動補正の合計（ボーナス）', badj],
      ['最終更新', upd]
    ]);
  } catch (err) {
    // サマリーの更新に失敗しても、記録の保存には影響しません。
  }
}

function doGet(e) {
  return handle_(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'bad json' });
  }
  return handle_(body);
}

function handle_(p) {
  if (!authOk_(p.token)) return out_({ ok: false, error: 'auth' });
  const action = p.action;
  if (action === 'ping') return out_({ ok: true, seq: count_() });
  if (action === 'pull') return pull_(parseInt(p.since || '0', 10));
  if (action === 'push') {
    let events = p.events;
    if (typeof p.data === 'string') {
      try { events = JSON.parse(p.data); } catch (err) { return out_({ ok: false, error: 'bad data' }); }
    }
    return push_(events);
  }
  return out_({ ok: false, error: 'unknown action' });
}

function authOk_(t) {
  if (TOKEN.length < 12 || TOKEN.indexOf('ここを') === 0) return false;
  return typeof t === 'string' && t === TOKEN;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    // 日付などが勝手に日付型へ変換されないよう、文字列の列を「書式なしテキスト」にします。
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('F:H').setNumberFormat('@');
  }
  return sh;
}

function count_() {
  return Math.max(0, getSheet_().getLastRow() - 1);
}

function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}

function safe_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max);
  if (/^[=+\-@]/.test(s)) s = ' ' + s; // 数式として解釈されないようにする
  return s;
}

function pull_(since) {
  const sh = getSheet_();
  const total = Math.max(0, sh.getLastRow() - 1);
  let reset = false;
  if (isNaN(since) || since < 0 || since > total) { since = 0; reset = true; } // 行が削除された場合は最初から読み直す
  const n = Math.min(MAX_PULL, total - since);
  const events = [];
  if (n > 0) {
    const rows = sh.getRange(since + 2, 1, n, 9).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      events.push({ id: str_(r[0]), t: Number(r[1]), dev: str_(r[3]), k: str_(r[4]), d: str_(r[5]), key: str_(r[6]), n: str_(r[7]), v: Number(r[8]) });
    }
  }
  return out_({ ok: true, events: events, seq: since + n, more: since + n < total, reset: reset });
}

function push_(events) {
  if (events && !Array.isArray(events) && typeof events === 'object') events = [events];
  if (!Array.isArray(events)) return out_({ ok: false, error: 'no events' });
  if (events.length > MAX_PUSH) return out_({ ok: false, error: 'too many events' });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet_();
    const last = sh.getLastRow();
    const seen = {};
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) seen[str_(ids[i][0])] = 1;
    }
    const rows = [];
    for (let i = 0; i < events.length; i++) {
      const e = events[i] || {};
      const id = typeof e.id === 'string' ? e.id : '';
      const t = Number(e.t);
      const v = Number(e.v);
      const d = typeof e.d === 'string' ? e.d : '';
      if (!id || id.length > 80 || !KINDS[e.k] || !isFinite(t) || !isFinite(v)) continue;
      if (d && !/^\d{4}-\d{2}-\d{2}$/.test(d)) continue;
      if (seen[id]) continue; // 同じ記録を 2 回送っても、1 回だけ保存します
      seen[id] = 1;
      rows.push([id, t, Utilities.formatDate(new Date(t), TZ, 'yyyy-MM-dd HH:mm:ss'), safe_(e.dev, 30), e.k, d, safe_(e.key, 60), safe_(e.n, 60), v]);
    }
    if (rows.length) {
      sh.getRange(last + 1, 1, rows.length, 9).setValues(rows);
      refreshSummary_();
    }
    return out_({ ok: true, added: rows.length, seq: Math.max(0, last - 1 + rows.length) });
  } finally {
    lock.releaseLock();
  }
}
