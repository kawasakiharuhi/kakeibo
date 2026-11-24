/**
 * 家計簿アプリ用のデータ保存API（Google Apps Script）
 *
 * いつもの「家計簿」スプレッドシートの月別シート（11月・10月…）を、そのままデータ置き場として使います。
 *  - 読み込み：各月シートの I〜L 列（日付・カテゴリー・内容・金額）を全部読みます
 *  - 追加：その月のシートの記録の一番下に、今までと同じ書き方で1行追加します
 *          （例：「支出・paypay」、振替は「支出・現金, 収入・paypay」）
 *  - 編集・削除：I〜L 列だけを書き換えます。左側の集計表（C〜G 列）には触りません
 *  - その月のシートがまだ無いときは「基盤」シートを複製して「2026年12月」のような名前で作ります
 *    （年が入るので、来年以降も古いシートを消さずに使い続けられます）
 *  - 今までの「6月」〜「11月」のようなシートは、そのまま 2026年の分として読み込みます
 *
 * 【大事】このファイルは「コード.gs」に上書きせず、新しいスクリプトファイルとして追加してください。
 *        「コード.gs」にある SHEETNAME() などの今の関数は、スプシの計算式で使っているので残します。
 *
 * 初回だけ：エディタ上部の関数選択で「setup」を選んで実行してください。
 *   アプリ接続用の「合言葉（キー）」を作って、実行ログに表示します。
 */

// ---- 設定 ----
const FIRST_YEAR = 2026;          // 一番古い月シート（一番右のシート）の年
const TEMPLATE_SHEET = '基盤';    // 新しい月を作るときに複製するシート
const COL = 9;                    // I 列（日付）から
const NCOL = 4;                   // I〜L の4列
const FIRST_ROW = 2;              // 1行目は見出し

// アプリの口座ID ⇔ スプシでの書き方
const ACCOUNTS = {
  cash: '現金', paypay: 'paypay', aupay: 'aupay', ufj: '三菱UFJ銀行', yucho: 'ゆうちょ銀行',
  chita: '知多信用金庫', starbucks: 'スターバックス', manaca: 'マナカ', credit: 'クレジットカード',
};
const NAME_TO_ID = Object.keys(ACCOUNTS).reduce((o, k) => { o[ACCOUNTS[k]] = k; return o; }, {});

/** 初回セットアップ：キー発行と、読み込みテスト */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('API_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('API_KEY', key);
  }
  const txs = readAll_();
  Logger.log('読み込めた記録：' + txs.length + '件（' + monthSheets_().map(m => m.sheet.getName()).join('・') + '）');
  Logger.log('アプリに入力するキー：' + key);
}

/** キーを作り直したいとき（前のキーは使えなくなります） */
function resetKey() {
  const key = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('API_KEY', key);
  Logger.log('新しいキー：' + key);
}

/** 読み込み：GET ?action=list&key=... */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!checkKey_(p.key)) return json_({ ok: false, error: 'unauthorized' });
  if (p.action === 'ping') return json_({ ok: true });
  if (p.action === 'list') return json_({ ok: true, txs: readAll_() });
  // 保存・削除も GET で受け付ける（ブラウザによっては POST の返事を受け取れず、
  // 保存できているのにエラーになることがあるため。読み込みと同じ通り道を使う）
  if (p.action === 'save' || p.action === 'delete') {
    let body;
    try { body = JSON.parse(p.payload || '{}'); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
    body.action = p.action;
    return json_(write_(body));
  }
  return json_({ ok: false, error: 'unknown_action' });
}

/**
 * 書き込み：POST
 *   追加 {key, action:'save', tx:{date,type,account,to?,amount,memo}}
 *   編集 {key, action:'save', tx:{id,...}, orig:{編集前の内容}}
 *   削除 {key, action:'delete', id, orig}
 * 編集・削除は「編集前の内容」と行の中身が一致するか確かめてから行います。
 * （スプシを直接いじって行がずれていたら conflict を返し、アプリは読み直します）
 */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  if (!checkKey_(body.key)) return json_({ ok: false, error: 'unauthorized' });
  return json_(write_(body));
}

/** 保存・削除の本体（GET・POST 共通） */
function write_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (body.action === 'save') {
      const tx = sanitize_(body.tx);
      if (!tx) return { ok: false, error: 'bad_tx' };
      if (!tx.id) return { ok: true, id: append_(tx) };                 // 追加
      const loc = locate_(tx.id, body.orig);                            // 編集
      if (!loc) return { ok: false, error: 'conflict' };
      const cur = monthSheets_().find(x => x.sheet.getName() === loc.sheet.getName());
      const d = ym_(tx.date);
      if (cur && cur.year === d.y && cur.month === d.m) {
        loc.sheet.getRange(loc.row, COL, 1, NCOL).setValues([toRow_(tx)]);
        return { ok: true, id: tx.id };
      }
      removeRow_(loc.sheet, loc.row);                                   // 月が変わったら移動
      return { ok: true, id: append_(tx) };
    }
    if (body.action === 'delete') {
      const loc = locate_(String(body.id || ''), body.orig);
      if (!loc) return { ok: false, error: 'conflict' };
      removeRow_(loc.sheet, loc.row);
      return { ok: true };
    }
    return { ok: false, error: 'unknown_action' };
  } finally {
    lock.releaseLock();
  }
}

// ---- 月シートの扱い ----

/**
 * 月シートの一覧を返す（古い順）。
 *  - 「2026年12月」のように年が入った名前 → その年
 *  - 「11月」のように年が無い名前（今までのシート）→ 一番右のシートを FIRST_YEAR として、
 *    右から順に「月が小さくなったら次の年」と数える
 */
function monthSheets_() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const out = [];
  let year = FIRST_YEAR, prev = null;
  for (let i = sheets.length - 1; i >= 0; i--) {
    const name = sheets[i].getName();
    let m = /^(\d{4})年(\d{1,2})月$/.exec(name);
    if (m) { out.push({ sheet: sheets[i], year: Number(m[1]), month: Number(m[2]) }); continue; }
    m = /^(\d{1,2})月$/.exec(name);
    if (!m) continue;
    const month = Number(m[1]);
    if (prev !== null && month < prev) year++;   // 12月→1月で年が変わる
    prev = month;
    out.push({ sheet: sheets[i], year: year, month: month });
  }
  out.sort((a, b) => (a.year - b.year) || (a.month - b.month));
  out.forEach((x, i) => x.order = i);
  return out;
}

function ym_(date) { return { y: Number(date.slice(0, 4)), m: Number(date.slice(5, 7)) }; }
function findMonthSheet_(y, m) { return monthSheets_().find(x => x.year === y && x.month === m) || null; }

/** 日付の月のシートを返す。無ければ「基盤」を複製して「YYYY年M月」を作る */
function sheetFor_(date) {
  const d = ym_(date);
  const found = findMonthSheet_(d.y, d.m);
  if (found) return found.sheet;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tpl = ss.getSheetByName(TEMPLATE_SHEET);
  if (!tpl) throw new Error('「' + TEMPLATE_SHEET + '」シートがありません');
  const name = d.y + '年' + d.m + '月';
  const sh = tpl.copyTo(ss).setName(name);

  // 置き場所：自分より新しい月シートの右、古い月シートの左（いつもの「新しい月が左」の並び）
  const all = ss.getSheets();
  const newer = monthSheets_().filter(x => x.sheet !== sh && (x.year * 12 + x.month) > (d.y * 12 + d.m));
  let pos = all.indexOf(ss.getSheetByName(TEMPLATE_SHEET)) + 2;
  newer.forEach(x => { pos = Math.max(pos, all.indexOf(x.sheet) + 2); });
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(Math.min(pos, all.length));

  fixFormulas_(sh, d.y, d.m);
  return sh;
}

/**
 * 複製した「基盤」の計算式を、年つきのシート名でも動くように直す。
 *  - A1（月の数字）は SHEETNAME() から取っているので、月の数字をそのまま書く
 *  - 残高の「先月のシート」を INDIRECT で月の数字から探している式は、
 *    先月のシート名（「11月」や「2026年12月」）を直接参照する式に書き換える
 */
function fixFormulas_(sh, y, m) {
  sh.getRange('A1').setValue(m);
  // 先月のシート（無ければ、それより前で一番新しい月のシート）から残高を引き継ぐ
  const target = y * 12 + m;
  const prev = monthSheets_().filter(x => x.sheet !== sh && (x.year * 12 + x.month) < target).pop();
  if (!prev) return;   // 前の月のシートが1つも無いときは元の式のまま
  const prevName = "'" + prev.sheet.getName().replace(/'/g, "''") + "'";
  const lastRow = Math.max(sh.getLastRow(), 1), lastCol = Math.max(sh.getLastColumn(), 1);
  const range = sh.getRange(1, 1, lastRow, lastCol);
  const formulas = range.getFormulas();
  const re = /INDIRECT\(\s*IF\(\s*A\$1\s*=\s*1\s*,\s*12\s*,\s*A\$1\s*-\s*1\s*\)\s*&\s*"月!([A-Za-z]+\d+)"\s*\)/gi;
  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      const f = formulas[r][c];
      if (f && re.test(f)) {
        re.lastIndex = 0;
        sh.getRange(r + 1, c + 1).setFormula(f.replace(re, (_, cell) => prevName + '!' + cell.toUpperCase()));
      }
      re.lastIndex = 0;
    }
  }
}

/** I〜L 列で、記録が入っている最後の行 */
function lastDataRow_(sh) {
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return FIRST_ROW - 1;
  const v = sh.getRange(FIRST_ROW, COL, last - FIRST_ROW + 1, NCOL).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i].some(x => x !== '' && x !== null)) return FIRST_ROW + i;
  return FIRST_ROW - 1;
}

function append_(tx) {
  const sh = sheetFor_(tx.date);
  const row = lastDataRow_(sh) + 1;
  const range = sh.getRange(row, COL, 1, NCOL);
  range.setValues([toRow_(tx)]);
  sh.getRange(row, COL).setNumberFormat('m"月"d"日"');
  sh.getRange(row, COL + 3).setNumberFormat('#,##0"円"');
  return sh.getName() + '!' + row;
}

/** 1行分（I〜L）を消して、下の記録を1行ずつ上に詰める（行そのものは削除しない） */
function removeRow_(sh, row) {
  const last = lastDataRow_(sh);
  if (row < last) {
    const below = sh.getRange(row + 1, COL, last - row, NCOL);
    below.copyTo(sh.getRange(row, COL, last - row, NCOL));
  }
  sh.getRange(last, COL, 1, NCOL).clearContent();
}

/** id（「11月!7」など）から行を探し、編集前の内容と一致するか確かめる */
function locate_(id, orig) {
  const m = /^(.+)!(\d+)(#[ab])?$/.exec(id || '');
  if (!m) return null;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(m[1]);
  if (!sh) return null;
  const row = Number(m[2]);
  if (orig) {
    const info = monthSheets_().find(x => x.sheet.getName() === m[1]);
    const cur = parseRow_(sh.getRange(row, COL, 1, NCOL).getValues()[0], info, row, false);
    const same = cur.some(t => t.date === orig.date && t.amount === Number(orig.amount) && (t.memo || '') === (orig.memo || '').replace(/\(戻り\)$/, ''));
    if (!same) return null;
  }
  return { sheet: sh, row: row };
}

// ---- 読み込み ----

function readAll_() {
  const months = monthSheets_();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const out = [];
  months.forEach((info, idx) => {
    const sh = info.sheet;
    const last = lastDataRow_(sh);
    if (last < FIRST_ROW) return;
    const values = sh.getRange(FIRST_ROW, COL, last - FIRST_ROW + 1, NCOL).getValues();
    values.forEach((r, i) => {
      parseRow_(r, info, FIRST_ROW + i, idx === 0, tz).forEach(t => out.push(t));
    });
  });
  return out;
}

/** 1行（日付・カテゴリー・内容・金額）をアプリの記録に変換 */
function parseRow_(r, info, row, isOldest, tz) {
  const id = info.sheet.getName() + '!' + row;
  let date = r[0];
  if (date instanceof Date) {
    date = Utilities.formatDate(date, tz || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  } else {
    const m = /(\d{1,2})月(\d{1,2})日/.exec(String(date));
    if (!m) return [];
    date = info.year + '-' + pad_(m[1]) + '-' + pad_(m[2]);
  }
  // 年は月シートの年に合わせる（スプシの日付の年がずれていても大丈夫なように）
  date = info.year + date.slice(4);
  const amount = Math.round(Number(String(r[3]).replace(/[^\d.-]/g, '')));
  if (!(amount > 0)) return [];
  const memo = String(r[2] || '');
  const parts = String(r[1] || '').split(/[,，、]/).map(s => s.trim()).filter(Boolean);
  const kinds = {};
  for (const p of parts) {
    const m = /^(収入|支出)・(.+)$/.exec(p);
    if (!m || !NAME_TO_ID[m[2]]) return [];
    (kinds[m[1]] = kinds[m[1]] || []).push(NAME_TO_ID[m[2]]);
  }
  const createdAt = info.order * 100000 + row;
  const base = { date: date, amount: amount, memo: memo, createdAt: createdAt };
  if (parts.length === 1) {
    const k = kinds['収入'] ? '収入' : '支出';
    const acc = kinds[k][0];
    if (isOldest && k === '収入' && memo === '') return [Object.assign({ id: id, type: 'opening', account: acc }, base, { memo: '初期残高' })];
    return [Object.assign({ id: id, type: k === '収入' ? 'income' : 'expense', account: acc }, base)];
  }
  if (kinds['支出'] && kinds['収入']) {
    const from = kinds['支出'][0], to = kinds['収入'][0];
    if (from === to) {   // 立て替えて戻ってきた等：支出と収入の両方として扱う
      return [
        Object.assign({ id: id + '#a', type: 'expense', account: from }, base),
        Object.assign({ id: id + '#b', type: 'income', account: to }, base, { memo: memo + '(戻り)', createdAt: createdAt + 0.5 }),
      ];
    }
    return [Object.assign({ id: id, type: 'transfer', account: from, to: to }, base)];
  }
  return [];
}

// ---- 書き込み用の変換 ----

function toRow_(t) {
  const y = Number(t.date.slice(0, 4)), m = Number(t.date.slice(5, 7)), d = Number(t.date.slice(8, 10));
  let cat;
  if (t.type === 'transfer') cat = '支出・' + ACCOUNTS[t.account] + ', 収入・' + ACCOUNTS[t.to];
  else cat = (t.type === 'expense' ? '支出・' : '収入・') + ACCOUNTS[t.account];
  return [new Date(y, m - 1, d), cat, t.memo || '', t.amount];
}

const TYPES = ['expense', 'income', 'transfer'];
function sanitize_(t) {
  if (!t || typeof t !== 'object') return null;
  const date = String(t.date || '');
  const amount = Math.round(Number(t.amount));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || TYPES.indexOf(t.type) < 0 || !(amount > 0)) return null;
  if (!ACCOUNTS[t.account]) return null;
  if (t.type === 'transfer' && (!ACCOUNTS[t.to] || t.to === t.account)) return null;
  const out = { date: date, type: t.type, account: t.account, amount: amount, memo: String(t.memo || '').slice(0, 200) };
  if (t.type === 'transfer') out.to = t.to;
  if (t.id) out.id = String(t.id).replace(/#[ab]$/, '');
  return out;
}

function checkKey_(key) {
  const real = PropertiesService.getScriptProperties().getProperty('API_KEY');
  return !!real && key === real;
}
function pad_(n) { return String(n).padStart(2, '0'); }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
