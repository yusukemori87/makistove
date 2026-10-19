/**
 * 薪ストーブ図鑑 メルマガ（Google Apps Script）
 *  登録（サイトのフォーム）→ 確認メール → 本登録 → 配信 → 配信停止
 *  エディタ（note風）：管理用デプロイのURLに ?a=admin を付けて開く（オーナーのみ）
 *  シート：「読者」「配信」「ログ」は setup() で自動作成
 */
const CFG = {
  siteName: '薪ストーブ図鑑',
  siteUrl: 'https://makistove.jp/',
  aboutUrl: 'https://makistove.jp/about/',
  privacyUrl: 'https://makistove.jp/privacy/',
  sender: '薪ストーブ図鑑（運営：合同会社NITOCRAFT）',
  address: '〒925-0054 石川県羽咋市千里浜町ト233-1',
  contact: 'ecobakacreation@gmail.com',
  owner: 'ecobakacreation@gmail.com',
  // 公開用（登録・確認・停止）のURL。停止リンクは必ずこのURLで作る
  publicUrl: 'https://script.google.com/macros/s/AKfycbzR5J-wGzeIXzv0Q02iJ4GaOC6-JshrEhK6TxyJTcZm1myuO3qZfjCP8hG7HC085Uc/exec',
  reserve: 5,          // 1日の送信枠のうち、確認メール用に残しておく通数
};
const SH = { users: '読者', casts: '配信', log: 'ログ' };
const U = { email: 1, pref: 2, src: 3, status: 4, token: 5, created: 6, confirmed: 7, stopped: 8, lastCast: 9 };
const C = { subject: 1, body: 2, state: 3, sent: 4, created: 5, started: 6, done: 7, editor: 8, pre: 9, sched: 10, updated: 11, mail: 12 };
const CAST_HEAD = ['件名', '本文（テキスト版）', '状態', '送信数', '作成日時', '配信開始', '配信完了', 'エディタHTML', 'プレヘッダー', '予約日時', '更新日時', 'メールHTML'];

/* ========== 初期設定（最初に1回だけ／更新時にもう一度実行） ========== */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const mk = (name, head) => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    s.setFrozenRows(1);
    return s;
  };
  mk(SH.users, ['メール', '都道府県', '登録元', '状態', 'トークン', '仮登録日時', '本登録日時', '停止日時', '最終配信No']);
  const casts = mk(SH.casts, CAST_HEAD);
  mk(SH.log, ['日時', '種別', '内容']);
  const def = ss.getSheetByName('シート1'); if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  casts.getRange('C2:C1000').clearDataValidations();
  casts.setColumnWidth(2, 360);
  // 予約配信と、枠切れの続きを1時間ごとに処理
  ScriptApp.getProjectTriggers().filter(t => ['processQueue', 'tick'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tick').timeBased().everyHours(1).create();
  imageFolder_(); // 画像フォルダを作る（Driveの権限確認も兼ねる）
  log_('setup', '初期設定完了');
}

/* ========== メニュー ========== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('メルマガ')
    .addItem('エディタを開く', 'openEditorFromSheet')
    .addItem('配信の続きを今すぐ送る', 'processQueue')
    .addItem('今日あと何通送れるか', 'showQuota')
    .addToUi();
}
function openEditorFromSheet() {
  const url = PropertiesService.getScriptProperties().getProperty('ADMIN_URL');
  const html = url
    ? '<p style="font-family:sans-serif">下のボタンでエディタが新しいタブで開きます。</p><p><a href="' + url + '?a=admin" target="_blank" style="display:inline-block;background:#1d3a5f;color:#fff;padding:12px 22px;border-radius:99px;text-decoration:none;font-family:sans-serif;font-weight:bold">エディタを開く</a></p>'
    : '<p style="font-family:sans-serif">エディタのURLが未設定です。</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(360).setHeight(150), 'メルマガエディタ');
}

/* ========== サイトからの登録（POST） ========== */
function doPost(e) {
  const p = (e && e.parameter) || {};
  const email = String(p.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 200) return json_({ ok: false, error: 'invalid' });
  if (p.website) return json_({ ok: true }); // ボット対策（隠し項目に入力があれば無視）
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_(SH.users);
    const row = findRow_(sh, email);
    const now = new Date();
    if (row) {
      const st = sh.getRange(row, U.status).getValue();
      if (st === '登録中') return json_({ ok: true, already: true });
      const token = Utilities.getUuid();
      sh.getRange(row, U.pref, 1, 5).setValues([[p.pref || '', p.src || '', '仮登録', token, now]]);
      sendConfirm_(email, token);
    } else {
      const token = Utilities.getUuid();
      sh.appendRow([email, p.pref || '', p.src || '', '仮登録', token, now, '', '', '']);
      sendConfirm_(email, token);
    }
    log_('登録申請', email);
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}

/* ========== 本登録・配信停止（リンクのクリック）／エディタ ========== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const a = p.a, t = String(p.t || '');
  if (a === 'admin') {
    if (!isOwner_()) return page_('このページは開けません', '管理者だけが使えるページです。', true);
    return HtmlService.createHtmlOutputFromFile('Editor').setTitle('メルマガエディタ｜薪ストーブ図鑑')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1');
  }
  if (!t) return page_('薪ストーブ図鑑', 'このページは メルマガの登録・停止用です。', true);
  const sh = sheet_(SH.users);
  const row = findRow_(sh, t, U.token);
  if (!row) return page_('リンクが無効です', 'お手数ですが、サイトからもう一度ご登録ください。', true);
  const email = sh.getRange(row, U.email).getValue();
  const now = new Date();
  if (a === 'confirm') {
    if (sh.getRange(row, U.status).getValue() !== '登録中') {
      sh.getRange(row, U.status).setValue('登録中');
      sh.getRange(row, U.confirmed).setValue(now);
      sh.getRange(row, U.stopped).setValue('');
      sendWelcome_(email, t);
      log_('本登録', email);
    }
    return page_('登録が完了しました', '薪ストーブ図鑑のメルマガにご登録いただき、ありがとうございます。<br>補助金一覧の更新や新しい機種の情報を、ときどきお届けします。', true);
  }
  if (a === 'unsub') {
    sh.getRange(row, U.status).setValue('停止');
    sh.getRange(row, U.stopped).setValue(now);
    log_('配信停止', email);
    return page_('配信を停止しました', email + ' へのメルマガ配信を停止しました。<br>これまでお読みいただき、ありがとうございました。', true);
  }
  return page_('薪ストーブ図鑑', '不明な操作です。', true);
}

/* ========== エディタ用API（オーナーのみ） ========== */
function apiInit() {
  owner_();
  const cs = sheet_(SH.casts), last = cs.getLastRow();
  const issues = [];
  if (last >= 2) {
    const v = cs.getRange(2, 1, last - 1, CAST_HEAD.length).getValues();
    v.forEach((r, i) => {
      if (!r[C.subject - 1] && !r[C.editor - 1] && !r[C.body - 1]) return;
      if (r[C.state - 1] === '削除') return;
      issues.push({
        no: i + 1, subject: r[C.subject - 1], state: r[C.state - 1] || '下書き', sent: r[C.sent - 1] || 0,
        created: iso_(r[C.created - 1]), done: iso_(r[C.done - 1]), sched: iso_(r[C.sched - 1]), updated: iso_(r[C.updated - 1]),
        excerpt: String(r[C.body - 1] || '').replace(/\s+/g, ' ').slice(0, 90),
      });
    });
  }
  issues.reverse();
  return { issues, stats: readerStats_(), quota: MailApp.getRemainingDailyQuota(), owner: CFG.owner };
}

function apiGet(no) {
  owner_();
  const r = sheet_(SH.casts).getRange(no + 1, 1, 1, CAST_HEAD.length).getValues()[0];
  let html = r[C.editor - 1];
  if (!html && r[C.body - 1]) html = String(r[C.body - 1]).split(/\n{2,}/).map(s => '<p>' + esc_(s).replace(/\n/g, '<br>') + '</p>').join('');
  return { no, subject: r[C.subject - 1], pre: r[C.pre - 1], html: html || '', state: r[C.state - 1] || '下書き', sent: r[C.sent - 1] || 0, sched: iso_(r[C.sched - 1]) };
}

function apiSave(d) {
  owner_();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const cs = sheet_(SH.casts); const now = new Date();
    let no = d.no;
    if (!no) {
      cs.appendRow(['', '', '下書き', 0, now]);
      no = cs.getLastRow() - 1;
    } else {
      const st = cs.getRange(no + 1, C.state).getValue();
      if (st === '配信中' || st === '配信済み') throw new Error('配信済みの号は編集できません（複製して書いてください）');
    }
    const row = no + 1;
    cs.getRange(row, C.subject).setValue(d.subject || '');
    cs.getRange(row, C.body).setValue(d.text || '');
    cs.getRange(row, C.editor).setValue(d.html || '');
    cs.getRange(row, C.pre).setValue(d.pre || '');
    cs.getRange(row, C.updated).setValue(now);
    cs.getRange(row, C.mail).setValue(d.mail || '');
    return { no, updated: now.toISOString() };
  } finally { lock.releaseLock(); }
}

function apiDuplicate(no) {
  owner_();
  const d = apiGet(no);
  return apiSave({ subject: d.subject + '（コピー）', pre: d.pre, html: d.html, text: sheet_(SH.casts).getRange(no + 1, C.body).getValue(), mail: sheet_(SH.casts).getRange(no + 1, C.mail).getValue() });
}

function apiDelete(no) {
  owner_();
  const cs = sheet_(SH.casts); const st = cs.getRange(no + 1, C.state).getValue();
  if (st === '配信中' || st === '配信済み') throw new Error('配信済みの号は削除できません');
  cs.getRange(no + 1, C.state).setValue('削除');
  return true;
}

function apiUploadImage(dataUrl, name) {
  owner_();
  const m = String(dataUrl).match(/^data:(image\/[a-z+]+);base64,(.+)$/);
  if (!m) throw new Error('画像の形式が不正です');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name || ('img_' + Date.now()));
  const f = imageFolder_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + f.getId();
}

function apiPreview(d) {
  owner_();
  return buildMail_(d.subject || '（件名未入力）', d.pre || '', d.mail || '', 'preview');
}

function apiTest(d) {
  owner_();
  const saved = apiSave(d);
  MailApp.sendEmail({ to: CFG.owner, subject: '【テスト】' + (d.subject || '（件名未入力）'), body: (d.text || '') + footer_('test'),
    htmlBody: buildMail_(d.subject, d.pre, d.mail, 'test'), name: CFG.siteName, replyTo: CFG.contact });
  return { no: saved.no, to: CFG.owner, quota: MailApp.getRemainingDailyQuota() };
}

function apiSend(d) {
  owner_();
  if (!d.subject) throw new Error('件名を入れてください');
  const saved = apiSave(d);
  const cs = sheet_(SH.casts);
  cs.getRange(saved.no + 1, C.state).setValue('配信中');
  cs.getRange(saved.no + 1, C.started).setValue(new Date());
  cs.getRange(saved.no + 1, C.sched).setValue('');
  log_('配信開始', 'No.' + saved.no + ' ' + d.subject);
  processQueue();
  return apiGet(saved.no);
}

function apiSchedule(d, isoTime) {
  owner_();
  if (!d.subject) throw new Error('件名を入れてください');
  const when = new Date(isoTime);
  if (isNaN(when) || when < new Date()) throw new Error('未来の日時を選んでください');
  const saved = apiSave(d);
  const cs = sheet_(SH.casts);
  cs.getRange(saved.no + 1, C.state).setValue('予約');
  cs.getRange(saved.no + 1, C.sched).setValue(when);
  log_('予約', 'No.' + saved.no + ' ' + when);
  return apiGet(saved.no);
}

function apiUnschedule(no) {
  owner_();
  const cs = sheet_(SH.casts);
  if (cs.getRange(no + 1, C.state).getValue() !== '予約') return apiGet(no);
  cs.getRange(no + 1, C.state).setValue('下書き');
  cs.getRange(no + 1, C.sched).setValue('');
  return apiGet(no);
}

function apiReaders() {
  owner_();
  const sh = sheet_(SH.users), last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, U.lastCast).getValues().map(r => ({
    email: r[0], src: r[U.src - 1], status: r[U.status - 1], created: iso_(r[U.created - 1]), confirmed: iso_(r[U.confirmed - 1]),
  })).reverse();
}

/* ========== 配信処理 ========== */
/** 1時間ごと：予約時刻を過ぎた号を配信開始し、配信中の続きを送る */
function tick() {
  const cs = sheet_(SH.casts), last = cs.getLastRow();
  if (last >= 2) {
    const v = cs.getRange(2, 1, last - 1, CAST_HEAD.length).getValues(); const now = new Date();
    v.forEach((r, i) => {
      if (r[C.state - 1] === '予約' && r[C.sched - 1] instanceof Date && r[C.sched - 1] <= now) {
        cs.getRange(i + 2, C.state).setValue('配信中');
        cs.getRange(i + 2, C.started).setValue(now);
        log_('予約配信開始', 'No.' + (i + 1));
      }
    });
  }
  processQueue();
}

/** 「配信中」の号を、今日の送信枠いっぱいまで送る（残りは次の tick で再開） */
function processQueue() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) return;
  try {
    const cs = sheet_(SH.casts), us = sheet_(SH.users);
    const cv = cs.getDataRange().getValues();
    for (let r = 2; r <= cv.length; r++) {
      if (cv[r - 1][C.state - 1] !== '配信中') continue;
      const no = r - 1; // 配信No＝行番号-1
      const row = cv[r - 1];
      const uv = us.getDataRange().getValues();
      let sent = Number(row[C.sent - 1]) || 0, left = false;
      for (let i = 1; i < uv.length; i++) {
        if (uv[i][U.status - 1] !== '登録中' || Number(uv[i][U.lastCast - 1]) >= no) continue;
        if (MailApp.getRemainingDailyQuota() <= CFG.reserve) { left = true; break; }
        try {
          sendCast_(uv[i][U.email - 1], row, uv[i][U.token - 1]);
          us.getRange(i + 1, U.lastCast).setValue(no);
          sent++;
        } catch (err) { log_('送信エラー', uv[i][U.email - 1] + ' ' + err); }
      }
      cs.getRange(r, C.sent).setValue(sent);
      if (left) { log_('配信', 'No.' + no + ' 今日の枠を使い切り。残りは自動で再開'); return; }
      cs.getRange(r, C.state).setValue('配信済み');
      cs.getRange(r, C.done).setValue(new Date());
      log_('配信', 'No.' + no + ' 完了 ' + sent + '通');
    }
  } finally { lock.releaseLock(); }
}

function showQuota() {
  SpreadsheetApp.getUi().alert('今日あと ' + MailApp.getRemainingDailyQuota() + ' 通送れます（Gmailの上限：1日100通）。\n登録中の読者：' + activeEmails_().length + ' 人');
}

/* ========== メール本文 ========== */
function unsubUrl_(token) { return token && token !== 'test' && token !== 'preview' ? CFG.publicUrl + '?a=unsub&t=' + token : ''; }

function footer_(token) {
  const u = unsubUrl_(token) || '（テスト送信のため停止リンクはありません）';
  return '\n\n――――――――――――――――\n' +
    '発行：' + CFG.sender + '\n' +
    '住所：' + CFG.address + '\n' +
    'サイト：' + CFG.siteUrl + '\n' +
    '運営者情報：' + CFG.aboutUrl + '\n' +
    'お問い合わせ：' + CFG.contact + '\n' +
    '配信停止はこちら（1クリックで停止できます）：\n' + u + '\n' +
    '※このメールは、薪ストーブ図鑑でメルマガ登録をされた方にお送りしています。';
}

/** HTMLメールの外枠（ヘッダー・タイトル・本文・フッター） */
function buildMail_(subject, pre, content, token) {
  const u = unsubUrl_(token);
  const f = 'font-family:-apple-system,BlinkMacSystemFont,\'Hiragino Sans\',\'Hiragino Kaku Gothic ProN\',Meiryo,sans-serif;';
  return '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + esc_(subject) + '</title></head>' +
    '<body style="margin:0;padding:0;background:#f4f1ea">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0">' + esc_(pre || '') + '&nbsp;</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f1ea"><tr><td align="center" style="padding:24px 12px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#ffffff;border-radius:10px;overflow:hidden">' +
    '<tr><td style="height:6px;background:#E5B42C;font-size:0;line-height:0">&nbsp;</td></tr>' +
    '<tr><td style="padding:22px 28px 0;' + f + '"><a href="' + CFG.siteUrl + '" style="text-decoration:none;color:#111"><span style="font-size:15px;font-weight:bold;letter-spacing:.08em">薪ストーブ図鑑</span>' +
    '<span style="font-size:10px;color:#999;letter-spacing:.2em;margin-left:8px">MAKI STOVE ZUKAN</span></a></td></tr>' +
    '<tr><td style="padding:18px 28px 6px;' + f + '"><h1 style="margin:0;font-size:24px;line-height:1.5;color:#111;font-weight:bold">' + esc_(subject) + '</h1></td></tr>' +
    '<tr><td style="padding:12px 28px 28px;' + f + 'font-size:16px;line-height:1.9;color:#222">' + (content || '') + '</td></tr>' +
    '<tr><td style="padding:22px 28px;background:#222;color:#bbb;font-size:12px;line-height:1.8;' + f + '">' +
    '発行：' + esc_(CFG.sender) + '<br>住所：' + esc_(CFG.address) + '<br>' +
    '<a href="' + CFG.siteUrl + '" style="color:#E5B42C">makistove.jp</a>　<a href="' + CFG.aboutUrl + '" style="color:#bbb">運営者情報</a>　お問い合わせ：<a href="mailto:' + CFG.contact + '" style="color:#bbb">' + CFG.contact + '</a><br>' +
    (u ? '<a href="' + u + '" style="color:#fff">配信停止はこちら</a>（1クリックで停止できます）' : '配信停止リンク（テストのため表示のみ）') +
    '<br><span style="color:#888">このメールは、薪ストーブ図鑑でメルマガ登録をされた方にお送りしています。</span></td></tr>' +
    '</table></td></tr></table></body></html>';
}

function sendConfirm_(email, token) {
  const url = CFG.publicUrl + '?a=confirm&t=' + token;
  const body = '薪ストーブ図鑑のメルマガにお申し込みいただき、ありがとうございます。\n\n' +
    '下のリンクを押すと登録が完了します。\n' + url + '\n\n' +
    '※お心当たりがない場合は、このメールを破棄してください。リンクを押さない限り登録されません。' +
    '\n\n――――――――――――――――\n発行：' + CFG.sender + '\n' + CFG.address + '\n' + CFG.siteUrl + '\nお問い合わせ：' + CFG.contact;
  MailApp.sendEmail({ to: email, subject: '【薪ストーブ図鑑】メルマガ登録の確認', body, name: CFG.siteName, replyTo: CFG.contact });
}

function sendWelcome_(email, token) {
  const body = '薪ストーブ図鑑のメルマガへの登録が完了しました。\n\n' +
    'これから、こんな情報をときどきお届けします。\n' +
    '・薪ストーブ補助金の一覧の更新（年2回・4月と10月ごろ）\n' +
    '・新しい機種や値上げの情報\n' +
    '・薪ストーブの最新情報やコラム\n\n' +
    'まずは補助金の一覧から、お住まいの自治体をチェックしてみてください。\n' + CFG.siteUrl + '#subsidy' +
    footer_(token);
  MailApp.sendEmail({ to: email, subject: '【薪ストーブ図鑑】ご登録ありがとうございます', body, name: CFG.siteName, replyTo: CFG.contact });
}

/** 1通送る（HTML版があればHTML＋テキスト、なければテキストのみ） */
function sendCast_(email, row, token) {
  const subject = row[C.subject - 1], text = String(row[C.body - 1] || ''), mail = row[C.mail - 1];
  const opt = { to: email, subject: subject, body: text + footer_(token), name: CFG.siteName, replyTo: CFG.contact };
  if (mail) opt.htmlBody = buildMail_(subject, row[C.pre - 1], mail, token);
  MailApp.sendEmail(opt);
}

/* ========== 共通 ========== */
function isOwner_() { try { return String(Session.getActiveUser().getEmail()).toLowerCase() === CFG.owner; } catch (e) { return false; } }
function owner_() { if (!isOwner_()) throw new Error('権限がありません'); }
function imageFolder_() {
  const p = PropertiesService.getScriptProperties(); const id = p.getProperty('IMG_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('薪ストーブ図鑑_メルマガ画像'); p.setProperty('IMG_FOLDER', f.getId()); return f;
}
function readerStats_() {
  const sh = sheet_(SH.users), last = sh.getLastRow(); const s = { active: 0, pending: 0, stopped: 0 };
  if (last < 2) return s;
  sh.getRange(2, U.status, last - 1, 1).getValues().forEach(r => { if (r[0] === '登録中') s.active++; else if (r[0] === '仮登録') s.pending++; else if (r[0] === '停止') s.stopped++; });
  return s;
}
function iso_(v) { return v instanceof Date ? v.toISOString() : (v ? String(v) : ''); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function findRow_(sh, value, col) {
  col = col || U.email;
  const last = sh.getLastRow(); if (last < 2) return 0;
  const v = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).toLowerCase() === String(value).toLowerCase()) return i + 2;
  return 0;
}
function activeEmails_() {
  const sh = sheet_(SH.users); const last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, U.status).getValues().filter(r => r[U.status - 1] === '登録中').map(r => r[0]);
}
function log_(kind, text) { try { sheet_(SH.log).appendRow([new Date(), kind, text]); } catch (e) {} }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function page_(title, msg, back) {
  const html = '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:sans-serif;max-width:560px;margin:60px auto;padding:0 20px;line-height:1.8;color:#222">' +
    '<div style="height:6px;background:#E5B42C;margin-bottom:28px"></div>' +
    '<h1 style="font-size:22px">' + title + '</h1><p>' + msg + '</p>' +
    (back ? '<p style="margin-top:32px"><a href="' + CFG.siteUrl + '" target="_top" style="color:#1d3a5f">薪ストーブ図鑑に戻る →</a></p>' : '') + '</div>';
  return HtmlService.createHtmlOutput(html).setTitle(title + '｜薪ストーブ図鑑').addMetaTag('viewport', 'width=device-width,initial-scale=1');
}
